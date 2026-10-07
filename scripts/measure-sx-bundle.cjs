// Locked local toolchain only. No install, dependency mutation, or baseline rebuild.
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const zlib = require('node:zlib');
const { spawnSync } = require('node:child_process');
const webpack = require('webpack');
const ts = require('typescript');
const root = path.resolve(__dirname, '..');
const label = process.env.SX_MEASURE_LABEL || 'task-6-final';
if (!/^[a-zA-Z0-9_-]+$/.test(label)) throw new Error('Invalid measurement label');
const output = path.resolve(process.env.SX_MEASURE_OUTPUT || path.join(root, '.docs/maintenance/evidence/use-sx'), label);
// Stable module resource paths keep Webpack deterministic IDs independent of the
// evidence label. Run this script sequentially; separate logs remain archived.
const work = path.join(root, 'temp/sx-bundle/current');
fs.mkdirSync(output, { recursive: true });
fs.mkdirSync(work, { recursive: true });
const report = { label, node: process.version, versions: {}, commands: [], bundles: [], caveats: [] };
report.runtimeStderr = [];
const writeStderr = process.stderr.write.bind(process.stderr);
process.stderr.write = function captureStderr(chunk, ...args) {
  report.runtimeStderr.push(String(chunk));
  return writeStderr(chunk, ...args);
};
const hash = data => crypto.createHash('sha256').update(data).digest('hex');
const save = () => fs.writeFileSync(path.join(output, 'measurement.json'), JSON.stringify(report, null, 2) + '\n');
function command(name, executable, args, options = {}, captureFailure = false) {
  const result = spawnSync(executable, args, { cwd: root, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024, ...options });
  const log = `${result.stdout || ''}${result.stderr || ''}`;
  fs.writeFileSync(path.join(output, `${name}.log`), log);
  const failures = log.split('\n').filter(line => /build failed|Exiting with exit code: [1-9]|Error -|error TS\d+|npm error/i.test(line));
  const warnings = log.split('\n').filter(line => /warn|outdated|over two months old|caniuse-lite/i.test(line));
  const item = { name, executable, args, status: result.status, signal: result.signal, error: result.error?.message, failures, warnings };
  report.commands.push(item); save();
  console.log(`${name}: exit ${result.status}; ${failures.length} failure markers; ${warnings.length} warning lines`);
  if ((result.status !== 0 || result.error || failures.length) && !captureFailure) throw new Error(`${name} failed; inspect ${output}/${name}.log`);
  return result;
}
function sizes(data) {
  return { raw: data.length, gzip: zlib.gzipSync(data, { level: 9 }).length,
    brotli: zlib.brotliCompressSync(data, { params: { [zlib.constants.BROTLI_PARAM_QUALITY]: 11 } }).length, sha256: hash(data) };
}
function filesSize(directory, names) {
  const files = [...new Set(names)].sort().map(name => ({ name, ...sizes(fs.readFileSync(path.join(directory, name))) }));
  return { files, totals: files.reduce((sum, f) => ({ raw: sum.raw + f.raw, gzip: sum.gzip + f.gzip, brotli: sum.brotli + f.brotli }), { raw: 0, gzip: 0, brotli: 0 }) };
}
function shipCapture(directory, allowedNames) {
  const manifests = fs.readdirSync(directory).filter(n => n.endsWith('.manifest.json'));
  const references = new Set();
  const host = manifests.map(name => {
    const manifest = JSON.parse(fs.readFileSync(path.join(directory, name)));
    for (const resource of Object.values(manifest.loaderConfig.scriptResources)) {
      if (resource.type === 'path') references.add(resource.path);
    }
    return { name, loaderConfig: manifest.loaderConfig };
  });
  // Webpack can leave identical lazy assets untouched, preserving old mtimes.
  // Follow the current entries' actual .u chunk filename tables as well as the
  // manifests; never infer reachability from a directory-wide filename pattern.
  const runtimeReferences = new Set();
  for (const name of references) {
    const ast = ts.createSourceFile(name, fs.readFileSync(path.join(directory, name), 'utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
    const initialIds = new Set(); const tables = [];
    function visit(node) {
      if (ts.isObjectLiteralExpression(node)) for (const p of node.properties) {
        if (ts.isPropertyAssignment(p) && ts.isNumericLiteral(p.name) && ts.isNumericLiteral(p.initializer) && p.initializer.text === '0') initialIds.add(p.name.text);
      }
      if (ts.isBinaryExpression(node) && ts.isPropertyAccessExpression(node.left) && node.left.name.text === 'u' && ts.isArrowFunction(node.right) && node.right.body.getText(ast).includes('"chunk."')) {
        const objects = [];
        function objectVisit(child) { if (ts.isObjectLiteralExpression(child)) objects.push(child); ts.forEachChild(child, objectVisit); }
        objectVisit(node.right.body); tables.push(objects);
      }
      ts.forEachChild(node, visit);
    }
    visit(ast);
    for (const objects of tables) {
      const chunkNames = {}; const hashes = {};
      for (const object of objects) for (const p of object.properties) {
        if (!ts.isPropertyAssignment(p) || !ts.isNumericLiteral(p.name) || !ts.isStringLiteral(p.initializer)) continue;
        (/^[a-f0-9]{20}$/.test(p.initializer.text) ? hashes : chunkNames)[p.name.text] = p.initializer.text;
      }
      for (const [id, chunkHash] of Object.entries(hashes)) {
        const file = `chunk.${chunkNames[id] || id}_${chunkHash}.js`;
        if (fs.existsSync(path.join(directory, file))) runtimeReferences.add(file);
        else if (!initialIds.has(id)) throw new Error(`Current entry ${name} references missing chunk ${file}`);
      }
    }
  }
  const names = new Set([...allowedNames.filter(n => n.endsWith('.js')), ...references]);
  for (const name of runtimeReferences) names.add(name);
  for (const name of names) {
    if (!name.endsWith('.js') || !fs.existsSync(path.join(directory, name))) throw new Error(`Missing reachable ship asset ${name}`);
  }
  return { ...filesSize(directory, [...names]), host, manifestReferences: [...references].sort(), runtimeReferences: [...runtimeReferences].sort(), selection: 'current-run JS union manifest path resources union exact current-entry Webpack runtime chunk tables; unique names; no maps or old unreferenced assets' };
}
const { dataCalls } = require('./bundle-contract.internal.cjs');

async function bundle(name, fixture, library) {
  const dir = path.join(work, name);
  const config = { mode: 'production', devtool: false, entry: path.join(root, 'tests/fixtures/sx-bundle', fixture),
    output: { path: dir, filename: 'entry.js', chunkFilename: '[name].js', library: { type: 'commonjs2' } },
    resolve: { extensions: ['.tsx', '.ts', '.js'], alias: { '@apvee/spfx-react-toolkit': library }, modules: [path.join(root, 'node_modules'), 'node_modules'] },
    resolveLoader: { modules: [path.join(root, 'node_modules')] },
    externals: [({ request }, done) => /^(react|react-dom)$/.test(request || '') || /^@microsoft\/sp-/.test(request || '') ? done(null, `commonjs ${request}`) : done()],
    module: { rules: [{ test: /\.tsx?$/, use: path.join(root, 'scripts/sx-browser-typescript-loader.cjs') }] },
    optimization: { concatenateModules: false, usedExports: true, minimize: true } };
  const stats = await new Promise((resolve, reject) => {
    const compiler = webpack(config);
    compiler.run((error, result) => compiler.close(closeError => error || closeError ? reject(error || closeError) : resolve(result)));
  });
  const detail = stats.toJson({ all: false, assets: true, chunks: true, modules: true, usedExports: true, reasons: true, errors: true, warnings: true });
  fs.writeFileSync(path.join(output, `${name}-stats.json`), JSON.stringify(detail, null, 2));
  if (stats.hasErrors() || detail.warnings.some(w => /export .*was not found|Module not found/.test(w.message))) throw new Error(`${name}: ${JSON.stringify([...detail.errors, ...detail.warnings])}`);
  // All emitted chunk assets are reachable from this single entry. Source maps
  // and license sidecars are excluded. A Set counts shared chunks only once.
  const names = [...new Set(detail.chunks.flatMap(chunk => chunk.files).filter(f => f.endsWith('.js')))];
  const source = names.map(n => fs.readFileSync(path.join(dir, n), 'utf8')).join('\n');
  for (const n of names) fs.copyFileSync(path.join(dir, n), path.join(output, `${name}-${n}`));
  const modules = detail.modules.filter(m => /helpers\/styles|hooks\/useSx|@griffel|react-theme|react-shared-contexts|@pnp\/sp/.test(m.name || ''))
    .map(m => ({ name: m.name.replace(library, '<tarball>'), sizeBeforeMinification: m.size, usedExports: m.usedExports, chunks: m.chunks }));
  const item = { name, fixture, ...filesSize(dir, names), warnings: detail.warnings, modules, retainedDataFactoryCalls: dataCalls(source),
    upstreamTokenNames: [...new Set(source.match(/var\(--[a-zA-Z0-9_-]+\)/g) || [])].sort(),
    upstreamTypographyNames: ['body1', 'body1Strong', 'body1Stronger', 'body2', 'caption1', 'caption1Strong', 'caption1Stronger', 'caption2', 'caption2Strong', 'subtitle1', 'subtitle2', 'subtitle2Stronger', 'title1', 'title2', 'title3', 'largeTitle', 'display'].filter(n => new RegExp(`(?:["']${n}["']|\\b${n}):`).test(source)) };
  report.bundles.push(item); save(); console.log(`${name}: ${JSON.stringify(item.totals)}`);
}
function cssGrowth() {
  const harness = require('../tests/fixtures/sx-harness.cjs');
  harness.createSxDom();
  const { createDOMRenderer } = require('@griffel/core');
  const { styles, load } = harness.loadSxModules();
  const useSx = load('../../hooks/useSx').useSx;
  const { getSxRendererCache } = load('cache.internal');
  const renderer = createDOMRenderer(document);
  const snapshots = [];
  function snapshot(stage) {
    const rules = harness.cssRules(renderer);
    const cache = getSxRendererCache(renderer);
    snapshots.push({ stage, rules: rules.length, cssBytes: Buffer.byteLength(rules.join('\n')), insertionCache: Object.keys(renderer.insertionCache).length,
      bindings: cache.bindings.size, assignments: cache.assignments.size });
  }
  snapshot('before mount');
  const mounted = harness.mountSx(useSx, { renderer, inputs: [styles.width.px(240), styles.typography.body1] });
  snapshot('cold mount: width240/body1');
  for (let i = 0; i < 100; i++) mounted.render({ renderer, inputs: [styles.width.px(240), styles.typography.body1] });
  snapshot('100 identical renders');
  for (let i = 0; i < 100; i++) mounted.render({ renderer, inputs: [styles.width.px(i % 2 ? 240 : 360), styles.typography.body1] });
  snapshot('100 alternating values240/360');
  for (let i = 0; i < 1000; i++) mounted.render({ renderer, inputs: [styles.width.px(1000 + i), styles.typography.body1] });
  snapshot('1000 new distinct widths1000..1999');
  mounted.unmount(); snapshot('after unmount');
  report.css = { environment: 'Real React17/ReactDOM act, real Griffel DOM renderer and insertionCache; JSDOM15 CSSOM, no CSS mocks. Not a browser computed-style/tenant validation.', snapshots };
  if (snapshots[1].rules !== snapshots[2].rules || snapshots[4].rules - snapshots[3].rules !== 1000 || snapshots[4].assignments - snapshots[3].assignments !== 1000 || snapshots[5].rules !== snapshots[4].rules) throw new Error('Unexpected CSS growth semantics');
  save();
}
async function main() {
  for (const name of ['typescript', 'webpack', 'react', '@griffel/core', '@griffel/react', '@fluentui/react-theme', '@fluentui/react-shared-contexts']) report.versions[name] = require(`${name}/package.json`).version;
  report.lockfileSha256 = hash(fs.readFileSync(path.join(root, 'package-lock.json')));
  report.fixtureHost = { react: 'host external17.0.1', spfx: 'all @microsoft/sp-* are host externals1.21.1', webpack: 'locked installed production,minimize,usedExports; module concatenation disabled for attribution; no custom sideEffects override', configSha256: hash(fs.readFileSync(path.join(root, 'apps/spfx-react-toolkit-test/config/config.json'))) };
  report.fixtureHost.manifests = fs.readdirSync(path.join(root, 'artifacts/sx-baseline/dist')).filter(n => n.endsWith('.manifest.json')).map(name => ({ name, sha256: hash(fs.readFileSync(path.join(root, 'artifacts/sx-baseline/dist', name))) }));
  report.npm = command('npm-version', 'npm', ['--version']).stdout.trim();
  const baselineTarball = path.join(root, 'artifacts/sx-baseline/library.tgz');
  report.baselineTarballSha256 = hash(fs.readFileSync(baselineTarball));
  if (report.baselineTarballSha256 !== 'd87c017581e27d58d2525be9133041953341dc85f9ea61601d68ba92303fa7d5') throw new Error('Frozen baseline tarball hash mismatch');
  command('build-library', 'npm', ['run', 'build:library']);
  command('typecheck-fixtures', process.execPath, [require.resolve('typescript/bin/tsc'), '--noEmit', '--strict', '--skipLibCheck', '--target', 'es2020', '--module', 'esnext', '--moduleResolution', 'node', '--jsx', 'react', ...fs.readdirSync(path.join(root, 'tests/fixtures/sx-bundle')).filter(f => /\.tsx?$/.test(f)).map(f => path.join(root, 'tests/fixtures/sx-bundle', f))]);
  const candidateDir = path.join(work, 'candidate-pack'); fs.mkdirSync(candidateDir, { recursive: true });
  if (!process.env.SX_MEASURE_CANDIDATE_TARBALL) command('pack-candidate', 'npm', ['pack', './packages/spfx-react-toolkit', '--ignore-scripts', '--pack-destination', candidateDir], { env: { ...process.env, npm_config_cache: path.join(work, 'npm-cache') } });
  const candidateTarball = process.env.SX_MEASURE_CANDIDATE_TARBALL ? path.resolve(process.env.SX_MEASURE_CANDIDATE_TARBALL) : path.join(candidateDir, fs.readdirSync(candidateDir).find(f => f.endsWith('.tgz')));
  report.candidateTarballSha256 = hash(fs.readFileSync(candidateTarball));
  fs.copyFileSync(candidateTarball, path.join(output, 'candidate.tgz'));
  const unpacked = {};
  for (const [name, tarball] of [['baseline', baselineTarball], ['candidate', candidateTarball]]) {
    const dir = path.join(work, name); fs.mkdirSync(dir, { recursive: true });
    command(`extract-${name}`, 'tar', ['-xzf', tarball, '-C', dir]); unpacked[name] = path.join(dir, 'package');
  }
  await bundle('unchanged-baseline', 'unchanged.ts', unpacked.baseline);
  await bundle('unchanged-candidate', 'unchanged.ts', unpacked.candidate);
  await bundle('minimum-direct', 'direct.tsx', unpacked.candidate);
  await bundle('minimum-wrapper', 'wrapper.tsx', unpacked.candidate);
  await bundle('minimum-wrapper-deep', 'wrapper-deep.tsx', unpacked.candidate);
  await bundle('already-fluent-direct', 'fluent-direct.tsx', unpacked.candidate);
  await bundle('already-fluent-wrapper', 'fluent-wrapper.tsx', unpacked.candidate);
  await bundle('full-ui', 'full.tsx', unpacked.candidate);
  if (!process.env.SX_MEASURE_CANDIDATE_TARBALL) {
    const unused = report.bundles.find(b => b.name === 'unchanged-candidate');
    if (unused.retainedDataFactoryCalls.length !== 0) throw new Error('Unused facade allocations survived in the unchanged consumer');
    for (const name of ['minimum-wrapper', 'minimum-wrapper-deep']) {
      const calls = report.bundles.find(b => b.name === name).retainedDataFactoryCalls;
      const expected = calls.length === 6 && calls.filter(c => c.kind === 'recipe').length === 1 && calls.some(c => c.kind === 'recipe' && c.name === 'body1') && calls.some(c => c.kind === 'declaration' && c.property === 'width' && c.value === '100%') && ['fontFamily', 'fontSize', 'fontWeight', 'lineHeight'].every(p => calls.some(c => c.kind === 'declaration' && c.property === p));
      if (!expected) throw new Error(`${name} retained unexpected descriptor members`);
    }
  }
  cssGrowth();
  const baselineMetadata = JSON.parse(fs.readFileSync(path.join(root, '.docs/maintenance/evidence/use-sx/baseline.json')));
  report.baselineShip = shipCapture(path.join(root, 'artifacts/sx-baseline/dist'), baselineMetadata.files.filter(f => f.currentBuild).map(f => f.name));
  report.baselineShip.status = baselineMetadata.commands;
  if (process.env.SX_MEASURE_SKIP_SHIP !== '1') {
    const start = Date.now();
    const preload = path.resolve(process.env.SX_BUILD_METADATA_PRELOAD || path.join(root, 'scripts/gulp-build-metadata-preload.cjs'));
    if (!fs.existsSync(preload)) throw new Error(`Missing visible-warning metadata preloader: ${preload}`);
    // The verified guard keeps exact known data-age advisories visible on stdout
    // in workers, without importing metadata there. Only the Gulp parent warms
    // metadata. Every other warning/error/stderr retains its original behavior.
    command('bundle-ship', process.execPath, ['--require', preload, require.resolve('gulp/bin/gulp.js'), 'bundle', '--ship'], { cwd: path.join(root, 'apps/spfx-react-toolkit-test'), env: { ...process.env, NODE_OPTIONS: `${process.env.NODE_OPTIONS || ''} --require ${JSON.stringify(preload)}`.trim() } }, true);
    const dist = path.join(root, 'apps/spfx-react-toolkit-test/dist');
    const current = fs.readdirSync(dist).filter(n => fs.statSync(path.join(dist, n)).mtimeMs >= start);
    report.candidateShip = shipCapture(dist, current);
    report.candidateShip.status = report.commands.find(c => c.name === 'bundle-ship');
    const archive = path.join(output, 'ship-dist'); fs.mkdirSync(archive, { recursive: true });
    for (const item of report.candidateShip.files) fs.copyFileSync(path.join(dist, item.name), path.join(archive, item.name));
    for (const item of report.candidateShip.host) fs.copyFileSync(path.join(dist, item.name), path.join(archive, item.name));
    report.candidateShip.clean = report.candidateShip.status.status === 0 && report.candidateShip.status.failures.length === 0;
    save();
    if (!report.candidateShip.clean) throw new Error('Ship emitted measured assets but reported failure; this is not a clean ship pass');
  } else report.caveats.push('Ship build skipped explicitly; intermediate measurement only.');
  save(); console.log(`Evidence: ${output}/measurement.json`);
}
main().catch(error => { report.error = error.stack; save(); console.error(error); process.exitCode = 1; });

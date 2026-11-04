const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const crypto = require('node:crypto');
const { spawnSync } = require('node:child_process');
const repository = path.resolve(__dirname, '..');
const fixture = path.join(__dirname, 'fixtures/documentation/artifact-set');
function mirror(t, real = false) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-build-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  if (real) {
    for (const route of ['README.md', 'LICENSE', 'AGENTS.md', 'package.json', 'docs', 'packages/spfx-react-toolkit', 'apps/spfx-react-toolkit-test/src', 'apps/spfx-react-toolkit-test/config', 'apps/spfx-react-toolkit-test/package.json', 'apps/spfx-react-toolkit-test/gulpfile.js', '.github/workflows/verify.yml', '.vscode/launch.json', '.vscode/settings.json', 'tests/fixtures']) {
      fs.mkdirSync(path.dirname(path.join(root, route)), { recursive: true });
      fs.cpSync(path.join(repository, route), path.join(root, route), { recursive: true, filter: source => !/(?:^|\/)(?:lib|node_modules|\.docs)(?:\/|$)/.test(source) });
    }
  } else fs.cpSync(fixture, root, { recursive: true });
  fs.copyFileSync(path.join(repository, 'packages/spfx-react-toolkit/tsdoc.json'), path.join(root, 'packages/spfx-react-toolkit/tsdoc.json'));
  fs.copyFileSync(path.join(repository, 'package-lock.json'), path.join(root, 'package-lock.json'));
  require('./fixtures/generated-tooling/helpers.cjs').copyEngineInputs(root, ['packages/spfx-react-toolkit/scripts/localization-resources.cjs', 'packages/spfx-react-toolkit/scripts/pack-library.cjs']);
  fs.symlinkSync(path.join(repository, 'node_modules'), path.join(root, 'node_modules'), 'dir');
  return root;
}
function run(root, args = [], mode, cwd = root, timeout) {
  const env = { ...process.env }; delete env.TOOLKIT_DOCS_MODE;
  if (mode !== undefined) env.TOOLKIT_DOCS_MODE = mode;
  return spawnSync(process.execPath, [path.join(root, 'scripts/docs/documentation.cjs'), ...args], { cwd, env, encoding: 'utf8', timeout });
}
function snapshot(root) {
  const result = {};
  function visit(route) {
    const file = path.join(root, route), stat = fs.statSync(file);
    result[route] = { mtime: stat.mtimeMs };
    if (stat.isDirectory()) for (const name of fs.readdirSync(file).sort()) visit(route + '/' + name);
    else result[route].hash = crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
  }
  visit('docs');
  for (const route of ['llms.txt', 'packages/spfx-react-toolkit/README.md']) if (fs.existsSync(path.join(root, route))) visit(route);
  return result;
}
test('CLI defaults to write from another cwd and environment check never repairs outputs', t => {
  const root = mirror(t);
  const written = run(root, [], undefined, os.tmpdir());
  assert.equal(written.status, 0, written.stderr);
  assert.ok(fs.existsSync(path.join(root, 'llms.txt')), 'default write must generate the complete set');
  fs.writeFileSync(path.join(root, 'llms.txt'), 'stale\n');
  const before = snapshot(root);
  const checked = run(root, [], 'check');
  assert.notEqual(checked.status, 0);
  assert.match(checked.stderr, /stale.*llms.txt/is);
  assert.deepEqual(snapshot(root), before);
  assert.equal(run(root, [], 'write').status, 0);
});
test('CLI --check overrides write, rejects unknown env and arguments, and validates before writes', t => {
  const root = mirror(t);
  const initial = snapshot(root);
  const check = run(root, ['--check'], 'write');
  assert.notEqual(check.status, 0);
  assert.deepEqual(snapshot(root), initial);
  for (const [args, mode] of [[[], 'repair'], [['--unknown'], undefined], [['--check'], 'repair']]) {
    const result = run(root, args, mode);
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /mode|usage/i);
    assert.deepEqual(snapshot(root), initial);
  }
  const source = path.join(root, 'packages/spfx-react-toolkit/src/index.ts');
  fs.appendFileSync(source, '\nexport const broken = ;\n');
  const invalid = run(root);
  assert.notEqual(invalid.status, 0);
  assert.match(invalid.stderr, /TS1109/);
  assert.deepEqual(snapshot(root), initial);
});
test('library build keeps compilation/localization and CI checks explicit generated artifacts before build', () => {
  const metadata = require('../packages/spfx-react-toolkit/package.json');
  assert.equal(metadata.scripts.build, 'node scripts/localization-resources.cjs --check-inventory && tsc -p tsconfig.json && node scripts/localization-resources.cjs --emit');
  assert.equal(metadata.scripts['pack:artifact'], 'node scripts/pack-library.cjs');
  const root = require('../package.json');
  assert.equal(root.scripts['pack:library'], 'npm run build:library && npm run pack:artifact --workspace @apvee/spfx-react-toolkit');
  assert.equal(root.scripts['docs:generate'], 'node scripts/docs/documentation.cjs');
  assert.equal(root.scripts['docs:check'], 'node scripts/docs/documentation.cjs --check');
  assert.equal(root.scripts['generated:sync'], 'node scripts/docs/generated.cjs');
  assert.equal(root.scripts['generated:check'], 'node scripts/docs/generated.cjs --check');
  const workflow = fs.readFileSync(path.join(repository, '.github/workflows/verify.yml'), 'utf8');
  assert.ok(!workflow.includes('TOOLKIT_DOCS_MODE'));
  assert.ok(workflow.indexOf('scripts/docs/generated.cjs --check') < workflow.indexOf('- run: npm run build'));
  assert.match(workflow, /network-namespace-probe.cjs/);
  for (const command of ['build', 'package:solution', 'verify:package']) assert.ok(workflow.includes('- run: npm run ' + command));
});
const generationTimeoutMs = 180000;
const realBuildTimeoutMs = 110000;
// Baseline generation, two real builds/checks and bounded mirror/assertion overhead.
const copiedBuildTimeoutMs = generationTimeoutMs + 2 * (realBuildTimeoutMs + generationTimeoutMs) + 10000;
test('real copied library build compiles/localizes stale or missing documents; separate generated check refuses without writes', { timeout: copiedBuildTimeoutMs }, t => {
  const started = process.hrtime.bigint(), boundaries = [];
  const elapsed = () => Number(process.hrtime.bigint() - started) / 1e6;
  let completed = false;
  t.after(() => {
    if (!completed) t.diagnostic(JSON.stringify({ diagnostic: 'copied library subprocesses', elapsedMs: elapsed(), boundaries }));
  });
  function normalExit(label, expectedStatus, operation) {
    const before = elapsed(), result = operation();
    const detail = { boundary: label, startedMs: before, elapsedMs: elapsed() - before,
      status: result.status, signal: result.signal, pid: result.pid,
      error: result.error === undefined ? undefined : { name: result.error.name, code: result.error.code, message: result.error.message.slice(0, 1024) },
      stdoutBytes: Buffer.byteLength(result.stdout ?? ''), stderrBytes: Buffer.byteLength(result.stderr ?? ''),
      stdout: result.stdout?.slice(0, 2048), stderr: result.stderr?.slice(0, 2048) };
    boundaries.push(detail);
    const context = label + ': ' + JSON.stringify(detail);
    assert.equal(result.error, undefined, context);
    // spawnSync uses the legacy Node null contract when no signal terminated the child.
    assert.equal(result.signal, null, context);
    assert.equal(result.status, expectedStatus, context);
    return result;
  }
  const root = mirror(t, true);
  const baseline = normalExit('baseline documentation write', 0, () => run(root, [], undefined, root, generationTimeoutMs));
  assert.equal(baseline.status, 0, baseline.stderr);
  const source = path.join(root, 'packages/spfx-react-toolkit/src/hooks/useStableCallback.ts');
  const original = fs.readFileSync(source, 'utf8');
  fs.writeFileSync(source, original.replace('/**', '/**\n * Build integration freshness probe.'));
  const developmentSource = path.join(root, 'docs/_sources/development/documentation.md');
  const originalDevelopment = fs.readFileSync(developmentSource, 'utf8');
  fs.appendFileSync(developmentSource, '\nBuild integration development freshness probe.\n');
  const developmentRoutes = Object.keys(snapshot(root)).filter(route => route.startsWith('docs/development/') && route.endsWith('.md')).sort();
  const expectedDevelopmentRoutes = [
    'docs/development/INDEX.md',
    'docs/development/documentation.md',
    'docs/development/native-executable-scenarios.md',
    'docs/development/setup.md',
    'docs/development/sharepoint-validation/INDEX.md',
    'docs/development/verification.md',
  ];
  assert.deepEqual(developmentRoutes, expectedDevelopmentRoutes);
  function build(label, expectedRoutes) {
    const before = snapshot(root);
    const result = normalExit(label + ':build', 0, () => spawnSync('npm', ['run', 'build:library'], { cwd: root, env: { ...process.env, TOOLKIT_DOCS_MODE: 'check' }, encoding: 'utf8', timeout: realBuildTimeoutMs }));
    assert.equal(result.status, 0, result.stdout + result.stderr);
    const checked = normalExit(label + ':check', 1, () => spawnSync(process.execPath, [path.join(root, 'scripts/docs/generated.cjs'), '--check'], { cwd: root, encoding: 'utf8', timeout: generationTimeoutMs }));
    assert.match(checked.stderr, /stale generated/i, checked.stderr);
    for (const route of expectedRoutes) assert.ok(checked.stderr.includes(route), label + ': missing stale route ' + route);
    assert.match(result.stdout, /Emitted|localization|resources/i);
    assert.ok(fs.existsSync(path.join(root, 'packages/spfx-react-toolkit/lib/index.js')), 'real tsc must run');
    const inventory = JSON.parse(fs.readFileSync(path.join(root, 'packages/spfx-react-toolkit/localization/resources.json'), 'utf8'));
    for (const resource of inventory.resources) for (const locale of resource.locales) {
      const output = path.join(root, 'packages/spfx-react-toolkit', resource.path.replace('{locale}', locale.toLowerCase()));
      assert.match(fs.readFileSync(output, 'utf8'), /define\(\[\], function/, 'real localization emission must run');
    }
    assert.deepEqual(snapshot(root), before, label + ': docs bytes and mtimes must survive');
    t.diagnostic(JSON.stringify({ label, status: result.status, documents: Object.keys(before).length, unchanged: true }));
  }
  build('stale JSDoc and development source', ['docs/api/hooks/react.md', 'docs/development/documentation.md']);
  fs.writeFileSync(source, original);
  fs.writeFileSync(developmentSource, originalDevelopment);
  fs.unlinkSync(path.join(root, 'llms.txt'));
  fs.rmSync(path.join(root, 'docs/development'), { recursive: true });
  build('missing development directory and llms.txt', [...expectedDevelopmentRoutes, 'llms.txt']);
  assert.equal(fs.existsSync(path.join(root, 'docs/development')), false, 'build must not recreate the missing directory');
  completed = true;
});

test('package-owned pack uses its library and repository artifacts from a neutral cwd', t => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'library-pack-boundary-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const library = path.join(root, 'packages/spfx-react-toolkit');
  fs.mkdirSync(path.join(library, 'scripts'), { recursive: true });
  fs.copyFileSync(path.join(repository, 'packages/spfx-react-toolkit/scripts/pack-library.cjs'), path.join(library, 'scripts/pack-library.cjs'));
  fs.writeFileSync(path.join(library, 'package.json'), JSON.stringify({ name: 'toolkit-pack-boundary', version: '1.0.0', files: ['payload.txt'] }));
  fs.writeFileSync(path.join(library, 'payload.txt'), 'owned library payload\n');
  const result = spawnSync(process.execPath, [path.join(library, 'scripts/pack-library.cjs')], { cwd: os.tmpdir(), encoding: 'utf8', timeout: 30000 });
  assert.equal(result.status, 0, result.stdout + result.stderr);
  const archive = path.join(root, 'artifacts/toolkit-pack-boundary-1.0.0.tgz');
  assert.ok(fs.existsSync(archive), result.stdout + result.stderr);
  const listing = spawnSync('tar', ['-tzf', archive], { encoding: 'utf8' });
  assert.equal(listing.status, 0, listing.stderr);
  assert.deepEqual(listing.stdout.trim().split('\n').sort(), ['package/package.json', 'package/payload.txt']);
  const payload = spawnSync('tar', ['-xOzf', archive, 'package/payload.txt'], { encoding: 'utf8' });
  assert.equal(payload.status, 0, payload.stderr);
  assert.equal(payload.stdout, 'owned library payload\n');
});

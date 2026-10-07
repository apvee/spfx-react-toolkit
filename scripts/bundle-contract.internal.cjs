const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const zlib = require('node:zlib');
const { createRequire } = require('node:module');
const { execFileSync } = require('node:child_process');
const { isDeepStrictEqual } = require('node:util');
const ts = require('typescript');
const webpack = require('webpack');
const root = path.resolve(__dirname, '..');
const toolkit = '@apvee/spfx-react-toolkit';
const hash = data => crypto.createHash('sha256').update(data).digest('hex');
const slash = value => value.replace(/\\/g, '/');
function normalizeResource(value, context) {
  value = slash(value || '').split('!').pop();
  const marker = '/node_modules/';
  if (value.includes(marker)) return value.slice(value.lastIndexOf(marker) + marker.length);
  return context && path.isAbsolute(value) ? slash(path.relative(context, value)) : value.replace(/^\.\//, '');
}
function collectEmittedModules(stats, context) {
  const result = new Map();
  function visit(module, chunkId) {
    if (module.filteredChildren > 0 || module.filteredModules > 0) throw new Error('Truncated emitted module stats');
    const name = module.nameForCondition || module.identifier || module.name;
    if (name) {
      const resource = normalizeResource(name, context);
      if (!result.has(resource)) result.set(resource, { resource, chunks: [], usedExports: module.usedExports ?? null, sizeBeforeMinification: module.size ?? null, reasonCounts: { total: (module.reasons ?? []).length, active: (module.reasons ?? []).filter(reason => reason.active).length } });
      const item = result.get(resource);
      if (!item.chunks.includes(chunkId)) item.chunks.push(chunkId);
    }
    for (const child of module.modules ?? []) visit(child, chunkId);
  }
  // Top-level stats.modules also contains parsed modules eliminated by DCE.
  // Only chunk membership establishes emission, including nested concatenation.
  for (const chunk of stats.chunks ?? []) for (const module of chunk.modules ?? []) visit(module, chunk.id);
  return [...result.values()].sort((a, b) => a.resource.localeCompare(b.resource));
}
function assertToolkitModuleOrigins(stats, toolkitDir) {
  const owners=new Map();
  function owner(directory) {
    if(owners.has(directory)) return owners.get(directory);
    const manifest=path.join(directory,'package.json');
    const parent=path.dirname(directory);
    const name=fs.existsSync(manifest) ? JSON.parse(fs.readFileSync(manifest,'utf8')).name : parent===directory ? undefined : owner(parent);
    owners.set(directory,name);return name;
  }
  function visit(module) {
    const actual=module.nameForCondition;
    if(actual && owner(path.dirname(actual))===toolkit && !actual.startsWith(toolkitDir+path.sep)) throw new Error(`Toolkit origin escaped consumer: ${actual}`);
    for(const child of module.modules ?? []) visit(child);
  }
  for(const chunk of stats.chunks ?? []) for(const module of chunk.modules ?? []) visit(module);
}
function collectReachableAssets(stats) {
  const chunks = new Map((stats.chunks ?? []).map(chunk => [chunk.id, chunk]));
  const initialIds = new Set(Object.values(stats.entrypoints ?? {}).flatMap(entry => entry.chunks ?? []));
  if (!initialIds.size) throw new Error('Missing entrypoint chunks');
  const reached = new Set();
  function visit(id) {
    if (reached.has(id)) return;
    const chunk = chunks.get(id);
    if (!chunk) throw new Error(`Missing reachable chunk ${id}`);
    reached.add(id);
    for (const child of chunk.children ?? []) visit(child);
    for (const group of Object.values(chunk.childrenByOrder ?? {})) for (const child of group) visit(child);
  }
  for (const id of initialIds) visit(id);
  const assets = ids => [...new Set([...ids].flatMap(id => chunks.get(id).files ?? []).filter(name => name.endsWith('.js')))].sort();
  const initial = assets(initialIds);
  const union = assets(reached);
  return { initial, async: union.filter(name => !initial.includes(name)), union };
}
function sizes(data) {
  return { raw: data.length, gzip: zlib.gzipSync(data, { level: 9 }).length,
    brotli: zlib.brotliCompressSync(data, { params: { [zlib.constants.BROTLI_PARAM_QUALITY]: 11 } }).length, sha256: hash(data) };
}
function renderConsumer(template, variant) {
  const slots = { ...variant.imports };
  if (variant.widthExport !== undefined || variant.widthExpression !== undefined) {
    if (!['width', 'full'].includes(variant.widthExport) || variant.widthExpression !== (variant.widthExport === 'width' ? 'width.full' : 'full')) throw new Error('Invalid width slot pair');
    slots.WIDTH_EXPORT = variant.widthExport; slots.WIDTH_FULL = variant.widthExpression;
  }
  let rendered = template;
  for (const [slot, value] of Object.entries(slots)) {
    if (typeof value !== 'string' || /['"\n\r]/.test(value)) throw new Error(`Invalid slot ${slot}`);
    rendered = rendered.split(`__${slot}__`).join(value);
  }
  if (/__[A-Z_]+__/.test(rendered)) throw new Error(`Unresolved consumer slot ${rendered.match(/__[A-Z_]+__/)[0]}`);
  return rendered;
}
function compareBundleReports(reference, candidate, { fixtureIds } = {}) {
  const failures = [];
  if (reference.errors?.length || candidate.errors?.length) failures.push('Measurement errors prevent reproducibility comparison');
  const expectedFixtures = reference.fixtures.filter(f => !fixtureIds || fixtureIds.includes(f.id.split('/')[0]));
  for (const fixture of expectedFixtures) if (!candidate.fixtures.some(item => item.id === fixture.id)) failures.push(`${fixture.id}: missing fixture`);
  const stableProvenance = report => {
    const { consumerRoot, toolkitDir, ...stable } = report.provenance;
    return stable;
  };
  if (JSON.stringify(stableProvenance(reference)) !== JSON.stringify(stableProvenance(candidate))) failures.push('Measurement provenance drift');
  if (reference.contractSha256 !== candidate.contractSha256) failures.push('Contract drift');
  for (const fixture of candidate.fixtures) {
    const previous = reference.fixtures.find(item => item.id === fixture.id);
    if (!previous || previous.status !== fixture.status) { failures.push(`${fixture.id}: availability drift`); continue; }
    if (fixture.status !== 'measured') continue;
    if (previous.fixtureSha256 !== fixture.fixtureSha256 || previous.configSha256 !== fixture.configSha256) failures.push(`${fixture.id}: fixture/config drift`);
    if (JSON.stringify(previous.assets) !== JSON.stringify(fixture.assets) || JSON.stringify(previous.totals) !== JSON.stringify(fixture.totals)) failures.push(`${fixture.id}: asset drift`);
    const emitted = item => item.emittedModules?.map(module => module.resource);
    if (JSON.stringify(emitted(previous)) !== JSON.stringify(emitted(fixture))) failures.push(`${fixture.id}: emitted module drift`);
    if (JSON.stringify(previous.retentionObservations) !== JSON.stringify(fixture.retentionObservations)) failures.push(`${fixture.id}: retention drift`);
  }
  return failures;
}
function canonicalContract() {
  return JSON.parse(fs.readFileSync(path.join(root,'tests/fixtures/tree-shaking/contract.json'),'utf8'));
}
function hasCanonicalCompletion(report) {
  try {
    const contract=canonicalContract();
    if(report.contractSha256!==hash(JSON.stringify(contract))) return false;
    const expected=contract.fixtures.flatMap(fixture=>Object.keys(fixture.variants).map(variant=>`${fixture.id}/${variant}`)).sort();
    if(!Array.isArray(report.fixtures) || !isDeepStrictEqual(report.fixtures.map(fixture=>fixture.id).sort(),expected)) return false;
    for(const item of report.fixtures) {
      if(item.status!=='measured' || !Number.isFinite(item.totals?.union?.gzip)) return false;
      const fixture=contract.fixtures.find(fixture=>item.id.startsWith(fixture.id+'/'));
      if(checkFixture(item,fixture,contract).length) return false;
    }
    const checked=checkComparisons(report.fixtures,contract);
    return checked.failures.length===0 && isDeepStrictEqual(report.comparisons,checked.comparisons);
  } catch {return false;}
}
function finalizeReport(report) {
  return {...report,passed:report.mode==='enforce'
    ? report.complete===true && report.failures?.length===0 && !(report.errors?.length) && hasCanonicalCompletion(report)
    : null};
}
function dataCalls(source, staticOnly = false) {
  const file = ts.createSourceFile('bundle.js', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
  const calls = [];
  function visit(node) {
    if (ts.isCallExpression(node) && node.arguments.length >= 2 && ts.isStringLiteral(node.arguments[0])) {
      const first = node.arguments[0].text;
      if (ts.isArrayLiteralExpression(node.arguments[1]) && (/^[a-zA-Z]+\./.test(first) || /^(body|caption|subtitle|title|largeTitle|display)/.test(first))) calls.push({ kind: 'recipe', name: first });
      else if (node.arguments.length === 3 && ts.isObjectLiteralExpression(node.arguments[2]) && node.arguments[2].properties.some(p => p.name?.getText(file) === 'fallback')) {
        calls.push({ kind: 'declaration', property: first, value: ts.isStringLiteral(node.arguments[1]) ? node.arguments[1].text : staticOnly ? undefined : node.arguments[1].getText(file) });
      }
    }
    ts.forEachChild(node, visit);
  }
  visit(file); return calls;
}
function retentionObservations(source) {
  return dataCalls(source, true).map(({ value, ...call }) => value === undefined ? call : { ...call, staticValue: value });
}
function validateContract(contract) {
  if (contract.schemaVersion !== 1 || !Array.isArray(contract.fixtures) || !Array.isArray(contract.comparisons) || !contract.moduleGroups) throw new Error('Invalid bundle contract');
  const ids = new Set();
  for (const fixture of contract.fixtures) {
    if (!/^[a-z0-9-]+$/.test(fixture.id) || ids.has(fixture.id) || !fixture.template || !fixture.variants) throw new Error('Invalid/duplicate fixture');
    ids.add(fixture.id);
    for (const group of [...fixture.forbiddenModuleGroups, ...fixture.requiredModuleGroups]) if (!contract.moduleGroups[group]) throw new Error(`Unknown module group ${group}`);
    for (const variant of Object.keys(fixture.variants)) if (!/^[a-z0-9-]+$/.test(variant)) throw new Error('Invalid variant name');
  }
  for (const patterns of Object.values(contract.moduleGroups)) for (const pattern of patterns) new RegExp(pattern);
  for (const comparison of contract.comparisons) if (comparison.metric !== 'gzip' || (!comparison.informational && (!Number.isFinite(comparison.maxIncreaseBytes) || comparison.maxIncreaseBytes < 0))) throw new Error('Invalid comparison');
}
function checkComparisons(fixtures, contract) {
  const comparisons = [], failures = [];
  for (const comparison of contract.comparisons) {
    const candidate = fixtures.find(f => f.id === comparison.candidate && f.status === 'measured');
    const reference = fixtures.find(f => f.id === comparison.reference && f.status === 'measured');
    if (!candidate || !reference) {
      comparisons.push({ ...comparison, status: 'unavailable' });
      failures.push(`Unavailable comparison ${comparison.candidate} / ${comparison.reference}`);
    } else {
      const increaseBytes = candidate.totals.union.gzip - reference.totals.union.gzip;
      const passed = comparison.informational ? null : increaseBytes <= comparison.maxIncreaseBytes;
      comparisons.push({ ...comparison, increaseBytes, passed });
      if (passed === false) failures.push(`Gzip comparison ${comparison.candidate} increased ${increaseBytes} > ${comparison.maxIncreaseBytes}`);
    }
  }
  return { comparisons, failures };
}
async function runInstalledGates({consumerRoot,outputDir,failures,bundle,runtime,mutations,independent}) {
  const reports={};
  for(const [name,verify] of [['bundle',bundle],['runtime',runtime],['mutations',mutations]]) {
    try {
      const report=await verify({consumerRoot,outputDir:path.join(outputDir,name)});
      reports[name]=report;
      if(report.complete!==true || report.passed!==true || name==='bundle' && (report.mode!=='enforce' || finalizeReport(report).passed!==true)) throw new Error(`${name}: complete ${name==='bundle'?'enforce ':''}PASS required`);
    } catch(error) {failures.push({args:[`production-${name}`],message:error.message});}
  }
  await independent();
  return reports;
}
function baselineSnapshot(report) {
  if (report.mode !== 'enforce' || report.passed !== true || finalizeReport(report).passed !== true) throw new Error('Baseline requires a canonical complete enforce PASS');
  return { schemaVersion: 1, purpose: 'Informational optimized production totals; pairwise budgets and exclusions are in contract.json', provenance: report.provenance, contractSha256: report.contractSha256,
    fixtures: report.fixtures.map(({id, fixtureSha256, configSha256, totals, assets}) => ({id, fixtureSha256, configSha256, totals, assets})) };
}
function checkFixture(item, fixture, contract) {
  const failures = [];
  const matches = group => item.emittedModules.filter(module => contract.moduleGroups[group].some(pattern => new RegExp(pattern).test(module.resource)));
  for (const group of fixture.forbiddenModuleGroups) if (matches(group).length) failures.push(`${item.id}: forbidden module group ${group}`);
  for (const group of fixture.requiredModuleGroups) if (!matches(group).length) failures.push(`${item.id}: missing required module group ${group}`);
  const expected = fixture.retainedDescriptors;
  if (expected) {
    for (const call of item.retentionObservations) {
      if (call.kind === 'recipe' && !expected.recipeNames.includes(call.name)) failures.push(`${item.id}: unexpected recipe ${call.name}`);
      if (call.kind === 'declaration' && (!expected.declarationProperties.includes(call.property) || expected.staticValues[call.property] && !expected.staticValues[call.property].includes(call.staticValue))) failures.push(`${item.id}: unexpected declaration ${call.property}=${call.staticValue}`);
    }
    if (item.retentionObservations.length > expected.maxCallCount) failures.push(`${item.id}: ${item.retentionObservations.length} descriptor calls exceeds ${expected.maxCallCount}`);
  }
  return failures;
}
function buildConfig(consumerRoot, entry, directory, attribution) {
  return { context: consumerRoot, mode: 'production', devtool: false, cache: false, entry: { main: entry },
    output: { path: directory, filename: 'entry.js', chunkFilename: '[id].[contenthash].js', library: { type: 'commonjs2' }, clean: true },
    resolve: { extensions: ['.tsx', '.ts', '.js'], modules: [path.join(consumerRoot, 'node_modules')] },
    externals: [({ request }, done) => /^(react|react-dom)(\/|$)/.test(request || '') || /^@microsoft\/sp-/.test(request || '') ? done(null, `commonjs ${request}`) : done()],
    module: { rules: [{ test: /\.tsx?$/, use: path.join(root, 'scripts/sx-browser-typescript-loader.cjs') }] },
    optimization: { usedExports: true, minimize: true, concatenateModules: !attribution, moduleIds: 'deterministic', chunkIds: 'deterministic' } };
}
function statsOptions() {
  return { all: false, assets: true, chunks: true, chunkModules: true, chunkRelations: true, nestedModules: true, modules: false, dependentModules: true, orphanModules: true, runtimeModules: true, cachedModules: true, modulesSpace: Infinity, chunkModulesSpace: Infinity, nestedModulesSpace: Infinity, groupModulesByAttributes: false, groupModulesByCacheStatus: false, groupModulesByLayer: false, groupModulesByType: false, groupModulesByPath: false, groupModulesByExtension: false, ids: true, reasons: true, usedExports: true, errors: true, warnings: true, entrypoints: true };
}
async function compile(config) {
  return new Promise((resolve, reject) => {
    const compiler = webpack(config);
    compiler.run((error, stats) => compiler.close(closeError => error || closeError ? reject(error || closeError) : resolve(stats)));
  });
}
function readRuntimeToolchain() {
  const npm = execFileSync('npm', ['--version'], { encoding: 'utf8' }).trim();
  if (!/^\d+\.\d+\.\d+$/.test(npm)) throw new Error('Invalid npm version provenance');
  return { node: process.version, npm };
}
function consumerProvenance(consumerRoot) {
  const consumerRequire = createRequire(path.join(consumerRoot, 'package.json'));
  const toolkitDir = path.join(consumerRoot, 'node_modules', toolkit);
  if (fs.lstatSync(toolkitDir).isSymbolicLink() || fs.realpathSync(toolkitDir) !== toolkitDir) throw new Error('Toolkit must originate from installed consumer tarball, not a symlink');
  const manifest = JSON.parse(fs.readFileSync(path.join(consumerRoot, 'package.json')));
  const dependency = manifest.dependencies?.[toolkit];
  if (!dependency?.startsWith('file:')) throw new Error('Consumer must declare a file tarball dependency');
  const tarball = path.resolve(consumerRoot, dependency.slice(5));
  if (!tarball.endsWith('.tgz') || !fs.existsSync(tarball)) throw new Error('Missing consumer tarball provenance');
  const lock = JSON.parse(fs.readFileSync(path.join(root, 'package-lock.json')));
  const versions = {};
  for (const name of ['typescript', 'webpack', 'react', 'react-dom', '@pnp/sp', '@pnp/core', '@pnp/queryable', '@griffel/core', '@griffel/react', '@fluentui/react-theme', '@fluentui/react-utilities', '@fluentui/react-shared-contexts', '@microsoft/sp-core-library']) {
    const installed = consumerRequire(name + '/package.json').version;
    if (installed !== lock.packages['node_modules/' + name].version) throw new Error(`Consumer ${name} differs from locked toolchain`);
    versions[name] = installed;
  }
  if (ts.version !== versions.typescript || require('webpack/package.json').version !== versions.webpack) throw new Error('Measurement tooling differs from locked consumer toolchain');
  if (!/^v22\./.test(process.version) || Number(process.versions.node.split('.')[1]) < 14) throw new Error('Unsupported Node version');
  return { ...readRuntimeToolchain(), versions, rootLockfileSha256: hash(fs.readFileSync(path.join(root, 'package-lock.json'))), consumerLockfileSha256: hash(fs.readFileSync(path.join(consumerRoot, 'package-lock.json'))), tarballSha256: hash(fs.readFileSync(tarball)), toolkitDir, consumerRoot, loaderSha256: hash(fs.readFileSync(path.join(root, 'scripts/sx-browser-typescript-loader.cjs'))), minifierVersion: require('terser/package.json').version, zlibVersion: process.versions.zlib, compression: { gzipLevel: 9, brotliQuality: 11 } };
}
async function measureTreeShaking({ consumerRoot, outputDir, contract, mode, attribution = false, fixtureIds }) {
  if (mode === 'enforce' && (attribution || fixtureIds)) throw new Error('Enforce requires the complete production contract');
  if (!['audit', 'enforce'].includes(mode)) throw new Error('mode must be audit or enforce');
  validateContract(contract);
  if(mode==='enforce') {
    const canonical=canonicalContract();
    if(!isDeepStrictEqual(contract,canonical)) throw new Error('Enforce requires the canonical complete production contract');
    // Semantically identical copies are accepted, then use canonical ordering
    // for the authoritative hash and complete measurement inventory.
    contract=canonical;
  }
  consumerRoot = fs.realpathSync(consumerRoot); outputDir = path.resolve(outputDir);
  fs.mkdirSync(outputDir, { recursive: true });
  if (fixtureIds?.some(id => !contract.fixtures.some(f => f.id === id))) throw new Error('Unknown selected fixture');
  const provenance = consumerProvenance(consumerRoot);
  const report = { schemaVersion: 1, mode, complete: false, provenance, attribution, contractSha256: hash(JSON.stringify(contract)), fixtures: [], comparisons: [], warnings: [], errors: [], failures: [] };
  const start = Date.now();
  const save = () => fs.writeFileSync(path.join(outputDir, 'bundle-report.json'), JSON.stringify(finalizeReport({ ...report, durationMs: Date.now() - start }), null, 2) + '\n');
  save();
  try {
  for (const fixture of contract.fixtures.filter(f => !fixtureIds || fixtureIds.includes(f.id))) {
    const templatePath = path.resolve(root, 'tests/fixtures/tree-shaking/consumers', fixture.template);
    const template = fs.readFileSync(templatePath, 'utf8');
    for (const [variantId, variant] of Object.entries(fixture.variants)) {
      const id = `${fixture.id}/${variantId}`;
      const rendered = renderConsumer(template, variant);
      const entry = path.join(consumerRoot, '.tree-shaking/entries', fixture.id, variantId + path.extname(templatePath));
      fs.mkdirSync(path.dirname(entry), { recursive: true }); fs.writeFileSync(entry, rendered);
      const missing = Object.values(variant.imports).filter(specifier => specifier.startsWith(toolkit)).filter(specifier => {
        try { createRequire(path.join(consumerRoot, 'package.json')).resolve(specifier); return false; } catch { return true; }
      });
      // Only the clean domain/probe aliases planned for Task3 may be unavailable.
      const allowed = new Set(['core', 'hooks', 'services', 'helpers', 'styles', 'hooks/useStableCallback', 'styles/useSx'].map(s => toolkit + '/' + s));
      if (missing.length) {
        if (mode === 'audit' && missing.every(s => allowed.has(s))) { report.fixtures.push({ id, status: 'unavailable', specifiers: missing, templateSha256: hash(template), fixtureSha256: hash(rendered) }); save(); continue; }
        report.errors.push(`${id}: missing imports ${missing.join(', ')}`); save(); throw new Error(report.errors.at(-1));
      }
      const directory = path.join(outputDir, 'assets', fixture.id, variantId);
      const config = buildConfig(consumerRoot, entry, directory, attribution);
      const normalizedConfig = { ...config, context: '<consumer>', entry: { main: slash(path.relative(consumerRoot, entry)) }, output: { ...config.output, path: '<output>' }, resolve: { ...config.resolve, modules: ['<consumer>/node_modules'] }, module: { rules: [{ test: '\\.tsx?$', use: 'scripts/sx-browser-typescript-loader.cjs' }] }, externals: ['react/react-dom and @microsoft/sp-* host externals'] };
      const stats = await compile(config);
      const detail = stats.toJson(statsOptions());
      const statsFile = `${fixture.id}-${variantId}-stats.json.gz`;
      fs.writeFileSync(path.join(outputDir, statsFile), zlib.gzipSync(JSON.stringify(detail), { level: 9 }));
      report.warnings.push(...detail.warnings.map(w => ({ id, message: w.message })));
      if (stats.hasErrors() || detail.warnings.some(w => /export .*was not found|Module not found/.test(w.message))) { report.errors.push(...[...detail.errors, ...detail.warnings].map(e => ({ id, message: e.message }))); save(); throw new Error(`Build failed: ${id}`); }
      const emittedModules = collectEmittedModules(detail, consumerRoot);
      // Check actual module origin before normalization, including concatenated
      // children. A normalized resource alone could hide a workspace escape.
      assertToolkitModuleOrigins(detail,provenance.toolkitDir);
      const reachable = collectReachableAssets(detail);
      const assets = reachable.union.map(name => {
        const file = path.join(directory, name);
        if (!fs.existsSync(file)) throw new Error(`Missing emitted asset ${name}`);
        return { name, ...sizes(fs.readFileSync(file)), initial: reachable.initial.includes(name) };
      });
      const totals = names => assets.filter(a => names.includes(a.name)).reduce((sum, asset) => ({ raw: sum.raw + asset.raw, gzip: sum.gzip + asset.gzip, brotli: sum.brotli + asset.brotli }), { raw: 0, gzip: 0, brotli: 0 });
      const item = { id, status: 'measured', statsFile, templateSha256: hash(template), fixtureSha256: hash(rendered), configSha256: hash(JSON.stringify(normalizedConfig)), assets, totals: { initial: totals(reachable.initial), async: totals(reachable.async), union: totals(reachable.union) }, emittedModules,
        retentionObservations: retentionObservations(reachable.union.map(name => fs.readFileSync(path.join(directory, name), 'utf8')).join('\n')), warnings: detail.warnings.map(w => w.message), errors: [] };
      item.failures = checkFixture(item, fixture, contract);
      report.fixtures.push(item); report.failures.push(...item.failures); save();
      console.log(`${id}: gzip ${item.totals.union.gzip}; ${item.assets.length} JS assets; ${item.failures.length} contract observations`);
    }
  }
  const comparisonResult = checkComparisons(report.fixtures, contract);
  report.comparisons = comparisonResult.comparisons;
  report.failures.push(...comparisonResult.failures);
  report.complete = true;
  save(); return finalizeReport({ ...report, durationMs: Date.now() - start });
  } catch (error) { report.errors.push(error.message); save(); throw error; }
}
module.exports = { assertToolkitModuleOrigins, measureTreeShaking, collectEmittedModules, collectReachableAssets, finalizeReport, renderConsumer, dataCalls, retentionObservations, sizes, normalizeResource, validateContract, checkFixture, buildConfig, statsOptions, compareBundleReports, readRuntimeToolchain, checkComparisons, baselineSnapshot, runInstalledGates };

// Mutations copy toolkit bytes; shared installed dependencies are read-only.
// Only the search prototype mutation owns a separate copy of @pnp/sp as well.
function copyMutationConsumer(consumerRoot, directory, copyPnp = false) {
  fs.mkdirSync(path.join(directory,'node_modules'),{recursive:true});
  for (const entry of fs.readdirSync(path.join(consumerRoot,'node_modules'),{withFileTypes:true})) {
    if (entry.name === '@apvee' || copyPnp && entry.name === '@pnp') continue;
    fs.symlinkSync(path.join(consumerRoot,'node_modules',entry.name),path.join(directory,'node_modules',entry.name));
  }
  if(copyPnp) for(const entry of fs.readdirSync(path.join(consumerRoot,'node_modules/@pnp'))) {
    const source=path.join(consumerRoot,'node_modules/@pnp',entry), target=path.join(directory,'node_modules/@pnp',entry);
    fs.mkdirSync(path.dirname(target),{recursive:true});
    if(entry==='sp') fs.cpSync(source,target,{recursive:true}); else fs.symlinkSync(source,target);
  }
  const packageDir=path.join(directory,'node_modules',toolkit);
  fs.cpSync(path.join(consumerRoot,'node_modules',toolkit),packageDir,{recursive:true});
  const manifest=JSON.parse(fs.readFileSync(path.join(consumerRoot,'package.json')));
  manifest.dependencies[toolkit]='file:'+path.resolve(consumerRoot,manifest.dependencies[toolkit].slice(5));
  fs.writeFileSync(path.join(directory,'package.json'),JSON.stringify(manifest));
  fs.copyFileSync(path.join(consumerRoot,'package-lock.json'),path.join(directory,'package-lock.json'));
  return packageDir;
}
async function verifyContractMutations({consumerRoot,outputDir}) {
  const {compileProbe,executeProbe}=require('./verify-production-runtime.cjs');
  const {assertEntrypointContract}=require('./package-entrypoints.cjs');
  const contract=JSON.parse(fs.readFileSync(path.join(root,'tests/fixtures/tree-shaking/contract.json')));
  const started=Date.now();
  const report={schemaVersion:1,complete:false,passed:false,mutations:[],failures:[],provenance:{consumerRoot,installation:'reuses installed consumer dependencies; separate toolkit copies; no npm install'}};
  fs.mkdirSync(outputDir,{recursive:true});
  const save=()=>fs.writeFileSync(path.join(outputDir,'mutation-report.json'),JSON.stringify(report,null,2)+'\n');save();
  const definitions=[
    {id:'unwanted-pnp',family:'helper',expected:/helper\/.*forbidden module group pnp/,change(packageDir){fs.appendFileSync(path.join(packageDir,'lib/helpers/spfx-graph-path.helpers.js'),"\nimport '@pnp/sp/webs';\n");}},
    {id:'unwanted-descriptor',family:'sx-width-only',expected:/unexpected declaration width=auto|descriptor calls exceeds/,change(packageDir){fs.appendFileSync(path.join(packageDir,'lib/helpers/styles/width.js'),'\nglobalThis.__contractMutation = auto;\n');}},
    {id:'missing-new-export',expected:/exports|entrypoint|contract/i,change(packageDir){const file=path.join(packageDir,'package.json'),manifest=JSON.parse(fs.readFileSync(file));delete manifest.exports['./styles/useSx'];fs.writeFileSync(file,JSON.stringify(manifest));}},
    {id:'missing-legacy-export',expected:/exports|entrypoint|contract/i,change(packageDir){const file=path.join(packageDir,'package.json'),manifest=JSON.parse(fs.readFileSync(file));delete manifest.exports['./lib/hooks/useSx'];fs.writeFileSync(file,JSON.stringify(manifest));}},
    {id:'workspace-origin',family:'helper',expected:/Toolkit origin escaped consumer/,change(packageDir){const file=path.join(packageDir,'lib/helpers/spfx-graph-path.helpers.js');fs.rmSync(file);fs.symlinkSync(path.join(root,'packages/spfx-react-toolkit/lib/helpers/spfx-graph-path.helpers.js'),file);}},
    {id:'byte-inflation',family:'helper',expected:/Gzip comparison helper\/root increased/,change(packageDir){const noise=crypto.randomBytes(16000).toString('hex');fs.appendFileSync(path.join(packageDir,'lib/index.js'),`\nimport { buildOneDriveAppDataPath as originalPath } from './helpers/spfx-graph-path.helpers';\nexport function buildOneDriveAppDataPath(name) { globalThis.__contractInflation = '${noise}'; return originalPath(name); }\n`);}},
    {id:'lost-list-registration',scenario:'list',expected:/getByTitle|lists|undefined/,change(packageDir){removeRegistration(path.join(packageDir,'lib/services/spfx-pnp-list.service.js'),'@pnp/sp/webs');}},
    {id:'lost-batch-registration',scenario:'batch',expected:/batched.*function/,change(packageDir){removeRegistration(path.join(packageDir,'lib/services/spfx-pnp.service.js'),'@pnp/sp/batching');}},
    {id:'lost-search-prototype',scenario:'search',copyPnp:true,expected:/search.*function/,change(packageDir,directory){const file=path.join(directory,'node_modules/@pnp/sp/search/index.js'),source=fs.readFileSync(file,'utf8');const neutralized=source.replace(/SPFI\.prototype\.search = function \(query\) \{[\s\S]*?\n\};/,'');if(source===neutralized)throw new Error('Missing search prototype mutation target');fs.writeFileSync(file,neutralized);}},
  ];
  function removeRegistration(file,registration) {
    const original=fs.readFileSync(file,'utf8'),mutated=original.replace(`import '${registration}';`,'');
    if(mutated===original)throw new Error(`Missing registration mutation target: ${registration}`);
    fs.writeFileSync(file,mutated);
  }
  for(const definition of definitions) {
    const directory=fs.realpathSync(fs.mkdtempSync(path.join(require('node:os').tmpdir(),'toolkit-contract-mutation-')));
    const output=path.join(outputDir,definition.id);
    let rejection, evidence;
    try {
      const packageDir=copyMutationConsumer(consumerRoot,directory,definition.copyPnp);
      definition.change(packageDir,directory);
      if(definition.scenario) {
        const bundle=await compileProbe({consumerRoot:directory,outputDir:output,scenario:definition.scenario,variant:'leaf'});
        evidence={bundleSha256:bundle.bundleSha256};
        executeProbe(bundle.bundlePath,definition.scenario,directory);
      } else if(definition.family) {
        const selected={...contract,fixtures:contract.fixtures.filter(f=>f.id===definition.family),comparisons:contract.comparisons.filter(c=>c.candidate.startsWith(definition.family+'/')&&c.reference.startsWith(definition.family+'/'))};
        const measured=await measureTreeShaking({consumerRoot:directory,outputDir:output,contract:selected,mode:'audit'});
        evidence={reportFile:path.join(output,'bundle-report.json'),comparisons:measured.comparisons};
        if(measured.failures.length) rejection=measured.failures.join('\n');
      } else {
        const files=[];function walk(dir){for(const entry of fs.readdirSync(dir,{withFileTypes:true})){const file=path.join(dir,entry.name);entry.isDirectory()?walk(file):files.push(slash(path.relative(packageDir,file)));}}walk(packageDir);
        assertEntrypointContract(JSON.parse(fs.readFileSync(path.join(packageDir,'package.json'))),files);
      }
    } catch(error) {rejection=error.message;evidence={...evidence,exit:error.exit??null,signal:error.signal??null,stderr:error.stderr,failureMarkers:error.failureMarkers};}
    finally {fs.rmSync(directory,{recursive:true,force:true});}
    const passed=typeof rejection==='string' && definition.expected.test(rejection) && !/Module parse failed|SyntaxError|Module not found/.test(rejection);
    const item={id:definition.id,expected:definition.expected.source,detected:passed,rejection,evidence};
    report.mutations.push(item);if(!passed)report.failures.push(`${definition.id}: ${rejection||'mutation incorrectly accepted'}`);save();
    console.log(`contract mutation ${definition.id}: ${passed?'detected':'FAILED'}`);
  }
  report.complete=true;report.passed=report.failures.length===0;report.durationMs=Date.now()-started;save();return report;
}
module.exports.verifyContractMutations=verifyContractMutations;

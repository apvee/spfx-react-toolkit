'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createRequire } = require('node:module');
const { spawnSync } = require('node:child_process');
const { scenarios, mutations } = require('../tests/fixtures/localization-consumers/spfx-scenarios.cjs');
const root = path.resolve(__dirname, '..');
const packageName = '@apvee/spfx-react-toolkit';
const prefix = 'ApveeSpfxReactToolkit';
const stripAnsi = text => text.replace(/\x1b\[[0-?]*[ -/]*[@-~]/g, '');

function readJson(file) { return JSON.parse(fs.readFileSync(file, 'utf8')); }
function writeJson(file, value) { fs.mkdirSync(path.dirname(file), { recursive: true }); fs.writeFileSync(file, JSON.stringify(value, undefined, 2) + '\n'); }
function filesBelow(directory) {
  if (!fs.existsSync(directory)) return [];
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap(item => item.isDirectory()
    ? filesBelow(path.join(directory, item.name)) : [path.join(directory, item.name)]);
}

function strictResult(result, label) {
  const output = stripAnsi((result.stdout ?? '') + '\n' + (result.stderr ?? ''));
  const failureMarkers = output.match(/The build failed[^\n]*|Exiting with exit code:\s*[1-9]\d*[^\n]*|Error - [^\n]*/g) ?? [];
  assert.ifError(result.error);
  assert.equal(result.status, 0, `${label} exited ${result.status}\n${output}`);
  assert.deepEqual(failureMarkers, [], `${label} reported an inner build failure\n${output}`);
  return { output, failureMarkers };
}

function readAmd(file) {
  let dictionary;
  vm.runInNewContext(fs.readFileSync(file, 'utf8'), { define: (dependencies, factory) => {
    assert.equal(dependencies.length, 0, `Resource dependencies in ${file}`);
    dictionary = factory();
  } }, { filename: file });
  assert.ok(dictionary && typeof dictionary.__locale === 'string', `Missing native locale metadata in ${file}`);
  return dictionary;
}

function inspectArtifacts(project, scenario, ids) {
  const emitted = filesBelow(path.join(project, 'dist'));
  const manifests = emitted.filter(file => file.endsWith('.manifest.json')).map(file => ({ file, value: readJson(file) }));
  assert.equal(manifests.length, 1, `Expected one emitted component manifest for ${scenario.name}`);
  const resources = manifests[0].value.loaderConfig.scriptResources;
  const actualIds = Object.keys(resources).filter(id => id.startsWith(prefix)).sort();
  assert.deepEqual(actualIds, [...scenario.resources].sort(), `Unexpected manifest resources for ${scenario.name}: ${actualIds.join(', ')}`);
  const localeAssets = emitted.filter(file => file.endsWith('.js') && path.basename(file).startsWith(prefix));
  for (const file of localeAssets) assert.ok(actualIds.some(id => path.basename(file).startsWith(id + '_')), `Unexpected emitted locale asset: ${file}`);
  const resourceFiles = [];
  for (const id of actualIds) {
    assert.ok(ids.includes(id), `Unknown native manifest ID ${id}`);
    const resource = resources[id];
    const selected = scenario.locale ?? 'en-us';
    // SPFx --locale produces one ordinary path; without it the same native ID
    // produces localizedPath with the complete culture map and defaultPath.
    assert.equal(resource.type, scenario.locale ? 'path' : 'localizedPath', `Unexpected native resource path type for ${id}`);
    if (resource.type === 'localizedPath') {
      assert.ok(typeof resource.defaultPath === 'string' && fs.existsSync(path.join(project, 'dist', resource.defaultPath)), `Missing emitted default resource ${id}`);
      readAmd(path.join(project, 'dist', resource.defaultPath));
    }
    const paths = resource.type === 'path' ? { [selected]: resource.path } : resource.paths;
    assert.ok(Object.keys(paths).some(locale => locale.toLowerCase() === selected), `Selected locale ${selected} did not emit ${id}`);
    for (const [locale, value] of Object.entries(paths)) {
      const relative = typeof value === 'string' ? value : value.path;
      const file = path.join(project, 'dist', relative);
      assert.ok(fs.existsSync(file), `Missing emitted resource ${id}/${locale}: ${relative}`);
      const dictionary = readAmd(file);
      assert.equal(dictionary.__locale.toLowerCase(), locale.toLowerCase(), `Selected locale ${locale.toLowerCase()} did not emit ${id}: actual ${dictionary.__locale}`);
      if (scenario.custom) assert.equal(dictionary.__consumerMapping, true, 'Custom consumer mapping must supply the emitted native data');
      resourceFiles.push({ id, locale, file: path.relative(project, file), bytes: fs.statSync(file).size, nativeLocale: dictionary.__locale });
    }
  }
  const scripts = emitted.filter(file => file.endsWith('.js')).filter(file => !resourceFiles.some(asset => path.join(project, asset.file) === file));
  const scriptEvidence = scripts.map(file => {
    const text = fs.readFileSync(file, 'utf8');
    const amd = text.match(/\bdefine\((?:"[^"]*",)?(\[[^\]]*\])/);
    const amdDependencies = amd ? JSON.parse(amd[1]) : [];
    const entryResource = resources[manifests[0].value.loaderConfig.entryModuleId];
    return { file: path.relative(project, file), role: entryResource && path.basename(file) === entryResource.path ? 'entrypoint' : 'chunk', bytes: Buffer.byteLength(text), amdDependencies,
      nativeAmdDependencies: amdDependencies.filter(id => ids.includes(id)), nativeIds: ids.filter(id => text.includes(id)) };
  });
  if (scenario.lazy) {
    assert.ok(scriptEvidence.length > 1, 'Lazy editor must emit a separate JavaScript chunk');
    assert.ok(scriptEvidence.some(script => script.nativeIds.includes(scenario.resources[0])), 'Lazy editor external must occur in emitted JavaScript');
  }
  return { manifest: path.relative(project, manifests[0].file), manifestResources: actualIds, scriptResources: resources, resourceFiles, scripts: scriptEvidence };
}

function createProject({ consumerRoot, directory, scenario, inventory, installed }) {
  const project = path.join(directory, scenario.name);
  fs.mkdirSync(path.join(project, 'src'), { recursive: true });
  const host = readJson(path.join(consumerRoot, 'package.json'));
  writeJson(path.join(project, 'package.json'), {
    name: `localization-${scenario.name}`, version: '1.0.0', private: true,
    engines: host.engines, dependencies: host.dependencies, devDependencies: host.devDependencies,
    scripts: { 'bundle:ship': 'gulp bundle --ship' },
  });
  // Every fixture shares the already installed consumer tree, never a workspace
  // alias or an additional peer installation. The toolkit itself must be real.
  fs.symlinkSync(path.join(consumerRoot, 'node_modules'), path.join(project, 'node_modules'), 'dir');
  const config = { version: '2.0', bundles: { 'localization-probe': { components: [{ entrypoint: './lib/probe.js', manifest: './src/probe.manifest.json' }] } }, externals: {}, localizedResources: {} };
  writeJson(path.join(project, 'config/config.json'), config);
  const cli = spawnSync(process.execPath, [path.join(installed, 'cli/spfx-react-toolkit.cjs'), 'configure-localization', '--project', project], { encoding: 'utf8' });
  strictResult(cli, 'Packed localization CLI');
  const configured = readJson(path.join(project, 'config/config.json'));
  assert.deepEqual(Object.keys(configured.localizedResources).sort(), inventory.resources.map(resource => resource.id).sort(), 'All eight resources must be registered');
  const pagination = inventory.resources.find(resource => resource.id === 'ApveeSpfxReactToolkitPaginationStrings');
  if (scenario.custom || scenario.kind === 'asset' || scenario.kind === 'locale') {
    const customDirectory = path.join(project, 'custom');
    fs.mkdirSync(customDirectory);
    for (const locale of pagination.locales) {
      if (scenario.kind === 'asset' || (scenario.kind === 'locale' && locale === 'it-IT')) continue;
      const sourceAsset = path.join(installed, pagination.path.replace('{locale}', locale.toLowerCase()));
      const targetAsset = path.join(customDirectory, locale.toLowerCase() + '.js');
      if (scenario.custom) fs.writeFileSync(targetAsset, `define([], function () { return ${JSON.stringify({ ...readAmd(sourceAsset), __consumerMapping: true })}; });\n`);
      else fs.copyFileSync(sourceAsset, targetAsset);
    }
    configured.localizedResources[pagination.id] = 'custom/{locale}.js';
  }
  if (scenario.kind === 'registration') delete configured.localizedResources[pagination.id];
  writeJson(path.join(project, 'config/config.json'), configured);
  if (scenario.custom) strictResult(spawnSync(process.execPath, [path.join(installed, 'cli/spfx-react-toolkit.cjs'), 'configure-localization', '--project', project, '--check'], { encoding: 'utf8' }), 'Packed CLI custom mapping check');
  const baseTsconfig = readJson(path.join(root, 'apps/spfx-react-toolkit-test/tsconfig.json'));
  baseTsconfig.compilerOptions.typeRoots = ['./node_modules/@types', './node_modules/@microsoft'];
  writeJson(path.join(project, 'tsconfig.json'), baseTsconfig);
  writeJson(path.join(project, 'config/write-manifests.json'), { cdnBasePath: '<!-- PATH TO CDN -->' });
  const manifest = readJson(path.join(root, 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/SpFxReactToolkitTestWebPart.manifest.json'));
  manifest.alias = 'LocalizationProbe';
  writeJson(path.join(project, 'src/probe.manifest.json'), manifest);
  let source = scenario.source ?? scenarios.find(item => item.name === 'pagination-control').source;
  if (scenario.kind === 'unknown') {
    source = "import strings from 'ApveeSpfxReactToolkitUnknownStrings';\nexport default strings;";
    fs.writeFileSync(path.join(project, 'src/native.d.ts'), "declare module 'ApveeSpfxReactToolkitUnknownStrings' { const value: { __locale: string }; export default value; }\n");
  }
  if (scenario.kind === 'registry') {
    source += '\n' + inventory.resources.map((resource, index) => `import resource${index} from '${resource.id}';`).join('\n');
    source += '\nexport const registry = [' + inventory.resources.map((_, index) => `resource${index}`).join(', ') + '];';
    fs.writeFileSync(path.join(project, 'src/native.d.ts'), inventory.resources.map(resource => `declare module '${resource.id}' { const value: { __locale: string }; export default value; }`).join('\n'));
  }
  fs.writeFileSync(path.join(project, 'src/probe.ts'), source + '\n');
  fs.writeFileSync(path.join(project, 'gulpfile.js'), "'use strict';\nconst build = require('@microsoft/sp-build-web');\nbuild.initialize(require('gulp'));\n");
  return { project, registeredIds: Object.keys(configured.localizedResources).sort() };
}

async function verifyLocalizationPackage({ consumerRoot, outputDir, only, provenance = 'isolated-installed-tarball-consumer' }) {
  const availableScenarios = [...scenarios, ...mutations];
  const requestedScenarios = only ?? availableScenarios.map(item => item.name);
  assert.ok(requestedScenarios.length > 0, 'At least one SPFx localization scenario must be selected');
  const knownNames = new Set(availableScenarios.map(item => item.name));
  const unknownNames = requestedScenarios.filter(name => !knownNames.has(name));
  assert.deepEqual(unknownNames, [], `Unknown SPFx localization scenarios: ${unknownNames.join(', ')}`);
  consumerRoot = fs.realpathSync(consumerRoot);
  outputDir = path.resolve(outputDir ?? path.join(root, '.docs/maintenance/evidence/localization/spfx-package'));
  fs.mkdirSync(outputDir, { recursive: true });
  const fromConsumer = createRequire(path.join(consumerRoot, 'package.json'));
  const installed = path.dirname(fromConsumer.resolve(packageName + '/package.json'));
  assert.equal(fs.lstatSync(installed).isSymbolicLink(), false, 'Localization gate requires a real packed installation');
  assert.notEqual(fs.realpathSync(installed), path.join(root, 'packages/spfx-react-toolkit'), 'Workspace library cannot serve as packed evidence');
  const inventory = readJson(path.join(installed, 'localization/resources.json'));
  const ids = inventory.resources.map(resource => resource.id);
  assert.equal(ids.length, 8);
  const assets = inventory.resources.flatMap(resource => resource.locales.map(locale => resource.path.replace('{locale}', locale.toLowerCase())));
  assert.equal(assets.length, 400);
  assert.equal(filesBelow(path.join(installed, 'lib/controls')).filter(file => /[/\\]loc[/\\][^/\\]+\.js$/.test(file)).length, 400, 'Packed installation must contain exactly 400 AMD assets');
  for (const asset of assets) readAmd(path.join(installed, asset));
  const directory = path.join(consumerRoot, 'localization-spfx-fixtures');
  assert.ok(!fs.existsSync(directory), `Localization fixture directory already exists: ${directory}`);
  fs.mkdirSync(directory);
  const report = { schemaVersion: 1, provenance, consumerRoot, installedPackage: installed, inventoryIds: ids, packedAmdAssets: assets.length, hostNetworkObserved: false, requestedScenarios, scenarios: [], mutations: [], complete: false, passed: false };
  writeJson(path.join(outputDir, 'report.json'), report);
  try {
    for (const scenario of availableScenarios.filter(item => requestedScenarios.includes(item.name))) {
      const start = Date.now();
      const { project, registeredIds } = createProject({ consumerRoot, directory, scenario, inventory, installed });
      console.log(`SPFx localization building ${scenario.name}${scenario.locale ? ` --locale ${scenario.locale}` : ''}`);
      const preload = path.join(root, 'scripts/gulp-build-metadata-preload.cjs');
      const args = ['--require', preload, fromConsumer.resolve('gulp/bin/gulp.js'), 'bundle', '--ship', ...(scenario.locale ? ['--locale', scenario.locale] : [])];
      const env = { ...process.env, NODE_OPTIONS: [process.env.NODE_OPTIONS, '--require', JSON.stringify(preload)].filter(Boolean).join(' ') };
      const result = spawnSync(process.execPath, args, { cwd: project, env, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
      fs.writeFileSync(path.join(outputDir, scenario.name + '.stdout.log'), result.stdout ?? '');
      fs.writeFileSync(path.join(outputDir, scenario.name + '.stderr.log'), result.stderr ?? '');
      const artifactDirectory = path.join(outputDir, 'artifacts', scenario.name);
      fs.mkdirSync(artifactDirectory, { recursive: true });
      fs.copyFileSync(path.join(project, 'config/config.json'), path.join(artifactDirectory, 'config.json'));
      if (fs.existsSync(path.join(project, 'dist'))) fs.cpSync(path.join(project, 'dist'), path.join(artifactDirectory, 'dist'), { recursive: true });
      let observed;
      let failure;
      try {
        strictResult(result, scenario.name);
        observed = inspectArtifacts(project, scenario, ids);
      } catch (error) { failure = error.message; }
      const output = stripAnsi((result.stdout ?? '') + '\n' + (result.stderr ?? ''));
      const failureMarkers = output.match(/The build failed[^\n]*|Exiting with exit code:\s*[1-9]\d*[^\n]*|Error - [^\n]*/g) ?? [];
      const entry = { name: scenario.name, command: process.execPath, args, status: result.status, signal: result.signal, failureMarkers, durationMs: Date.now() - start, registeredIds, artifactDirectory: path.relative(outputDir, artifactDirectory), ...observed, failure };
      if (scenario.kind) {
        report.mutations.push(entry);
        assert.ok(failure, `${scenario.name} mutation unexpectedly passed`);
        assert.match(failure, scenario.expected, `${scenario.name} failed for an unintended reason`);
        entry.rejectedForIntendedReason = true;
      } else {
        report.scenarios.push(entry);
        assert.equal(failure, undefined, `${scenario.name} failed; evidence: ${outputDir}\n${failure}`);
      }
      writeJson(path.join(outputDir, 'report.json'), report);
      console.log(`SPFx localization ${scenario.name}: ${scenario.kind ? 'rejected for intended reason' : 'passed'} (${entry.durationMs}ms)`);
    }
    report.complete = true;
    report.passed = true;
    return report;
  } finally {
    writeJson(path.join(outputDir, 'report.json'), report);
    if (process.env.SPFX_KEEP_LOCALIZATION_FIXTURES !== '1') fs.rmSync(directory, { recursive: true, force: true });
  }
}

module.exports = { verifyLocalizationPackage, strictResult, inspectArtifacts };
if (require.main === module) {
  const consumerRoot = process.argv[2];
  assert.ok(consumerRoot, 'Usage: node scripts/verify-localization-package.cjs <installed-consumer> [evidence-directory] [scenario,...]');
  verifyLocalizationPackage({ consumerRoot, outputDir: process.argv[3], only: process.argv[4]?.split(','), provenance: process.env.SPFX_LOCALIZATION_PROVENANCE }).catch(error => { console.error(error); process.exitCode = 1; });
}

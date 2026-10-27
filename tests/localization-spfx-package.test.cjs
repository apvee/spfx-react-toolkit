'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { scenarios, mutations } = require('./fixtures/localization-consumers/spfx-scenarios.cjs');
const { strictResult, inspectArtifacts, verifyLocalizationPackage } = require('../scripts/verify-localization-package.cjs');
const inventory = require('../packages/spfx-react-toolkit/localization/resources.json');
const paginationId = 'ApveeSpfxReactToolkitPaginationStrings';
const ids = inventory.resources.map(resource => resource.id);

test('packed SPFx matrix covers import aliases, independent families, lazy editor, custom mappings and explicit locales', () => {
  assert.deepEqual(scenarios.map(scenario => scenario.name), [
    'helper-only', 'pagination-root', 'pagination-domain', 'pagination-control',
    'two-families', 'lazy-editor', 'custom-mapping', 'pagination-it-it', 'pagination-ar-sa',
  ]);
  assert.deepEqual(scenarios.filter(scenario => scenario.name.startsWith('pagination-')).map(scenario => scenario.resources), Array(5).fill([paginationId]));
  assert.deepEqual(mutations.map(mutation => mutation.kind), ['asset', 'registration', 'unknown', 'locale', 'registry']);
});

test('strict Gulp result rejects nested failures even when the wrapper exits zero', () => {
  for (const marker of ['The build failed.', 'Exiting with exit code: 1', 'Error - [webpack] missing native asset']) {
    assert.throws(() => strictResult({ status: 0, stdout: marker, stderr: '' }, 'probe'), /inner build failure/);
  }
  assert.throws(() => strictResult({ status: 1, stdout: '', stderr: 'failure' }, 'probe'), /exited 1/);
  assert.deepEqual(strictResult({ status: 0, stdout: 'Finished subtask webpack', stderr: '' }, 'probe').failureMarkers, []);
});

for (const [description, only, expected] of [
  ['unknown scenario names', ['pagination-typo'], /Unknown SPFx localization scenarios: pagination-typo/],
  ['an empty scenario selection', [], /At least one SPFx localization scenario must be selected/],
]) {
  test(`packed runner rejects ${description} before creating fixture or evidence files`, async t => {
    const consumerRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'localization-invalid-selection-'));
    t.after(() => fs.rmSync(consumerRoot, { recursive: true, force: true }));
    await assert.rejects(verifyLocalizationPackage({ consumerRoot, outputDir: path.join(consumerRoot, 'evidence'), only }), expected);
    assert.deepEqual(fs.readdirSync(consumerRoot), []);
  });
}

function artifactFixture(t) {
  const project = fs.mkdtempSync(path.join(os.tmpdir(), 'localization-artifact-inspector-'));
  t.after(() => fs.rmSync(project, { recursive: true, force: true }));
  const deploy = path.join(project, 'dist');
  fs.mkdirSync(deploy, { recursive: true });
  const manifestFile = path.join(deploy, 'probe.manifest.json');
  const manifest = { loaderConfig: { scriptResources: { [paginationId]: { type: 'localizedPath', defaultPath: 'pagination_it-it.js', paths: { 'en-US': 'pagination_en-us.js', 'it-IT': 'pagination_it-it.js' } } } } };
  fs.writeFileSync(manifestFile, JSON.stringify(manifest));
  fs.writeFileSync(path.join(deploy, 'pagination_it-it.js'), 'define([], function () { return { __locale: "it-IT" }; });');
  fs.writeFileSync(path.join(deploy, 'pagination_en-us.js'), 'define([], function () { return { __locale: "en-US" }; });');
  fs.writeFileSync(path.join(deploy, 'probe.js'), `define(["${paginationId}"], function (strings) { return strings; });`);
  return { project, deploy, manifest, manifestFile };
}

test('artifact inspector reads manifest resources, actual AMD metadata and emitted external references', t => {
  const { project } = artifactFixture(t);
  const result = inspectArtifacts(project, { name: 'inspector-test', resources: [paginationId] }, ids);
  assert.deepEqual(result.manifestResources, [paginationId]);
  assert.equal(result.resourceFiles.find(asset => asset.locale === 'it-IT').nativeLocale, 'it-IT');
  assert.deepEqual(result.scripts[0].nativeIds, [paginationId]);
});

test('artifact inspector accepts the single native path emitted by an explicit SPFx locale', t => {
  const { project, manifest, manifestFile } = artifactFixture(t);
  manifest.loaderConfig.scriptResources[paginationId] = { type: 'path', path: 'pagination_it-it.js' };
  fs.writeFileSync(manifestFile, JSON.stringify(manifest));
  const result = inspectArtifacts(project, { name: 'selected-locale', resources: [paginationId], locale: 'it-it' }, ids);
  assert.equal(result.resourceFiles.length, 1);
  assert.equal(result.resourceFiles[0].nativeLocale, 'it-IT');
});

test('artifact inspector rejects missing assets, selected locale fallback and retained unrelated resources', t => {
  const { project, deploy, manifest, manifestFile } = artifactFixture(t);
  manifest.loaderConfig.scriptResources[paginationId] = { type: 'path', path: 'pagination_it-it.js' };
  fs.writeFileSync(manifestFile, JSON.stringify(manifest));
  assert.throws(() => inspectArtifacts(project, { name: 'missing-selected', resources: [paginationId], locale: 'ar-sa' }, ids), /Selected locale ar-sa did not emit/);
  manifest.loaderConfig.scriptResources[paginationId] = { type: 'localizedPath', defaultPath: 'pagination_it-it.js', paths: { 'en-US': 'pagination_en-us.js', 'it-IT': 'pagination_it-it.js' } };
  manifest.loaderConfig.scriptResources.ApveeSpfxReactToolkitStageStrings = manifest.loaderConfig.scriptResources[paginationId];
  fs.writeFileSync(manifestFile, JSON.stringify(manifest));
  assert.throws(() => inspectArtifacts(project, { name: 'root-registry', resources: [paginationId] }, ids), /Unexpected manifest resources/);
  delete manifest.loaderConfig.scriptResources.ApveeSpfxReactToolkitStageStrings;
  fs.writeFileSync(manifestFile, JSON.stringify(manifest));
  fs.rmSync(path.join(deploy, 'pagination_it-it.js'));
  assert.throws(() => inspectArtifacts(project, { name: 'missing-asset', resources: [paginationId] }, ids), /Missing emitted default resource/);
});

// Normal root tests check the strict evidence machinery above. The genuine Gulp
// matrix shares the one installed tarball consumer in verify:package; it never
// installs a fresh SPFx peer tree per test. This opt-in also supports focused reruns.
test('genuine packed SPFx Gulp ship matrix', { skip: !process.env.SPFX_LOCALIZATION_CONSUMER, timeout: 1800000 }, async () => {
  const report = await verifyLocalizationPackage({ consumerRoot: process.env.SPFX_LOCALIZATION_CONSUMER, outputDir: process.env.SPFX_LOCALIZATION_EVIDENCE });
  assert.equal(report.passed, true);
  assert.equal(report.scenarios.length, 9);
  assert.equal(report.mutations.length, 5);
});

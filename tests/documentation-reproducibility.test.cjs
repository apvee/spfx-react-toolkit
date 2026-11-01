const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const test = require('node:test');
const root = path.resolve(__dirname, '..');
const proof = () => require('../scripts/verify-documentation-reproducibility.cjs');

test('hashes full ordinal paths and exact bytes with unambiguous framing', () => {
  const { hashArtifacts } = proof();
  const result = hashArtifacts({ 'b.md': 'second\n', 'a.md': 'first\n' });
  assert.deepEqual(result.paths, ['a.md', 'b.md']);
  assert.deepEqual(result, hashArtifacts({ 'a.md': Buffer.from('first\n'), 'b.md': 'second\n' }));
  assert.notEqual(result.digest, hashArtifacts({ 'a.md': 'first\n', 'b.md': 'Second\n' }).digest);
  assert.notEqual(hashArtifacts({ ab: 'c' }).digest, hashArtifacts({ a: 'bc' }).digest);
});

test('lock admission rejects loose and nested unlocked packages, workspace links and escaping files', () => {
  const { createDependencyAdmission } = require('../scripts/documentation-inputs.cjs');
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'docs-lock-'));
  try {
    const put = (route, text) => { fs.mkdirSync(path.dirname(path.join(temp, route)), { recursive: true }); fs.writeFileSync(path.join(temp, route), text); };
    const records = { 'node_modules/locked': { version: '1.0.0' }, 'node_modules/locked/node_modules/@scope/nested': { version: '2.0.0' }, 'node_modules/workspace': { link: true, resolved: 'src' } };
    put('package-lock.json', JSON.stringify({ lockfileVersion: 3, packages: records }));
    for (const [route, record] of Object.entries(records)) if (!record.link) put(route + '/package.json', JSON.stringify({ name: route.split('/node_modules/').at(-1).replace('node_modules/', ''), version: record.version }));
    put('node_modules/locked/index.d.ts', 'export {};');
    put('node_modules/loose/index.d.ts', 'export {};');
    put('node_modules/loose-file.d.ts', 'export {};');
    put('node_modules/locked/node_modules/loose/index.d.ts', 'export {};');
    put('src/index.ts', 'export {};');
    fs.symlinkSync(path.join(temp, 'src'), path.join(temp, 'node_modules/workspace'));
    fs.symlinkSync(path.join(temp, 'src/index.ts'), path.join(temp, 'node_modules/locked/escape.d.ts'));
    const admission = createDependencyAdmission(temp);
    assert.ok(admission.admitted(path.join(temp, 'node_modules/locked/index.d.ts')));
    assert.ok(admission.admitted(path.join(temp, 'node_modules/locked/node_modules/@scope/nested/package.json')));
    for (const route of ['node_modules/loose/index.d.ts', 'node_modules/loose-file.d.ts', 'node_modules/locked/node_modules/loose/index.d.ts', 'node_modules/workspace/index.ts', 'node_modules/locked/escape.d.ts']) assert.equal(admission.admitted(path.join(temp, route)), false, route);
  } finally { fs.rmSync(temp, { recursive: true, force: true }); }
});

test('offline guard rejects every deliberate network/process route in fresh children', () => {
  const directory = path.join(__dirname, 'fixtures/documentation/reproducibility');
  for (const route of ['fetch', 'http', 'https', 'http2', 'net', 'net-prototype', 'tls', 'dgram', 'dns', 'dns-promises', 'child', 'worker']) {
    const child = spawnSync(process.execPath, ['--require', path.join(directory, 'offline-guard.cjs'), path.join(directory, 'offline-probe.cjs'), route], { encoding: 'utf8', env: { ...process.env, NODE_OPTIONS: '', NODE_PATH: '' } });
    assert.notEqual(child.status, 0, route);
    assert.match(child.stderr, /DOC_OFFLINE_FORBIDDEN/, route + ': ' + child.stderr);
  }
});

test('copied tiny engine reconstructs development only from sources and restores its deleted directory', { timeout: 120000 }, t => {
  const fixture = fs.mkdtempSync(path.join(os.tmpdir(), 'docs-repro-input-'));
  t.after(() => fs.rmSync(fixture, { recursive: true, force: true }));
  fs.cpSync(path.join(__dirname, 'fixtures/documentation/artifact-set'), fixture, { recursive: true });
  fs.mkdirSync(path.join(fixture, 'docs/development'), { recursive: true });
  fs.writeFileSync(path.join(fixture, 'docs/development/stale.md'), '# Poisoned old output\n\nNever copy this file.\n');
  fs.writeFileSync(path.join(fixture, 'docs/development/INDEX.md'), '# Poisoned old index\n\nNever copy this target through link closure.\n');
  // Exercise target-closure traversal as well as the initial docs-tree copy.
  fs.appendFileSync(path.join(fixture, 'docs/_sources/development/INDEX.md'), '\n[Old output](../../development/INDEX.md). [Contributor input](../../../contributor.md).\n');
  fs.writeFileSync(path.join(fixture, 'contributor.md'), '# Contributor input\n\nA source-owned target outside the standard copied roots.\n');
  const result = proof().verifyReproducibility(root, { fixtureRoot: fixture, packageRoot: 'packages/spfx-react-toolkit', sourceRoot: 'packages/spfx-react-toolkit/src', tsconfig: 'packages/spfx-react-toolkit/tsconfig.json' });
  assert.equal(result.ok, true);
  assert.equal(result.environments.length, 4);
  assert.ok(result.hashes.paths.includes('llms.txt'));
  assert.ok(result.audit.reads > 0);
  assert.equal(result.sourceMutationChanged, true);
  assert.ok(!result.copiedInputs.some(route => route.startsWith('docs/development/')), 'cold copies must exclude all public development files');
  assert.equal(result.developmentSourceMutationChanged, true);
  assert.equal(result.developmentRebuild.directoryAbsentBefore, true);
  assert.equal(result.developmentRebuild.sourcesRetained, true);
  assert.ok(result.copiedInputs.includes('scripts/generated-artifacts.cjs'));
  assert.ok(result.copiedInputs.includes('scripts/documentation-context.cjs'));
  assert.ok(result.copiedInputs.includes('docs/_sources/development/INDEX.md'));
  assert.ok(result.copiedInputs.includes('docs/_sources/development/sharepoint-validation/INDEX.md'));
  assert.ok(result.copiedInputs.includes('contributor.md'));
  assert.deepEqual(result.developmentRebuild.hashes.paths, result.hashes.paths.filter(route => route.startsWith('docs/development/')));
  for (const [route, hash] of Object.entries(result.developmentRebuild.hashes.perFile)) assert.equal(hash, result.hashes.perFile[route]);
  assert.equal(result.offlineProbes, 12);
  assert.ok(result.copiedInputs.includes('LICENSE'));
  assert.ok(result.copiedInputs.includes('packages/spfx-react-toolkit/LICENSE'));
  const licenseInputs = Object.fromEntries(['LICENSE', 'packages/spfx-react-toolkit/LICENSE'].map(route => [route, fs.readFileSync(path.join(__dirname, 'fixtures/documentation/artifact-set', route))]));
  assert.notDeepEqual(licenseInputs.LICENSE, licenseInputs['packages/spfx-react-toolkit/LICENSE']);
  for (const [route, hash] of Object.entries(proof().hashArtifacts(licenseInputs).perFile)) assert.equal(result.copiedInputHashes.perFile[route], hash);
});

test('stale development files and anchors cannot rescue a removed source target', t => {
  const fixture = fs.mkdtempSync(path.join(os.tmpdir(), 'docs-repro-poison-'));
  t.after(() => fs.rmSync(fixture, { recursive: true, force: true }));
  fs.cpSync(path.join(__dirname, 'fixtures/documentation/artifact-set'), fixture, { recursive: true });
  const config = { packageRoot: 'packages/spfx-react-toolkit', sourceRoot: 'packages/spfx-react-toolkit/src', tsconfig: 'packages/spfx-react-toolkit/tsconfig.json' };
  const { generateArtifacts, applyArtifacts } = require('../scripts/documentation-artifacts.cjs');
  const artifacts = generateArtifacts(fixture, config);
  applyArtifacts(fixture, artifacts, 'write', config);
  const source = path.join(fixture, 'docs/_sources/development/sharepoint-validation/INDEX.md');
  const index = path.join(fixture, 'docs/_sources/development/INDEX.md');
  const original = fs.readFileSync(index, 'utf8');
  fs.appendFileSync(index, '\n[Host section](sharepoint-validation/INDEX.md#host-checks)\n');
  fs.writeFileSync(source, '# Replacement host\n\nA different natural heading.\n');
  assert.throws(() => generateArtifacts(fixture, config), error => error.diagnostics?.some(item => item.code === 'DOC_LINK_ANCHOR' && item.file === 'docs/development/INDEX.md'));
  fs.writeFileSync(index, original);
  fs.unlinkSync(source);
  assert.throws(() => generateArtifacts(fixture, config), error => error.diagnostics?.some(item => item.code === 'DOC_LINK_TARGET' && item.file === 'docs/development/INDEX.md'));
  for (const [route, text] of Object.entries(artifacts)) assert.equal(fs.readFileSync(path.join(fixture, route), 'utf8'), text, route);
});

test('graph and examples share clean toolkit mapping and reject unknown private/native-from-root imports', () => {
  const fixture = path.join(__dirname, 'fixtures/documentation/reproducibility/self-import');
  const config = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' };
  const { buildApiGraph } = require('../scripts/documentation-model.cjs');
  const { buildDocumentationModel } = require('../scripts/documentation-contract.cjs');
  const { checkDocumentationExamples } = require('../scripts/documentation-examples.cjs');
  const graph = buildApiGraph(fixture, config);
  assert.deepEqual(graph.diagnostics, []);
  const control = graph.entities.find(entity => entity.qualifiedName === 'control');
  const options = graph.entities.find(entity => entity.qualifiedName === 'NativeOptions');
  assert.ok(control.signatures[0].parameters[0].type.references.includes(options.key));
  assert.equal(options.source.file, 'src/native.ts');
  assert.equal(options.visibility, 'public');
  const model = buildDocumentationModel(graph, config);
  assert.deepEqual(checkDocumentationExamples(fixture, model, config), []);
  const negative = { ...model, examples: [
    { ownerKey: control.key, language: 'ts', source: { file: 'negative.ts', line: 1 }, code: "import type { NativeOptions } from '@fixture/cold';\nimport { secret } from '@fixture/cold/private';" },
  ] };
  const diagnostics = checkDocumentationExamples(fixture, negative, config);
  assert.ok(diagnostics.some(item => item.code === 'TS2459' && item.line === 1));
  assert.ok(diagnostics.some(item => item.code === 'TS2307' && item.line === 2));
});

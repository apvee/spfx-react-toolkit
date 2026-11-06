const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const original = process.argv[2];
const root = path.resolve(__dirname, '../../../..');
const { createDocumentationInputs } = require(path.join(root, 'scripts/docs/documentation-inputs.cjs'));
const { buildApiGraph } = require(path.join(root, 'scripts/docs/documentation-model.cjs'));
const { checkDocumentationExamples } = require(path.join(root, 'scripts/docs/documentation-examples.cjs'));
const { hashArtifacts } = require(path.join(root, 'scripts/docs/verify-documentation-reproducibility.cjs'));
const config = require(path.join(root, 'scripts/docs/documentation.config.cjs'));
const snapshot = () => hashArtifacts(Object.fromEntries(JSON.parse(process.argv[3]).map(route => [route, fs.readFileSync(path.join(root, route))])));
const before = snapshot();
const inputs = createDocumentationInputs(root, config);
const files = ['package.json', 'scripts/docs/documentation.config.cjs', config.sourceRoot + '/index.ts'];
for (const route of files) {
  assert.throws(() => inputs.readFile(path.join(original, route)), error => error.code === 'DOC_FORBIDDEN_INPUT');
  assert.throws(() => fs.readFileSync(path.join(original, route)), error => error.code === 'DOC_FORBIDDEN_INPUT');
}
const alias = path.join(root, config.sourceRoot, 'original-source-alias.ts');
fs.symlinkSync(path.join(original, config.sourceRoot, 'index.ts'), alias);
try { assert.throws(() => inputs.readFile(alias), error => error.code === 'DOC_FORBIDDEN_INPUT'); }
finally { fs.unlinkSync(alias); }
assert.ok(inputs.readFile(require.resolve('typescript/package.json')));
// Both copied canonical consumers must reject a direct escaped config before
// constructing any graph/program or changing an output.
const escapedConfig = { ...config, tsconfig: path.join(original, config.tsconfig) };
assert.throws(() => buildApiGraph(root, escapedConfig), error => error.code === 'DOC_FORBIDDEN_INPUT');
assert.throws(() => checkDocumentationExamples(root, { examples: [] }, escapedConfig), error => error.code === 'DOC_FORBIDDEN_INPUT');
// Exercise actual resolution in both consumers using a tiny versioned source
// fixture. The symlink points at original checkout source and stays unimportable.
const fixture = path.join(__dirname, 'self-import');
const fixtureConfig = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' };
const fixtureSource = path.join(fixture, 'src/index.ts');
const saved = fs.readFileSync(fixtureSource);
const importedAlias = path.join(fixture, 'escaped.ts');
fs.symlinkSync(path.join(original, config.sourceRoot, 'index.ts'), importedAlias);
try {
  fs.appendFileSync(fixtureSource, '\nimport * as escaped from "../escaped";\nexport const forbidden = escaped;\n');
  const graph = buildApiGraph(fixture, fixtureConfig);
  assert.ok(graph.diagnostics.some(diagnostic => diagnostic.code === 'TS2307' && /escaped/.test(diagnostic.message)), JSON.stringify(graph.diagnostics));
  const diagnostics = checkDocumentationExamples(fixture, { graph, scenarios: [], examples: [{ ownerKey: 'escape', language: 'ts', source: { file: 'escape.ts', line: 1 }, code: `import * as escaped from ${JSON.stringify(path.join(original, config.sourceRoot, 'index.ts'))};` }] }, fixtureConfig);
  assert.ok(diagnostics.some(diagnostic => diagnostic.code === 'TS2307' && diagnostic.file === 'escape.ts'), JSON.stringify(diagnostics));
  assert.ok(diagnostics.some(diagnostic => diagnostic.code === 'TS2307' && diagnostic.file === 'src/index.ts' && /escaped/.test(diagnostic.message)), JSON.stringify(diagnostics));
} finally { fs.writeFileSync(fixtureSource, saved); fs.unlinkSync(importedAlias); }
assert.deepEqual(snapshot(), before, 'Copied consumer negative probes must preserve all generated document bytes.');
process.stdout.write(JSON.stringify({ directCompilerReadsDenied: 3, directFilesystemReadsDenied: 3, sourceSymlinkDenied: true, lockedCompilerReadable: true, copiedGraphConfigDenied: true, copiedExamplesConfigDenied: true, copiedGraphSymlinkTS2307: true, copiedExamplesSourceTS2307: true, outputSnapshotUnchanged: true }));

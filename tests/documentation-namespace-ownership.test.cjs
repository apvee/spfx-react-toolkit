const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const { buildApiGraph } = require('../scripts/documentation-model.cjs');
const { buildDocumentationModel } = require('../scripts/documentation-contract.cjs');
const { collectExecutableExamples } = require('../scripts/executable-examples-model.cjs');
const { renderApiPages } = require('../scripts/documentation-render.cjs');

const fixture = path.join(__dirname, 'fixtures/documentation/namespace-ownership');
const config = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' };
let graph;
let model;
let result;
const binding = name => graph.bindings.find(item => item.specifier === '@fixture/namespace' &&
  [item.exportedName, ...item.memberPath].join('.') === name);
test.before(() => {
  graph = buildApiGraph(fixture, config, { requestedSourceEntries: ['src/private.ts'] });
  assert.deepEqual(graph.diagnostics, []);
  model = buildDocumentationModel(graph, config);
  result = collectExecutableExamples(fixture, model, config);
});

test('namespace runtime routes retain every erased boundary, shadowing and genuine duplicate paths', () => {
  for (const name of ['width', 'renamedWidth', 'aliasWidth', 'nested', 'nested.inner', 'duplicateReal',
    'cycleReal', 'cycleReal.next', 'imports', 'overloads']) {
    assert.equal(binding(name)?.runtimeExport, true, name);
  }
  for (const name of ['TypeWidth', 'hiddenWidth', 'ImportedWidth', 'NamedImportWidth', 'NamedTypeWidth',
    'OtherTypeWidth', 'erased', 'starErased', 'importErased', 'shadowed', 'empty', 'typeOnly',
    'ambient', 'ambientFile', 'cycle', 'cycle.next', 'TypeNested', 'TypeNested.inner', 'bypassedWidth']) {
    assert.equal(binding(name)?.runtimeExport, false, name);
  }
  assert.ok(!graph.bindings.some(item => item.entityKey === 'src/private.ts#<module>'));
  assert.equal(graph.entities.find(item => item.key === 'src/private.ts#<module>').visibility, 'support');
  assert.equal(graph.entryQueries['src/private.ts'][0].specifier, '');
});

test('namespace runtime descendants include implemented classes, enums and uninitialized exports but exclude erased const enums', () => {
  for (const name of ['classOnly', 'enumOnly', 'uninitialized']) assert.equal(binding(name)?.runtimeExport, true, name);
  assert.equal(binding('constEnumOnly')?.runtimeExport, false);
});

test('namespace runtime descendants retain merged functions and destructured implemented values', () => {
  assert.equal(binding('mergedFunction')?.runtimeExport, true);
  assert.equal(binding('destructured')?.runtimeExport, true);
});

test('namespace module ownership uses the unique public runtime entity before parsing any file contract', () => {
  const contract = model.contracts['src/width.ts'];
  assert.equal(contract.examples[0].ownerKey, 'src/width.ts#<module>');
  assert.strictEqual(contract, model.contracts['src/width.ts#<module>']);
  assert.deepEqual(contract.source, { file: 'src/width.ts', line: 1 });
  assert.deepEqual(contract.examples[0].source, { file: 'src/width.ts', line: 7 });
  assert.equal(model.examples.filter(item => item.title === 'Width recipe').length, 1);
  const reordered = buildDocumentationModel({ ...graph, entities: [...graph.entities].reverse() }, config);
  assert.equal(reordered.contracts['src/width.ts'].examples[0].ownerKey, 'src/width.ts#<module>');
  assert.equal(model.contracts['src/width.ts#px'].examples[0].ownerKey, 'src/width.ts#<module>::member:px');
  assert.equal(model.contracts['src/width.ts#auto'].examples[0].code, 'const ordinary = 42;\n');
  for (const file of ['erased', 'empty', 'types', 'ambient', 'cycle-a', 'private', 'declaration', 'control']) {
    assert.equal(model.contracts[`src/${file}.ts`].examples[0].ownerKey, `src/${file}.ts`, file);
  }
  assert.equal(model.graph.entities.find(item => item.key === 'src/declaration.ts#Declared').namespaceRole, 'declaration');
});

test('namespace executable owner and names exclude erased aliases without losing genuine binding recipes', () => {
  assert.deepEqual(result.examples.map(item => item.title).sort(), [
    'Control declaration recipe', 'Declared namespace recipe', 'Nonnamespace binding recipe', 'Runtime binding recipe', 'Width recipe',
  ]);
  const width = result.examples.find(item => item.title === 'Width recipe');
  assert.equal(width.ownerKey, 'src/width.ts#<module>');
  assert.equal(width.featureKey, 'src/width.ts#<module>');
  assert.equal(width.featureTitle, 'width');
  assert.equal(width.groupKey, 'helpers');
  assert.deepEqual(width.publicNames, ['aliasWidth', 'renamedWidth', 'width']);
  const validBinding = result.examples.find(item => item.title === 'Runtime binding recipe');
  assert.equal(validBinding.ownerKey, '@fixture/namespace#aliasWidth');
  assert.equal(validBinding.featureKey, width.featureKey);
  // This snippet is self-contained; type-only import errors cannot hide the owner bug.
  const erased = model.examples.find(item => item.title === 'Erased binding recipe');
  assert.doesNotMatch(erased.code, /TypeWidth/);
  assert.deepEqual(result.diagnostics.filter(item => item.symbol === '@fixture/namespace#TypeWidth'), [{
    code: 'EXECUTABLE_RUNTIME_OWNER', file: 'src/index.ts', line: erased.source.line,
    symbol: '@fixture/namespace#TypeWidth', message: 'Executable examples require a runtime API owner.',
  }]);
  assert.ok(result.diagnostics.every(item => item.code === 'EXECUTABLE_RUNTIME_OWNER'), JSON.stringify(result.diagnostics));
  assert.deepEqual(result.examples.find(item => item.title === 'Nonnamespace binding recipe').publicNames, ['Control', 'TypeControl']);
  assert.equal(result.examples.find(item => item.title === 'Declared namespace recipe').featureKey, 'src/declaration.ts#Declared');
});

test('namespace owner normalization preserves published recipe bodies, introductions and member ownership', () => {
  // Negative ownership fixtures deliberately include unrelated support modules;
  // render the positive owner's real contracts rather than their invalid pages.
  const entities = graph.entities.filter(item => item.source.file === 'src/width.ts');
  const keys = new Set(entities.map(item => item.key));
  const positive = { ...graph, entities, bindings: graph.bindings.filter(item => keys.has(item.entityKey)) };
  const current = buildDocumentationModel(positive, config);
  const oldGraph = structuredClone(positive);
  oldGraph.bindings.forEach(item => { delete item.runtimeExport; });
  const previous = buildDocumentationModel(oldGraph, config);
  assert.equal(previous.contracts['src/width.ts#px'].examples[0].ownerKey, 'src/width.ts#<module>::member:px');
  assert.deepEqual(renderApiPages(current, config), renderApiPages(previous, config));
});

test('namespace ownership fails closed for contradictory physical module owners', () => {
  const malformed = structuredClone(graph);
  const original = malformed.entities.find(item => item.key === 'src/width.ts#<module>');
  malformed.entities.push({ ...original, key: 'src/width.ts#<contradictory-module>' });
  malformed.bindings.push({ ...binding('width'), entityKey: 'src/width.ts#<contradictory-module>' });
  const invalid = buildDocumentationModel(malformed, config);
  assert.equal(invalid.contracts['src/width.ts'].examples[0].ownerKey, 'src/width.ts');
  assert.ok(invalid.diagnostics.some(item => item.code === 'DOC_AMBIGUOUS_MODULE_OWNER' &&
    item.file === 'src/width.ts' && item.line === 1));
});

test('namespace identities and source bytes survive alias order, line changes and relocation', t => {
  const changed = structuredClone(model);
  changed.graph.bindings.reverse();
  changed.examples.forEach(item => { item.source.line += 20; });
  const again = collectExecutableExamples(fixture, changed, config);
  const original = result.examples.find(item => item.title === 'Width recipe');
  const perturbed = again.examples.find(item => item.title === 'Width recipe');
  assert.ok(original && perturbed);
  assert.equal(perturbed.id, original.id);
  assert.deepEqual(perturbed.publicNames, original.publicNames);
  assert.equal(perturbed.code, original.code);
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'namespace-owner-'));
  t.after(() => fs.rmSync(temporary, { recursive: true, force: true }));
  fs.cpSync(fixture, temporary, { recursive: true });
  fs.symlinkSync(path.resolve(__dirname, '../node_modules'), path.join(temporary, 'node_modules'), 'dir');
  const relocated = buildDocumentationModel(buildApiGraph(temporary, config), config);
  const relocatedResult = collectExecutableExamples(temporary, relocated, config);
  const moved = relocatedResult.examples.find(item => item.title === 'Width recipe');
  assert.ok(moved, JSON.stringify(relocatedResult.diagnostics));
  assert.equal(moved.id, original.id);
  assert.equal(moved.code, original.code);
  assert.deepEqual(moved.source, { file: 'src/width.ts', line: 7 });
});

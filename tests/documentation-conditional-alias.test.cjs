const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const { buildApiGraph, resolvePublicReference } = require('../scripts/docs/documentation-model.cjs');
const { buildDocumentationModel, validateDocumentation } = require('../scripts/docs/documentation-contract.cjs');
const { checkDocumentationExamples } = require('../scripts/docs/documentation-examples.cjs');
const config = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' };
const root = path.join(__dirname, 'fixtures/documentation/conditional-alias');
let graph, model;
test.before(() => {
  graph = buildApiGraph(root, config);
  assert.deepEqual(graph.diagnostics, []);
  model = buildDocumentationModel(graph, config);
});
const guarded = current => current.entities.find(entity => entity.key === 'src/index.ts#GuardedObject');

test('conditional branches preserve the complete named alias and its source documentation', () => {
  const alias = guarded(graph);
  assert.match(alias.rawDoc, /Guards object attributes while preserving nonobject shorthand/);
  assert.deepEqual(alias.source, { file: 'src/index.ts', line: 10 });
  assert.deepEqual(alias.docSource, { file: 'src/index.ts', line: 1 });
  assert.deepEqual(alias.typeParameters, [{ name: 'Props' }, { name: 'Owned', constraint: 'PropertyKey' }]);
  assert.match(alias.typeExpression, /^Props extends object\s+\? Omit<Props, Owned> &/);
  assert.match(alias.typeExpression, /\{ \[Key in Owned\]\?: never \}\s+: Props$/);
  assert.equal(alias.visibility, 'support');
  assert.ok(!graph.bindings.some(binding => binding.entityKey === alias.key));
  assert.equal(graph.entities.filter(entity => entity.key === alias.key).length, 1);
  const contract = model.contracts[alias.key];
  assert.equal(contract.summary, 'Guards object attributes while preserving nonobject shorthand.');
  assert.deepEqual(contract.typeParams.map(parameter => parameter.name), ['Props', 'Owned']);
  assert.deepEqual(contract.source, { file: 'src/index.ts', line: 1 });
  assert.equal(contract.examples.length, 1);
  assert.deepEqual(contract.examples[0].source, { file: 'src/index.ts', line: 6 });
  assert.ok(model.examples.some(example => example.ownerKey === alias.key));
  assert.deepEqual(validateDocumentation(model).filter(diagnostic => diagnostic.symbol === alias.key), []);
  assert.deepEqual(validateDocumentation(model), []);
  assert.deepEqual(checkDocumentationExamples(root, model, config), []);
});

test('resolved conditional branches keep forbidden attributes and nested callback ownership', () => {
  const object = resolvePublicReference(graph, 'ObjectSurface');
  assert.deepEqual(object.members.map(member => member.name), ['call', 'callback', 'options', 'rawHtml', 'text']);
  for (const name of ['call', 'rawHtml']) {
    const member = object.members.find(member => member.name === name);
    assert.equal(member.optional, true);
    assert.equal(member.forbidden, true);
    assert.equal(member.declaredType, 'never');
    assert.equal(member.ownerKey, 'src/index.ts#GuardedObject');
  }
  assert.equal(object.members.find(member => member.name === 'text').optional, false);
  assert.equal(resolvePublicReference(graph, 'ShorthandSurface').type.text, 'string');
  const nested = graph.entities.find(entity => entity.qualifiedName === 'GuardedObject.options');
  assert.equal(nested.rawDoc, '/** Nested options must be omitted or undefined. */');
  assert.deepEqual(nested.members.map(member => member.name), ['label']);
  assert.deepEqual(nested.docSource, { file: 'src/index.ts', line: 16 });
  const callback = graph.entities.find(entity => entity.qualifiedName === 'GuardedObject.callback');
  assert.match(callback.rawDoc, /Callback must be omitted or undefined/);
  assert.deepEqual(callback.signatures[0].typeParameters, [{ name: 'Item', constraint: 'Props' }]);
  assert.deepEqual(callback.signatures[0].lexicalTypeParameters, [{ name: 'Props' }, { name: 'Owned', constraint: 'PropertyKey' }]);
  assert.equal(callback.signatures[0].returns.text, 'Item');
  assert.equal(model.contracts[callback.key].summary, 'Callback must be omitted or undefined.');
  assert.deepEqual(callback.docSource, { file: 'src/index.ts', line: 21 });
});

test('branch-first registration matches alias-first registration after relocation', () => {
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-conditional-alias-'));
  try {
    fs.cpSync(root, temporary, { recursive: true });
    fs.appendFileSync(path.join(temporary, 'src/index.ts'), '\nexport type { GuardedObject as ACanonical };\n');
    const relocated = buildApiGraph(temporary, config);
    assert.deepEqual(relocated.diagnostics, []);
    const { visibility: originalVisibility, ...original } = guarded(graph);
    const { visibility: relocatedVisibility, ...canonical } = guarded(relocated);
    assert.equal(originalVisibility, 'support');
    assert.equal(relocatedVisibility, 'public');
    assert.deepEqual(original, canonical);
  } finally {
    fs.rmSync(temporary, { recursive: true, force: true });
  }
});

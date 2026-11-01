const assert = require('node:assert/strict');
const path = require('node:path');
const test = require('node:test');
const { buildApiGraph, resolvePublicReference } = require('../scripts/documentation-model.cjs');
const { buildDocumentationModel, validateDocumentation, resolveDocumentationReference } = require('../scripts/documentation-contract.cjs');
const { renderApiPages, renderCatalog } = require('../scripts/documentation-render.cjs');
const { markdownAnchors } = require('../scripts/documentation.cjs');
const config = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' };
const root = path.join(__dirname, 'fixtures/documentation/pilot-metadata');
let graph, model;
test.before(() => {
  graph = buildApiGraph(root, config);
  assert.deepEqual(graph.diagnostics, []);
  model = buildDocumentationModel(graph, config);
});
const entity = reference => resolvePublicReference(graph, reference);
const contract = reference => model.contracts[entity(reference).key];

test('source-file namespaces consume their module contract with exact provenance and keep local binding overrides', () => {
  const namespace = entity('measures');
  assert.equal(namespace.namespaceRole, 'sourceFile');
  assert.equal(entity('measures.Nested').namespaceRole, 'declaration');
  assert.equal(entity('measures.module').namespaceRole, 'declaration');
  assert.equal(contract('measures.module').summary, 'Declared module-named namespace purpose.');
  assert.equal(contract('measures').summary, 'Source measure namespace purpose.');
  assert.deepEqual(contract('measures').source, { file: 'src/measures.ts', line: 1 });
  assert.equal(resolveDocumentationReference(model, 'alternate').contract.summary, 'A local binding override.');
  assert.equal(contract('measures.Nested').summary, 'Declared nested namespace owns its purpose.');
  assert.equal(contract('measures.unit').remarks.length, 0);
  const page = renderApiPages(model)['docs/api/shared.md'];
  assert.equal(page.split('Source measure namespace purpose.').length - 1, 1);
  assert.equal(page.split('Namespace requirements belong here only.').length - 1, 1);
  assert.ok(!validateDocumentation(model).some(d => d.symbol === namespace.key));
});

test('runtime and type headings share one page with visible natural type qualification, exact references and retained aliases', () => {
  assert.ok(!model.diagnostics.some(d => d.code === 'DOC_ANCHOR_COLLISION'));
  assert.equal(resolveDocumentationReference(model, 'compose').anchor, 'compose');
  assert.equal(resolveDocumentationReference(model, 'Compose').anchor, 'compose-type');
  assert.equal(resolveDocumentationReference(model, 'assemble').anchor, 'assemble');
  const page = renderApiPages(model)['docs/api/shared.md'];
  assert.match(page, /\n## Compose \(type\)\n/);
  assert.match(page, /\n## compose\n/);
  assert.match(page, /\n## assemble\n/);
  assert.ok(markdownAnchors(page).has('compose-type'));
  assert.match(page, /See \[Compose\]\(shared.md#compose-type\) and \[assemble\]\(shared.md#assemble\)/);
  assert.match(page, /import type \{ Compose \} from '@fixture\/pilot';/);
  assert.match(page, /import \{ assemble \} from '@fixture\/pilot';/);
  assert.match(renderCatalog(model), /\[`Compose`\]\(shared.md#compose-type\)/);
});

test('anonymous union branches retain named alias identity, full shape, docs and parameter-only ownership', () => {
  const choice = graph.entities.find(e => e.qualifiedName === 'Choice');
  assert.equal(choice.kind, 'type');
  assert.match(choice.rawDoc, /Preserves both alternatives/);
  assert.match(choice.typeExpression, /kind: 'many'/);
  assert.equal(choice.variants.length, 2);
  assert.deepEqual(choice.typeParameters.map(p => p.name), ['T']);
  assert.equal(graph.entities.filter(e => e.key === choice.key).length, 1);
  const options = graph.entities.find(e => e.qualifiedName === 'Compose.options');
  assert.ok(options);
  const own = model.contracts[options.key];
  assert.equal(own.summary, 'Exactly one required result configuration.');
  assert.equal(own.returns, undefined);
  assert.equal(own.typeParams.length, 0);
  assert.equal(own.examples.length, 0);
  assert.deepEqual(own.source, model.contracts[entity('Compose').signatures[2].key].source);
  const explicit = graph.entities.find(e => e.qualifiedName === 'configure.options');
  assert.equal(model.contracts[explicit.key].summary, 'Explicit parameter type purpose.');
  assert.equal(model.contracts[explicit.key].params.length, 0);
});

test('local typed const overloads keep local signature contracts across files while external aliases retain toolkit contracts', () => {
  for (const name of ['compose', 'Compose']) {
    const signatures = entity(name).signatures;
    assert.equal(signatures.length, 3);
    const contracts = signatures.map(s => model.contracts[s.key]);
    assert.deepEqual(contracts.map(c => c.params.map(p => p.name)), [[], ['value'], ['options']]);
    assert.deepEqual(contracts.map(c => c.typeParams.map(p => p.name)), [[], ['TValue'], ['TItem']]);
    assert.deepEqual(contracts.map(c => c.returns), ['Empty result.', 'Original value.', 'Configured result.']);
    assert.ok(contracts.every(c => c.source.file === 'src/contracts.ts'));
  }
  const alias = entity('externalAlias');
  const call = model.contracts[alias.signatures[0].key];
  assert.equal(call.summary, 'Toolkit alias contract with local timing requirements.');
  assert.equal(call.params[0].description, 'Required local value.');
  assert.equal(call.source.file, 'src/index.ts');
  const diagnostics = validateDocumentation(model);
  assert.ok(!diagnostics.some(d => /MISSING_(PARAM|TYPE_PARAM|RETURNS)|UNKNOWN_(PARAM|TYPE_PARAM)/.test(d.code)), JSON.stringify(diagnostics));
  const page = renderApiPages(model)['docs/api/shared.md'];
  assert.doesNotMatch(page, /Upstream callable|upstream provider/);
  assert.ok(page.includes('Required original value.') && page.includes('Configured result.'));
});

test('outer and inner generics validate at their authoring scope without accepting undeclared names', () => {
  for (const name of ['Selector', 'SeparateSelector']) {
    const value = entity(name);
    assert.ok(!validateDocumentation(model).some(d => d.symbol?.startsWith(value.key) && /TYPE_PARAM/.test(d.code)));
  }
  const invalid = structuredClone(graph);
  invalid.entities.find(e => e.key === entity('Selector').key).rawDoc = contractRaw('Selector').replace('@typeParam TOuter', '@typeParam TUnknown');
  const diagnostics = validateDocumentation(buildDocumentationModel(invalid, config));
  assert.ok(diagnostics.some(d => d.code === 'DOC_UNKNOWN_TYPE_PARAM' && d.message.includes('TUnknown')));
  assert.ok(diagnostics.some(d => d.code === 'DOC_MISSING_TYPE_PARAM' && d.message.includes('TOuter')));
  const missingInner = structuredClone(graph);
  const separate = missingInner.entities.find(e => e.key === entity('SeparateSelector').key);
  separate.signatures[0].rawDoc = separate.signatures[0].rawDoc.replace('   * @typeParam TInner - Required selected key.\n', '');
  assert.ok(validateDocumentation(buildDocumentationModel(missingInner, config)).some(d => d.code === 'DOC_MISSING_TYPE_PARAM' && d.message.includes('TInner')));
  const unknownInner = structuredClone(graph);
  const inner = unknownInner.entities.find(e => e.key === entity('SeparateSelector').key);
  inner.signatures[0].rawDoc = inner.signatures[0].rawDoc.replace('@typeParam TInner', '@typeParam TUnknown');
  assert.ok(validateDocumentation(buildDocumentationModel(unknownInner, config)).some(d => d.code === 'DOC_UNKNOWN_TYPE_PARAM' && d.message.includes('TUnknown')));
  function contractRaw(name) { return entity(name).rawDoc; }
});

test('the independent fixture has complete normalized contracts and self-contained source-only examples', () => {
  assert.deepEqual(validateDocumentation(model), []);
  const { checkDocumentationExamples } = require('../scripts/documentation-examples.cjs');
  assert.equal(model.examples.length, 15);
  assert.deepEqual(checkDocumentationExamples(root, model, config), []);
});


test('instantiated typed consts validate local signature lexical generics and render the closed shape without reopening them', () => {
  for (const reference of ['numberHandler', 'nameSelector']) {
    const projected = entity(reference);
    const diagnostics = validateDocumentation(model).filter(d => d.symbol?.startsWith(projected.key));
    assert.deepEqual(diagnostics, []);
    assert.deepEqual(projected.typeParameters, []);
    assert.deepEqual(projected.signatures[0].lexicalTypeParameters.map(p => p.name), ['TOuter']);
  }
  const number = entity('numberHandler').signatures[0];
  assert.equal(number.text, '(value: number): number');
  assert.equal(number.parameters[0].type.text, 'number');
  assert.equal(number.returns.text, 'number');
  const selector = entity('nameSelector').signatures[0];
  assert.equal(selector.text, '<TKey extends "name">(value: TKey): TKey');
  const page = renderApiPages(model)['docs/api/shared.md'];
  for (const reference of ['numberHandler', 'nameSelector']) {
    const marker = '\n## ' + reference + '\n';
    const start = page.indexOf(marker);
    const next = page.indexOf('\n## ', start + marker.length);
    const section = page.slice(start, next === -1 ? undefined : next);
    assert.ok(section.includes('Source declaration type parameters'));
    assert.ok(section.includes('These describe the source declaration; the callable uses the instantiated signature above.'));
    assert.ok(section.includes('`TOuter`:'));
    assert.ok(!/### Type parameters[^#]*`TOuter`:/.test(section));
    assert.ok(section.includes(reference === 'numberHandler' ? '(value: number): number' : '<TKey extends "name">(value: TKey): TKey'));
  }
  const withoutDuplicateOuter = structuredClone(graph);
  for (const projected of withoutDuplicateOuter.entities) {
    for (const signature of projected.signatures) {
      signature.rawDoc = signature.rawDoc.replace(/\n\s*\* @typeParam TOuter[^\n]*/g, '');
    }
  }
  assert.deepEqual(validateDocumentation(buildDocumentationModel(withoutDuplicateOuter, config)), []);
  const withoutInner = structuredClone(graph);
  const projectedSelector = withoutInner.entities.find(e => e.key === entity('nameSelector').key);
  projectedSelector.signatures[0].rawDoc = projectedSelector.signatures[0].rawDoc.replace(/\n\s*\* @typeParam TKey[^\n]*/, '');
  assert.ok(validateDocumentation(buildDocumentationModel(withoutInner, config)).some(d => d.symbol?.startsWith(projectedSelector.key) && d.code === 'DOC_MISSING_TYPE_PARAM' && d.message.includes('TKey')));
  const invalid = structuredClone(graph);
  for (const reference of ['numberHandler', 'nameSelector']) {
    const projected = invalid.entities.find(e => e.key === entity(reference).key);
    projected.signatures[0].rawDoc = projected.signatures[0].rawDoc.replace('@typeParam TOuter', '@typeParam TUndeclared');
  }
  const failures = validateDocumentation(buildDocumentationModel(invalid, config));
  for (const reference of ['numberHandler', 'nameSelector']) {
    assert.ok(failures.some(d => d.symbol?.startsWith(entity(reference).key) && d.code === 'DOC_UNKNOWN_TYPE_PARAM' && d.message.includes('TUndeclared')));
  }
});

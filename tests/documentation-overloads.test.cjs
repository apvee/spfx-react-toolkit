const assert = require('node:assert/strict');
const path = require('node:path');
const test = require('node:test');
const { buildApiGraph, resolvePublicReference } = require('../scripts/docs/documentation-model.cjs');
const { buildDocumentationModel, validateDocumentation, bindingContractKey } = require('../scripts/docs/documentation-contract.cjs');
const { checkDocumentationExamples } = require('../scripts/docs/documentation-examples.cjs');
const root = path.join(__dirname, 'fixtures/documentation/overload-contracts');
const config = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' };
let graph;
test.before(() => { graph = buildApiGraph(root, config); });

test('each overload owns its distinct parameter names and optionality at its physical source comment', () => {
  assert.deepEqual(graph.diagnostics, []);
  const model = buildDocumentationModel(graph, config);
  const entity = resolvePublicReference(graph, 'select');
  const signatures = entity.signatures.map(signature => model.contracts[signature.key]);
  assert.deepEqual(signatures.map(contract => contract.params.map(parameter => parameter.name)), [['options'], ['siteUrl', 'options']]);
  assert.deepEqual(signatures.map(contract => contract.typeParams.map(parameter => parameter.name)), [['TSelected'], ['TLegacy']]);
  assert.deepEqual(signatures.map(contract => contract.returns), ['The selected value.', 'The legacy value, or undefined when omitted.']);
  assert.deepEqual(signatures.map(contract => contract.examples[0].title), ['Scoped', 'Legacy']);
  assert.deepEqual(model.contracts[entity.key].source, signatures[0].source);
  assert.notDeepEqual(signatures[0].source, signatures[1].source);
  assert.deepEqual(validateDocumentation(model), []);
});

test('both overload examples and the binding-local example compile from source', () => {
  const model = buildDocumentationModel(graph, config);
  assert.equal(model.examples.length, 3);
  assert.deepEqual(checkDocumentationExamples(root, model, config), []);
});

test('each overload still rejects missing, foreign, duplicate and empty parameter documentation', () => {
  const cases = [
    [0, '@param options - Required selected value.', '', 'DOC_MISSING_PARAM', 'options'],
    [1, '@param siteUrl - Optional site URL; omitted or undefined leaves selection idle.', '', 'DOC_MISSING_PARAM', 'siteUrl'],
    [1, '@param options - Optional legacy value; omitted or undefined returns undefined.', '', 'DOC_MISSING_PARAM', 'options'],
    [0, '@param options', '@param siteUrl - Foreign legacy parameter.\n * @param options', 'DOC_UNKNOWN_PARAM', 'siteUrl'],
    [1, '@param options', '@param selected - Foreign scoped parameter.\n * @param options', 'DOC_UNKNOWN_PARAM', 'selected'],
    [0, '@param options', '@param options - Duplicate value.\n * @param options', 'DOC_DUPLICATE_PARAM', 'options'],
    [1, '@param options', '@param options - Duplicate value.\n * @param options', 'DOC_DUPLICATE_PARAM', 'options'],
    [0, '@param options - Required selected value.', '@param options -', 'DOC_EMPTY_TAG', 'options'],
    [1, '@param options - Optional legacy value; omitted or undefined returns undefined.', '@param options -', 'DOC_EMPTY_TAG', 'options'],
  ];
  for (const [index, before, after, code, name] of cases) {
    const { diagnostics, signature } = mutateSignature(index, before, after);
    assert.ok(diagnostics.some(item => item.symbol === signature.key && item.code === code && item.message.includes(name)), `${index}: ${code} ${name}: ${JSON.stringify(diagnostics)}`);
  }
});

test('optional omission, returns and generics remain required on the particular overload', () => {
  const cases = [
    [1, '@param options - Optional legacy value; omitted or undefined returns undefined.', '@param options - Legacy value.', 'DOC_MISSING_OMISSION', 'options'],
    [1, '@param siteUrl - Optional site URL; omitted or undefined leaves selection idle.', '@param siteUrl - Site URL.', 'DOC_MISSING_OMISSION', 'siteUrl'],
    [0, '@returns The selected value.', '', 'DOC_MISSING_RETURNS', undefined],
    [1, '@returns The legacy value, or undefined when omitted.', '', 'DOC_MISSING_RETURNS', undefined],
    [0, '@typeParam TSelected - Required selected value type.', '', 'DOC_MISSING_TYPE_PARAM', 'TSelected'],
    [1, '@typeParam TLegacy - Optional legacy value type.', '', 'DOC_MISSING_TYPE_PARAM', 'TLegacy'],
    [0, '@typeParam TSelected', '@typeParam TForeign', 'DOC_UNKNOWN_TYPE_PARAM', 'TForeign'],
    [1, '@typeParam TLegacy', '@typeParam TForeign', 'DOC_UNKNOWN_TYPE_PARAM', 'TForeign'],
  ];
  for (const [index, before, after, code, name] of cases) {
    const { diagnostics, signature } = mutateSignature(index, before, after);
    assert.ok(diagnostics.some(item => item.symbol === signature.key && item.code === code && (!name || item.message.includes(name))), `${index}: ${code}: ${JSON.stringify(diagnostics)}`);
  }
});

test('a separate binding-local overview keeps name completeness and optional omission checks', () => {
  for (const [before, after, code] of [
    ['@param siteUrl - Legacy site URL; omitted or undefined leaves selection idle.', '', 'DOC_MISSING_PARAM'],
    ['@param options', '@param absent', 'DOC_UNKNOWN_PARAM'],
    ['@param options - Scoped or legacy value; omitted or undefined in the legacy form returns undefined.', '@param options - Selected value.', 'DOC_MISSING_OMISSION'],
  ]) {
    const copy = structuredClone(graph);
    const binding = copy.bindings.find(item => item.exportedName === 'alternate');
    assert.ok(binding.rawDoc.includes(before));
    binding.rawDoc = binding.rawDoc.replace(before, after);
    const diagnostics = validateDocumentation(buildDocumentationModel(copy, config));
    assert.ok(diagnostics.some(item => item.symbol === bindingContractKey(binding) && item.code === code), `${code}: ${JSON.stringify(diagnostics)}`);
  }
});

function mutateSignature(index, before, after) {
  const copy = structuredClone(graph);
  const entity = resolvePublicReference(copy, 'select');
  const signature = entity.signatures[index];
  assert.ok(signature.rawDoc.includes(before));
  const changed = signature.rawDoc.replace(before, after);
  if (signature.rawDoc === entity.rawDoc) entity.rawDoc = changed;
  signature.rawDoc = changed;
  return { signature, diagnostics: validateDocumentation(buildDocumentationModel(copy, config)) };
}

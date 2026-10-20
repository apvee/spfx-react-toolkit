const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const ts = require('typescript');
const { assertCompatibleDeclaration } = require('../scripts/api-compatibility.helpers.cjs');
const baseline = require('./fixtures/api-baseline.json');

for (const [moduleName, interfaceName] of [
  ['hooks/useSPFxPnPSearch.d.ts', 'UseSPFxPnPSearchOptions'],
  ['services/spfx-pnp-search.service.d.ts', 'SPFxPnPSearchOptions'],
  ['services/spfx-pnp-search.service.d.ts', 'SPFxPnPSearchRequestOptions']
]) {
  test(`compatibility permits only the optional AND/OR addition in ${interfaceName}`, () => {
    const original = baseline[moduleName].declaration;
    const anchor = `export interface ${interfaceName} {`;
    const field = "refinerOperator?: 'and' | 'or';";
    const current = original.replace(anchor, anchor + '\n' + field);
    assert.doesNotThrow(() => assertCompatibleDeclaration(moduleName, current, original));
    for (const invalid of [
      "refinerOperator: 'and' | 'or';", 'refinerOperator?: string;',
      "refinerOperator?: 'and' | 'or' | 'xor';", 'refinerOperator?: any;',
      field + '\n' + field
    ]) {
      assert.throws(() => assertCompatibleDeclaration(moduleName,
        original.replace(anchor, anchor + '\n' + invalid), original));
    }
    assert.throws(() => assertCompatibleDeclaration(moduleName,
      current.replace('pageSize?: number;', 'pageSize?: string;'), original));
    // This approval cannot hide unrelated additions to result/state types.
    const resultAnchor = moduleName.startsWith('hooks') ? 'export interface SearchRefiner {' : 'export interface SPFxPnPSearchRefiner {';
    assert.throws(() => assertCompatibleDeclaration(moduleName,
      original.replace(resultAnchor, resultAnchor + '\n' + field), original));
  });
}

test('finished search feature requires each operator option without accepting omissions', () => {
  for (const moduleName of ['hooks/useSPFxPnPSearch.d.ts', 'services/spfx-pnp-search.service.d.ts']) {
    const original = baseline[moduleName].declaration;
    assert.throws(() => assertCompatibleDeclaration(moduleName, original, original,
      { requireSearchRefinerOperator: true }), /Missing approved search refiner operator/);
  }
});

test('search operators remain optional precise choices in hook and service consumers', () => {
  const program = ts.createProgram([path.resolve(__dirname, 'fixtures/pnp-search-options-types.ts')], {
    noEmit: true, strict: true, skipLibCheck: true, target: ts.ScriptTarget.ES2020,
    module: ts.ModuleKind.ESNext, moduleResolution: ts.ModuleResolutionKind.NodeJs, jsx: ts.JsxEmit.React,
    lib: ['lib.es2020.d.ts', 'lib.dom.d.ts']
  });
  const diagnostics = ts.getPreEmitDiagnostics(program);
  assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: file => file, getCurrentDirectory: () => process.cwd(), getNewLine: () => '\n'
  }));
});

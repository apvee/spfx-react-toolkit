const assert = require('node:assert/strict');
const ts = require('typescript');

/** Check executable compound usage rather than treating two exported values as full coverage. */
function assertFormExampleCoverage(formSurface, registry, panel) {
  for (const symbol of ['Form', 'createForm']) {
    assert.equal([...registry.matchAll(new RegExp("symbol: ['\"]" + symbol + "['\"]", 'g'))].length, 1, `${symbol} needs one registry entry`);
  }
  for (const marker of ["key: 'form'", "webpackChunkName: 'spfx-demo-form'", "'./panels/FormPanel'"]) assert.ok(registry.includes(marker), `Form lazy registry missing ${marker}`);
  const expected = formSurface.filter(item => ['FormSuite', 'ScopedFormSuite'].includes(item.name))
    .flatMap(item => item.members.map(member => member.name.split('.').at(-1)));
  const source = ts.createSourceFile('FormPanel.tsx', panel, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  const rendered = new Set(), called = new Set();
  function visit(node) {
    if (ts.isJsxOpeningElement(node) || ts.isJsxSelfClosingElement(node)) {
      if (ts.isPropertyAccessExpression(node.tagName)) rendered.add(node.tagName.name.text);
    }
    if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression)) called.add(node.expression.name.text);
    ts.forEachChild(node, visit);
  }
  visit(source);
  for (const member of new Set(expected)) assert.ok((member.startsWith('use') ? called : rendered).has(member), `Executable Form sample missing ${member}`);
  for (const marker of ['direction', 'resize', 'values', 'dirty', 'touched', 'errors', 'submit', 'reset', 'new-record', 'dirty-refresh', 'save-delay', 'save-fail', 'request-submit', 'settings-group', 'settings-reference', 'settings-enabled']) {
    assert.ok(panel.includes(`data-form-demo="${marker}"`), `Observable Form sample missing ${marker}`);
  }
  assert.doesNotMatch(panel, /\bstyle\s*=|\bLayout\./, 'Form sample must compose class descriptors');
}
module.exports = { assertFormExampleCoverage };

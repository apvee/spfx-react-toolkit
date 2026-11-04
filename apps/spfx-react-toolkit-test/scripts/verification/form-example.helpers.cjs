const assert = require('node:assert/strict');
const ts = require('typescript');

/** Check executable compound usage rather than treating two exported values as full coverage. */
function assertFormExampleCoverage(formSurface, catalogue, sourceText) {
  for (const symbol of ['Form', 'createForm']) {
    assert.ok(catalogue.includes('"featureTitle": "' + symbol + '"'), `${symbol} needs canonical scenario availability`);
  }
  for (const title of ['Edit typed reports across native fields and composition', 'Validate and transform quantity with recoverable failures', 'Edit a dynamic report with explicit save and reset']) {
    assert.ok(catalogue.includes('"title": "' + title + '"'), title);
  }
  const expected = formSurface.filter(item => ['FormSuite', 'ScopedFormSuite'].includes(item.name))
    .flatMap(item => item.members.map(member => member.name.split('.').at(-1)));
  const executable = [...sourceText.matchAll(/\* ```tsx\n([\s\S]*?) \* ```/g)].map(match => match[1].split('\n').map(line => line.replace(/^ \* ?/, '')).join('\n')).filter(code => code.includes('export function Demo(')).join('\n');
  const source = ts.createSourceFile('form-examples.tsx', executable, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
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
    assert.ok(sourceText.includes(`data-form-demo="${marker}"`) || sourceText.includes(`'data-form-demo': '${marker}'`), `Observable Form sample missing ${marker}`);
  }
  assert.doesNotMatch(sourceText, /\bstyle\s*=|\bLayout\./, 'Form sample must compose class descriptors');
}
module.exports = { assertFormExampleCoverage };

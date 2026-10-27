const assert = require('node:assert/strict');
function declarations(element, property) {
  const values = [];
  function visit(rules) {
    for (const rule of Array.from(rules)) {
      if (rule.selectorText && element.matches(rule.selectorText)) {
        const value = rule.style.getPropertyValue(property); if (value) values.push(value);
      }
      if (rule.cssRules) visit(rule.cssRules);
    }
  }
  for (const sheet of Array.from(element.ownerDocument.styleSheets)) visit(sheet.cssRules);
  return values;
}
exports.assertIconSizing = (svg, size) => {
  assert.equal(svg.getAttribute('width'), '80', 'consumer SVG is not cloned or rewritten');
  assert.equal(svg.getAttribute('height'), '80');
  assert.ok(declarations(svg, 'inline-size').includes('100%'), 'SVG has a matching rule to fill its icon container');
  assert.ok(declarations(svg, 'block-size').includes('100%'), 'SVG height fills its icon container');
  const inner = svg.parentElement, shell = inner.parentElement;
  assert.ok(declarations(inner, 'inline-size').includes('100%'), 'introduced slot container fills the owned shell');
  assert.ok(declarations(inner, 'block-size').includes('100%'));
  assert.ok(declarations(shell, 'inline-size').includes(`${size}px`), 'owned shell retains documented size');
  assert.ok(declarations(shell, 'block-size').includes(`${size}px`));
  for (const node of [svg, inner, shell]) assert.equal(node.hasAttribute('style'), false);
};

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { loadSxModules, createSxDom, mountSx, cssRules, activeRules } = require('./fixtures/sx-harness.cjs');
createSxDom();
const { JSDOM } = require('jsdom');
const { createDOMRenderer } = require('@griffel/core');
function fixture() {
  assert.ok(fs.existsSync(path.resolve(__dirname, '../packages/spfx-react-toolkit/src/hooks/useSx.ts')), 'The public useSx hook is missing');
  const f = loadSxModules();
  return { ...f, useSx: f.load('../../hooks/useSx').useSx };
}
test('default renderer inserts during React17 render without any toolkit provider', t => {
  const { useSx, styles: s } = fixture();
  const f = mountSx(useSx, { inputs: [s.width.px(241)] });
  t.after(() => f.unmount());
  assert.ok(f.element.className);
  assert.ok([...document.querySelectorAll('style[data-make-styles-bucket]')].some(element => [...element.sheet.cssRules].some(rule => rule.cssText.includes('241px'))));
});
test('renderer replacement and separate documents/salts use independent initialized factories', t => {
  const { useSx, styles: s } = fixture();
  const firstDoc = new JSDOM('<html><head></head><body></body></html>').window.document;
  const secondDoc = new JSDOM('<html><head></head><body></body></html>').window.document;
  const firstRenderer = createDOMRenderer(firstDoc, { classNameHashSalt: 'first' });
  const secondRenderer = createDOMRenderer(secondDoc, { classNameHashSalt: 'second' });
  const errors = [];
  const original = console.error;
  console.error = (...args) => errors.push(args.join(' '));
  t.after(() => { console.error = original; });
  const f = mountSx(useSx, { renderer: firstRenderer, inputs: [s.width.px(240)] }, firstDoc);
  t.after(() => f.unmount());
  const firstClass = f.element.className;
  const firstCount = cssRules(firstRenderer).length;
  f.render({ renderer: secondRenderer, targetDocument: secondDoc, inputs: [s.width.px(240)] });
  assert.notEqual(f.element.className, firstClass);
  assert.ok(cssRules(secondRenderer).some(rule => rule.includes('240px')));
  assert.equal(cssRules(firstRenderer).length, firstCount);
  assert.ok(firstDoc.querySelectorAll('style').length > 0);
  assert.ok(secondDoc.querySelectorAll('style').length > 0);
  // Core 1.19.2 logs this warning on the first call of each salted factory.
  // Characterize that upstream defect without suppressing it in production.
  assert.equal(errors.length, 4);
  assert.ok(errors.every(error => error.includes('different \"classNameHashSalt\"')));
});
test('Fluent direction changes and explicit override produce consistent LTR/RTL sequences', t => {
  const { useSx, createDeclaration } = fixture();
  const renderer = createDOMRenderer(document);
  const padding = createDeclaration('paddingLeft', '8px', { property: 'paddingLeft', fallback: '0px' });
  const f = mountSx(useSx, { renderer, dir: 'ltr', inputs: [padding] });
  t.after(() => f.unmount());
  const ltr = f.element.className;
  assert.ok(activeRules(renderer, ltr).some(rule => /padding-left:\s*var/.test(rule)));
  f.render({ renderer, dir: 'rtl', inputs: [padding] });
  assert.notEqual(f.element.className, ltr);
  assert.ok(activeRules(renderer, f.element.className).some(rule => /padding-right:\s*var/.test(rule)));
  f.render({ renderer, dir: 'rtl', options: { dir: 'ltr' }, inputs: [padding] });
  assert.equal(f.element.className, ltr);
});
test('host nonce/insertion point are respected and unmount preserves host styles', () => {
  const { useSx, styles: s } = fixture();
  const target = new JSDOM('<html><head><meta id="point"><style id="host">.host {color:red}</style></head><body></body></html>').window.document;
  const insertionPoint = target.getElementById('point');
  const renderer = createDOMRenderer(target, { insertionPoint, styleElementAttributes: { nonce: 'host-nonce', 'data-owner': 'host' } });
  const f = mountSx(useSx, { renderer, inputs: [s.width.px(240)] }, target);
  const styles = [...target.querySelectorAll('style[data-make-styles-bucket]')];
  assert.ok(styles.length > 0);
  for (const style of styles) {
    assert.equal(style.getAttribute('nonce'), 'host-nonce');
    assert.equal(style.getAttribute('data-owner'), 'host');
    assert.ok(insertionPoint.compareDocumentPosition(style) & target.defaultView.Node.DOCUMENT_POSITION_FOLLOWING);
    assert.ok(style.compareDocumentPosition(target.getElementById('host')) & target.defaultView.Node.DOCUMENT_POSITION_FOLLOWING);
  }
  const before = cssRules(renderer);
  f.unmount();
  assert.deepEqual(cssRules(renderer), before);
  assert.equal(target.getElementById('host').textContent, '.host {color:red}');
});

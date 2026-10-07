const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { loadSxModules, createSxDom, mountSx, cssRules, activeRules } = require('./fixtures/sx-harness.cjs');
createSxDom();
const { createDOMRenderer, makeStyles, mergeClasses } = require('@griffel/core');
function fixture() {
  assert.ok(fs.existsSync(path.resolve(__dirname, '../packages/spfx-react-toolkit/src/hooks/useSx.ts')), 'The public useSx hook is missing');
  const f = loadSxModules();
  return { ...f, useSx: f.load('../../hooks/useSx').useSx };
}
test('empty sx is empty and changing 240 to 360 to 240 reuses classes without inline styles', t => {
  const { useSx, styles: s } = fixture();
  const renderer = createDOMRenderer(document);
  const f = mountSx(useSx, { renderer, inputs: [s.width.px(240)] });
  t.after(() => f.unmount());
  const first = f.element.className;
  const before = cssRules(renderer).length;
  assert.equal(f.sx(false, null, undefined, ''), '');
  assert.equal(f.element.hasAttribute('style'), false);
  f.render({ renderer, inputs: [s.width.px(360)] });
  assert.notEqual(f.element.className, first);
  assert.ok(activeRules(renderer, f.element.className).some(rule => rule.includes('360px')));
  f.render({ renderer, inputs: [s.width.px(240)] });
  assert.equal(f.element.className, first);
  const after = cssRules(renderer).length;
  assert.ok(after > before);
  f.render({ renderer, inputs: [s.width.px(240)] });
  assert.equal(cssRules(renderer).length, after);
});
test('siblings with different values keep their own assignment classes', t => {
  const { useSx, styles: s } = fixture();
  const renderer = createDOMRenderer(document);
  const f = mountSx(useSx, { renderer, inputs: [s.width.px(240)], sibling: [s.width.px(360)] });
  t.after(() => f.unmount());
  const parent = activeRules(renderer, f.element.className);
  const child = activeRules(renderer, f.element.firstChild.className);
  assert.ok(parent.some(rule => rule.includes('240px')));
  assert.ok(!parent.some(rule => rule.includes('360px')));
  assert.ok(child.some(rule => rule.includes('360px')));
  assert.ok(!child.some(rule => rule.includes('240px')));
  assert.equal(f.container.querySelectorAll('[style]').length, 0);
});
test('intercalated external Griffel width wins until a later width descriptor rebinds it', t => {
  const { useSx, styles: s, createDeclaration } = fixture();
  const renderer = createDOMRenderer(document);
  const external = makeStyles({ root: { width: '400px' } })({ renderer, dir: 'ltr' }).root;
  const gap = createDeclaration('columnGap', '12px', { property: 'columnGap', fallback: 'normal' });
  const f = mountSx(useSx, { renderer, inputs: [s.width.px(240), external, gap] });
  t.after(() => f.unmount());
  let rules = activeRules(renderer, f.element.className);
  assert.ok(rules.some(rule => /width:\s*400px/.test(rule)));
  assert.ok(!rules.some(rule => /\{width:\s*var\(/.test(rule)));
  f.render({ renderer, inputs: [s.width.px(240), external, gap, s.width.px(360)] });
  rules = activeRules(renderer, f.element.className);
  assert.ok(!rules.some(rule => /width:\s*400px/.test(rule)));
  assert.ok(rules.some(rule => /\{width:\s*var\(/.test(rule)));
  assert.ok(rules.some(rule => /--apvee-sx-width-base:\s*360px/.test(rule)));
  assert.equal(mergeClasses(f.element.className, external).includes(external.split(' ').at(-1)), true);
});
test('two separate sx results preserve independent state assignments during merge', t => {
  const { useSx, styles: s } = fixture();
  const renderer = createDOMRenderer(document);
  const f = mountSx(useSx, { renderer, inputs: [] });
  t.after(() => f.unmount());
  const combined = mergeClasses(f.sx(s.hover(s.width.px(240))), f.sx(s.width.px(360)));
  const rules = activeRules(renderer, combined);
  assert.ok(rules.some(rule => /:hover.*--apvee-sx-width-base-hover:\s*240px/.test(rule)));
  assert.ok(rules.some(rule => /--apvee-sx-width-base:\s*360px/.test(rule)));
  assert.ok(rules.some(rule => /:where\(/.test(rule) && /initial/.test(rule)));
});
// Documentless real core CSS supports queries that JSDOM's CSSOM cannot parse.
function retainedCoreRules(renderer, className) {
  const classes = new Set(className.split(/\s+/));
  return Object.keys(renderer.insertionCache).filter(rule =>
    [...rule.matchAll(/\.([a-zA-Z0-9_-]+)/g)].some(match => classes.has(match[1])));
}
const foreignScopes = [
  ['hover', ':hover', s => s.hover],
  ['active', ':active', s => s.active],
  ['focus-visible', ':focus-visible', s => s.focusVisible],
  ['viewport', '@media (min-width: 640px)', s => s.viewport.medium],
  ['container', '@container apvee-sx (min-inline-size: 640px)', s => s.responsive.medium],
  ['viewport-hover', '@media (min-width: 640px)', s => d => s.viewport.medium(s.hover(d)), ':hover'],
  ['container-hover', '@container apvee-sx (min-inline-size: 640px)', s => d => s.responsive.medium(s.hover(d)), ':hover']
];
for (const [name, selector, descriptorScope, stateSelector] of foreignScopes) {
  test(`later native or sx width wins in the same foreign Griffel ${name} scope`, () => {
    const { styles: s, load } = fixture();
    const renderer = createDOMRenderer(null);
    const environment = { renderer, dir: 'ltr' };
    const resolve = (...inputs) => load('resolve.internal').resolveSxInputs(environment, inputs);
    const scopedWidth = stateSelector ? { [stateSelector]: { width: '400px' } } : { width: '400px' };
    const external = makeStyles({ root: { [selector]: scopedWidth } })(environment).root;
    const descriptor = descriptorScope(s)(s.width.px(360));
    for (const compose of [inputs => resolve(...inputs), inputs => mergeClasses(...inputs.map(input =>
      typeof input === 'string' ? input : resolve(input)))]) {
      const sxLast = retainedCoreRules(renderer, compose([external, descriptor]));
      assert.ok(!sxLast.some(rule => /width:\s*400px/.test(rule)), 'later descriptor must replace the native property in this scope');
      assert.ok(sxLast.some(rule => /width:\s*var\(/.test(rule)), 'native scoped property binding must remain');
      const nativeLast = retainedCoreRules(renderer, compose([descriptor, external]));
      assert.ok(nativeLast.some(rule => /width:\s*400px/.test(rule)), 'later native property must remain');
      assert.ok(!nativeLast.some(rule => /width:\s*var\(/.test(rule)), 'later native property must replace the scoped binding');
    }
  });
  test(`a ${name}-only descriptor preserves a foreign base width`, () => {
    const { styles: s, load } = fixture();
    const renderer = createDOMRenderer(null);
    const environment = { renderer, dir: 'ltr' };
    const external = makeStyles({ root: { width: '400px' } })(environment).root;
    const className = load('resolve.internal').resolveSxInputs(environment,
      [external, descriptorScope(s)(s.width.px(360))]);
    const rules = retainedCoreRules(renderer, className);
    assert.ok(rules.some(rule => /width:\s*400px/.test(rule)), 'foreign base declaration must survive scoped sx input');
    assert.ok(!rules.some(rule => /^\.[a-z0-9]+\{width:\s*var\(/.test(rule)), 'scoped input must not emit an unscoped native property');
  });
}
test('live assignments have greater specificity than local resets; container queries use inline size', () => {
  const { styles: s, load } = fixture();
  // Inspect the real core renderer insertion cache rather than relying on JSDOM
  // to implement container queries or CSS variable computed values.
  // The core runtime accepts a documentless target and stores real CSS rules.
  const renderer = createDOMRenderer(null);
  const classes = load('resolve.internal').resolveSxInputs({ renderer, dir: 'ltr' }, [s.width.px(240), s.responsive.medium(s.hover(s.width.px(360))), s.viewport.large(s.width.full)]);
  assert.ok(classes);
  const rules = Object.keys(renderer.insertionCache);
  const base = rules.find(rule => /--apvee-sx-width-base:\s*240px/.test(rule));
  assert.match(base, /^\.([a-z0-9]+)\.\1\s*\{/);
  assert.ok(rules.some(rule => /^@container apvee-sx \(min-inline-size:\s*640px\)/.test(rule)));
  assert.ok(rules.some(rule => /^@media \(min-width:\s*1024px\)/.test(rule)));
});
test('container inlineSize supplies only the named inline-size query boundary', () => {
  const { styles: s, load } = fixture();
  assert.ok(s.container, 'The container facade is missing');
  const declarations = load('normalize.internal').normalizeSxInputs([s.container.inlineSize])[0].declarations;
  assert.deepEqual(declarations.map(({ property, value }) => [property, value]), [
    ['containerType', 'inline-size'], ['containerName', 'apvee-sx']
  ]);
  assert.ok(Object.isFrozen(s.container.inlineSize));
});
test('removing query and state inputs after rerender removes their assignment classes', t => {
  const { useSx, styles: s } = fixture();
  // JSDOM's CSSOM cannot parse @container. Use the real documentless renderer
  // insertion cache to check emitted rules and mounted React class membership;
  // computed query/state behavior is established by the standalone browser fixture.
  const renderer = createDOMRenderer(null);
  const f = mountSx(useSx, { renderer, inputs: [s.width.px(111), s.responsive.medium(s.width.px(222)), s.hover(s.width.px(333))] });
  t.after(() => f.unmount());
  const mountedRules = () => {
    const classes = new Set(f.element.className.split(/\s+/));
    return Object.keys(renderer.insertionCache).filter(rule =>
      [...rule.matchAll(/\.([a-zA-Z0-9_-]+)/g)].some(match => classes.has(match[1])));
  };
  assert.ok(mountedRules().some(rule => /--apvee-sx-width-container-640:\s*222px/.test(rule)), 'the initial container assignment must be observable');
  assert.ok(mountedRules().some(rule => /--apvee-sx-width-base-hover:\s*333px/.test(rule)), 'the initial hover assignment must be observable');
  f.render({ renderer, inputs: [s.width.px(111)] });
  const rules = mountedRules();
  assert.ok(!rules.some(rule => /222px|333px/.test(rule)));
  assert.ok(rules.some(rule => /--apvee-sx-width-base:\s*111px/.test(rule)));
});

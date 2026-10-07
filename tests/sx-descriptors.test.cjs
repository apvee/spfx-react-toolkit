const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { loadStyles } = require('./fixtures/sx-harness.cjs');

function load() {
  assert.ok(fs.existsSync(path.resolve(__dirname, '../packages/spfx-react-toolkit/src/helpers/styles/index.ts')),
    'The public style descriptor exports are missing');
  return loadStyles();
}
function values(recipe) {
  return Object.fromEntries(recipe.declarations.map(item => [item.property, item.value]));
}

test('width and minimum width descriptors contain explicit CSS lengths', () => {
  const { width, minWidth } = load();
  assert.equal(width.full.property, 'width');
  assert.equal(width.full.value, '100%');
  assert.equal(width.auto.value, 'auto');
  assert.equal(width.px(240).value, '240px');
  assert.equal(width.px(0).value, '0px');
  assert.equal(minWidth.zero.property, 'minWidth');
  assert.equal(minWidth.zero.value, '0px');
  assert.equal(minWidth.px(10).value, '10px');
  assert.ok(Object.isFrozen(width.full));
});
test('pixel factories reject non-finite and negative values', () => {
  const { width, minWidth } = load();
  assert.throws(() => width.px(NaN), RangeError);
  assert.throws(() => width.px(-1), RangeError);
  assert.throws(() => width.px(Infinity), RangeError);
  assert.throws(() => minWidth.px(NaN), RangeError);
  assert.throws(() => minWidth.px(-1), RangeError);
  assert.throws(() => minWidth.px(-Infinity), RangeError);
});
test('subtle foreground retains the Fluent token reference', () => {
  const { foreground } = load();
  assert.equal(foreground.subtle.property, 'color');
  assert.equal(foreground.subtle.value, 'var(--colorNeutralForeground2)');
});
test('body1 includes only the four official typography properties', () => {
  const { typography } = load();
  assert.equal(typography.body1.kind, 'recipe');
  assert.deepEqual(values(typography.body1), {
    fontFamily: 'var(--fontFamilyBase)', fontSize: 'var(--fontSizeBase300)',
    fontWeight: 'var(--fontWeightRegular)', lineHeight: 'var(--lineHeightBase300)'
  });
  assert.ok(Object.isFrozen(typography.body1.declarations));
});
test('state and responsive scopes preserve inputs and inclusive breakpoints', () => {
  const { responsive, viewport, hover, active, focusVisible, foreground, width } = load();
  const state = hover(foreground.subtle);
  assert.equal(state.kind, 'state');
  assert.equal(state.state, 'hover');
  assert.deepEqual(state.inputs, [foreground.subtle]);
  const scoped = responsive.medium(state);
  assert.equal(scoped.kind, 'responsive');
  assert.equal(scoped.target, 'container');
  assert.equal(scoped.breakpoint, 640);
  assert.deepEqual(scoped.inputs, [state]);
  assert.equal(responsive.small(width.full).breakpoint, 480);
  assert.equal(responsive.large(width.full).breakpoint, 1024);
  assert.equal(viewport.medium(width.full).target, 'viewport');
  assert.equal(active(width.full).state, 'active');
  assert.equal(focusVisible(width.full).state, 'focus-visible');
  assert.ok(Object.isFrozen(scoped.inputs));
});

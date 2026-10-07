const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { loadSxModules } = require('./fixtures/sx-harness.cjs');
function fixture() {
  assert.ok(fs.existsSync(path.resolve(__dirname, '../packages/spfx-react-toolkit/src/helpers/styles/normalize.internal.ts')), 'The ordered sx normalizer is missing');
  const f = loadSxModules();
  return { ...f, normalize: f.load('normalize.internal').normalizeSxInputs };
}
test('normalization preserves intercalated class boundaries and the last declaration in each scope', () => {
  const { styles: s, normalize } = fixture();
  const result = normalize([s.width.px(240), s.width.full, 'external', s.width.px(360), s.hover(s.width.px(180)), s.responsive.medium(s.width.px(500)), false, null, undefined]);
  assert.equal(result.length, 3);
  assert.deepEqual(result.map(segment => segment.kind), ['declarations', 'class', 'declarations']);
  assert.equal(result[1].className, 'external');
  assert.deepEqual(result[0].declarations.map(d => [d.property, d.value]), [['width', '100%']]);
  assert.deepEqual(result[2].declarations.map(d => d.value), ['360px', '180px', '500px']);
});
test('logical padding/margin tuple recipes resolve per-longhand overrides without crossing strings', () => {
  const { createDeclaration: declaration, createRecipe: recipe, normalize } = fixture();
  const d = (property, value) => declaration(property, value, { property, fallback: '0px' });
  const padding = recipe('padding tuple', [d('paddingBlockStart', '4px'), d('paddingInlineEnd', '8px'), d('paddingBlockEnd', '12px'), d('paddingInlineStart', '16px')]);
  const margin = recipe('margin tuple', [d('marginBlockStart', '1px'), d('marginInlineEnd', '2px'), d('marginBlockEnd', '3px'), d('marginInlineStart', '4px')]);
  const [segment] = normalize([padding, margin, d('paddingInlineEnd', '20px')]);
  assert.deepEqual(segment.declarations.map(({ property, value }) => [property, value]), [
    ['paddingBlockStart', '4px'], ['paddingInlineEnd', '20px'], ['paddingBlockEnd', '12px'], ['paddingInlineStart', '16px'],
    ['marginBlockStart', '1px'], ['marginInlineEnd', '2px'], ['marginBlockEnd', '3px'], ['marginInlineStart', '4px']
  ]);
});
test('empty and falsy inputs produce no segments; scopes preserve canonical values', () => {
  const { normalize, styles: s } = fixture();
  assert.deepEqual(normalize([]), []);
  assert.deepEqual(normalize([false, null, undefined, '']), []);
  const [segment] = normalize([s.width.px(240), s.hover(s.width.px(180)), s.viewport.medium(s.focusVisible(s.width.px(360)))]);
  assert.deepEqual(segment.declarations.map(({ property, value, query, state }) => ({ property, value, query, state })), [
    { property: 'width', value: '240px', query: { target: 'base' }, state: undefined },
    { property: 'width', value: '180px', query: { target: 'base' }, state: 'hover' },
    { property: 'width', value: '360px', query: { target: 'viewport', breakpoint: 640 }, state: 'focus-visible' }
  ]);
});

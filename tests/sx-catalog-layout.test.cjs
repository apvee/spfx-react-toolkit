const test = require('node:test');
const assert = require('node:assert/strict');
const { loadSxModules } = require('./fixtures/sx-harness.cjs');
const load = name => loadSxModules().load(name);
function values(input) {
  return input.kind === 'recipe'
    ? Object.fromEntries(input.declarations.map(item => [item.property, item.value]))
    : { [input.property]: input.value };
}
const scales = [
  ['none', 'var(--spacingVerticalNone)', 'var(--spacingHorizontalNone)'],
  ['extraSmall', 'var(--spacingVerticalXS)', 'var(--spacingHorizontalXS)'],
  ['small', 'var(--spacingVerticalS)', 'var(--spacingHorizontalS)'],
  ['medium', 'var(--spacingVerticalM)', 'var(--spacingHorizontalM)'],
  ['large', 'var(--spacingVerticalL)', 'var(--spacingHorizontalL)'],
  ['extraLarge', 'var(--spacingVerticalXL)', 'var(--spacingHorizontalXL)']
];
const spacing = {
  gap: [['rowGap', 'Vertical'], ['columnGap', 'Horizontal']],
  padding: [['paddingBlockStart', 'Vertical'], ['paddingBlockEnd', 'Vertical'], ['paddingInlineStart', 'Horizontal'], ['paddingInlineEnd', 'Horizontal']],
  'padding-inline': [['paddingInlineStart', 'Horizontal'], ['paddingInlineEnd', 'Horizontal']],
  'padding-block': [['paddingBlockStart', 'Vertical'], ['paddingBlockEnd', 'Vertical']],
  'padding-inline-start': [['paddingInlineStart', 'Horizontal']],
  'padding-inline-end': [['paddingInlineEnd', 'Horizontal']],
  'padding-block-start': [['paddingBlockStart', 'Vertical']],
  'padding-block-end': [['paddingBlockEnd', 'Vertical']],
  margin: [['marginBlockStart', 'Vertical'], ['marginBlockEnd', 'Vertical'], ['marginInlineStart', 'Horizontal'], ['marginInlineEnd', 'Horizontal']],
  'margin-inline': [['marginInlineStart', 'Horizontal'], ['marginInlineEnd', 'Horizontal']],
  'margin-block': [['marginBlockStart', 'Vertical'], ['marginBlockEnd', 'Vertical']],
  'margin-inline-start': [['marginInlineStart', 'Horizontal']],
  'margin-inline-end': [['marginInlineEnd', 'Horizontal']],
  'margin-block-start': [['marginBlockStart', 'Vertical']],
  'margin-block-end': [['marginBlockEnd', 'Vertical']]
};
test('flex, items, grid and logical alignment expose only intended properties', () => {
  const flex = load('flex');
  assert.deepEqual(values(flex.row), { display: 'flex', flexDirection: 'row' });
  assert.deepEqual(values(flex.column), { display: 'flex', flexDirection: 'column' });
  assert.deepEqual(values(flex.wrap), { flexWrap: 'wrap' });
  assert.deepEqual(values(flex.noWrap), { flexWrap: 'nowrap' });
  const item = load('flex-item');
  for (const [name, expected] of [['grow', { flexGrow: '1' }], ['noGrow', { flexGrow: '0' }], ['shrink', { flexShrink: '1' }], ['noShrink', { flexShrink: '0' }]]) {
    assert.deepEqual(values(item[name]), expected);
  }
  assert.deepEqual(values(load('grid').columns(3)), { display: 'grid', gridTemplateColumns: 'repeat(3, minmax(0, 1fr))' });
  for (const [file, property, entries] of [
    ['align-items', 'alignItems', ['start', 'center', 'end', 'stretch', 'baseline']],
    ['justify-content', 'justifyContent', ['start', 'center', 'end', 'spaceBetween']],
    ['align-self', 'alignSelf', ['auto', 'start', 'center', 'end', 'stretch']]
  ]) for (const name of entries) assert.deepEqual(values(load(file)[name]), { [property]: name === 'spaceBetween' ? 'space-between' : name });
});
test('grid columns accepts positive integers and rejects invalid counts', () => {
  const { columns } = load('grid');
  assert.deepEqual(values(columns(1)), { display: 'grid', gridTemplateColumns: 'repeat(1, minmax(0, 1fr))' });
  for (const count of [0, -1, 1.5, NaN, Infinity, -Infinity]) assert.throws(() => columns(count), RangeError);
});
test('grid columns serializes large finite integers in CSS decimal integer syntax', () => {
  const { columns } = load('grid');
  assert.deepEqual(values(columns(1e21)), {
    display: 'grid', gridTemplateColumns: 'repeat(1000000000000000000000, minmax(0, 1fr))'
  });
  assert.deepEqual(values(columns(Number.MAX_VALUE)), {
    display: 'grid', gridTemplateColumns: `repeat(17976931348623157${'0'.repeat(292)}, minmax(0, 1fr))`
  });
});
for (const [file, properties] of Object.entries(spacing)) {
  test(`${file} uses literal axis token references and validated pixels`, () => {
    const facade = load(file);
    assert.deepEqual(Object.keys(facade).sort(), ['none', 'extraSmall', 'small', 'medium', 'large', 'extraLarge', 'px'].sort());
    for (const [name, verticalToken, horizontalToken] of scales) {
      const expected = Object.fromEntries(properties.map(([property, axis]) => [property, axis === 'Vertical' ? verticalToken : horizontalToken]));
      assert.deepEqual(values(facade[name]), expected);
      assert.ok(Object.isFrozen(facade[name]));
    }
    for (const pixels of [0, 2.5, 240]) assert.deepEqual(values(facade.px(pixels)), Object.fromEntries(properties.map(([property]) => [property, `${pixels}px`])));
    for (const pixels of [-1, NaN, Infinity, -Infinity]) assert.throws(() => facade.px(pixels), RangeError);
  });
}
test('global, axis and side spacing overrides share canonical logical properties', () => {
  const { load } = loadSxModules();
  const normalize = load('normalize.internal').normalizeSxInputs;
  for (const family of ['padding', 'margin']) {
    const normalized = normalize([load(family).px(4), load(`${family}-inline`).px(8), load(`${family}-inline-start`).px(12)]);
    assert.deepEqual(Object.fromEntries(normalized[0].declarations.map(item => [item.property, item.value])), {
      [`${family}BlockStart`]: '4px', [`${family}BlockEnd`]: '4px', [`${family}InlineStart`]: '12px', [`${family}InlineEnd`]: '8px'
    });
    const reversed = normalize([load(`${family}-inline-start`).px(12), load(`${family}-inline`).px(8), load(family).px(4)]);
    assert.deepEqual(Object.fromEntries(reversed[0].declarations.map(item => [item.property, item.value])), {
      [`${family}BlockStart`]: '4px', [`${family}BlockEnd`]: '4px', [`${family}InlineStart`]: '4px', [`${family}InlineEnd`]: '4px'
    });
  }
});
test('dimension facades use literal lengths and reject invalid pixels', () => {
  for (const [file, property, constants] of [
    ['width', 'width', { auto: 'auto', full: '100%' }], ['height', 'height', { auto: 'auto', full: '100%' }],
    ['min-width', 'minWidth', { zero: '0px' }], ['min-height', 'minHeight', { zero: '0px' }],
    ['max-width', 'maxWidth', { none: 'none' }], ['max-height', 'maxHeight', { none: 'none' }]
  ]) {
    const facade = load(file);
    for (const [name, value] of Object.entries(constants)) assert.deepEqual(values(facade[name]), { [property]: value });
    assert.deepEqual(values(facade.px(0)), { [property]: '0px' });
    assert.deepEqual(values(facade.px(12.5)), { [property]: '12.5px' });
    for (const value of [-1, NaN, Infinity, -Infinity]) assert.throws(() => facade.px(value), RangeError);
  }
});

test('spacing references exist at the exact react-theme 9.2.0 dependency floor', () => {
  const themePackage = require('@fluentui/react-theme/package.json');
  assert.equal(themePackage.version, '9.2.0');
  assert.equal(themePackage.dependencies['@fluentui/tokens'], '1.0.0-alpha.22');
  const { tokens } = require('@fluentui/react-theme');
  for (const [, verticalToken, horizontalToken] of scales) {
    for (const reference of [verticalToken, horizontalToken]) {
      const key = reference.slice(6, -1);
      assert.equal(tokens[key], reference);
    }
  }
});

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadSxModules } = require('./fixtures/sx-harness.cjs');
function declarations(f, inputs) {
  return Object.fromEntries(f.load('normalize.internal').normalizeSxInputs(inputs)[0].declarations.map(({ property, value }) => [property, value]));
}
// Literal oracles from react-theme 9.2.0 / tokens 1.0.0-alpha.22, independent of our factories.
const typographyOracle = {
  body1: ['Base300', 'Regular'], body1Strong: ['Base300', 'Semibold'], body1Stronger: ['Base300', 'Bold'],
  body2: ['Base400', 'Regular'], caption1: ['Base200', 'Regular'], caption1Strong: ['Base200', 'Semibold'],
  caption1Stronger: ['Base200', 'Bold'], caption2: ['Base100', 'Regular'], caption2Strong: ['Base100', 'Semibold'],
  subtitle1: ['Base500', 'Semibold'], subtitle2: ['Base400', 'Semibold'], subtitle2Stronger: ['Base400', 'Bold'],
  title1: ['Hero800', 'Semibold'], title2: ['Hero700', 'Semibold'], title3: ['Base600', 'Semibold'],
  largeTitle: ['Hero900', 'Semibold'], display: ['Hero1000', 'Semibold']
};
test('all 17 typography selections apply only the four official token properties', () => {
  const f = loadSxModules(), typography = f.load('typography');
  assert.deepEqual(Object.keys(typography).sort(), Object.keys(typographyOracle).sort());
  for (const [name, [size, weight]] of Object.entries(typographyOracle)) {
    const wanted = { fontFamily: 'var(--fontFamilyBase)', fontSize: `var(--fontSize${size})`,
      fontWeight: `var(--fontWeight${weight})`, lineHeight: `var(--lineHeight${size})` };
    assert.deepEqual(declarations(f, [typography[name]]), wanted, name);
    assert.deepEqual(require('@fluentui/react-theme').typographyStyles[name], wanted, `installed floor ${name}`);
    assert.ok(Object.isFrozen(typography[name]));
  }
});
test('text selections preserve logical alignment and truncate only the documented longhands', () => {
  const f = loadSxModules(), alignment = f.load('text-align'), text = f.load('text');
  for (const value of ['start', 'center', 'end']) assert.deepEqual(declarations(f, [alignment[value]]), { textAlign: value });
  assert.deepEqual(declarations(f, [text.wrap]), { whiteSpace: 'normal' });
  assert.deepEqual(declarations(f, [text.noWrap]), { whiteSpace: 'nowrap' });
  assert.deepEqual(declarations(f, [text.truncate]), { whiteSpace: 'nowrap', overflowX: 'hidden', overflowY: 'hidden', textOverflow: 'ellipsis' });
  assert.deepEqual(declarations(f, [text.truncate, text.wrap]), { whiteSpace: 'normal', overflowX: 'hidden', overflowY: 'hidden', textOverflow: 'ellipsis' });
});
const foregroundOracle = { primary: 'colorNeutralForeground1', subtle: 'colorNeutralForeground2', muted: 'colorNeutralForeground3',
  disabled: 'colorNeutralForegroundDisabled', brand: 'colorBrandForeground1', onBrand: 'colorNeutralForegroundOnBrand',
  inverted: 'colorNeutralForegroundInverted', link: 'colorBrandForegroundLink' };
const backgroundOracle = { canvas: 'colorNeutralBackground1', alternative: 'colorNeutralBackground2', subtle: 'colorSubtleBackground',
  transparent: 'colorTransparentBackground', brand: 'colorBrandBackground', brandTint: 'colorBrandBackground2',
  inverted: 'colorNeutralBackgroundInverted', disabled: 'colorNeutralBackgroundDisabled' };
const presetOracle = { canvas: ['colorNeutralBackground1', 'colorNeutralForeground1'], alternative: ['colorNeutralBackground2', 'colorNeutralForeground1'],
  subtle: ['colorSubtleBackground', 'colorNeutralForeground2'], transparent: ['colorTransparentBackground'],
  brand: ['colorBrandBackground', 'colorNeutralForegroundOnBrand'], brandTint: ['colorBrandBackground2', 'colorBrandForeground2'],
  inverted: ['colorNeutralBackgroundInverted', 'colorNeutralForegroundInverted'], success: ['colorStatusSuccessBackground1', 'colorStatusSuccessForeground1'],
  warning: ['colorStatusWarningBackground1', 'colorStatusWarningForeground1'], danger: ['colorStatusDangerBackground1', 'colorStatusDangerForeground1'],
  disabled: ['colorNeutralBackgroundDisabled', 'colorNeutralForegroundDisabled'] };
test('foreground and background apply the approved Fluent role references without implicit states', () => {
  const f = loadSxModules();
  for (const [module, property, oracle] of [['foreground', 'color', foregroundOracle], ['background', 'backgroundColor', backgroundOracle]]) {
    const selections = f.load(module);
    assert.deepEqual(Object.keys(selections).sort(), Object.keys(oracle).sort());
    for (const [name, token] of Object.entries(oracle)) assert.deepEqual(declarations(f, [selections[name]]), { [property]: `var(--${token})` });
  }
});
test('presets contain only their approved color properties and transparent preserves inherited color', () => {
  const f = loadSxModules(), presets = f.load('presets');
  assert.deepEqual(Object.keys(presets).sort(), Object.keys(presetOracle).sort());
  for (const [name, [background, foreground]] of Object.entries(presetOracle)) {
    const wanted = { backgroundColor: `var(--${background})` };
    if (foreground) wanted.color = `var(--${foreground})`;
    assert.deepEqual(declarations(f, [presets[name]]), wanted, name);
  }
});
test('preset and foreground overrides resolve property by property in argument order', () => {
  const f = loadSxModules(), presets = f.load('presets'), foreground = f.load('foreground');
  assert.deepEqual(declarations(f, [presets.canvas, foreground.subtle]), {
    backgroundColor: 'var(--colorNeutralBackground1)', color: 'var(--colorNeutralForeground2)' });
  assert.deepEqual(declarations(f, [foreground.subtle, presets.canvas]), {
    color: 'var(--colorNeutralForeground1)', backgroundColor: 'var(--colorNeutralBackground1)' });
});
test('every catalog color token exists in the installed light dark and high contrast floor', () => {
  const theme = require('@fluentui/react-theme');
  assert.equal(require('@fluentui/react-theme/package.json').version, '9.2.0');
  assert.equal(require('@fluentui/tokens/package.json').version, '1.0.0-alpha.22');
  const tokens = [...Object.values(foregroundOracle), ...Object.values(backgroundOracle), ...Object.values(presetOracle).flat(),
    'colorNeutralStroke1', 'colorNeutralStroke2', 'colorBrandStroke1', 'colorNeutralStrokeDisabled', 'colorNeutralStrokeAccessible'];
  for (const name of ['webLightTheme', 'webDarkTheme', 'teamsHighContrastTheme']) {
    for (const token of tokens) assert.equal(typeof theme[name][token], 'string', `${name}.${token}`);
  }
});
test('border geometry and color recipes expand all four sides to independent longhands', () => {
  const f = loadSxModules(), width = f.load('border-width'), style = f.load('border-style'), color = f.load('border-color');
  for (const [name, value] of Object.entries({ none: '0px', thin: 'var(--strokeWidthThin)', thick: 'var(--strokeWidthThick)' })) {
    assert.deepEqual(declarations(f, [width[name]]), { borderTopWidth: value, borderRightWidth: value, borderBottomWidth: value, borderLeftWidth: value });
  }
  assert.deepEqual(declarations(f, [style.solid]), { borderTopStyle: 'solid', borderRightStyle: 'solid', borderBottomStyle: 'solid', borderLeftStyle: 'solid' });
  for (const [name, token] of Object.entries({ primary: 'colorNeutralStroke1', subtle: 'colorNeutralStroke2', brand: 'colorBrandStroke1', disabled: 'colorNeutralStrokeDisabled' })) {
    const value = `var(--${token})`;
    assert.deepEqual(declarations(f, [color[name]]), { borderTopColor: value, borderRightColor: value, borderBottomColor: value, borderLeftColor: value });
  }
  const result = declarations(f, [width.thick, style.solid, color.primary, width.none, color.brand]);
  assert.equal(result.borderTopWidth, '0px');
  assert.equal(result.borderTopStyle, 'solid');
  assert.equal(result.borderTopColor, 'var(--colorBrandStroke1)');
});
test('radius and shadow selections use the exact supported tokens without other geometry', () => {
  const f = loadSxModules(), radius = f.load('border-radius'), shadow = f.load('box-shadow');
  for (const [name, token] of Object.entries({ none: 'borderRadiusNone', small: 'borderRadiusSmall', medium: 'borderRadiusMedium',
    large: 'borderRadiusLarge', extraLarge: 'borderRadiusXLarge', circular: 'borderRadiusCircular' })) {
    const value = `var(--${token})`;
    assert.deepEqual(declarations(f, [radius[name]]), { borderTopLeftRadius: value, borderTopRightRadius: value, borderBottomLeftRadius: value, borderBottomRightRadius: value });
    assert.equal(typeof require('@fluentui/react-theme').webLightTheme[token], 'string');
  }
  for (const [name, value] of Object.entries({ none: 'none', small: 'var(--shadow4)', medium: 'var(--shadow8)' })) {
    assert.deepEqual(declarations(f, [shadow[name]]), { boxShadow: value });
  }
  for (const token of ['strokeWidthThin', 'strokeWidthThick', 'shadow4', 'shadow8']) {
    assert.equal(typeof require('@fluentui/react-theme').webLightTheme[token], 'string');
  }
});
test('overflow recipes keep axis selection independent and leave native coupling to CSS', () => {
  const f = loadSxModules(), overflow = f.load('overflow');
  for (const value of ['visible', 'hidden', 'auto']) {
    assert.deepEqual(declarations(f, [overflow[value]]), { overflowX: value, overflowY: value });
    assert.deepEqual(declarations(f, [overflow.horizontal[value]]), { overflowX: value });
    assert.deepEqual(declarations(f, [overflow.vertical[value]]), { overflowY: value });
  }
  assert.deepEqual(declarations(f, [overflow.hidden, overflow.horizontal.auto]), { overflowX: 'auto', overflowY: 'hidden' });
});
test('scrollbar exposes only fluent and changes appearance without choosing scrolling behavior', () => {
  const f = loadSxModules(), scrollbar = f.load('scrollbar');
  assert.deepEqual(Object.keys(scrollbar), ['fluent']);
  assert.deepEqual(declarations(f, [scrollbar.fluent]), { scrollbarWidth: 'thin', scrollbarColor: 'var(--colorNeutralStrokeAccessible) transparent' });
  const { createDOMRenderer } = require('@griffel/core');
  const renderer = createDOMRenderer(null);
  f.load('resolve.internal').resolveSxInputs({ renderer, dir: 'ltr' }, [scrollbar.fluent]);
  const css = Object.keys(renderer.insertionCache).join('\n');
  assert.match(css, /@media\s*\(forced-colors:\s*active\).*scrollbar-color:\s*auto/);
  assert.ok(!/overflow|scrollbar-gutter|overscroll|scroll-behavior|-webkit-scrollbar|forced-color-adjust/.test(css));
});
test('forced colors binding overrides participate in caching and emit their own media rule', () => {
  const f = loadSxModules(), { createDOMRenderer } = require('@griffel/core');
  const renderer = createDOMRenderer(null), resolve = f.load('resolve.internal').resolveSxInputs;
  const create = forcedColors => f.createDeclaration('scrollbarColor', 'red transparent', {
    property: 'scrollbarColor', fallback: 'auto', ...(forcedColors === undefined ? {} : { forcedColors }) });
  resolve({ renderer, dir: 'ltr' }, [create()]);
  resolve({ renderer, dir: 'ltr' }, [create('auto')]);
  resolve({ renderer, dir: 'ltr' }, [create('initial')]);
  const css = Object.keys(renderer.insertionCache).join('\n');
  assert.match(css, /@media\s*\(forced-colors:\s*active\).*scrollbar-color:\s*auto/);
  assert.match(css, /@media\s*\(forced-colors:\s*active\).*scrollbar-color:\s*initial/);
});

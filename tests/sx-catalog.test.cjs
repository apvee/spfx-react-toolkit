const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { loadSxModules } = require('./fixtures/sx-harness.cjs');
const stylesDirectory = path.resolve(__dirname, '../packages/spfx-react-toolkit/src/helpers/styles');
const directTokenCounts = {
  foreground: 8, background: 8, presets: 21, 'border-color': 16,
  'border-width': 8, 'border-radius': 24, 'box-shadow': 2, scrollbar: 1
};
function inspectThemeReferences(source, filename) {
  const ast = ts.createSourceFile(filename, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
  const theme = require('@fluentui/react-theme');
  const violations = [], directTokens = [];
  const imported = new Set();
  for (const statement of ast.statements) {
    if (ts.isImportDeclaration(statement) && statement.moduleSpecifier.text === '@fluentui/react-theme') {
      const bindings = statement.importClause?.namedBindings;
      if (bindings && ts.isNamedImports(bindings)) {
        for (const element of bindings.elements) {
          if (!element.propertyName) imported.add(element.name.text);
        }
      }
    }
  }
  function visit(node) {
    // AST literal nodes exclude comments/JSDoc and the independent test oracles.
    if (ts.isStringLiteralLike(node) || ts.isTemplateHead(node) || ts.isTemplateMiddle(node) || ts.isTemplateTail(node)) {
      for (const match of node.text.matchAll(/var\(\s*--([\w-]+)/g)) {
        if (!match[1].startsWith('apvee-sx-')) violations.push(`handwritten var(--${match[1]})`);
      }
    }
    if (ts.isPropertyAccessExpression(node) && ts.isIdentifier(node.expression) && node.expression.text === 'tokens') {
      const key = node.name.text;
      directTokens.push(key);
      if (!imported.has('tokens') || !Object.hasOwn(theme.tokens, key)) violations.push(`unofficial tokens.${key}`);
    }
    if (ts.isPropertyAccessExpression(node) && ts.isPropertyAccessExpression(node.expression)
      && ts.isIdentifier(node.expression.expression) && node.expression.expression.text === 'typographyStyles') {
      const recipe = node.expression.name.text, property = node.name.text;
      if (!imported.has('typographyStyles') || !Object.hasOwn(theme.typographyStyles[recipe] || {}, property)) {
        violations.push(`unofficial typographyStyles.${recipe}.${property}`);
      }
    }
    // Existing typed tokens[`spacingVertical${suffix}`] accesses remain valid;
    // their union keys are checked by TypeScript and their values by layout tests.
    ts.forEachChild(node, visit);
  }
  visit(ast);
  return { violations, directTokens };
}
test('catalog theme references use official Fluent accesses instead of handwritten variable literals', () => {
  const violations = [], counts = {};
  for (const filename of fs.readdirSync(stylesDirectory).filter(name => name.endsWith('.ts'))) {
    const result = inspectThemeReferences(fs.readFileSync(path.join(stylesDirectory, filename), 'utf8'), filename);
    violations.push(...result.violations.map(message => `${filename}: ${message}`));
    const module = filename.slice(0, -3);
    if (Object.hasOwn(directTokenCounts, module)) counts[module] = result.directTokens.length;
  }
  assert.deepEqual(violations, [], `${violations.length} invalid catalog theme references`);
  assert.deepEqual(counts, directTokenCounts);
});
test('theme reference policy catches strings and template fragments while allowing private variables and official recipes', () => {
  const source = [
    "import { tokens, typographyStyles } from '@fluentui/react-theme';",
    '/** var(--colorNeutralForeground1) */',
    "const privateValue = 'var(--apvee-sx-color, inherit)';",
    'const official = tokens.colorNeutralForeground1;',
    'const typography = typographyStyles.body1.fontFamily;',
    'const spacing = tokens[`spacingVertical${suffix}`];',
    "const literal = 'var(--colorNeutralForeground1)';",
    'const template = `var(--colorNeutralBackground1) ${tokens.colorNeutralForeground1} var(--shadow4)`;',
    'const invalid = tokens.missingToken;'
  ].join('\n');
  assert.deepEqual(inspectThemeReferences(source, 'policy.ts').violations, [
    'handwritten var(--colorNeutralForeground1)', 'handwritten var(--colorNeutralBackground1)',
    'handwritten var(--shadow4)', 'unofficial tokens.missingToken'
  ]);
  assert.deepEqual(inspectThemeReferences('const value = tokens.colorNeutralForeground1;', 'unimported.ts').violations,
    ['unofficial tokens.colorNeutralForeground1']);
});
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
  assert.ok(!/overflow|scrollbar-gutter|overscroll|scroll-behavior|forced-color-adjust/.test(css));
});
test('fluent scrollbar emits equally sized vendor axes only outside forced colors', () => {
  const f = loadSxModules(), { createDOMRenderer } = require('@griffel/core');
  const renderer = createDOMRenderer(null);
  f.load('resolve.internal').resolveSxInputs({ renderer, dir: 'ltr' }, [f.styles.scrollbar.fluent]);
  const rules = Object.keys(renderer.insertionCache);
  const vendor = rules.filter(rule => rule.includes('::-webkit-scrollbar'));
  assert.ok(vendor.length > 0, 'The recipe must size both vendor scrollbar axes');
  assert.ok(vendor.every(rule => rule.includes('@supports selector(::-webkit-scrollbar)')
    && /@media\s*\(forced-colors:\s*none\)/.test(rule)));
  assert.ok(vendor.some(rule => /::-webkit-scrollbar\{width:6px/.test(rule)), 'The vertical scrollbar uses the approved 6px width');
  assert.ok(vendor.some(rule => /::-webkit-scrollbar\{height:6px/.test(rule)), 'The horizontal scrollbar uses the approved 6px height');
  assert.ok(vendor.some(rule => /::-webkit-scrollbar-thumb\{background-color:var\(--colorNeutralStrokeAccessible\)/.test(rule)));
  assert.ok(vendor.some(rule => /::-webkit-scrollbar-thumb\{border-radius:3px/.test(rule)), 'The 6px thumb retains rounded caps');
  for (const [property, variable] of [
    ['width', '--apvee-sx-scrollbarWidth-vendor-base'],
    ['color', '--apvee-sx-scrollbarColor-vendor-base']
  ]) {
    assert.ok(rules.some(rule => rule.includes('@supports selector(::-webkit-scrollbar)')
      && /@media\s*\(forced-colors:\s*none\)/.test(rule) && rule.includes(`${variable}:auto`)));
    assert.ok(rules.some(rule => rule.includes(`scrollbar-${property}:var(${variable},`)));
    assert.ok(rules.some(rule => rule.includes(':where(') && rule.includes(`${variable}:initial`)),
      'Every vendor override blocks inherited assignments locally');
  }
});
test('vendor scrollbar binding caches stay distinct from generic bindings in either insertion order', () => {
  for (const recipeFirst of [false, true]) {
    const f = loadSxModules(), { createDOMRenderer } = require('@griffel/core');
    const renderer = createDOMRenderer(null), resolve = f.load('resolve.internal').resolveSxInputs;
    const generic = f.createRecipe('generic-scrollbar', [
      f.createDeclaration('scrollbarWidth', 'thin', { property: 'scrollbarWidth', fallback: 'auto' }),
      f.createDeclaration('scrollbarColor', 'var(--colorNeutralStrokeAccessible) transparent', {
        property: 'scrollbarColor', fallback: 'auto', forcedColors: 'auto' })
    ]);
    const classes = new Map();
    for (const input of recipeFirst ? [f.styles.scrollbar.fluent, generic] : [generic, f.styles.scrollbar.fluent]) {
      classes.set(input, resolve({ renderer, dir: 'ltr' }, [input]));
    }
    const { activeRules } = require('./fixtures/sx-harness.cjs');
    assert.ok(activeRules(renderer, classes.get(f.styles.scrollbar.fluent)).some(rule => rule.includes('::-webkit-scrollbar')));
    assert.ok(!activeRules(renderer, classes.get(generic)).some(rule => rule.includes('::-webkit-scrollbar')));
  }
});
test('vendor scrollbar selectors stay inside responsive and interaction scopes', () => {
  const f = loadSxModules(), { createDOMRenderer } = require('@griffel/core');
  for (const [input, scope, state] of [
    [f.styles.responsive.medium(f.styles.hover(f.styles.scrollbar.fluent)), '@container apvee-sx (min-inline-size: 640px)', ':hover'],
    [f.styles.viewport.small(f.styles.focusVisible(f.styles.scrollbar.fluent)), '@media (min-width: 480px)', ':focus-visible']
  ]) {
    const renderer = createDOMRenderer(null);
    f.load('resolve.internal').resolveSxInputs({ renderer, dir: 'ltr' }, [input]);
    const vendor = Object.keys(renderer.insertionCache).filter(rule => rule.includes('::-webkit-scrollbar'));
    assert.ok(vendor.length > 0);
    assert.ok(vendor.every(rule => rule.includes(scope)
      && (rule.includes(`${state}::-webkit-scrollbar`) || rule.includes(`${state}{--apvee-sx-`))));
  }
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

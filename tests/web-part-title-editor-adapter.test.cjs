const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const base = 'controls/web-part-shell/title/';
const roles = ['primary', 'subtle', 'brand', 'neutralMuted', 'brandAlt', 'red', 'orange', 'yellow', 'green', 'teal', 'blue', 'purple', 'pink', 'brown', 'neutralDark', 'neutralLight', 'neutralLighter', 'brandDarkest', 'brandDarker', 'brandDark', 'brandLight', 'darkRed', 'amber', 'yellowGreen', 'cyan', 'navy'];
test('engine adapter preserves every allowed color identity and Automatic absence', t => {
  const h = createHarness(); t.after(() => h.close());
  const adapter = h.load(base + 'web-part-title-editor-adapter.internal.ts');
  for (const color of [...roles, undefined]) {
    const title = { type: 'heading', level: 2, content: [{ type: 'text', text: 'Title', ...(color ? { marks: { color } } : {}) }] };
    const engine = adapter.titleDocumentToEditorContent(title);
    assert.deepEqual(adapter.editorContentToTitleDocument(engine), title);
    assert.equal(engine.content[0].content[0].marks?.[0]?.attrs.color, color);
  }
});
test('adapter refuses unregistered or arbitrary color values on both protocol boundaries', t => {
  const h = createHarness(); t.after(() => h.close());
  const adapter = h.load(base + 'web-part-title-editor-adapter.internal.ts');
  for (const color of ['#f00', 'custom', 'colorPalettePinkForeground2', '__proto__', 'constructor']) {
    assert.throws(() => adapter.titleDocumentToEditorContent({ type: 'paragraph', content: [{ type: 'text', text: 'T', marks: { color } }] }), /Invalid/);
    assert.throws(() => adapter.editorContentToTitleDocument({ type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'T', marks: [{ type: 'titleStyle', attrs: { color } }] }] }] }), /Invalid/);
  }
});

test('picker offers twelve theme and ten standard foreground colors in approved order', t => {
  const h = createHarness(); t.after(() => h.close());
  const format = h.load(base + 'web-part-title-format.internal.ts');
  const themes = require('@fluentui/react-theme');
  const expected = ['primary', 'neutralDark', 'subtle', 'neutralMuted', 'neutralLight', 'neutralLighter', 'brandDarkest', 'brandDarker', 'brandDark', 'brandAlt', 'brand', 'brandLight', 'darkRed', 'red', 'amber', 'yellow', 'yellowGreen', 'green', 'cyan', 'blue', 'navy', 'purple'];
  for (const theme of [themes.webLightTheme, themes.webDarkTheme]) {
    const colors = format.getTitlePickerColors(undefined, theme);
    const themeExpected = theme === themes.webDarkTheme ? expected.filter(id => id !== 'red') : expected;
    assert.deepEqual(Array.from(colors, item => item.id), themeExpected);
    assert.equal(colors.filter(item => item.group === 'theme').length, 12);
    assert.equal(colors.filter(item => item.group === 'palette').length, theme === themes.webDarkTheme ? 9 : 10);
  }
  assert.equal(format.getTitleColorValueKey('yellow', themes.webLightTheme), themes.webLightTheme.colorPaletteYellowForeground3);
  assert.equal(format.getTitleColorValueKey('yellowGreen', themes.webLightTheme), '#bdc704');
  for (const legacy of ['orange', 'teal', 'pink', 'brown']) {
    assert.equal(format.getTitlePickerColors(legacy, themes.webLightTheme).filter(item => item.id === legacy).length, 1);
    assert.equal(format.getTitlePickerColors(undefined, themes.webLightTheme).some(item => item.id === legacy), false);
  }
});
test('picker deduplicates mixes and endpoints while preserving selected identity and custom CSS', t => {
  const h = createHarness(); t.after(() => h.close());
  const format = h.load(base + 'web-part-title-format.internal.ts');
  const themes = require('@fluentui/react-theme');
  const allSame = Object.fromEntries(Object.entries(themes.webLightTheme).map(([key, value]) => [key, key.includes('Foreground') ? '#112233' : value]));
  assert.deepEqual(Array.from(format.getTitlePickerColors(undefined, allSame), item => item.id), ['primary']);
  assert.deepEqual(Array.from(format.getTitlePickerColors('brandLight', allSame), item => item.id), ['brandLight']);
  assert.deepEqual(Array.from(format.getTitlePickerColors('pink', allSame), item => item.id), ['pink']);
  const systemTheme = { ...allSame, colorNeutralForeground1: 'CanvasText', colorNeutralForeground2: 'CanvasText' };
  assert.equal(format.getTitleColorValueKey('neutralDark', systemTheme), 'canvastext');
  const cssTheme = { ...themes.webLightTheme, colorBrandForeground1: 'var(--hostBrand)', colorNeutralForeground1: 'CanvasText' };
  assert.equal(format.getTitleColorValueKey('brandDarkest', cssTheme), 'color-mix(in srgb, var(--hostBrand) 25%, canvastext)');
  assert.ok(format.getTitlePickerColors('brandDarkest', cssTheme).some(item => item.id === 'brandDarkest'));
  for (const selected of [undefined, 'brand', 'yellowGreen']) {
    const colors = format.getTitlePickerColors(selected, themes.teamsHighContrastTheme);
    assert.equal(new Set(Array.from(colors, item => format.getTitleColorValueKey(item.id, themes.teamsHighContrastTheme))).size, colors.length);
    if (selected) assert.ok(colors.some(item => item.id === selected));
  }
});

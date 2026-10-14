const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const harness = require('./fixtures/formatted-text-harness.cjs');
const { cssRules } = require('./fixtures/sx-harness.cjs');
const base = '../packages/spfx-react-toolkit/src/controls/web-part-shell/title/';
let api, codec, styles, React, ReactDOM, act, renderer, host, dom;
test.before(async () => {
  dom = harness.createSxDom();
  api = await harness.loadModule(path.resolve(__dirname, base + 'web-part-title-renderer.internal.tsx'));
  codec = await harness.loadModule(path.resolve(__dirname, base + 'web-part-title-codec.internal.ts'));
  styles = await harness.loadModule(path.resolve(__dirname, base + 'web-part-title-styles.internal.ts'));
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  renderer = require('@griffel/core').createDOMRenderer(document);
  host = document.createElement('div'); document.body.appendChild(host);
});
test.after(async () => {
  if (host) act(() => { ReactDOM.unmountComponentAtNode(host); });
  await harness.cleanup(); if (dom) dom.window.close();
});
function render(document, theme = require('@fluentui/react-theme').webLightTheme, dir = 'ltr') {
  const { RendererProvider } = require('@griffel/react');
  const { FluentProvider } = require('@fluentui/react-provider');
  act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer },
    React.createElement(FluentProvider, { theme, dir, 'data-provider': '' }, React.createElement(api.WebPartTitleReadRenderer, { document }))), host); });
  return host.querySelector('h1,h2,h3,h4,p');
}
const content = [{ type: 'text', text: 'Title' }];
test('Read uses semantic H1-H4 and paragraph with explicit size independent of heading', () => {
  for (const level of [1, 2, 3, 4]) {
    const block = render({ type: 'heading', level, content: [{ type: 'text', text: 'Title', marks: { size: 28 } }] });
    assert.ok(block, 'Read must render a semantic title block');
    assert.equal(block.tagName, 'H' + level);
    assert.equal(block.querySelector('span').className, 'wps-size-28');
  }
  assert.equal(render({ type: 'paragraph', content }).tagName, 'P');
});
test('Read composes marks and breaks with literal escaped text and no input mutation', () => {
  const ast = Object.freeze({ type: 'heading', level: 2, content: Object.freeze([
    Object.freeze({ type: 'text', text: '<img src=x>', marks: Object.freeze({ bold: true, italic: true, underline: true, color: 'brand', size: 20 }) }),
    Object.freeze({ type: 'hardBreak' }), Object.freeze({ type: 'text', text: 'End' })
  ]) });
  const block = render(ast);
  assert.equal(block.querySelector('strong em u span').textContent, '<img src=x>');
  assert.equal(block.querySelectorAll('br').length, 1);
  assert.ok(!block.querySelector('img'));
  assert.equal(block.querySelector('span').className, 'wps-color-brand wps-size-20');
  assert.equal(ast.content.length, 3);
});
test('empty Read mounts through JSX without any heading or text', () => {
  for (const content of [[], [{ type: 'text', text: ' \u00a0\n' }, { type: 'hardBreak' }]]) {
    assert.equal(render({ type: 'heading', level: 2, content }), null);
    assert.equal(host.querySelector('[data-provider]').textContent, '');
  }
});
test('scoped CSS covers finite sizes, lines, spacing toggles and declared defaults', () => {
  const block = render({ type: 'heading', level: 2, lineSpacing: 2.4, spaceBefore: true, spaceAfter: true, content });
  assert.equal(block.className, 'wps-line-24 wps-space-after wps-space-before');
  const rules = cssRules(renderer).join('\n').replace(/:\s+/g, ':').replace(/\{\s+/g, '{');
  for (let size = 8; size <= 96; size++) assert.match(rules, new RegExp('\\.wps-size-' + size + '[^{]*\\{font-size:' + size + 'px'));
  for (const [line, value] of [['10', '1'], ['12', '1.2'], ['14', '1.4'], ['16', '1.6'], ['20', '2'], ['24', '2.4']]) {
    assert.match(rules, new RegExp('\\.wps-line-' + line + '[^{]*\\{line-height:' + value.replace('.', '\\.') + '[;}]'));
  }
  for (const marker of ['space-before', 'space-after']) assert.match(rules, new RegExp('\\.wps-' + marker + '[^{]*\\{margin-(?:block-start|block-end):var\\(--spacingVerticalS\\)'));
  assert.equal(window.getComputedStyle(block).lineHeight, '2.4');
  for (const [tag, sizeToken, lineToken] of [
    ['h1', 'fontSizeHero800', 'lineHeightHero800'],
    ['h2', 'fontSizeHero700', 'lineHeightHero700'],
    ['h3', 'fontSizeBase600', 'lineHeightBase600'],
    ['h4', 'fontSizeBase500', 'lineHeightBase500'],
    ['p', 'fontSizeBase300', 'lineHeightBase300'],
  ]) {
    const document = tag === 'p' ? { type: 'paragraph', content } : { type: 'heading', level: Number(tag.slice(1)), content };
    const current = render(document);
    // JSDOM cannot resolve CSS variables; native source/compiled browser checks
    // assert the resulting pixel values in both Read and Edit.
    assert.match(rules, new RegExp(' ' + tag + '[^{]*\\{font-size:var\\(--' + sizeToken + '\\)'));
    assert.match(rules, new RegExp(' ' + tag + '[^{]*\\{line-height:var\\(--' + lineToken + '\\)'));
    assert.equal(window.getComputedStyle(current).marginBlockStart, '0');
    assert.equal(window.getComputedStyle(current).marginBlockEnd, '0');
  }
  const explicit = render({ type: 'heading', level: 1, content: [{ type: 'text', text: 'X', marks: { size: 8 } }] });
  assert.equal(window.getComputedStyle(explicit.querySelector('span')).fontSize, '8px');
  assert.doesNotMatch(rules, /wps-size-(?:7|97)(?!\d)[^{]*\{/);
  for (const rule of cssRules(renderer).filter(rule => rule.includes('wps-'))) assert.match(rule, /\.[a-zA-Z0-9_-]+ \.wps-/);
  assert.ok(!host.querySelector('[style]'));
});
test('theme palette updates under light, dark, high contrast and RTL without changing content/classes', () => {
  const themes = require('@fluentui/react-theme');
  const ast = { type: 'paragraph', content: ['primary', 'subtle', 'brand', 'neutralMuted', 'brandAlt', 'red', 'orange', 'yellow', 'green', 'teal', 'blue', 'purple', 'pink', 'brown', 'neutralDark', 'neutralLight', 'neutralLighter', 'brandDarkest', 'brandDarker', 'brandDark', 'brandLight', 'darkRed', 'amber', 'yellowGreen', 'cyan', 'navy'].map(color => ({ type: 'text', text: color, marks: { color } })) };
  let html;
  for (const theme of [themes.webLightTheme, themes.webDarkTheme, themes.teamsHighContrastTheme, { ...themes.webLightTheme, colorPalettePinkForeground2: '#abcd01', colorBrandForeground2: '#1b2c3d' }]) for (const dir of ['ltr', 'rtl']) {
    const block = render(ast, theme, dir);
    html ??= block.outerHTML;
    assert.equal(block.outerHTML, html);
    const provider = host.querySelector('[data-provider]');
    for (const token of ['colorNeutralForeground1', 'colorNeutralForeground2', 'colorBrandForeground1', 'colorNeutralForeground3', 'colorBrandForeground2', 'colorPaletteRedForeground3', 'colorPaletteDarkOrangeForeground3', 'colorPaletteYellowForeground3', 'colorPaletteGreenForeground3', 'colorPaletteTealForeground2', 'colorPaletteBlueForeground2', 'colorPalettePurpleForeground2', 'colorPalettePinkForeground2', 'colorPaletteBrownForeground2']) {
      // Fluent injects theme variables in a provider-scoped rule; JSDOM cannot resolve CSS var colors.
      const matchingRules = [...document.styleSheets].flatMap(sheet => [...sheet.cssRules])
        .filter(rule => rule.selectorText && provider.matches(rule.selectorText));
      const effective = matchingRules.map(rule => rule.style.getPropertyValue('--' + token)).filter(Boolean).pop();
      assert.equal(effective?.trim(), theme[token], token);
    }
  }
  const rules = cssRules(renderer).join('\n').replace(/:\s+/g, ':').replace(/\{\s+/g, '{');
  for (const [role, token] of [['primary','colorNeutralForeground1'],['subtle','colorNeutralForeground2'],['brand','colorBrandForeground1'],['neutralMuted','colorNeutralForeground3'],['brandAlt','colorBrandForeground2'],['red','colorPaletteRedForeground3'],['orange','colorPaletteDarkOrangeForeground3'],['yellow','colorPaletteYellowForeground3'],['green','colorPaletteGreenForeground3'],['teal','colorPaletteTealForeground2'],['blue','colorPaletteBlueForeground2'],['purple','colorPalettePurpleForeground2'],['pink','colorPalettePinkForeground2'],['brown','colorPaletteBrownForeground2']]) assert.match(rules, new RegExp('\\.wps-color-' + role + '[^{]*\\{color:var\\(--' + token + '\\)'));
});
test('only protocol attributes survive hostile imported markup', () => {
  const parsed = codec.parseTitleHTML('<h2 id="unsafe" onclick="bad()" style="color:red"><span class="alien" title="unsafe">Safe</span><script>bad()</script></h2>', document);
  const block = render({ ...parsed.document, id: 'not-forwarded', style: { color: 'red' }, onClick: () => { throw new Error('not-forwarded'); } });
  assert.equal(block.textContent, 'Safe');
  for (const node of [block.parentElement, block, ...block.querySelectorAll('*')]) {
    assert.deepEqual([...node.attributes].map(a => a.name), ['class']);
  }
});
test('invalid AST throws rather than fallback rendering and Read modules never import engine or properties', () => {
  const previousConsoleError = console.error;
  const preventExpected = event => { if (event.error?.message === 'Invalid title document') event.preventDefault(); };
  window.addEventListener('error', preventExpected);
  console.error = () => {};
  try { assert.throws(() => render({ type: 'heading', level: 5, content }), /Invalid title document/); }
  finally { console.error = previousConsoleError; window.removeEventListener('error', preventExpected); }
  assert.equal(render({ type: 'paragraph', content }).textContent, 'Title');
  for (const name of ['web-part-title-renderer.internal.tsx', 'web-part-title-styles.internal.ts']) {
    const source = fs.readFileSync(path.resolve(__dirname, base + name), 'utf8');
    assert.doesNotMatch(source, /@tiptap|prosemirror|dangerouslySetInnerHTML|useSPFxProperties|createDOMRenderer/);
  }
  assert.equal(typeof styles.useWebPartTitleStyles, 'function');
});

test('shared editor wrapper scopes identical markers and repeated rendering reuses host CSS', () => {
  const { RendererProvider } = require('@griffel/react');
  const { FluentProvider } = require('@fluentui/react-provider');
  const { webLightTheme } = require('@fluentui/react-theme');
  function EditorDocumentProbe() {
    const classes = styles.useWebPartTitleStyles();
    return React.createElement('div', { className: classes.root, 'data-editor-wrapper': '' },
      React.createElement('div', {}, React.createElement('p', { className: 'wps-line-16' },
        React.createElement('span', { className: 'wps-size-96 wps-color-subtle' }, 'Draft'))));
  }
  act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer },
    React.createElement(FluentProvider, { theme: webLightTheme }, React.createElement(EditorDocumentProbe))), host); });
  const draft = host.querySelector('[data-editor-wrapper] p');
  assert.equal(window.getComputedStyle(draft).lineHeight, '1.6');
  assert.equal(window.getComputedStyle(draft.firstChild).fontSize, '96px');
  const count = Object.keys(renderer.insertionCache).length;
  render({ type: 'paragraph', content });
  render({ type: 'paragraph', content });
  assert.equal(Object.keys(renderer.insertionCache).length, count);
  const outside = document.createElement('span'); outside.className = 'wps-size-96'; outside.textContent = 'Outside';
  document.body.appendChild(outside);
  assert.notEqual(window.getComputedStyle(outside).fontSize, '96px');
  outside.remove();
});

test('derived title colors use the same foreground-only class resolver as swatches', async () => {
  const format = await harness.loadModule(path.resolve(__dirname, base + 'web-part-title-format.internal.ts'));
  render({ type: 'paragraph', content: [{ type: 'text', text: 'Mixed', marks: { color: 'yellowGreen' } }] });
  const rules = cssRules(renderer).join('\n').replace(/:\s+/g, ':').replace(/\{\s+/g, '{');
  for (const { id } of format.titleColorDescriptors) {
    const source = format.titleFormatRegistry.colors[id];
    const sources = typeof source === 'string' ? [source] : [source.first, source.second];
    for (const token of sources) assert.match(token, /^color.*Foreground/);
    const expected = format.getTitleColorCss(id);
    assert.ok(rules.replace(/\s+/g, '').includes('.wps-color-' + id + '{color:' + expected.replace(/\s+/g, '')), id);
  }
  assert.equal(format.getTitleColorCss('yellowGreen'), 'color-mix(in srgb, var(--colorPaletteYellowForeground3) 73%, var(--colorPaletteGreenForeground3))');
  assert.ok(!host.querySelector('[style]'));
});

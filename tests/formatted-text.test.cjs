const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const entry = path.resolve(__dirname, '../packages/spfx-react-toolkit/src/controls/formatted-text/formatted-text.tsx');
let FormattedText, React, ReactDOM, act, renderer, host;
test.before(async () => {
  assert.ok(fs.existsSync(entry), 'FormattedText implementation is missing');
  const harness = require('./fixtures/formatted-text-harness.cjs');
  harness.createSxDom();
  ({ FormattedText } = await harness.loadModule(entry));
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  renderer = require('@griffel/core').createDOMRenderer(document);
  host = document.createElement('div'); document.body.appendChild(host);
  test.after(async () => { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); await harness.cleanup(); });
});
function render(props) {
  const { RendererProvider } = require('@griffel/react');
  const { FluentProvider } = require('@fluentui/react-provider');
  const { webLightTheme } = require('@fluentui/react-theme');
  act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer },
    React.createElement(FluentProvider, { theme: webLightTheme }, React.createElement(FormattedText, { 'data-root': '', ...props }))), host); });
  return host.querySelector('[data-root]');
}
test('every typography variant preserves span semantics unless as is supplied', () => {
  const { typography } = require('./fixtures/sx-harness.cjs').loadStyles();
  assert.equal(Object.keys(typography).length, 17);
  for (const variant of Object.keys(typography)) {
    assert.equal(render({ variant, children: 'Text' }).tagName, 'SPAN');
    for (const as of ['span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']) {
      assert.equal(render({ variant, as }).tagName, as.toUpperCase());
    }
  }
});
test('native ref, events, attributes and consumer styles reach the root', () => {
  const ref = React.createRef(); let clicks = 0;
  const root = render({ as: 'h2', ref, id: 'native', title: 'Title', 'aria-label': 'Label', dir: 'rtl',
    className: 'consumer', style: { color: 'red' }, onClick: () => clicks++, children: 'Content' });
  assert.equal(ref.current, root); root.click(); assert.equal(clicks, 1);
  assert.equal(root.id, 'native'); assert.equal(root.title, 'Title'); assert.equal(root.dir, 'rtl');
  assert.equal(root.getAttribute('aria-label'), 'Label'); assert.ok(root.classList.contains('consumer'));
  assert.equal(root.style.color, 'red');
});
test('empty content retains its root and raw HTML cannot replace children', () => {
  for (const children of [undefined, null, false, '']) assert.equal(render({ children }).tagName, 'SPAN');
  const root = render({ children: '**safe**', dangerouslySetInnerHTML: { __html: '<img src=x>' } });
  assert.equal(root.textContent, 'safe'); assert.equal(root.querySelector('img'), null);
});
test('Markdown toggle and mark composition preserve owned content', () => {
  let root = render({ highlight: true, subtle: true, children: '**Strong** and `code` [link](./docs)' });
  assert.equal(root.querySelectorAll('mark').length, 1);
  assert.equal(root.querySelector('mark strong').textContent, 'Strong');
  assert.equal(root.querySelector('code').textContent, 'code'); assert.equal(root.querySelector('a').getAttribute('href'), './docs');
  root = render({ markdown: false, children: '**Literal**' });
  assert.equal(root.textContent, '**Literal**'); assert.equal(root.querySelector('strong'), null);
  for (const attribute of ['variant', 'font', 'align', 'markdown', 'highlight', 'subtle', 'maxlines', 'truncate', 'wrap']) {
    assert.equal(root.hasAttribute(attribute), false);
  }
});
test('clamp accepts positive integers on any root and consumer style overrides its count', () => {
  for (const as of ['span', 'p', 'h2']) for (const maxLines of [1, 2, 4]) {
    const root = render({ as, maxLines, wrap: false, truncate: true, children: 'Long text' });
    assert.equal(root.tagName, as.toUpperCase());
    assert.equal(root.hasAttribute('style'), false);
    const rules = require('./fixtures/sx-harness.cjs').activeRules(renderer, root.className).join('\n');
    assert.match(rules, new RegExp(String.raw`-webkit-line-clamp:\s*${maxLines}[;}]`));
  }
  assert.equal(render({ maxLines: 2, style: { WebkitLineClamp: 5 } }).style.WebkitLineClamp, '5');
  for (const maxLines of [0, -1, 1.5, NaN, Infinity, undefined]) {
    assert.equal(render({ maxLines }).hasAttribute('style'), false);
  }
});
test('overflow and combined decorations emit the requested independent CSS', () => {
  const { activeRules } = require('./fixtures/sx-harness.cjs');
  const css = props => activeRules(renderer, render(props).className).join('\n');
  assert.match(css({ truncate: true, wrap: true }), /nowrap/);
  assert.match(css({ truncate: true }), /ellipsis/);
  assert.doesNotMatch(css({ wrap: false }), /(?:overflow-x|overflow-y):[^;}]*hidden/);
  assert.match(css({ underline: true, strikethrough: true }), /underline line-through/);
});

test('line-clamp classes update and reuse inserted CSS without generating inline styles', () => {
  const first = render({ maxLines: 2 }).className;
  const second = render({ maxLines: 3 }).className;
  assert.notEqual(first, second);
  const count = Object.keys(renderer.insertionCache).length;
  const root = render({ maxLines: 2 });
  assert.equal(root.className, first);
  assert.equal(root.hasAttribute('style'), false);
  assert.equal(Object.keys(renderer.insertionCache).length, count);
  assert.equal(render({}).hasAttribute('style'), false);
});

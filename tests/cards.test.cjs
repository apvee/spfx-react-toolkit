const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, ReactDOM, act, dom, renderer, ContentCard, PreviewCard;
test.before(async () => {
  process.env.NODE_ENV = 'test'; dom = harness.createSxDom(); global.Node = dom.window.Node; global.NodeFilter = dom.window.NodeFilter;
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  renderer = require('@griffel/core').createDOMRenderer(document);
  ({ ContentCard } = await harness.loadModule('packages/spfx-react-toolkit/src/controls/content-card/content-card.tsx'));
  ({ PreviewCard } = await harness.loadModule('packages/spfx-react-toolkit/src/controls/preview-card/preview-card.tsx'));
});
test.after(async () => { dom?.window.close(); await harness.cleanup(); });
function mount(element, dark = false, rtl = false) {
  const host = document.createElement('div'); document.body.appendChild(host);
  const { RendererProvider } = require('@griffel/react');
  const { FluentProvider, webLightTheme, webDarkTheme } = require('@fluentui/react-components');
  function render(next = element, nextDark = dark, nextRtl = rtl) {
    element = next; dark = nextDark; rtl = nextRtl;
    act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer }, React.createElement(FluentProvider, {
      theme: dark ? webDarkTheme : webLightTheme, dir: rtl ? 'rtl' : 'ltr' }, element)), host); });
  }
  render(); return { host, render, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
test('both cards own headings and shells while preserving inner refs/events and literal strings', () => {
  for (const Card of [ContentCard, PreviewCard]) {
    const root = React.createRef(), title = React.createRef(), action = React.createRef(); let clicks = 0, calls = 0;
    const f = mount(React.createElement(Card, { ref: root, id: 'card', title: { ref: title, children: '**Literal**', className: 'caller', 'data-title': 'yes' },
      headingLevel: 2, description: '*Description*', headerActions: { ref: action, children: (Tag, props) => { calls++; return React.createElement(Tag, props, React.createElement('button', { onClick: () => clicks++ }, 'Header')); } },
      footer: 'Supplemental', actions: React.createElement('button', {}, 'Footer'), dangerouslySetInnerHTML: { __html: '<b id="unsafe">Unsafe</b>' }, as: 'a', href: '/unsafe' }, '**Body**'));
    try {
      assert.equal(root.current.tagName, 'DIV'); assert.equal(root.current.hasAttribute('href'), false);
      assert.equal(title.current.parentElement.tagName, 'H2'); assert.equal(title.current.textContent, '**Literal**'); assert.ok(title.current.classList.contains('caller'));
      assert.equal(f.host.querySelector('strong'), null); assert.equal(f.host.querySelector('#unsafe'), null); assert.equal(calls, 1);
      assert.equal(action.current.tagName, 'DIV'); act(() => action.current.querySelector('button').click()); assert.equal(clicks, 1);
      assert.equal(f.host.querySelector('[data-card-footer]').textContent, 'Supplemental'); assert.equal(f.host.querySelector('[data-card-actions]').textContent, 'Footer');
    } finally { f.close(); }
  }
});
test('orientation/theme/direction/width changes retain caller controls and default h3', () => {
  const tree = orientation => React.createElement(PreviewCard, { title: 'Title', orientation, preview: 'Preview' }, React.createElement('input', { defaultValue: 'Initial' }));
  const f = mount(tree('vertical'));
  try {
    assert.equal(f.host.querySelector('h3').textContent, 'Title'); const input = f.host.querySelector('input'); input.value = 'Edited';
    f.render(tree('horizontal'), true, true); assert.equal(f.host.querySelector('input'), input); assert.equal(input.value, 'Edited');
    assert.equal(input.closest('[dir]').getAttribute('dir'), 'rtl');
    assert.equal(f.host.querySelector('[data-card-detail]').hasAttribute('style'), false);
  } finally { f.close(); }
});
test('defined preview suppresses generated image and logo; empty props/callback keep region without preflight', () => {
  let calls = 0, logoCalls = 0; const logo = { children: (Tag, props) => { logoCalls++; return React.createElement(Tag, props, 'Logo'); } };
  const tree = preview => React.createElement(PreviewCard, { title: 'Title', image: { src: '/image', alt: 'Image' }, preview, logo });
  for (const preview of [false, null, '', React.createElement(React.Fragment)]) {
    const f = mount(tree(preview)); try { assert.equal(f.host.querySelector('[data-card-preview]'), null); assert.equal(f.host.querySelector('img'), null); } finally { f.close(); }
  }
  assert.equal(logoCalls, 0);
  const f = mount(tree({ children: () => { calls++; return undefined; } }));
  try { assert.ok(f.host.querySelector('[data-card-preview]')); assert.equal(f.host.querySelector('img'), null); assert.equal(calls, 1); assert.equal(logoCalls, 1);
    f.render(tree({})); assert.ok(f.host.querySelector('[data-card-preview]')); } finally { f.close(); }
});
test('generated image failures preserve alt semantics and only a new src resets failure', () => {
  let fallbackCalls = 0; const tree = (src, alt, fallback) => React.createElement(PreviewCard, { title: 'Title', image: { src, alt }, previewFallback: fallback });
  const f = mount(tree('/first', 'Cover'));
  try {
    const image = f.host.querySelector('img'); assert.equal(image.alt, 'Cover'); act(() => { image.dispatchEvent(new window.Event('error')); });
    assert.equal(f.host.querySelector('img'), null); assert.equal(f.host.querySelector('[role="img"]').getAttribute('aria-label'), 'Cover');
    f.render(tree('/first', 'Changed', { children: () => { fallbackCalls++; return undefined; } })); assert.equal(fallbackCalls, 1); assert.equal(f.host.querySelector('img'), null);
    f.render(tree('/second', '', { children: () => { fallbackCalls++; return undefined; } })); assert.ok(f.host.querySelector('img')); assert.equal(fallbackCalls, 1);
    act(() => { f.host.querySelector('img').dispatchEvent(new window.Event('error')); }); assert.equal(f.host.querySelector('[data-card-image-fallback]').getAttribute('aria-hidden'), 'true');
  } finally { f.close(); }
});
test('native Card selection, focus, disabled, checkbox/floatingAction and click parity', () => {
  const { Card } = require('@fluentui/react-components');
  function observe(Control) {
    const changes = [], clicks = []; const ref = React.createRef();
    const f = mount(React.createElement(Control, { ...(Control === Card ? {} : { title: 'Native' }), ref,
      defaultSelected: false, focusMode: 'tab-exit', onSelectionChange: (_, data) => changes.push(data.selected), onClick: () => clicks.push(true),
      checkbox: { 'aria-label': 'Select' }, floatingAction: React.createElement('button', { id: 'floating' }, 'Float') }, React.createElement('button', { id: 'child' }, 'Child')));
    try {
      act(() => ref.current.click()); const selected = ref.current.getAttribute('aria-selected');
      act(() => { f.host.querySelector('#child').dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true })); });
      return { changes, selected, clicks: clicks.length, checkbox: !!f.host.querySelector('input[type=checkbox]'), floating: !!f.host.querySelector('#floating'), tabIndex: ref.current.tabIndex };
    } finally { f.close(); }
  }
  assert.deepEqual(observe(ContentCard), observe(Card)); assert.deepEqual(observe(PreviewCard), observe(Card));
  const f = mount(React.createElement(ContentCard, { title: 'Disabled', disabled: true }, React.createElement('button', {}, 'Still enabled')));
  try { assert.equal(f.host.querySelector('button').disabled, false); } finally { f.close(); }
});
test('logo introduced div/img containers retain inner refs and safe bounded tag', () => {
  const ref = React.createRef(); const f = mount(React.createElement(PreviewCard, { title: 'Logo', preview: {}, logo: { as: 'img', src: '/logo', alt: 'Brand', ref } }));
  try { assert.equal(ref.current.tagName, 'IMG'); assert.equal(ref.current.alt, 'Brand'); assert.ok(ref.current.className); assert.equal(ref.current.parentElement.classList.contains('fui-CardPreview__logo'), true); } finally { f.close(); }
});
test('required titles reject meaningless static content without probing opaque callbacks', () => {
  const { renderToStaticMarkup } = require('react-dom/server');
  for (const Card of [ContentCard, PreviewCard]) {
    for (const title of [undefined, null, false, true, '', '  ', {}, { children: '' }, React.createElement(React.Fragment)]) {
      assert.throws(() => renderToStaticMarkup(React.createElement(Card, { title })), /meaningful phrasing/);
    }
    let calls = 0;
    const f = mount(React.createElement(Card, { title: { children: () => { calls++; return undefined; } } }));
    try { assert.ok(f.host.querySelector('h3')); assert.equal(calls, 1); } finally { f.close(); }
    const named = mount(React.createElement(Card, { title: { 'aria-label': 'Named title' } }));
    try { assert.equal(named.host.querySelector('h3 > span').getAttribute('aria-label'), 'Named title'); } finally { named.close(); }
  }
});
test('guarded slot objects keep fixed heading shell and omit raw HTML', () => {
  const ref = React.createRef(); let clicks = 0;
  const f = mount(React.createElement(ContentCard, { title: { as: 'h1', ref, children: 'Safe title', style: { color: 'red' }, onClick: () => clicks++, dangerouslySetInnerHTML: { __html: '<b id="slot-injected">Unsafe</b>' } }, headingLevel: 4 }));
  try {
    assert.equal(ref.current.tagName, 'SPAN'); assert.equal(ref.current.parentElement.tagName, 'H4'); assert.equal(ref.current.style.color, 'red');
    assert.equal(f.host.querySelector('#slot-injected'), null); act(() => { ref.current.click(); }); assert.equal(clicks, 1);
  } finally { f.close(); }
});
test('native checkbox remains a selectable Card slot when floating action is absent', () => {
  const { Card } = require('@fluentui/react-components');
  function observe(Control) {
    const changes = []; const ref = React.createRef();
    const f = mount(React.createElement(Control, { ...(Control === Card ? {} : { title: 'Selectable' }), checkbox: { ref, 'aria-label': 'Choose card' }, defaultSelected: false, onSelectionChange: (_, data) => changes.push(data.selected) }));
    try {
      assert.equal(ref.current.type, 'checkbox'); assert.equal(ref.current.getAttribute('aria-label'), 'Choose card');
      act(() => { ref.current.click(); }); return { changes, checked: ref.current.checked };
    } finally { f.close(); }
  }
  const native = observe(Card); assert.equal(native.checked, true); assert.ok(native.changes.length > 0);
  assert.deepEqual(observe(ContentCard), native); assert.deepEqual(observe(PreviewCard), native);
});

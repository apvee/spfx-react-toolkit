const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, ReactDOM, act, api, useSx, dom;
test.before(async () => {
  dom = harness.createSxDom();
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  api = await harness.loadModule('packages/spfx-react-toolkit/src/controls/content-slot.internal.tsx');
  ({ useSx } = await harness.loadModule('packages/spfx-react-toolkit/src/hooks/useSx.ts'));
});
test.after(async () => { dom?.window.close(); await harness.cleanup(); });
function mount(value, options = { shorthand: 'preserve' }, tag = 'span') {
  const host = document.createElement('div'); document.body.appendChild(host);
  function View() { return React.createElement('section', {}, api.renderContentSlot(value, tag, useSx(), options)); }
  act(() => { ReactDOM.render(React.createElement(View), host); });
  return { host, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
test('shorthand keeps identity and DOM while wrapping is explicit and booleans are absent', () => {
  const node = React.createElement('input', { defaultValue: 'caller' });
  assert.equal(api.renderContentSlot(node, 'span', () => '', { shorthand: 'preserve' }), node);
  for (const value of [undefined, null, false, true]) {
    const f = mount(value, { shorthand: 'wrap' });
    try { assert.equal(f.host.querySelector('section').childNodes.length, 0); } finally { f.close(); }
  }
  for (const value of ['', 0, React.createElement(React.Fragment)]) {
    const f = mount(value, { shorthand: 'wrap' });
    try { assert.ok(f.host.querySelector('section > span')); } finally { f.close(); }
  }
  const f = mount('**Literal**');
  try { assert.equal(f.host.querySelector('section').innerHTML, '**Literal**'); } finally { f.close(); }
});
test('object props and callbacks preserve refs/classes/style/events and apply owned guards', () => {
  const consumerRef = React.createRef(), ownedRef = React.createRef(); let clicks = 0, received;
  const f = mount({ as: 'h1', ref: consumerRef, className: 'consumer', style: { color: 'red' }, id: 'replace',
    unknownProp: 'remove', href: '/remove', 'aria-labelledby': 'remove', dangerouslySetInnerHTML: { __html: '<img data-injected>' },
    onClick: () => clicks++, children: (Component, props) => { received = props; return React.createElement(Component, props, 'Content'); } },
    { shorthand: 'preserve', className: 'toolkit', defaultProps: { title: 'Default' }, ownedProps: { id: 'owned', ref: ownedRef }, omitProps: ['aria-labelledby'] });
  try {
    const el = f.host.querySelector('section > span');
    assert.ok(el); assert.equal(consumerRef.current, el); assert.equal(ownedRef.current, el);
    assert.equal(el.id, 'owned'); assert.equal(el.title, 'Default');
    assert.ok(el.classList.contains('toolkit')); assert.ok(el.classList.contains('consumer'));
    assert.equal(el.style.color, 'red'); assert.equal(el.hasAttribute('href'), false); assert.equal(el.hasAttribute('unknownprop'), false);
    assert.equal(received['aria-labelledby'], undefined); assert.equal(received.dangerouslySetInnerHTML, undefined); assert.equal(received.as, undefined);
    act(() => el.click()); assert.equal(clicks, 1);
  } finally { f.close(); }
  assert.equal(consumerRef.current, null); assert.equal(ownedRef.current, null);
});
test('presence and content inspection never execute callbacks or components', () => {
  let calls = 0;
  const render = () => { calls++; return undefined; };
  const object = { children: render };
  assert.equal(api.hasContentSlot(object), true);
  assert.equal(api.hasContentSlotContent(object), true);
  assert.equal(api.inspectContentSlot(object).kind, 'render');
  assert.equal(calls, 0);
  assert.equal(api.hasContentSlot({ children: '' }), true);
  assert.equal(api.hasContentSlotContent({ children: '' }), false);
  assert.equal(api.hasContentSlotContent({ 'aria-label': 'Meaningful' }), true);
  assert.equal(api.hasContentSlotContent({ 'aria-label': '  ' }), false);
  assert.equal(api.inspectContentSlot({ children: 'Label' }).children, 'Label');
  for (const value of [undefined, null, false, true, '', React.createElement(React.Fragment), [false, '']]) assert.equal(api.hasContentSlot(value), false);
  for (const value of [0, [0], new Set(['item']), React.createElement(render)]) assert.equal(api.hasContentSlot(value), true);
  assert.equal(calls, 0);
});
test('portals and iterables render as shorthand, explicit empty objects retain containers', () => {
  const portalHost = document.createElement('div'); document.body.appendChild(portalHost);
  const portal = ReactDOM.createPortal(React.createElement('b', {}, 'Portal'), portalHost);
  assert.equal(api.hasContentSlot(portal), true);
  let f = mount(portal);
  try { assert.equal(portalHost.textContent, 'Portal'); assert.equal(f.host.querySelector('section').childNodes.length, 0); } finally { f.close(); portalHost.remove(); }
  f = mount(new Set(['A', 'B']));
  try { assert.equal(f.host.querySelector('section').textContent, 'AB'); } finally { f.close(); }
  f = mount({ children: undefined, 'data-empty': 'yes' });
  try { assert.equal(f.host.querySelector('span').getAttribute('data-empty'), 'yes'); } finally { f.close(); }
});
test('bounded alternative tags retain img props and reject other as values', () => {
  const f = mount({ as: 'img', src: '/logo.png', alt: 'Logo' }, { shorthand: 'preserve', allowedTags: ['img'] }, 'div');
  try { assert.equal(f.host.querySelector('img').alt, 'Logo'); } finally { f.close(); }
  const guarded = mount({ as: 'a', href: '/unsafe', children: 'Logo' }, { shorthand: 'preserve', allowedTags: ['img'] }, 'div');
  try { assert.ok(guarded.host.querySelector('section > div')); assert.equal(guarded.host.querySelector('a'), null); } finally { guarded.close(); }
});

test('container defaults retain a toolkit ref when the consumer also supplies a ref', () => {
  const toolkitRef = React.createRef(), consumerRef = React.createRef();
  const f = mount({ ref: consumerRef, children: 'Refs' }, { shorthand: 'preserve', defaultProps: { ref: toolkitRef } });
  try {
    const element = f.host.querySelector('span');
    assert.equal(consumerRef.current, element);
    assert.ok(toolkitRef.current === element, 'toolkit default ref must be preserved');
  } finally { f.close(); }
});

test('static child mapping preserves formatting ownership and opaque callback execution', () => {
  let maps = 0, calls = 0;
  const mapChildren = children => { maps++; return typeof children === 'string' ? React.createElement('em', {}, children) : children; };
  let f = mount({ children: '**Literal**' }, { shorthand: 'preserve', mapChildren });
  try { assert.equal(f.host.querySelector('span > em').textContent, '**Literal**'); assert.equal(maps, 1); } finally { f.close(); }
  const callback = { children: () => { calls++; return undefined; } };
  assert.equal(api.hasContentSlot(callback), true); assert.equal(calls, 0);
  f = mount(callback, { shorthand: 'preserve', mapChildren });
  try { assert.equal(calls, 1); assert.equal(maps, 1); assert.equal(f.host.querySelector('section').childNodes.length, 0); } finally { f.close(); }
});

test('independent host renderers/themes/directions retain their slot props and class composition', async () => {
  const { ContentSlotHost } = await harness.loadModule('tests/fixtures/content-slot-host.tsx');
  const { FluentProvider, webLightTheme, webDarkTheme } = require('@fluentui/react-components');
  const { RendererProvider } = require('@griffel/react');
  const { createDOMRenderer } = require('@griffel/core');
  const first = createDOMRenderer(document), second = createDOMRenderer(document);
  const host = document.createElement('div'); document.body.appendChild(host);
  function instance(renderer, theme, dir, name) {
    return React.createElement(RendererProvider, { renderer }, React.createElement(FluentProvider, { theme, dir }, React.createElement(ContentSlotHost, { name })));
  }
  try {
    act(() => { ReactDOM.render(React.createElement(React.Fragment, {}, instance(first, webLightTheme, 'ltr', 'Light'), instance(second, webDarkTheme, 'rtl', 'Dark')), host); });
    const light = host.querySelector('[data-slot-host="Light"]'), dark = host.querySelector('[data-slot-host="Dark"]');
    assert.equal(light.closest('[dir]').getAttribute('dir'), 'ltr');
    assert.equal(dark.closest('[dir]').getAttribute('dir'), 'rtl');
    assert.ok(light.classList.contains('consumer')); assert.ok(dark.classList.contains('consumer'));
    assert.equal(light.textContent, 'Light'); assert.equal(dark.textContent, 'Dark');
    assert.ok(Object.keys(first.stylesheets).length > 0); assert.ok(Object.keys(second.stylesheets).length > 0);
    assert.equal(light.hasAttribute('style'), false); assert.equal(dark.hasAttribute('style'), false);
  } finally { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); }
});

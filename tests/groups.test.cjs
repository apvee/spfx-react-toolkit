const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, ReactDOM, act, SettingsGroup, AccordionGroup, FluentProvider, renderer, dom;
test.before(async () => {
  process.env.NODE_ENV = 'test';
  dom = harness.createSxDom();
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  ({ FluentProvider } = require('@fluentui/react-components'));
  renderer = require('@griffel/core').createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  ({ SettingsGroup } = await harness.loadModule('packages/spfx-react-toolkit/src/controls/settings-group/settings-group.tsx'));
  ({ AccordionGroup } = await harness.loadModule('packages/spfx-react-toolkit/src/controls/accordion-group/accordion-group.tsx'));
});
test.after(async () => { dom?.window.close(); await harness.cleanup(); });
function mount(element) {
  const host = document.createElement('div'); document.body.appendChild(host);
  const { RendererProvider } = require('@griffel/react');
  const { webLightTheme } = require('@fluentui/react-components');
  function render(next = element) {
    element = next;
    act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer },
      React.createElement(FluentProvider, { theme: webLightTheme }, element)), host); });
  }
  render();
  return { host, render, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
test('SettingsGroup preserves native refs/events and caller input identity across keyed reorder', () => {
  const ref = React.createRef(); let clicks = 0;
  const tree = reverse => React.createElement(SettingsGroup, { ref, id: 'settings', title: '**Preferences**', description: '*Details*', onClick: () => clicks++ },
    (reverse ? ['b', 'a'] : ['a', 'b']).map(value => React.createElement(SettingsGroup.Item, { key: value, title: value },
      React.createElement('input', { id: value, 'aria-label': value, defaultValue: value }))));
  const f = mount(tree(false));
  try {
    assert.equal(ref.current.id, 'settings'); assert.equal(f.host.querySelector('h2').textContent, '**Preferences**');
    assert.equal(f.host.querySelector('h2 strong'), null);
    const input = f.host.querySelector('#a'); input.value = 'Changed'; act(() => input.click()); assert.equal(clicks, 1);
    f.render(tree(true)); assert.equal(f.host.querySelector('#a'), input); assert.equal(input.value, 'Changed');
    assert.equal(input.hasAttribute('aria-labelledby'), false);
  } finally { f.close(); }
});
test('empty group slots leave no header or text wrapper and JSX remains caller-owned', () => {
  const f = mount(React.createElement(SettingsGroup, { title: '', description: false }, React.createElement(SettingsGroup.Item, {})));
  try {
    assert.equal(f.host.querySelector('[data-group-header]'), null); assert.equal(f.host.querySelector('[data-group-text]'), null);
    f.render(React.createElement(SettingsGroup, { title: React.createElement('em', { id: 'custom' }, '**Literal**'), headingLevel: 4 }));
    assert.equal(f.host.querySelector('#custom').textContent, '**Literal**'); assert.equal(f.host.querySelector('#custom strong'), null);
    assert.equal(f.host.querySelector('h4').firstElementChild, f.host.querySelector('#custom'));
  } finally { f.close(); }
});
test('owned roots reject raw HTML and preserve hidden attributes', () => {
  for (const Group of [SettingsGroup, AccordionGroup]) {
    const f = mount(React.createElement(Group, { hidden: true, id: 'owned', dangerouslySetInnerHTML: { __html: '<img id="injected">' } },
      React.createElement(Group.Item, { value: 'a', hidden: true, id: 'row', dangerouslySetInnerHTML: { __html: '<img id="row-injected">' } }, 'Safe')));
    try {
      assert.equal(f.host.querySelector('#injected'), null); assert.equal(f.host.querySelector('#row-injected'), null);
      assert.equal(f.host.querySelector('#owned').hidden, true); assert.equal(f.host.querySelector('#row').hidden, true);
    } finally { f.close(); }
  }
});
test('AccordionGroup delegates uncontrolled toggle payload and retains a closed input', () => {
  const changes = []; const ref = React.createRef();
  const f = mount(React.createElement(AccordionGroup, { onToggle: (_, data) => changes.push(data), defaultOpenItems: ['a'] },
    React.createElement(AccordionGroup.Item, { ref, value: 'a', title: '**Profile**', description: '*Details*', headingLevel: 5 },
      React.createElement('input', { defaultValue: 'Initial' }))));
  try {
    const button = f.host.querySelector('button'); const input = f.host.querySelector('input'); input.value = 'Changed';
    assert.equal(ref.current.querySelector('h5 button'), button); assert.equal(button.querySelector('[data-group-text]').textContent, '**Profile***Details*');
    assert.equal(button.querySelector('strong'), null); assert.equal(button.querySelector('em'), null);
    const panel = f.host.querySelector('#' + button.getAttribute('aria-controls'));
    assert.ok(panel); assert.equal(panel.hasAttribute('inert'), false);
    act(() => button.click()); assert.deepEqual(changes.at(-1), { value: 'a', openItems: [] });
    assert.equal(panel.getAttribute('inert'), ''); assert.equal(f.host.querySelector('input'), input);
    act(() => button.click()); assert.equal(input.value, 'Changed'); assert.equal(panel.hasAttribute('inert'), false);
  } finally { f.close(); }
});
test('controlled expansion remains authoritative and disabled triggers stay native', () => {
  const requests = [];
  const tree = openItems => React.createElement(AccordionGroup, { openItems, multiple: true, collapsible: false, onToggle: (_, data) => requests.push(data) },
    React.createElement(AccordionGroup.Item, { value: 'a', title: 'A' }, 'A body'),
    React.createElement(AccordionGroup.Item, { value: 2, title: 'B' }, 'B body'),
    React.createElement(AccordionGroup.Item, { value: 'disabled', title: 'Disabled', disabled: true }, 'Disabled body'));
  const f = mount(tree(['a']));
  try {
    const buttons = f.host.querySelectorAll('button'); assert.equal(buttons[0].getAttribute('aria-disabled'), 'true');
    act(() => buttons[1].click()); assert.equal(buttons[1].getAttribute('aria-expanded'), 'false');
    assert.deepEqual(requests.at(-1), { value: 2, openItems: [2, 'a'] });
    f.render(tree(['a', 2])); assert.equal(buttons[1].getAttribute('aria-expanded'), 'true'); assert.equal(buttons[2].disabled, true);
  } finally { f.close(); }
});
test('unmountOnClose defers initially closed descendants and resets after closing', () => {
  let mounts = 0, cleanups = 0;
  function Content() { React.useEffect(() => { mounts++; return () => cleanups++; }, []); return React.createElement('input', { defaultValue: 'Initial' }); }
  const f = mount(React.createElement(AccordionGroup, {}, React.createElement(AccordionGroup.Item, { value: 'a', title: 'A', unmountOnClose: true }, React.createElement(Content))));
  try {
    const button = f.host.querySelector('button'); assert.equal(mounts, 0); assert.equal(f.host.querySelector('input'), null);
    act(() => button.click()); const input = f.host.querySelector('input'); input.value = 'Changed'; assert.equal(mounts, 1);
    act(() => button.click()); assert.equal(f.host.querySelector('input'), null); assert.equal(cleanups, 1);
    act(() => button.click()); assert.notEqual(f.host.querySelector('input'), input); assert.equal(f.host.querySelector('input').value, 'Initial');
  } finally { f.close(); }
});

test('text slots wrap caller JSX in semantic presets without cloning or replacing its DOM', () => {
  const childRef = React.createRef(); let childClicks = 0;
  const title = React.createElement('em', { ref: childRef, className: 'caller-text', onClick: () => childClicks++ }, 'Caller title');
  const tree = level => React.createElement(SettingsGroup, { title, headingLevel: level },
    React.createElement(SettingsGroup.Item, { title: React.createElement('em', { id: 'item-title' }, 'Item'),
      description: React.createElement('span', { id: 'item-description' }, 'Description') }));
  const f = mount(tree(2));
  try {
    const ownedChild = childRef.current;
    assert.equal(ownedChild.parentElement.tagName, 'H2'); assert.equal(ownedChild.className, 'caller-text');
    assert.ok(ownedChild.parentElement.className); assert.equal(ownedChild.hasAttribute('style'), false);
    assert.equal(f.host.querySelector('#item-title').parentElement.tagName, 'SPAN');
    assert.equal(f.host.querySelector('#item-description').parentElement.tagName, 'SPAN');
    act(() => ownedChild.click()); assert.equal(childClicks, 1);
    f.render(tree(2)); assert.equal(childRef.current, ownedChild);
  } finally { f.close(); }
});

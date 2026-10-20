const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, ReactDOM, act, DescriptionList, FluentProvider, renderer, dom;
test.before(async () => {
  process.env.NODE_ENV = 'test';
  dom = harness.createSxDom();
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  ({ FluentProvider } = require('@fluentui/react-components'));
  renderer = require('@griffel/core').createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  const entry = 'packages/spfx-react-toolkit/src/controls/description-list/description-list.tsx';
  assert.ok(fs.existsSync(path.resolve(__dirname, '..', entry)), 'DescriptionList production control must exist');
  ({ DescriptionList } = await harness.loadModule(entry));
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

test('data pairs preserve native dl/dt/dd semantics, zero and literal text', () => {
  const ref = React.createRef(); let clicks = 0;
  const f = mount(React.createElement(DescriptionList, { ref, id: 'metadata', 'aria-label': 'Record', title: 'Native title',
    onClick: () => clicks++, items: [{ id: 'zero', label: 0, value: 0 }, { id: 'literal', label: '**Label**', value: '*Value*' }] }));
  try {
    const list = f.host.querySelector('dl');
    assert.equal(ref.current, list); assert.equal(list.title, 'Native title');
    assert.equal(list.getAttribute('aria-label'), 'Record');
    assert.deepEqual([...list.querySelectorAll('dt')].map(n => n.textContent), ['0', '**Label**']);
    assert.deepEqual([...list.querySelectorAll('dd')].map(n => n.textContent), ['0', '*Value*']);
    assert.equal(list.querySelector('strong,em'), null);
    act(() => list.click()); assert.equal(clicks, 1);
    for (const row of list.children) assert.deepEqual([...row.children].map(n => n.tagName), ['DT', 'DD']);
    assert.equal(list.querySelector('[style]'), null);
  } finally { f.close(); }
});

test('empty data or absent children leave no list while a declared empty fragment remains composable', () => {
  const f = mount(React.createElement(DescriptionList, { items: [] }));
  try {
    assert.equal(f.host.querySelector('dl'), null);
    for (const children of [undefined, false, true, null]) {
      f.render(React.createElement(DescriptionList, {}, children)); assert.equal(f.host.querySelector('dl'), null);
    }
    f.render(React.createElement(DescriptionList, {}, React.createElement(React.Fragment)));
    assert.ok(f.host.querySelector('dl')); assert.equal(f.host.querySelector('dt'), null);
  } finally { f.close(); }
});

test('slots retain wrappers, callback props, refs and invariant native cells', () => {
  const labelRef = React.createRef(); let callbackValue, callbackClicks = 0;
  const itemRef = React.createRef();
  const f = mount(React.createElement(DescriptionList, {}, React.createElement(DescriptionList.Item,
    { ref: itemRef, id: 'pair', label: { as: 'div', ref: labelRef, className: 'caller-label', title: 'Slot title', children: 'Owner' } },
    { className: 'caller-value', ref: node => { callbackValue = node; }, onClick: () => callbackClicks++,
      children: (Component, props) => React.createElement(Component, { ...props, 'data-custom-render': '' }, React.createElement('em', {}, 'Ada')) })));
  try {
    const row = f.host.querySelector('#pair'); assert.equal(itemRef.current, row);
    assert.equal(labelRef.current.parentElement.tagName, 'DT'); assert.equal(labelRef.current.title, 'Slot title');
    assert.ok(labelRef.current.classList.contains('caller-label'));
    assert.equal(callbackValue.parentElement.tagName, 'DD'); assert.ok(callbackValue.classList.contains('caller-value'));
    assert.ok(callbackValue.hasAttribute('data-custom-render')); assert.equal(callbackValue.querySelector('em').textContent, 'Ada');
    act(() => callbackValue.click()); assert.equal(callbackClicks, 1);
    assert.deepEqual([...row.children].map(n => n.tagName), ['DT', 'DD']);
  } finally { f.close(); }
  assert.equal(labelRef.current, null); assert.equal(callbackValue, null);
});

test('owned raw HTML pathways are removed and hidden native roots preserve mounted content', () => {
  const raw = { __html: '<img data-injected="">' };
  const f = mount(React.createElement(DescriptionList, { hidden: true, dangerouslySetInnerHTML: raw },
    React.createElement(DescriptionList.Item, { hidden: true, dangerouslySetInnerHTML: raw,
      label: { dangerouslySetInnerHTML: raw, children: 'Safe label' } }, { dangerouslySetInnerHTML: raw, children: 'Safe value' })));
  try {
    assert.equal(f.host.querySelector('[data-injected]'), null);
    assert.equal(f.host.querySelector('dl').hidden, true); assert.equal(f.host.querySelector('dl > div').hidden, true);
    assert.equal(f.host.querySelector('dt').textContent, 'Safe label'); assert.equal(f.host.querySelector('dd').textContent, 'Safe value');
  } finally { f.close(); }
});

test('keyed data reordering and layout changes retain caller controls and independent lists', () => {
  const input = id => React.createElement('input', { id, defaultValue: id, 'aria-label': id });
  const entries = [{ id: 'a', label: 'A', value: input('a') }, { id: 'b', label: 'B', value: input('b') }];
  const tree = (layout, reverse) => React.createElement(React.Fragment, {},
    React.createElement(DescriptionList, { id: 'first', layout, items: reverse ? [...entries].reverse() : entries }),
    React.createElement(DescriptionList, { id: 'second', layout: 'vertical' }, React.createElement(DescriptionList.Item, { label: 'Independent' }, input('c'))));
  const f = mount(tree('vertical', false));
  try {
    const a = f.host.querySelector('#a'), c = f.host.querySelector('#c'); a.value = 'Changed'; c.value = 'Other';
    const secondClass = f.host.querySelector('#second').className;
    f.render(tree('horizontal', true));
    assert.equal(f.host.querySelector('#a'), a); assert.equal(a.value, 'Changed'); assert.equal(f.host.querySelector('#c'), c);
    assert.equal(c.value, 'Other'); assert.equal(f.host.querySelector('#second').className, secondClass);
    assert.deepEqual([...f.host.querySelector('#first').querySelectorAll('dt')].map(n => n.textContent), ['B', 'A']);
  } finally { f.close(); }
});

test('compound fragments, JSX and undefined or legacy null definitions keep semantic cells', () => {
  const childRef = React.createRef();
  const f = mount(React.createElement(DescriptionList, {}, React.createElement(React.Fragment, {},
    React.createElement(DescriptionList.Item, { label: React.createElement('strong', { ref: childRef, className: 'owned-child' }, 'Rich') }, undefined),
    React.createElement(DescriptionList.Item, { label: 'Legacy' }, null))));
  try {
    assert.equal(childRef.current.className, 'owned-child'); assert.equal(childRef.current.parentElement.parentElement.tagName, 'DT');
    assert.deepEqual([...f.host.querySelectorAll('dd')].map(n => n.textContent), ['', '']);
  } finally { f.close(); }
});

test('an orphan Item fails with the missing-parent contract', () => {
  const host = document.createElement('div'); document.body.appendChild(host);
  const original = console.error; console.error = () => {};
  try {
    assert.throws(() => act(() => { ReactDOM.render(React.createElement(DescriptionList.Item, { label: 'Orphan' }, 'Value'), host); }), /inside DescriptionList/);
  } finally { console.error = original; act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); }
});

test('slot guards keep divs fixed, filter unsupported props and normalize conditions', () => {
  const f = mount(React.createElement(DescriptionList, { items: [
    { id: 'unsafe', label: { as: 'h1', href: '/unsupported', unknownSlotProp: 'remove', children: 'Safe term' }, value: { as: 'button', children: 'Safe value' } },
    { id: 'false', label: 'Conditional', value: false },
    { id: 'true', label: 'Conditional true', value: true },
    { id: 'empty', label: 'Explicit empty', value: { children: undefined, 'data-empty-container': '' } },
  ] }));
  try {
    assert.ok(!f.host.querySelector('h1,button'), 'content tags cannot replace fixed divs');
    const term = f.host.querySelector('dt > div');
    assert.equal(term.textContent, 'Safe term');
    assert.equal(term.hasAttribute('href'), false);
    assert.equal(term.hasAttribute('unknownslotprop'), false);
    const definitions = [...f.host.querySelectorAll('dd')];
    assert.equal(definitions[1].children.length, 0, 'false creates no content container');
    assert.equal(definitions[2].children.length, 0, 'true creates no content container');
    assert.ok(definitions[3].querySelector('div[data-empty-container]'));
  } finally { f.close(); }
});

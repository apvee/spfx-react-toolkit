const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const { Simulate } = require('react-dom/test-utils');
const fluent = require('@fluentui/react-components');
const { createDOMRenderer } = require('@griffel/core');
const { RendererProvider } = require('@griffel/react');
window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
const cases = [
  ['provider-webpart.tsx', 'SPFxWebPartProvider', 'Compare WebPart provider availability with the current host', 'WebPart'],
  ['provider-application-customizer.tsx', 'SPFxApplicationCustomizerProvider', 'Compare Application Customizer provider availability with the current host', 'AppCustomizer'],
  ['provider-field-customizer.tsx', 'SPFxFieldCustomizerProvider', 'Compare Field Customizer provider availability with the current host', 'FieldCustomizer'],
  ['provider-listview-commandset.tsx', 'SPFxListViewCommandSetProvider', 'Compare Command Set provider availability with the current host', 'CommandSet'],
];
function fixture(t, name, title) {
  let kind = 'WebPart', revision = 7, regions, props = {}, invocations = 0;
  const hookCalls = { identity: 0, page: 0, properties: 0 };
  const toolkit = { useSPFxContext: () => ({ kind }), useSPFxInstanceInfo: () => { hookCalls.identity++; return { kind, id: 'actual-extension' }; }, useSPFxPageContext: () => { hookCalls.page++; return { web: { title: 'Actual web' } }; }, useSPFxProperties: () => { hookCalls.properties++; return { properties: props }; } };
  for (const [, owner] of cases) toolkit[owner] = () => { invocations++; throw Error('No provider may be mounted by selection'); };
  const h = createHarness({ '@apvee/spfx-react-toolkit': toolkit });
  const { Demo } = h.load(scenario(title, `packages/spfx-react-toolkit/src/core/${name}#${cases.find(x=>x[0]===name)[1]}`).file);
  let card = true, signal;
  const { DemoCard } = h.load(path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard.tsx'));
  function Root() { return h.React.createElement(RendererProvider, { renderer: createDOMRenderer(null) }, h.React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme }, h.React.createElement(Demo, { resetRevision: revision, applicationCustomizerTop: signal, ...(card ? { render: r => { regions = r; return h.React.createElement(DemoCard, { title, iconName: 'Info', ...r }); } } : {}) }))); }
  const view = h.mountComponent(Root);
  const render = () => h.act(() => { require('react-dom').render(h.React.createElement(Root), view.container); });
  t.after(() => { view.unmount(); h.close(); });
  return { h, view, toolkit, render, hookCalls, get regions() { return regions; }, get invocations() { return invocations; }, kind(value) { kind = value; render(); }, reset() { revision++; render(); }, signal(value) { signal = value; render(); }, properties(value) { props = value; render(); }, standalone() { card = false; render(); }, select(value) { h.act(() => Simulate.change(view.container.querySelector('select'), { target: { value } })); } };
}
for (const [file, owner, title, baselineKind] of cases) test(`${owner} owns all availability choices, its baseline and reset lifetime`, t => {
  const f = fixture(t, file, title), select = f.view.container.querySelector('select');
  assert.equal(select.value, owner, 'initial revision 7 is not a reset');
  assert.deepEqual([...select.options].map(x => x.value), cases.map(x => x[1]));
  const disclosure = [...f.view.container.querySelectorAll('button')].find(x => x.textContent.includes('Advanced diagnostics'));
  assert.equal(disclosure.getAttribute('aria-expanded'), 'false');
  for (const [, choice, , required] of cases) {
    f.select(choice); f.kind(required);
    assert.match(f.view.container.textContent, /Host matchYes/);
    assert.match(f.view.container.textContent, /Export availableYes/);
    f.toolkit[choice] = false; f.render();
    assert.match(f.view.container.textContent, /Export availableNo/);
    assert.match(f.view.container.textContent, /Host matchYes/);
    f.kind('ACE'); assert.match(f.view.container.textContent, /Host matchNo: requires another host/);
    f.toolkit[choice] = () => { throw Error('Must not mount'); };
  }
  const oldChange = f.regions.controls.props.children.props.onChange;
  f.reset(); assert.equal(f.view.container.querySelector('select'), select); assert.equal(select.value, owner);
  f.h.act(() => oldChange(undefined, { value: cases.find(x=>x[1]!==owner)[1] })); assert.equal(select.value, owner, 'captured old change is obsolete');
  f.kind(baselineKind); assert.match(f.view.container.textContent, /Host matchYes/);
  f.standalone(); assert.equal(f.view.container.querySelector('select').value, owner);
  f.h.act(() => f.view.container.querySelector('button').click()); assert.match(f.view.container.textContent, /Selection does not mount/);
  assert.equal(f.view.container.querySelectorAll('[data-providers-demo="result"]').length, 1);
  assert.equal(f.invocations, 0);
});
test('actual Top observation checks the prerequisite before hooks, kind and exact nullish text semantics', t => {
  const f = fixture(t, 'provider-application-customizer.tsx', 'Inspect the actual Top placeholder provider');
  assert.match(f.view.container.textContent, /native Top placeholder/);
  assert.deepEqual(f.hookCalls, { identity: 0, page: 0, properties: 0 }, 'no-port wrapper invokes no operational hooks');
  f.signal({ placeholder: 'Top' }); assert.equal(f.hookCalls.page, 0); assert.equal(f.hookCalls.properties, 0); assert.match(f.view.container.textContent, /requires AppCustomizer/); assert.doesNotMatch(f.view.container.textContent, /Provider active/);
  f.kind('AppCustomizer'); assert.match(f.view.container.textContent, /Provider active: AppCustomizer actual-extension \| Actual web \| No message/);
  // Last two inputs are legacy null/malformed false robustness only; the source contract remains required string testMessage.
  for (const [properties, message] of [[undefined, 'No message'], [{}, 'No message'], [{ testMessage: undefined }, 'No message'], [{ testMessage: '' }, ''], [{ testMessage: 'Test message' }, 'Test message'], [{ testMessage: null }, 'No message'], [{ testMessage: false }, 'false']]) {
    f.properties(properties); assert.equal(f.view.container.querySelector('[data-providers-demo="native-top"]').textContent, `SPFxApplicationCustomizerProvider active: AppCustomizer actual-extension | Actual web | ${message}`);
    f.reset(); assert.equal(f.view.container.querySelector('[data-providers-demo="native-top"]').textContent.endsWith(`| ${message}`), true);
  }
  assert.equal(f.invocations, 0);
});

const test = require('node:test');
const assert = require('node:assert/strict');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const title = 'Select native site references';
const owner = 'packages/spfx-react-toolkit/src/property-pane/pickers/property-pane-site-picker.ts#PropertyPaneSitePicker';
test('native Site exact executable availability with truthful staging metadata', () => {
  const selected = scenario(title, owner);
  assert.deepEqual(selected.entry.metadata, { host: 'webPart', kind: 'read', contentPlacement: 'configuration', hostAdapter: 'propertyPane', observe: "Prepare the native Site picker, open this WebPart's actual property pane, select site references and inspect the saved values; selection and Reset do not open the pane or query." });
});

const { createNativeSiteFixture } = require('./fixtures/executable-examples/native-site-scenario.cjs');
for (const compiled of [false, true]) {
  const mode = compiled ? 'COMPILED' : 'SOURCE';
  test(`${mode} Site missing native port reports exact staging and exposes no enabled prepare, factory or query`, async t => {
    const f = createNativeSiteFixture(t, { compiled, connected: false }); await f.flush();
    assert.match(f.container.textContent, /initialized native scenario host is not connected/);
    assert.match(f.container.textContent, /existing full integrated configuration independently/);
    assert.equal(f.container.querySelectorAll('button').length, 0); assert.equal(f.fields.length, 0); assert.equal(f.calls.length, 0);
    f.setRevision(8); await f.flush(); assert.equal(f.fields.length, 0); assert.equal(f.calls.length, 0);
  });
  test(`${mode} Site actual public factory/provider subject mounts only in native render and observes live SDK replacements`, async t => {
    const f = createNativeSiteFixture(t, { compiled }); await f.flush();
    const original = f.action('Prepare native scenario');
    assert.equal(f.fields.length, 0); assert.equal(f.calls.length, 0);
    f.click('Prepare native scenario'); assert.equal(f.fields.length, 0); assert.equal(f.calls.length, 0);
    const configuration = f.getter(), group = configuration.pages[0].groups[0];
    assert.equal(configuration.pages.length, 1); assert.equal(group.isGroupNameHidden, true); assert.equal(Object.hasOwn(group, 'groupName'), false); assert.equal(Object.hasOwn(group, 'isCollapsed'), false); assert.equal(group.groupFields.length, 1);
    assert.equal(f.fields.length, 1); assert.equal(f.fields[0].args[0], 'pickerSites'); assert.equal(f.fields[0].args[1].key, 'picker-sites'); assert.equal(f.fields[0].args[1].instance, f.instance); assert.deepEqual(f.fields[0].args[1].value, []);
    assert.equal(f.calls.length, 0); const field = group.groupFields[0]; f.mount(field); await f.flush();
    const trigger = f.nativeDom.querySelector('.fui-Dropdown__button'); assert.ok(trigger); assert.equal(f.handlers.size, 1);
    const references = [{ kind: 'site', url: 'https://tenant.test/sites/root', text: 'Current collection' }];
    f.bag = { pickerSites: references }; f.click('Inspect saved site references'); assert.deepEqual(JSON.parse(f.container.querySelector('pre').textContent), references);
    f.setRevision(8); f.h.act(() => original()); await f.flush(); assert.equal(f.nativeDom.querySelector('.fui-Dropdown__button'), trigger); assert.equal(f.adapter.candidate?.generation !== undefined, true);
    f.setRevision(9); f.click('Inspect saved site references'); assert.match(f.container.textContent, /Open the prepared native scenario before inspecting/);
    const next = f.getter().pages[0].groups[0].groupFields[0]; f.mount(next); await f.flush();
    assert.deepEqual(f.fields.at(-1).args[1].value, references); assert.equal(f.nativeDom.querySelector('.fui-Dropdown__button'), trigger);
    f.h.act(() => field.properties.onDispose(f.nativeDom)); assert.equal(f.handlers.size, 1);
    f.close(next); assert.equal(f.handlers.size, 0); assert.equal(f.nativeDom.textContent, ''); f.origins();
  });
  test(`${mode} Site repeated Reset invalidates old commands, preserves Demo identity and admits fresh preparation`, async t => {
    const f = createNativeSiteFixture(t, { compiled }); const firstButton = f.container.querySelector('button'), original = f.action('Prepare native scenario');
    f.setRevision(8); f.h.act(() => original()); assert.equal(f.adapter.candidate, undefined);
    f.click('Prepare native scenario'); assert.ok(f.adapter.candidate); f.setRevision(9); assert.equal(f.adapter.candidate, undefined);
    f.setRevision(10);
    f.click('Prepare native scenario'); assert.ok(f.adapter.candidate, 'fresh callback after repeated Reset');
    assert.equal(f.container.querySelector('button'), firstButton); assert.equal(f.fields.length, 0); assert.equal(f.calls.length, 0);
  });
}

const { pinnedComponents } = require('./fixtures/native-property-pane/pinned-components.cjs');
const ReactDOM = require('react-dom');
for (const compiled of [false, true]) test(`${compiled ? 'COMPILED' : 'SOURCE'} real pinned unnamed Site subject replaces persisted collapsed named group and witnesses before native parent Rendered`, async t => {
  const f = createNativeSiteFixture(t, { compiled }), { PropertyPane } = pinnedComponents(), trace = [];
  const nativeRoot = document.createElement('div'); document.body.appendChild(nativeRoot);
  let configuration = { hideTitleBar: true, pages: [{ displayGroupsAsAccordion: true, groups: [{ groupName: 'Old selection', groupFields: [] }] }] };
  const data = () => ({ webPartId: 'same-native-instance', title: 'Native Site', configuration, properties: f.bag, isReactive: true, onRendered: () => {
    trace.push(f.adapter.witnesses.size ? 'child-before-parent' : 'parent-without-child'); f.adapter.rendered();
  }, onPropertyPaneFieldChanged: (name, value) => { f.bag = { ...f.bag, [name]: value }; f.adapter.fieldChanged(name, undefined, value); }, onConfigurationEvent: () => {} });
  function render() { f.h.act(() => { ReactDOM.render(f.h.React.createElement(PropertyPane, { data: data() }), nativeRoot); }); }
  t.after(() => { f.h.act(() => { ReactDOM.unmountComponentAtNode(nativeRoot); }); nativeRoot.remove(); });
  render(); const oldHeader = [...nativeRoot.querySelectorAll('button')].find(node => node.textContent === 'Old selection'); assert.ok(oldHeader);
  f.h.act(() => oldHeader.click()); assert.equal(oldHeader.getAttribute('aria-expanded'), 'false');
  f.click('Prepare native scenario'); const actual = f.getter(); configuration = { ...actual, hideTitleBar: true, pages: [{ ...actual.pages[0], displayGroupsAsAccordion: true, groups: [...actual.pages[0].groups, { groupName: 'Other settings', groupFields: [] }] }] };
  render(); await f.flush();
  assert.ok(nativeRoot.querySelector('.fui-Dropdown__button'), 'actual public native subject mounts under actual pinned group/custom host');
  assert.equal(trace.at(-1), 'child-before-parent'); assert.equal(nativeRoot.querySelectorAll('.ms-CustomFieldHost').length, 1);
  assert.deepEqual([...nativeRoot.querySelectorAll('button[aria-expanded]')].filter(node => !node.closest('.ms-CustomFieldHost')).map(node => node.textContent), ['Other settings']);
  const business = f.adapter.nativeView.controller, other = [...nativeRoot.querySelectorAll('button')].find(node => node.textContent === 'Other settings');
  f.h.act(() => other.click()); assert.equal(other.getAttribute('aria-expanded'), 'false'); assert.equal(f.adapter.nativeView.controller, business);
  const trigger = nativeRoot.querySelector('.fui-Dropdown__button'); f.h.act(() => trigger.click()); await f.flush();
  const current = [...document.querySelectorAll('[role="option"]')].find(node => node.textContent.includes('Current collection'));
  assert.ok(current, 'real discovered current-site option'); await f.h.act(async () => { current.click(); await f.saved; }); await f.flush();
  assert.deepEqual(f.bag.pickerSites, [{ kind: 'site', url: 'https://tenant.test/sites/root', siteId: '22222222-2222-4222-8222-000000000003' }]);
  f.click('Inspect saved site references'); assert.deepEqual(JSON.parse(f.container.querySelector('pre').textContent), f.bag.pickerSites); f.origins();
  assert.deepEqual(f.dataRoles.map(row => row.id), ['ApveeSpfxReactToolkitPickersStrings']);
  t.diagnostic(JSON.stringify({ actualPrivateAdapter: f.adapterFile, actualSiteNativeData: f.dataRoles }));
});

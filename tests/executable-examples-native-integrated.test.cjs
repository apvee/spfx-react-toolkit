const test = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const owner = 'packages/spfx-react-toolkit/src/property-pane/pickers/property-pane-site-picker.ts#PropertyPaneSitePicker';
test('integrated native case has one source owner and observable saved compatibility command', () => {
  const selected = scenario('Configure native metadata and file cascade', owner);
  assert.deepEqual(selected.entry.publicNames, ['PropertyPaneSitePicker']);
  assert.equal(selected.entry.metadata.hostAdapter, 'propertyPane');
  assert.match(selected.source, /Inspect saved compatibility mapping/);
  assert.doesNotMatch(selected.source, /components\/shared|from ['"]\./);
});
test('actual WebPart connects one terminal instance host and forwards every genuine native hook', () => {
  const source = fs.readFileSync(path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/SpFxReactToolkitTestWebPart.ts'), 'utf8');
  assert.equal((source.match(/new NativePropertyPaneDemoHost/g) ?? []).length, 1);
  for (const hook of ['fieldChanged', 'configurationStart', 'rendered', 'configurationComplete', 'changesApplied', 'dispose']) assert.match(source, new RegExp('nativeHost\\.' + hook + '\\('));
  assert.doesNotMatch(source, /PropertyPaneSitePicker|PropertyPaneFileFolderPicker|pickerCascadeIdentity|_pickerSiteRevision/);
});
const { loadIntegratedBusiness } = require('./fixtures/executable-examples/integrated-native-business.cjs');
const allFields = part => part.getPropertyPaneConfiguration().pages[0].groups.flatMap(group => group.groupFields);
for (const compiled of [false, true]) {
  const mode = compiled ? 'COMPILED' : 'SOURCE';
  test(`${mode} integrated live saved compatibility false empty undefined and source policy matrix is read-only`, () => {
    const part = loadIntegratedBusiness({ compiled });
    const defined = ['fileFolderTitle', 'fileFolderLayout', 'fileFolderGraph', 'fileFolderNavigationVisible', 'fileFolderInitialView', 'fileFolderNavigationPreset', 'fileFolderSiteUrl', 'fileFolderFromLink', 'fileFolderDeviceUpload', 'fileFolderDeviceLabel', 'fileFolderMyFiles', 'fileFolderRecent', 'fileFolderShared', 'fileFolderQuickAccess', 'fileFolderOrganization', 'fileFolderSite'];
    const inspect = settings => { part.properties = settings; const before = JSON.stringify(settings), refreshes = part.refreshes; const result = part.inspectCompatibility(); assert.equal(JSON.stringify(part.properties), before); assert.equal(part.refreshes, refreshes); assert.deepEqual(result.effective, { preset: settings.fileFolderPreset === 'extended' ? 'extended' : 'focused' }); return result; };
    for (const settings of [{}, Object.fromEntries(defined.map(key => [key, undefined])), { fileFolderUpload: true, fileFolderCreateFolder: true }]) {
      const result = inspect(settings); assert.equal(result.hasLegacySettings, false); assert.equal(result.savedPreset, 'focused');
    }
    for (const [key, value] of [['fileFolderGraph', false], ['fileFolderNavigationVisible', false], ['fileFolderTitle', ''], ['fileFolderDeviceLabel', ''], ...defined.slice(10).map(key => [key, false])]) {
      const result = inspect({ [key]: value }); assert.equal(result.hasLegacySettings, true, key); assert.equal(result.savedPreset, 'legacy');
    }
    assert.equal(inspect({ fileFolderPreset: 'legacy' }).compatibility.visibleLocations, undefined);
    assert.deepEqual(inspect({ fileFolderMyFiles: false }).compatibility.visibleLocations, []);
    assert.deepEqual(inspect(Object.fromEntries(defined.slice(10).map(key => [key, false]))).compatibility.visibleLocations, []);
    assert.deepEqual(inspect({ fileFolderSite: false, fileFolderScope: 'oneDrive' }).compatibility.visibleLocations, ['myFiles']);
    assert.deepEqual(inspect({ fileFolderSite: false }).compatibility.visibleLocations, []);
    for (const kind of ['siteCollection', 'library', 'folder', 'oneDrive']) {
      const settings = { fileFolderScope: kind, fileFolderMyFiles: true, fileFolderQuickAccess: true, fileFolderSite: true, fileFolderInitialView: 'recent', fileFolderNavigationPreset: 'oneDrive', fileFolderSiteUrl: ' https://tenant.sharepoint.com/sites/hr ', fileFolderGraph: false, fileFolderNavigationVisible: false, fileFolderLayout: 'tiles', fileFolderTitle: '', fileFolderDeviceLabel: '' };
      const compatibility = inspect(settings).compatibility;
      assert.deepEqual(compatibility.visibleLocations, kind === 'oneDrive' ? ['myFiles', 'site'] : ['quickAccess', 'site']);
      assert.equal(compatibility.initialView, kind === 'oneDrive' ? 'recent' : 'location');
      assert.equal(compatibility.navigationPreset, 'oneDrive'); assert.equal(compatibility.siteNavigationUrl, 'https://tenant.sharepoint.com/sites/hr');
      assert.equal(compatibility.graphEnabled, false); assert.equal(compatibility.navigationVisible, false); assert.equal(compatibility.layout, 'tiles'); assert.equal(compatibility.title, undefined); assert.equal(compatibility.localFilesLabel, undefined);
      assert.equal(inspect({ ...settings, fileFolderFolder: 'Policies' }).compatibility.initialView, 'location');
      assert.equal(inspect({ ...settings, fileFolderInitialFolder: ' ' }).compatibility.initialView, 'location');
    }
    for (const preset of ['focused', 'extended']) assert.deepEqual(inspect({ fileFolderPreset: preset, fileFolderGraph: false }).compatibility, { preset });
    for (const preset of ['auto', undefined, 'oneDrive', 'current-site']) assert.equal(inspect({ fileFolderPreset: 'legacy', fileFolderNavigationPreset: preset }).compatibility.navigationPreset, ['oneDrive', 'current-site'].includes(preset) ? preset : undefined);
    for (const upload of [undefined, false, true]) for (const device of [undefined, false, true]) assert.equal(inspect({ fileFolderPreset: 'legacy', fileFolderUpload: upload, fileFolderDeviceUpload: device, fileFolderDeviceLabel: 'Device' }).compatibility.localFilesMode, upload && device ? 'upload' : 'disabled');
    part.properties = { fileFolderGraph: false, fileFolderDeviceLabel: 'new live bag' }; assert.equal(part.inspectCompatibility().compatibility.localFilesLabel, 'new live bag'); assert.deepEqual(part.inspectSaved(), part.properties);
  });
  test(`${mode} integrated full subject keys policy generic composition and all command ABA identity`, () => {
    const part = loadIntegratedBusiness({ compiled }), fixture = part.fixture;
    const original = ['Prepare native scenario', 'Inspect saved JSON', 'Inspect saved compatibility mapping'].map(fixture.action);
    const first = part.getPropertyPaneConfiguration(), group = first.pages[0].groups[0];
    assert.equal(group.isGroupNameHidden, true); assert.equal(Object.hasOwn(group, 'groupName'), false); assert.equal(Object.hasOwn(group, 'isCollapsed'), false); assert.equal(group.groupFields.length, 1); assert.equal(group.groupFields[0].targetProperty, 'pickerSites');
    assert.equal(allFields(part).filter(field => field.targetProperty === 'showTitle').length, 1); assert.equal(allFields(part).filter(field => field.targetProperty === 'description').length, 1);
    assert.equal(allFields(part).find(field => field.targetProperty === 'showTitle').properties.checked, true);
    assert.match(first.pages[0].header.description, /Localized host instructions.*Configure metadata/);
    assert.deepEqual(allFields(part).find(field => field.targetProperty === 'pickerPrincipals').properties.principalTypes, ['user', 'sharePointGroup', 'microsoft365Group', 'securityGroup', 'distributionList']);
    assert.deepEqual(allFields(part).find(field => field.targetProperty === 'pickerPrincipals').properties.userIdentity, { mode: 'available' });
    const controller = part.nativeHost.nativeView.controller;
    fixture.demo.render({ resetRevision: 3 }); fixture.demo.render({ resetRevision: 7 });
    fixture.h.act(() => original.forEach(command => command()));
    assert.equal(part.nativeHost.nativeView.controller, controller); assert.match(fixture.regions.status.props.children, /Local state reset/);
    fixture.h.act(() => fixture.action('Inspect saved compatibility mapping')()); assert.match(fixture.regions.status.props.children, /Open the prepared native scenario/);
    part.open = false; part.onPropertyPaneConfigurationComplete();
    const closed = part.getPropertyPaneConfiguration(); assert.deepEqual(closed.pages[0].groups.flatMap(group => group.groupFields).map(field => field.targetProperty), ['showTitle', 'description']);
  });
}
const { createNativeReferenceFixture } = require('./fixtures/executable-examples/native-site-scenario.cjs');
for (const compiled of [false, true]) test(`${compiled ? 'COMPILED' : 'SOURCE'} integrated actual production instance pinned builtins callbacks native subjects and finite transports`, { timeout: 30000 }, async t => {
  const f = createNativeReferenceFixture(t, { subject: 'Integrated', compiled });
  assert.equal(f.fields.length, 0); assert.equal(f.calls.length, 0);
  const initialButton = f.container.querySelector('button');
  f.click('Inspect saved compatibility mapping'); assert.match(f.container.textContent, /Open the prepared native scenario/); assert.equal(f.fields.length, 0);
  f.click('Prepare native scenario'); assert.equal(f.fields.length, 0); assert.equal(f.calls.length, 0);
  let configuration = f.getter();
  assert.equal(configuration.pages[0].groups[0].groupFields[0].targetProperty, 'pickerSites');
  assert.deepEqual(f.fields.map(field => field.name), ['Site', 'List', 'Principal', 'FileFolder']);
  f.mountPane(configuration); await f.flush();
  const siteRoot = f.nativeDom.querySelector('.fui-Dropdown__button'); assert.ok(siteRoot);
  f.open(); await f.flush();
  // Allow the real dropdown's scheduled focus transfer before delivering an ordinary selection.
  await f.h.act(async () => { await new Promise(resolve => window.requestAnimationFrame(resolve)); });
  assert.equal(document.activeElement === document.querySelector('.fui-SearchBox input'), true, 'actual dropdown focus reached search input before selection');
  const savedSite = f.nextSaved(); f.choose('Current collection');
  await f.h.act(async () => { await savedSite; }); await f.flush();
  assert.deepEqual(f.bag.pickerSites, [{ kind: 'site', url: 'https://tenant.sharepoint.com/sites/hr', siteId: '22222222-2222-4222-8222-222222222222' }]);
  assert.deepEqual(f.bag.pickerLists, []); assert.deepEqual(f.bag.fileFolderValues, []);
  assert.deepEqual(f.nativeCallbacks[0], { name: 'Site', values: ['pickerSites', f.bag.pickerSites, true] }); assert.equal(f.changes[0].name, 'pickerSites');
  const owned = f.adapter.nativeView.controller;
  const dispatches = [], changed = f.instance.onPropertyPaneFieldChanged.bind(f.instance);
  f.instance.onPropertyPaneFieldChanged = (name, old, value) => { assert.equal(f.bag[name], value, 'real SDK mutated live bag before protected forward'); dispatches.push({ name, old, value }); changed(name, old, value); };
  const refresh = f.refreshes, automatic = f.automaticRefreshes;
  f.builtin('Presentation for all pickers', 'tagPicker');
  assert.equal(f.bag.pickerPresentation, 'tagPicker'); assert.equal(dispatches.at(-1).value, 'tagPicker'); assert.equal(f.refreshes, refresh + 1); assert.equal(f.automaticRefreshes, automatic + 1);
  f.bag = { ...f.bag, pickerMultiselect: true, fileFolderValues: [], fileFolderGraph: false, fileFolderDeviceLabel: 'live replacement' };
  configuration = f.getter(); f.mountPane(configuration); await f.flush();
  f.builtin('Multiple selection for all pickers', false); assert.equal(f.bag.pickerMultiselect, false); assert.equal(dispatches.at(-1).value, false);
  // Open the genuine remaining named options accordion for its actual text field.
  const options = [...f.nativeDom.querySelectorAll('button')].find(node => node.textContent === 'File picker options'); assert.ok(options); f.h.act(() => options.click());
  f.builtin('Field description (hint)', ''); assert.equal(f.bag.fileFolderHint, ''); assert.equal(dispatches.at(-1).value, '');
  f.bag = { ...f.bag, fileFolderExtensions: '.pdf', fileFolderInitialWeb: 'https://tenant.sharepoint.com/sites/hr', fileFolderInitialLibrary: 'Documents', fileFolderInitialFolder: 'Start' };
  configuration = f.getter(); f.mountPane(configuration); await f.flush();
  f.builtin('File category', 'images'); assert.equal(f.bag.fileFolderExtensions, undefined);
  f.builtin('Selectable content', 'folders'); assert.equal(f.bag.fileFolderCategory, undefined);
  f.builtin('Selectable content', 'files'); f.builtin('Preset', 'extended'); assert.equal(f.bag.fileFolderPreset, 'extended'); f.builtin('Preset', 'focused');
  f.builtin('File picker source', 'folder'); assert.equal(f.bag.fileFolderInitialFolder, undefined); assert.deepEqual(f.bag.fileFolderValues, []);
  configuration = f.getter(); f.mountPane(configuration); await f.flush();
  f.builtin('Existing folder path', 'Policies'); assert.equal(f.bag.fileFolderFolder, 'Policies');
  f.builtin('File picker source', 'siteCollection'); configuration = f.getter(); f.mountPane(configuration); await f.flush();

  f.builtin('Show title', false); assert.equal(f.bag.showTitle, false);
  const beforeGeneric = f.refreshes; f.builtin('Description', 'Current description'); assert.equal(f.bag.description, 'Current description'); assert.equal(f.refreshes, beforeGeneric);
  f.bag = { ...f.bag, fileFolderPreset: undefined };
  f.click('Inspect saved compatibility mapping'); assert.match(f.container.querySelector('[aria-label="Saved compatibility mapping"]').textContent, /live replacement/);
  f.click('Inspect saved JSON'); assert.equal(JSON.parse(f.container.querySelector('[aria-label="Saved native JSON"]').textContent).pickerMultiselect, false);
  assert.equal(f.adapter.nativeView.controller, owned); assert.equal(f.container.querySelector('button'), initialButton);
  f.setRevision(3); f.setRevision(7); assert.equal(f.adapter.nativeView.controller, owned);
  f.closePane(); assert.equal(f.handlers.size, 0); assert.equal(f.nativeDom.textContent, '');
  const observations = f.instance.nativeObservations.read();
  assert.ok(observations.some(event => event.event === 'getter' && event.open)); assert.ok(observations.some(event => event.event === 'field-dispose' && event.open)); assert.equal(observations.at(-1).event, 'complete'); assert.equal(observations.at(-1).open, false);
  f.click('Prepare native scenario'); configuration = f.getter(); f.mountPane(configuration); await f.flush();
  assert.equal(f.fields.at(-1).args[1].valueRevision, 0, 'new fully retired session starts a fresh baseline');
  f.browse(); await f.flush(); assert.equal(f.hosted.posts.length, 1, JSON.stringify({ text: f.nativeDom.textContent, calls: f.calls, unexpected: f.unexpected })); const opening = f.hosted.opening(); opening.initialize();
  const saved = f.nextSaved(), confirmed = opening.result('pick'); opening.command('pick', { command: 'pick', items: [opening.raw('FileA')] }); await f.h.act(async () => { const result = await confirmed; assert.equal(result.data.result, 'success', JSON.stringify({ result, calls: f.calls, unexpected: f.unexpected })); await saved; }); await f.flush();
  assert.equal(f.bag.fileFolderValues[0].itemId, 'FileA');
  f.closePane(); f.origins();
  t.diagnostic(JSON.stringify({ mode: compiled ? 'Compiled' : 'Source', implementationFiles: f.origins(), data: f.dataRoles, attempts: f.calls, messages: f.hosted.messages, nativeCallbacks: f.nativeCallbacks, observations }));
});
for (const compiled of [false, true]) test(`${compiled ? 'COMPILED' : 'SOURCE'} integrated all eight actual factories dependent parents and native cascade revisions`, { timeout: 30000 }, async t => {
  const f = createNativeReferenceFixture(t, { subject: 'Integrated', compiled });
  const ReactDOM = require('react-dom');
  f.h.act(() => { ReactDOM.render(f.h.React.createElement(f.pinned.PropertyPane, { data: { webPartId: f.instance.context.instanceId, title: 'Previous selection', properties: {}, isReactive: true,
    configuration: { hideTitleBar: true, pages: [{ displayGroupsAsAccordion: true, groups: [{ groupName: 'Old selection', groupFields: [] }] }] }, onRendered() {}, onConfigurationEvent() {} } }), f.nativeDom); });
  const oldHeader = [...f.nativeDom.querySelectorAll('button')].find(node => node.textContent === 'Old selection'); assert.ok(oldHeader); f.h.act(() => oldHeader.click()); assert.equal(oldHeader.getAttribute('aria-expanded'), 'false');

  const site = { kind: 'site', url: 'https://tenant.sharepoint.com/sites/HR/Subweb', text: 'Known site' };
  const list = { kind: 'list', webUrl: site.url, listId: '22222222-2222-4222-8222-000000000010', text: 'List 00' };
  f.bag = { pickerSites: [site], pickerLists: [list], pickerContentTypeScope: 'list' };
  f.click('Prepare native scenario'); let configuration = f.getter();
  assert.deepEqual(f.fields.map(field => field.name), ['Site', 'List', 'Field', 'View', 'ListItem', 'ContentType', 'Principal', 'FileFolder']);
  assert.deepEqual(f.fields.map(field => field.args[1].key), ['picker-sites', 'picker-lists', 'picker-fields', 'picker-views', 'picker-items', 'picker-content-types', 'picker-principals', 'file-folder-values']);
  for (const name of ['Field', 'View', 'ListItem', 'ContentType']) { const props = f.fields.find(field => field.name === name).args[1]; assert.equal(props.webUrl, site.url); assert.deepEqual(props.list, { kind: 'id', id: list.listId }); }
  f.mountPane(configuration); await f.flush(); assert.equal(f.nativeDom.querySelectorAll('.ms-CustomFieldHost').length, 8);
  assert.equal([...f.nativeDom.querySelectorAll('button')].some(node => node.textContent === 'Old selection'), false);
  const referenceHeader = [...f.nativeDom.querySelectorAll('button')].find(node => node.textContent === 'Picker references (JSON)'); assert.ok(referenceHeader); f.h.act(() => referenceHeader.click()); assert.equal(referenceHeader.getAttribute('aria-expanded'), 'false');
  assert.ok(f.nativeDom.querySelector('.fui-Dropdown__button'), 'unnamed real Site stays rendered above independently collapsed references'); f.h.act(() => referenceHeader.click());
  const trigger = f.nativeDom.querySelector('.fui-Dropdown__button'), first = configuration.pages[0].groups[0].groupFields[0];
  f.builtin('Content type scope', 'site'); assert.deepEqual(f.bag.pickerContentTypes, []);
  configuration = f.getter(); f.mountPane(configuration); await f.flush();
  assert.equal(f.nativeDom.querySelector('.fui-Dropdown__button') === trigger, true); assert.equal(f.fields.filter(field => field.name === 'ContentType').at(-1).args[1].valueRevision, 1);
  f.h.act(() => first.properties.onDispose(f.nativeDom)); assert.ok(f.nativeDom.querySelector('.fui-Dropdown__button'));
  const refreshes = f.refreshes;
  // A real current native Site callback is SDK-authoritative and clears every dependent array.
  const current = f.fields.filter(field => field.name === 'Site').at(-1).actual;
  const root = f.adapter.mounted.keys().next().value;
  const replacement = { kind: 'site', url: 'https://tenant.sharepoint.com/sites/other', text: 'Other' };
  f.bag = { ...f.bag, pickerSites: [replacement], pickerFields: [{ kind: 'field', webUrl: site.url, fieldId: '22222222-2222-4222-8222-000000000100', internalName: 'Column0' }], pickerViews: [], pickerItems: [], pickerContentTypes: [], pickerPrincipals: [], fileFolderValues: [] };
  f.instance.onPropertyPaneFieldChanged('pickerSites', [site], [replacement]);
  assert.equal(f.refreshes, refreshes + 1); for (const key of ['pickerLists', 'pickerFields', 'pickerViews', 'pickerItems', 'pickerContentTypes', 'fileFolderValues']) assert.deepEqual(f.bag[key], []);
  configuration = f.getter();
  assert.equal(f.fields.filter(field => field.name === 'List').at(-1).args[1].valueRevision, 1); assert.equal(f.fields.filter(field => field.name === 'FileFolder').at(-1).args[1].valueRevision, 1);
  assert.equal(current.properties.key, 'picker-sites'); assert.ok(root);
  f.closePane(); assert.equal(f.handlers.size, 0); f.origins();
});
for (const compiled of [false, true]) test(`${compiled ? 'COMPILED' : 'SOURCE'} two real production instances generic settings terminal callbacks and isolated bounded diagnostics`, () => {
  const a = loadIntegratedBusiness({ compiled }), b = loadIntegratedBusiness({ compiled });
  assert.notEqual(a.nativeHost, b.nativeHost); assert.notEqual(a.nativeHost.port, b.nativeHost.port);
  a.render(); b.render();
  for (const part of [a, b]) {
    const rendered = part.fixture.content[0]; assert.equal(rendered.event, 'render'); assert.equal(rendered.dom, part.domElement);
    assert.equal(rendered.element.props.value.port, part.nativeHost.port); assert.equal(rendered.element.props.children.props.instance, part);
  }
  const aOwner = a.nativeHost.nativeView.controller, bOwner = b.nativeHost.nativeView.controller;
  a.onPropertyPaneConfigurationComplete(); a.onPropertyPaneConfigurationStart(); assert.equal(a.nativeHost.nativeView.controller, aOwner); assert.equal(b.nativeHost.nativeView.controller, bOwner);
  const bRefreshes = b.refreshes; a.properties = { pickerPresentation: 'tagPicker', showTitle: false, description: 'Live generic' };
  a.onPropertyPaneFieldChanged('pickerPresentation', 'combobox', 'tagPicker'); assert.equal(a.refreshes, 1); assert.equal(b.refreshes, bRefreshes);
  const configured = allFields(a); assert.equal(configured.find(field => field.targetProperty === 'showTitle').properties.checked, false);
  a.onPropertyPaneFieldChanged('showTitle', true, false); a.onPropertyPaneFieldChanged('description', '', 'Live generic'); assert.equal(a.refreshes, 1);
  for (let i = 0; i < 80; i++) a.getPropertyPaneConfiguration();
  const trace = a.nativeObservations.read(); assert.equal(trace.length, 64); assert.ok(trace[0].ordinal > 1); assert.ok(trace.every(event => Object.keys(event).every(key => ['ordinal', 'event', 'open', 'fieldKey'].includes(key))));
  const otherTrace = b.nativeObservations.read(); assert.notDeepEqual(trace, otherTrace); a.nativeObservations.clear(); assert.deepEqual(a.nativeObservations.read(), []); assert.deepEqual(b.nativeObservations.read(), otherTrace);
  const ordinal = trace.at(-1).ordinal; a.getPropertyPaneConfiguration(); assert.ok(a.nativeObservations.read()[0].ordinal > ordinal);
  a.onDispose(); assert.equal(a.fixture.content.at(-1).terminal, true); assert.equal(a.nativeHost.environment, undefined); assert.equal(a.nativeHost.candidate, undefined); assert.equal(a.nativeHost.nativeView, undefined); assert.equal(a.nativeHost.produced, undefined);
  a.onPropertyPaneConfigurationStart(); a.onPropertyPaneRendered(); a.onPropertyPaneConfigurationComplete(); a.onAfterPropertyPaneChangesApplied(); a.onPropertyPaneFieldChanged('pickerSites', [], []);
  assert.equal(a.nativeHost.port.prepare(() => assert.fail('terminal factory cannot run')).status, 'blocked');
  assert.deepEqual(allFields(a).map(field => field.targetProperty), ['showTitle', 'description']); assert.equal(b.nativeHost.nativeView.controller, bOwner);
});
for (const compiled of [false, true]) test(`${compiled ? 'COMPILED' : 'SOURCE'} integrated outside-site rejection and latest native disposal cancel pending hosted publication`, { timeout: 30000 }, async t => {
  const f = createNativeReferenceFixture(t, { subject: 'Integrated', compiled }); f.click('Prepare native scenario'); f.mountPane(f.getter()); await f.flush();
  f.browse(); await f.flush(); const opening = f.hosted.opening(); opening.initialize();
  f.hosted.items.FileA.sharepointIds.siteId = '33333333-3333-4333-8333-333333333333';
  const rejected = opening.result('outside'); opening.command('outside', { command: 'pick', items: [opening.raw('FileA')] });
  await f.h.act(async () => { const result = await rejected; assert.equal(result.data.result, 'error'); assert.equal(result.data.error.code, 'unusableItem'); });
  assert.equal(f.bag.fileFolderValues, undefined); assert.equal(f.changes.length, 0);
  const latest = f.adapter.produced.configuration;
  const fields = latest.pages[0].groups.flatMap(group => group.groupFields).filter(field => field.type === 1);
  f.closePane(); assert.equal(opening.port.closed, true); const saved = JSON.stringify(f.bag), calls = f.calls.length;
  for (const field of fields) f.h.act(() => field.properties.onRender(f.nativeDom, undefined, () => assert.fail('retired native callback')));
  opening.command('late', { command: 'pick', items: [opening.raw('FileA')] }); await f.flush();
  assert.equal(JSON.stringify(f.bag), saved); assert.equal(f.changes.length, 0); assert.equal(f.calls.length, calls + 1, 'only explicit test command recorded; disposed channel performs no request'); assert.equal(f.handlers.size, 0); f.origins();
});
for (const compiled of [false, true]) test(`${compiled ? 'COMPILED' : 'SOURCE'} explicit host diagnostics publish scalar lifecycle facts into DOM without subscriptions or native effects`, () => {
  const part = loadIntegratedBusiness({ compiled }), { h } = part.fixture;
  const module = h.load(path.resolve(__dirname, '../apps/spfx-react-toolkit-test', compiled ? 'lib' : 'src', 'webparts/spFxReactToolkitTest/components/shared/NativePropertyPaneContext.' + (compiled ? 'js' : 'tsx')));
  const view = h.mountComponent(() => h.React.createElement(module.NativePropertyPaneContext.Provider, { value: part.nativeContext }, h.React.createElement(module.NativePropertyPaneDiagnostics)));
  const read = () => view.container.querySelector('pre').textContent, click = label => h.act(() => [...view.container.querySelectorAll('button')].find(button => button.textContent === label).click());
  assert.match(read(), /Inspect actual getter/); const refreshes = part.refreshes;
  click('Inspect native host lifecycle'); const first = read(); assert.ok(JSON.parse(first).some(event => event.event === 'getter'));
  part.getPropertyPaneConfiguration(); assert.equal(read(), first, 'no subscription publishes during native getter/commit');
  click('Inspect native host lifecycle'); assert.ok(JSON.parse(read()).length > JSON.parse(first).length);
  click('Clear native host observations'); assert.deepEqual(part.nativeObservations.read(), []); assert.match(read(), /External post-getter SDK or React failure requires reload/); assert.equal(part.refreshes, refreshes);
});
for (const compiled of [false, true]) test(`${compiled ? 'COMPILED' : 'SOURCE'} integrated pending hosted result is inert after latest native close and resource retirement`, { timeout: 30000 }, async t => {
  const f = createNativeReferenceFixture(t, { subject: 'Integrated', compiled, hold: request => request.url?.includes('/items/FileA') });
  f.click('Prepare native scenario'); const configuration = f.getter(); f.mountPane(configuration); await f.flush();
  f.browse(); await f.flush(); const opening = f.hosted.opening(); opening.initialize(); opening.command('pending', { command: 'pick', items: [opening.raw('FileA')] }); await f.flush();
  assert.equal(f.pending.length, 1); const before = JSON.stringify(f.bag); f.closePane(); assert.equal(opening.port.closed, true); assert.equal(f.adapter.mounted.size, 0); assert.equal(f.handlers.size, 0);
  await f.settle(f.pending[0]); await f.flush(); assert.equal(JSON.stringify(f.bag), before); assert.equal(f.changes.length, 0); assert.equal(f.adapter.nativeView, undefined); assert.equal(f.adapter.produced, undefined); f.origins();
});

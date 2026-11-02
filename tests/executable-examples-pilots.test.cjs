const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const { pilotModules, button } = require('./fixtures/executable-examples/pilot-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const ReactDOM = require('react-dom');
function fixture(t, title, initialRevision = 7) {
  // These source hooks import the SPFx enum but generic pilots never use them.
  const h = createHarness({ '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 } } });
  t.after(h.close);
  const modules = pilotModules(h);
  const { Demo, entry } = modules.scenario(title);
  let revision = initialRevision;
  const render = regions => h.React.createElement('div', { 'data-placement': entry.metadata.contentPlacement },
    regions.controls, regions.commands, regions.status, regions.children, regions.diagnostics);
  function Root() { return h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Demo, { resetRevision: revision, render })); }
  const view = h.mountComponent(Root);
  function rerender() { h.act(() => { ReactDOM.render(h.React.createElement(Root), view.container); }); }
  return { ...h, ...view, modules, entry, rerender, reset() { revision++; rerender(); },
    click(text) { h.act(() => Simulate.click(button(view.container, text))); } };
}
const ids = h => [...h.container.querySelectorAll('[data-order-id]')].map(row => Number(row.dataset.orderId));

test('real catalogue groups uniquely cover source-owned scenarios and preserve every independently specified accepted pilot owner', t => {
  const h = fixture(t, 'Browse controlled rows');
  const { executableExamples: entries, executableGroups: groups } = h.modules.catalogue;
  assert.ok(entries.length > 0); assert.ok(groups.length > 0);
  assert.ok(groups.every(group => group.key.trim().length > 0 && group.features.length > 0));
  assert.equal(new Set(groups.map(group => group.key)).size, groups.length, 'group keys are unique');
  const features = groups.flatMap(group => group.features);
  assert.ok(features.every(feature => feature.key.trim().length > 0 && feature.scenarios.length > 0));
  assert.equal(new Set(features.map(feature => feature.key)).size, features.length, 'feature keys are globally unique');
  assert.equal(new Set(entries.map(entry => entry.id)).size, entries.length, 'global scenario ids are unique');
  assert.equal(new Set(entries.map(entry => JSON.stringify([entry.ownerKey, entry.title]))).size, entries.length, 'canonical owner-title pairs are unique');
  const groupedEntries = features.flatMap(feature => feature.scenarios);
  assert.equal(new Set(groupedEntries.map(entry => entry.id)).size, groupedEntries.length, 'each scenario appears in one feature');
  assert.deepEqual(groupedEntries.map(entry => entry.id).sort(), entries.map(entry => entry.id).sort(), 'groups cover exactly the actual catalogue');
  const byId = new Map(entries.map(entry => [entry.id, entry]));
  for (const group of groups) for (const feature of group.features) for (const entry of feature.scenarios) {
    assert.equal(entry, byId.get(entry.id), 'grouped scenarios retain the complete canonical entry');
    assert.equal(entry.groupKey, group.key);
    assert.equal(entry.featureKey, feature.key);
    assert.equal(entry.featureTitle, feature.title);
    assert.equal(entry.ownerKey, feature.key);
  }
  // Independent accepted pilot cases remain mandatory as other families migrate.
  const { pilotTitles, filterPilotGroups } = require('./fixtures/demo-presentation/jsdoc-pilots.cjs');
  const directServiceOwners = {
    createSPFxPnPContextService: 'packages/spfx-react-toolkit/src/services/spfx-pnp-context.service.ts#createSPFxPnPContextService',
    createSPFxPnPService: 'packages/spfx-react-toolkit/src/services/spfx-pnp.service.ts#createSPFxPnPService',
    createSPFxPnPListService: 'packages/spfx-react-toolkit/src/services/spfx-pnp-list.service.ts#createSPFxPnPListService',
  };
  const searchOwners = {
    useSPFxPnPSearch: 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPSearch.ts#useSPFxPnPSearch',
    SearchVerticals: 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPSearch.ts#SearchVerticals',
  };
  const trimOwners = { EntraGroupTrim: 'packages/spfx-react-toolkit/src/controls/entra-group-trim/entra-group-trim.tsx#EntraGroupTrim', SharePointGroupTrim: 'packages/spfx-react-toolkit/src/controls/sharepoint-group-trim/sharepoint-group-trim.tsx#SharePointGroupTrim', SharePointPermissionTrim: 'packages/spfx-react-toolkit/src/controls/sharepoint-permission-trim/sharepoint-permission-trim.tsx#SharePointPermissionTrim' };
  const pickerOwners = {
    FileFolderPicker: 'packages/spfx-react-toolkit/src/controls/file-folder-picker/file-folder-picker.tsx#FileFolderPicker',
    SitePicker: 'packages/spfx-react-toolkit/src/controls/pickers/site-picker.tsx#SitePicker',
    ListPicker: 'packages/spfx-react-toolkit/src/controls/pickers/list-picker.tsx#ListPicker',
    FieldPicker: 'packages/spfx-react-toolkit/src/controls/pickers/field-picker.tsx#FieldPicker',
    ContentTypePicker: 'packages/spfx-react-toolkit/src/controls/pickers/content-type-picker.tsx#ContentTypePicker',
    ViewPicker: 'packages/spfx-react-toolkit/src/controls/pickers/view-picker.tsx#ViewPicker',
    ListItemPicker: 'packages/spfx-react-toolkit/src/controls/pickers/list-item-picker.tsx#ListItemPicker',
    PrincipalPicker: 'packages/spfx-react-toolkit/src/controls/pickers/principal-picker.tsx#PrincipalPicker',
  };
  const pilotOwners = {
    ...pickerOwners,
    ...trimOwners,
    ...searchOwners,
    ...directServiceOwners,
    useSPFxPnPListById: 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPListById.ts#useSPFxPnPListById',
    useSPFxPnPListByUrl: 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPListByUrl.ts#useSPFxPnPListByUrl',
    useSPFxPnPListByPath: 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPListByPath.ts#useSPFxPnPListByPath',

    useSPFxPnP: 'packages/spfx-react-toolkit/src/hooks/useSPFxPnP.ts#useSPFxPnP',
    useSPFxPnPContext: 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPContext.ts#useSPFxPnPContext',
    useSPFxConfigurationStore: 'packages/spfx-react-toolkit/src/hooks/useSPFxConfigurationStore.ts#useSPFxConfigurationStore',
    createSPFxConfigurationStoreService: 'packages/spfx-react-toolkit/src/services/spfx-configuration-store.service.ts#createSPFxConfigurationStoreService',
    useSPFxTenantKeyValueStore: 'packages/spfx-react-toolkit/src/hooks/useSPFxTenantKeyValueStore.ts#useSPFxTenantKeyValueStore',
    createSPFxTenantKeyValueStoreService: 'packages/spfx-react-toolkit/src/services/spfx-tenant-key-value-store.service.ts#createSPFxTenantKeyValueStoreService',
    useSPFxTenantProperty: 'packages/spfx-react-toolkit/src/hooks/useSPFxTenantProperty.ts#useSPFxTenantProperty',
    createSPFxTenantPropertyService: 'packages/spfx-react-toolkit/src/services/spfx-tenant-property.service.ts#createSPFxTenantPropertyService',
    useSPFxSiteKeyValueStore: 'packages/spfx-react-toolkit/src/hooks/useSPFxSiteKeyValueStore.ts#useSPFxSiteKeyValueStore',
    createSPFxSiteKeyValueStoreService: 'packages/spfx-react-toolkit/src/services/spfx-site-key-value-store.service.ts#createSPFxSiteKeyValueStoreService',
    useSPFxUserPhoto: 'packages/spfx-react-toolkit/src/hooks/useSPFxUserPhoto.ts#useSPFxUserPhoto',
    useSPFxOneDriveAppData: 'packages/spfx-react-toolkit/src/hooks/useSPFxOneDriveAppData.ts#useSPFxOneDriveAppData',
    WebPartShell: 'packages/spfx-react-toolkit/src/controls/web-part-shell/web-part-shell.tsx#WebPartShell',
    Pagination: 'packages/spfx-react-toolkit/src/controls/pagination/pagination.tsx#Pagination',
    useStableCallback: 'packages/spfx-react-toolkit/src/hooks/useStableCallback.ts#useStableCallback',
    useSx: 'packages/spfx-react-toolkit/src/hooks/useSx.ts#useSx',
    useSPFxPnPList: 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPList.ts#useSPFxPnPList',
    // The accepted Context cohort owns 21 APIs, including the two cross-site cases.
    useSPFxContext: 'packages/spfx-react-toolkit/src/core/context.internal.tsx#useSPFxContext',
    useSPFxPageContext: 'packages/spfx-react-toolkit/src/hooks/useSPFxPageContext.ts#useSPFxPageContext',
    useSPFxInstanceInfo: 'packages/spfx-react-toolkit/src/hooks/useSPFxInstanceInfo.ts#useSPFxInstanceInfo',
    useSPFxDisplayMode: 'packages/spfx-react-toolkit/src/hooks/useSPFxDisplayMode.ts#useSPFxDisplayMode',
    useSPFxIsEdit: 'packages/spfx-react-toolkit/src/hooks/useSPFxDisplayMode.ts#useSPFxIsEdit',
    useSPFxEnvironmentInfo: 'packages/spfx-react-toolkit/src/hooks/useSPFxEnvironmentInfo.ts#useSPFxEnvironmentInfo',
    useSPFxUserInfo: 'packages/spfx-react-toolkit/src/hooks/useSPFxUserInfo.ts#useSPFxUserInfo',
    useSPFxSiteInfo: 'packages/spfx-react-toolkit/src/hooks/useSPFxSiteInfo.ts#useSPFxSiteInfo',
    useSPFxListInfo: 'packages/spfx-react-toolkit/src/hooks/useSPFxListInfo.ts#useSPFxListInfo',
    useSPFxLocaleInfo: 'packages/spfx-react-toolkit/src/hooks/useSPFxLocaleInfo.ts#useSPFxLocaleInfo',
    useSPFxHubSiteInfo: 'packages/spfx-react-toolkit/src/hooks/useSPFxHubSiteInfo.ts#useSPFxHubSiteInfo',
    useSPFxTeams: 'packages/spfx-react-toolkit/src/hooks/useSPFxTeams.ts#useSPFxTeams',
    useSPFxPageType: 'packages/spfx-react-toolkit/src/hooks/useSPFxPageType.ts#useSPFxPageType',
    useSPFxCorrelationInfo: 'packages/spfx-react-toolkit/src/hooks/useSPFxCorrelationInfo.ts#useSPFxCorrelationInfo',
    useSPFxPermissions: 'packages/spfx-react-toolkit/src/hooks/useSPFxPermissions.ts#useSPFxPermissions',
    useSPFxCrossSitePermissions: 'packages/spfx-react-toolkit/src/hooks/useSPFxCrossSitePermissions.ts#useSPFxCrossSitePermissions',
    useSPFxContainerInfo: 'packages/spfx-react-toolkit/src/hooks/useSPFxContainerInfo.ts#useSPFxContainerInfo',
    useSPFxContainerSize: 'packages/spfx-react-toolkit/src/hooks/useSPFxContainerSize.ts#useSPFxContainerSize',
    useSPFxThemeInfo: 'packages/spfx-react-toolkit/src/hooks/useSPFxThemeInfo.ts#useSPFxThemeInfo',
    useSPFxFluent9ThemeInfo: 'packages/spfx-react-toolkit/src/hooks/useSPFxFluent9ThemeInfo.ts#useSPFxFluent9ThemeInfo',
    useSPFxServiceScope: 'packages/spfx-react-toolkit/src/hooks/useSPFxServiceScope.ts#useSPFxServiceScope',
    // Six distinct Clients owners remain mandatory beside the original five pilots.
    useSPFxHttpClient: 'packages/spfx-react-toolkit/src/hooks/useSPFxHttpClient.ts#useSPFxHttpClient',
    useSPFxSPHttpClient: 'packages/spfx-react-toolkit/src/hooks/useSPFxSPHttpClient.ts#useSPFxSPHttpClient',
    useSPFxMSGraphClient: 'packages/spfx-react-toolkit/src/hooks/useSPFxMSGraphClient.ts#useSPFxMSGraphClient',
    useSPFxAadHttpClient: 'packages/spfx-react-toolkit/src/hooks/useSPFxAadHttpClient.ts#useSPFxAadHttpClient',
    useSPFxAadTokenProvider: 'packages/spfx-react-toolkit/src/hooks/useSPFxAadTokenProvider.ts#useSPFxAadTokenProvider',
    useSPFxApiPermissionPrecheck: 'packages/spfx-react-toolkit/src/hooks/useSPFxApiPermissionPrecheck.ts#useSPFxApiPermissionPrecheck',
  };
  assert.equal(Object.keys(pilotOwners).filter(name => !Object.hasOwn(directServiceOwners, name) && !Object.hasOwn(searchOwners, name) && !Object.hasOwn(pickerOwners, name) && !Object.hasOwn(trimOwners, name)).length, 47, 'three PnP selectors, one PnP operation, one PnP context, two configuration, five original, 21 Context, six Clients, two GraphData two Site two tenant-property and two tenant-list owners');
  assert.equal(Object.keys(pilotOwners).filter(name => !Object.hasOwn(searchOwners, name) && !Object.hasOwn(pickerOwners, name) && !Object.hasOwn(trimOwners, name)).length, 50, 'complete accepted predecessor owner cohort');
  assert.deepEqual(Object.keys(pilotTitles).sort(), Object.keys(pilotOwners).sort(), 'every accepted pilot has an independent canonical owner expectation');
  const pilots = filterPilotGroups(groups);
  const native = pilots.flatMap(group => group.features).filter(feature => feature.scenarios.some(entry => entry.metadata.hostAdapter === 'propertyPane'));
  assert.deepEqual(native.map(feature => feature.title), ['PropertyPaneSitePicker']);
  assert.equal(native.flatMap(feature => feature.scenarios).length, 1);
  const prior = pilots.flatMap(group => group.features).filter(feature => !native.includes(feature));
  assert.deepEqual(prior.map(feature => feature.title).sort(), Object.keys(pilotTitles).sort());
  for (const group of pilots) for (const feature of group.features.filter(feature => prior.includes(feature))) {
    assert.equal(feature.key, pilotOwners[feature.title]);
    assert.equal(group.key, feature.title === 'WebPartShell' || feature.title === 'Pagination' || (Object.hasOwn(pickerOwners, feature.title) || Object.hasOwn(trimOwners, feature.title)) ? 'controls' : ['createSPFxPnPContextService', 'createSPFxPnPService', 'createSPFxPnPListService', 'createSPFxSiteKeyValueStoreService', 'createSPFxTenantPropertyService', 'createSPFxTenantKeyValueStoreService', 'createSPFxConfigurationStoreService'].includes(feature.title) ? 'services' : 'hooks');
    assert.deepEqual(feature.scenarios.map(entry => entry.title).sort(), [...pilotTitles[feature.title]].sort());
    for (const entry of feature.scenarios) {
      assert.equal(entry, byId.get(entry.id), 'filtered scenarios retain their actual canonical catalogue entry');
      const expectedId = require('node:crypto').createHash('sha256').update(JSON.stringify([pilotOwners[feature.title], entry.title])).digest('hex');
      assert.equal(entry.id, expectedId, 'accepted owner/title retains its canonical scenario identity');
    }
  }
});
for (const title of ['Report current count through a retained callback']) {
  test(`${title}: original committed callback observes state; reset retains its identity`, t => {
    const h = fixture(t, title);
    assert.equal(h.entry.metadata.contentPlacement, 'configuration');
    h.click('Increment counter'); h.click('Increment counter'); h.click('Invoke original callback');
    assert.match(h.container.textContent, /Observed counter: 2/);
    assert.match(h.container.textContent, /Current count: 2/);
    assert.match(h.container.textContent, /Same callback identity: yes/);
    h.click('Increment counter'); h.rerender();
    assert.match(h.container.textContent, /Observed counter: 2/, 'observation changes only upon invocation');
    const originalButton = button(h.container, 'Invoke original callback');
    h.reset();
    assert.equal(button(h.container, 'Invoke original callback'), originalButton, 'same mounted Demo');
    assert.match(h.container.textContent, /Current count: 0/);
    assert.match(h.container.textContent, /Observed counter: not invoked/);
    h.click('Invoke original callback');
    assert.match(h.container.textContent, /Observed counter: 0/);
    assert.match(h.container.textContent, /Same callback identity: yes/);
    assert.equal(h.container.querySelector('[data-demo-subject]'), null);
  });
}
test('identity pilot owns real rerenders, original-reference invocation and reset without report state', t => {
  const h = fixture(t, 'Retain callback identity through rerenders');
  assert.equal(h.entry.metadata.contentPlacement, 'configuration');
  h.click('Rerender component'); h.click('Rerender component');
  assert.match(h.container.textContent, /Rerender revision: 2/);
  h.click('Invoke original callback');
  assert.match(h.container.textContent, /Original invocations: 1/);
  assert.match(h.container.textContent, /Same callback identity: yes/);
  const originalButton = button(h.container, 'Invoke original callback');
  h.reset();
  assert.equal(button(h.container, 'Invoke original callback'), originalButton);
  assert.match(h.container.textContent, /Rerender revision: 0/);
  assert.match(h.container.textContent, /Original invocations: 0/);
  h.click('Invoke original callback');
  assert.match(h.container.textContent, /Original invocations: 1/);
  assert.match(h.container.textContent, /Same callback identity: yes/);
  assert.equal([...h.container.querySelectorAll('button')].some(node => node.textContent === 'Report'), false);
  assert.doesNotMatch(h.container.textContent, /Observed counter:|Current count:/);
});
test('controlled known-total pilot slices exact deterministic orders and resets to five', t => {
  const h = fixture(t, 'Browse controlled rows');
  assert.deepEqual(ids(h), [1, 2, 3, 4, 5]);
  assert.match(h.container.textContent, /1–5 of 23/);
  assert.equal(h.container.querySelectorAll('[data-order-id]')[0].textContent.includes('Hardware'), true);
  assert.equal(h.container.querySelectorAll('[data-order-id]')[1].textContent.includes('Software'), true);
  h.click('Next page'); assert.deepEqual(ids(h), [6, 7, 8, 9, 10]);
  h.click('Last page'); assert.deepEqual(ids(h), [21, 22, 23]);
  h.reset(); assert.deepEqual(ids(h), [1, 2, 3, 4, 5]);
});
test('manual completion retains rows, locks duplicate requests and invalidates stale completion on reset and unmount', async t => {
  const h = fixture(t, 'Browse controlled rows with pending completion');
  h.click('Next page');
  assert.deepEqual(ids(h), [1, 2, 3, 4, 5]);
  assert.ok(button(h.container, 'Next page').disabled);
  h.click('Next page'); assert.match(h.container.textContent, /Requests: 1/);
  h.click('Complete pending request');
  await h.act(async () => { await Promise.resolve(); });
  assert.deepEqual(ids(h), [6, 7, 8, 9, 10]);
  h.click('Next page');
  // A completion scheduled before reset must not republish into its new baseline.
  h.act(() => { Simulate.click(button(h.container, 'Complete pending request')); h.reset(); });
  await h.act(async () => { await Promise.resolve(); });
  assert.deepEqual(ids(h), [1, 2, 3, 4, 5]);
  assert.ok(button(h.container, 'Complete pending request').disabled);
  h.click('Next page'); assert.deepEqual(ids(h), [1, 2, 3, 4, 5], 'manual baseline survives reset');
  const Empty = h.modules.scenario('Show an empty controlled result').Demo;
  h.act(() => {
    Simulate.click(button(h.container, 'Complete pending request'));
    ReactDOM.render(h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Empty)), h.container);
  });
  await h.act(async () => { await Promise.resolve(); });
  assert.deepEqual(ids(h), []); assert.match(h.container.textContent, /0–0 of 0/);
});
test('empty result remains empty after reset and disables every real navigation action', t => {
  const h = fixture(t, 'Show an empty controlled result');
  assert.deepEqual(ids(h), []); assert.match(h.container.textContent, /0–0 of 0/);
  for (const text of ['First page', 'Previous page', 'Next page', 'Last page']) assert.ok(button(h.container, text).disabled);
  h.reset(); assert.deepEqual(ids(h), []); assert.match(h.container.textContent, /0–0 of 0/);
});

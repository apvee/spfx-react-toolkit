const assert = require('node:assert/strict');
const test = require('node:test');
const { pilotTitles, filterPilotGroups } = require('./fixtures/demo-presentation/jsdoc-pilots.cjs');
const { createContextFixture } = require('./fixtures/executable-examples/context-scenarios.cjs');
const { checkLegacyPermissionResult } = require('../scripts/jsdoc-pilots-browser.helpers.cjs');

test('PnP operations acceptance retains the exact independent hook case', () => {
  assert.deepEqual(pilotTitles.useSPFxPnP, ['Compare explicit invoke and batch reads']);
});

test('PnP context acceptance retains the exact independent hook case', () => {
  assert.deepEqual(pilotTitles.useSPFxPnPContext, ['Read current or cross-site web information']);
});

test('configuration acceptance retains distinct exact hook and factory scenarios', () => {
  assert.deepEqual(pilotTitles.useSPFxConfigurationStore, ['Compare layered configuration and physical inventories']);
  assert.deepEqual(pilotTitles.createSPFxConfigurationStoreService, ['Compose tenant defaults and site overrides with the service']);
});

test('tenant list acceptance contains the two exact new hook and factory cases', () => {
  assert.deepEqual(pilotTitles.useSPFxTenantKeyValueStore, ['Explore tenant records and exact text persistence']);
  assert.deepEqual(pilotTitles.createSPFxTenantKeyValueStoreService, ['Round trip numeric looking text through a tenant service']);
});

test('browser pilot filter preserves its prior cohort and adds exact service and Search cases without mutation', () => {
  const direct = { createSPFxPnPContextService: ['Explore PnP context caching and request headers'], createSPFxPnPService: ['Read through a supplied PnP service client'], createSPFxPnPListService: ['Read and page an explicitly selected list service'] };
  const search = { useSPFxPnPSearch: ['Search SharePoint text with explicit execution', 'Filter search results with native refiner tokens', 'Request debounced suggestions after explicit start'], SearchVerticals: ['Search selected verticals with builder paging'] };
  const pickers = {
    FileFolderPicker: ['Explore hosted files and folders with explicit source policies'],
    SitePicker: ['Explore site selection and contextual search'], ListPicker: ['Explore list catalogs and independent consumer choices'],
    FieldPicker: ['Explore site columns and list fields'], ContentTypePicker: ['Explore site and list content type references'],
    ViewPicker: ['Explore scoped view catalogs'], ListItemPicker: ['Explore list item labels and query pages'],
    PrincipalPicker: ['Explore principal discovery and explicit identity resolution'],
  };
  const spGroup = { SharePointGroupTrim: ['Stage direct SharePoint group membership and observe protected lifetimes'] };
  const trims = { ...spGroup, SharePointPermissionTrim: ['Stage resource permissions and observe protected lifetimes', 'Compose permission and Entra checks with nested AND'] };
  const additions = { ...direct, ...search, ...pickers, ...trims };
  for (const [name, titles] of Object.entries(additions)) assert.deepEqual(pilotTitles[name], titles);
  const features = Object.entries(pilotTitles).map(([title, titles]) => ({ key: title, title, scenarios: [...titles.map(title => ({ title })), { title: 'Future scenario' }] }));
  const expectedGroup = title => (['WebPartShell', 'Pagination'].includes(title) || Object.hasOwn(pickers, title) || Object.hasOwn(trims, title)) ? 'controls' : ['createSPFxPnPContextService', 'createSPFxPnPService', 'createSPFxPnPListService', 'createSPFxSiteKeyValueStoreService', 'createSPFxTenantPropertyService', 'createSPFxTenantKeyValueStoreService', 'createSPFxConfigurationStoreService'].includes(title) ? 'services' : 'hooks';
  const groups = ['controls', 'hooks', 'services'].map(key => ({ key, features: features.filter(feature => expectedGroup(feature.title) === key) })).concat([{ key: 'future', features: [{ title: 'Future feature', scenarios: [] }] }]);
  const original = structuredClone(groups);
  const filtered = filterPilotGroups(groups);
  assert.equal(filtered.length, 3);
  const retained = filtered.flatMap(group => group.features).filter(feature => !Object.hasOwn(additions, feature.title));
  assert.equal(retained.length, 47, 'exact earlier cohort retained');
  assert.equal(retained.flatMap(feature => feature.scenarios).length, 55);
  assert.equal(filtered.flatMap(group => group.features).length, retained.length + Object.keys(additions).length);
  assert.equal(filtered.flatMap(group => group.features).flatMap(feature => feature.scenarios).length, 55 + Object.values(additions).flat().length);
  assert.deepEqual(groups, original, 'filter never mutates any app catalogue group, feature or scenario');

  const preceding = filtered.flatMap(group => group.features).filter(feature => !Object.hasOwn(search, feature.title) && !Object.hasOwn(pickers, feature.title) && !Object.hasOwn(trims, feature.title));
  assert.equal(preceding.length, 50, 'complete accepted predecessor cohort');
  assert.equal(preceding.flatMap(feature => feature.scenarios).length, 58);
  const g1 = filtered.flatMap(group => group.features).filter(feature => !Object.hasOwn(pickers, feature.title) && !Object.hasOwn(trims, feature.title)).map(feature => ({ ...feature, scenarios: feature.scenarios.filter(entry => !['Filter search results with native refiner tokens','Request debounced suggestions after explicit start'].includes(entry.title)) }));
  assert.equal(g1.length, 52, 'complete accepted g1 feature cohort');
  assert.equal(g1.flatMap(feature => feature.scenarios).length, 60, 'complete accepted g1 scenario cohort');
  for (const title of ['Filter search results with native refiner tokens','Request debounced suggestions after explicit start']) {
    const missing = structuredClone(groups), duplicate = structuredClone(groups);
    missing.flatMap(g => g.features).find(f => f.title === 'useSPFxPnPSearch').scenarios = missing.flatMap(g => g.features).find(f => f.title === 'useSPFxPnPSearch').scenarios.filter(entry => entry.title !== title);
    assert.throws(() => filterPilotGroups(missing), /Missing pilot scenario or duplicate/);
    duplicate.flatMap(g => g.features).find(f => f.title === 'useSPFxPnPSearch').scenarios.push({ title });
    assert.throws(() => filterPilotGroups(duplicate), /Missing pilot scenario or duplicate/);
  }
  const missingFeature = structuredClone(groups);
  for (const group of missingFeature) group.features = group.features.filter(feature => feature.title !== 'WebPartShell');
  assert.throws(() => filterPilotGroups(missingFeature), /Missing pilot feature or duplicate: WebPartShell/);
  const missingScenario = structuredClone(groups);
  const shell = missingScenario.flatMap(group => group.features).find(feature => feature.title === 'WebPartShell');
  shell.scenarios = shell.scenarios.filter(scenario => scenario.title !== 'Configure the actual web part Shell');
  assert.throws(() => filterPilotGroups(missingScenario), /Missing pilot scenario or duplicate: WebPartShell\/Configure the actual web part Shell/);
  const acceptedPermission = filtered.flatMap(group => group.features).filter(feature => !Object.hasOwn(spGroup, feature.title));
  assert.equal(acceptedPermission.length,61,'all accepted 5c1 features preserved');
  assert.equal(acceptedPermission.flatMap(feature=>feature.scenarios).length,72,'all accepted 5c1 scenarios preserved');
  const prior = filtered.flatMap(group => group.features).filter(feature => !Object.hasOwn(trims, feature.title));
  assert.equal(prior.length,60,'complete final accepted predecessor feature cohort');
  assert.equal(prior.flatMap(feature=>feature.scenarios).length,70,'complete final accepted predecessor scenario cohort');
  assert.equal(filtered.flatMap(group=>group.features).length,60+Object.keys(trims).length);
  assert.equal(filtered.flatMap(group=>group.features).flatMap(feature=>feature.scenarios).length,70+Object.values(trims).flat().length);
  for (const [name,titles] of Object.entries(trims)) for (const title of titles) {
    const missing = structuredClone(groups), duplicate = structuredClone(groups);
    const missingOwner = missing.flatMap(group => group.features).find(feature => feature.title === name);
    missingOwner.scenarios = missingOwner.scenarios.filter(entry => entry.title !== title);
    assert.throws(() => filterPilotGroups(missing), new RegExp('Missing pilot scenario or duplicate: '+name));
    duplicate.flatMap(group => group.features).find(feature => feature.title === name).scenarios.push({ title });
    assert.throws(() => filterPilotGroups(duplicate), new RegExp('Missing pilot scenario or duplicate: '+name));
  }
  for (const owner of [...Object.keys(trims), ...Object.keys(pickers), 'useSPFxPnPSearch', 'SearchVerticals', 'createSPFxPnPContextService', 'createSPFxPnPService', 'createSPFxPnPListService', 'useSPFxPnPListById', 'useSPFxPnPListByUrl', 'useSPFxPnPListByPath', 'useSPFxPnP', 'useSPFxPnPContext', 'useSPFxConfigurationStore', 'createSPFxConfigurationStoreService', 'useSPFxTenantProperty', 'createSPFxTenantPropertyService', 'useSPFxTenantKeyValueStore', 'createSPFxTenantKeyValueStoreService', 'createSPFxConfigurationStoreService']) {
    const missing = structuredClone(groups);
    for (const group of missing) group.features = group.features.filter(feature => feature.title !== owner);
    assert.throws(() => filterPilotGroups(missing), new RegExp('Missing pilot feature or duplicate: ' + owner));
    const duplicate = structuredClone(groups);
    const group = duplicate.find(group => group.features.some(feature => feature.title === owner));
    group.features.push(structuredClone(group.features.find(feature => feature.title === owner)));
    assert.throws(() => filterPilotGroups(duplicate), new RegExp('Missing pilot feature or duplicate: ' + owner));
    const duplicateScenario = structuredClone(groups);
    const duplicated = duplicateScenario.flatMap(group => group.features).find(feature => feature.title === owner);
    duplicated.scenarios.push({ ...duplicated.scenarios[0] });
    assert.throws(() => filterPilotGroups(duplicateScenario), new RegExp('Missing pilot scenario or duplicate: ' + owner));
    const missingScenario = structuredClone(groups);
    missingScenario.flatMap(group => group.features).find(feature => feature.title === owner).scenarios = [];
    assert.throws(() => filterPilotGroups(missingScenario), new RegExp('Missing pilot scenario or duplicate: ' + owner));
  }
  const missingService = structuredClone(groups);
  missingService.find(group => group.key === 'services').features = missingService.find(group => group.key === 'services').features.filter(feature => feature.title !== 'createSPFxSiteKeyValueStoreService');
  assert.throws(() => filterPilotGroups(missingService), /Missing pilot feature or duplicate: createSPFxSiteKeyValueStoreService/);
  assert.deepEqual(groups, original, 'negative checks also leave the original catalogue unchanged');
});

// Execute the checker's locator contract against actual emitted-Demo DOM. This
// adapter has no browser polling/layout and is not a Playwright execution claim.
function domLocator(resolve) {
  const descendants = selector => resolve().flatMap(root => [...root.querySelectorAll(selector)]);
  return {
    getByLabel(label, { exact }) {
      assert.equal(exact, true);
      return domLocator(() => descendants('[aria-label]').filter(node => node.getAttribute('aria-label') === label));
    },
    getByText(text, { exact }) {
      return domLocator(() => descendants('*').filter(node => {
        const content = node.textContent.replace(/\s+/g, ' ').trim();
        return exact ? content === text : content.toLowerCase().includes(text.toLowerCase());
      }));
    },
    getByRole(role) { return domLocator(() => descendants(`[role="${role}"]`)); },
    locator(selector) { return domLocator(() => descendants(selector)); },
    first() { return domLocator(() => resolve().slice(0, 1)); },
    async waitFor() { assert.ok(resolve().some(node => node.isConnected), 'No visible locator match'); },
    async count() { return resolve().length; },
    async allTextContents() { return resolve().map(node => node.textContent); },
  };
}
const check = (label, actual, expected) => assert.deepEqual(actual, expected, label);
async function legacyFixture(t) {
  const f = createContextFixture(t, 'Check staged legacy site web and list permissions');
  f.change('Site URL', 'https://tenant.test/sites/target');
  f.click('Check / repeat');
  await f.flush();
  assert.equal(f.calls.length, 2);
  return { ...f, panel: domLocator(() => [f.container]) };
}
for (const state of ['pending', 'error']) test(`legacy checker contract cannot accept actual ${state} Unavailable masks`, async t => {
  const f = await legacyFixture(t);
  if (state === 'error') await f.flush(() => f.calls[0].reject(new Error('Native mask failure')));
  assert.match(f.container.textContent, /Site mask: Unavailable; Web mask: Unavailable; List mask: Unavailable/);
  await assert.rejects(() => checkLegacyPermissionResult(f.panel, check), /No visible locator match/);
});
test('legacy checker contract accepts successful native masks and verifies actual decision values', async t => {
  const f = await legacyFixture(t);
  await f.flush(() => f.calls.forEach(call => call.resolve({ High: '0', Low: '31' })));
  await checkLegacyPermissionResult(f.panel, check);
});
test('legacy checker contract rejects available masks with unexpected native decisions', async t => {
  const f = await legacyFixture(t);
  await f.flush(() => f.calls.forEach(call => call.resolve({ High: '0', Low: '0' })));
  await assert.rejects(() => checkLegacyPermissionResult(f.panel, check), /legacy decision values/);
});

test('isolated Styles suggestions route imports the exact canonical emitted case and retains its presentation tier', () => {
  const fs = require('node:fs'), path = require('node:path');
  const selected = require('./fixtures/executable-examples/cards-description-list.cjs').scenario('Request debounced suggestions after explicit start','packages/spfx-react-toolkit/src/hooks/useSPFxPnPSearch.ts#useSPFxPnPSearch');
  const source = fs.readFileSync(path.resolve('tests/fixtures/styles-panel-browser/index.tsx'),'utf8');
  assert.ok(source.includes('/generated/jsdoc-examples/' + selected.entry.id + "'"));
  assert.match(source, /<SuggestionsDemo render=/);
  assert.match(fs.readFileSync(path.resolve('tests/fixtures/styles-panel-browser/toolkit-boundary.ts'),'utf8'), /Isolated suggestion presentation double/);
});

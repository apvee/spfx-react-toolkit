const assert = require('node:assert/strict');
const test = require('node:test');
const { pilotTitles, filterPilotGroups } = require('./fixtures/demo-presentation/jsdoc-pilots.cjs');
const { createContextFixture } = require('./fixtures/executable-examples/context-scenarios.cjs');
const { checkLegacyPermissionResult } = require('../scripts/jsdoc-pilots-browser.helpers.cjs');

test('browser pilot filter selects exactly 36 features and 43 independently named scenarios without mutation', () => {
  const features = Object.entries(pilotTitles).map(([title, titles]) => ({ key: title, title, scenarios: [...titles.map(title => ({ title })), { title: 'Future scenario' }] }));
  const expectedGroup = title => ['WebPartShell', 'Pagination'].includes(title) ? 'controls' : title === 'createSPFxSiteKeyValueStoreService' ? 'services' : 'hooks';
  const groups = ['controls', 'hooks', 'services'].map(key => ({ key, features: features.filter(feature => expectedGroup(feature.title) === key) })).concat([{ key: 'future', features: [{ title: 'Future feature', scenarios: [] }] }]);
  const original = structuredClone(groups);
  const filtered = filterPilotGroups(groups);
  assert.equal(filtered.length, 3);
  assert.equal(filtered.flatMap(group => group.features).length, 36);
  assert.equal(filtered.flatMap(group => group.features).flatMap(feature => feature.scenarios).length, 43);
  assert.deepEqual(groups, original, 'filter never mutates any app catalogue group, feature or scenario');

  const missingFeature = structuredClone(groups);
  for (const group of missingFeature) group.features = group.features.filter(feature => feature.title !== 'WebPartShell');
  assert.throws(() => filterPilotGroups(missingFeature), /Missing pilot feature or duplicate: WebPartShell/);
  const missingScenario = structuredClone(groups);
  const shell = missingScenario.flatMap(group => group.features).find(feature => feature.title === 'WebPartShell');
  shell.scenarios = shell.scenarios.filter(scenario => scenario.title !== 'Configure the actual web part Shell');
  assert.throws(() => filterPilotGroups(missingScenario), /Missing pilot scenario or duplicate: WebPartShell\/Configure the actual web part Shell/);
  const missingService = structuredClone(groups);
  missingService.find(group => group.key === 'services').features = [];
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

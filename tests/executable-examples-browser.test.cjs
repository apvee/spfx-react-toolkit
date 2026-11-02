const assert = require('node:assert/strict');
const test = require('node:test');
const { pilotTitles, filterPilotGroups } = require('./fixtures/demo-presentation/jsdoc-pilots.cjs');
const { createContextFixture } = require('./fixtures/executable-examples/context-scenarios.cjs');
const { checkLegacyPermissionResult } = require('../scripts/jsdoc-pilots-browser.helpers.cjs');

test('browser pilot filter selects exactly 26 features and 33 independently named scenarios', () => {
  const features = Object.entries(pilotTitles).map(([title, titles]) => ({ key: title, title, scenarios: [...titles.map(title => ({ title })), { title: 'Future scenario' }] }));
  const groups = [{ key: 'controls', features: features.slice(0, 2) }, { key: 'hooks', features: features.slice(2) }, { key: 'future', features: [{ title: 'Future feature', scenarios: [] }] }];
  const filtered = filterPilotGroups(groups);
  assert.equal(filtered.length, 2);
  assert.equal(filtered.flatMap(group => group.features).length, 26);
  assert.equal(filtered.flatMap(group => group.features).flatMap(feature => feature.scenarios).length, 33);
  assert.equal(groups[0].features[0].scenarios.length, 4, 'filter never mutates the app catalogue');
  assert.throws(() => filterPilotGroups(groups.slice(0, 1)), /Missing pilot feature/);
  assert.throws(() => filterPilotGroups([{ ...groups[0], features: [{ ...features[0], scenarios: [] }, features[1]] }, groups[1]]), /Missing pilot scenario/);
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

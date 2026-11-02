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

test('real catalog derives nine features and sixteen source-owned scenarios', t => {
  const h = fixture(t, 'Browse controlled rows');
  assert.equal(h.modules.catalogue.executableExamples.length, 16);
  assert.equal(h.modules.catalogue.executableGroups.flatMap(group => group.features).length, 9);
  assert.deepEqual(h.modules.catalogue.executableGroups.map(group => group.key), ['controls', 'hooks']);
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

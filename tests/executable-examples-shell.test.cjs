const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./fixtures/web-part-shell-harness.cjs');
const { pilotModules, button } = require('./fixtures/executable-examples/pilot-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const ReactDOM = require('react-dom');
const { RendererProvider } = require('@griffel/react');
const { createDOMRenderer } = require('@griffel/core');
function fixture(t, title, mode = 1, options = {}) {
  const h = createHarness(); t.after(h.close);
  const modules = pilotModules(h);
  let Demo = modules.scenario(title).Demo, scenarioKey = title;
  const { WebPartShellDemoHost } = modules.load('WebPartShellDemoHost.tsx');
  const { useExecutableShellPort } = modules.load('shared/useExecutableShellPort.ts');
  const { useExampleExplorerState } = modules.load('shared/ExampleExplorerState.tsx');
  const instance = h.instance('source-pilot'); instance.displayMode = mode;
  instance.properties = { title: 'Persisted title', titleHTML: undefined, showTitle: true, extra: 'keep', ...options.properties };
  let revision = 9, port, navigation, propertiesInfo, refreshes = 0, mounts = 0;
  const { useSPFxProperties } = h.load('hooks/useSPFxProperties.ts');
  instance.context.propertyPane.refresh = () => { refreshes++; };
  const renderer = createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  const container = document.createElement('div'); document.body.appendChild(container);
  t.after(() => { h.act(() => ReactDOM.unmountComponentAtNode(container)); container.remove(); });
  function Content() {
    port = useExecutableShellPort(); navigation = useExampleExplorerState(); propertiesInfo = useSPFxProperties();
    h.React.useEffect(() => { mounts++; }, []);
    return h.React.createElement(Demo, { key: scenarioKey, shell: port, resetRevision: revision });
  }
  function render() { h.act(() => ReactDOM.render(h.React.createElement(RendererProvider, { renderer, targetDocument: document },
    h.React.createElement(WebPartShellDemoHost, { instance, 'data-real-shell': '' }, h.React.createElement(Content))), container)); }
  const navigationControl = document.createElement('button'); navigationControl.textContent = 'Scenario selector';
  document.body.appendChild(navigationControl); t.after(() => navigationControl.remove());
  if (options.focusNavigation) navigationControl.focus();
  render();
  return { ...h, container, instance, modules, render, navigationControl,
    setProperties(patch) { h.act(() => propertiesInfo.setProperties(patch)); },
    get port() { return port; }, get navigation() { return navigation; }, get refreshes() { return refreshes; }, get mounts() { return mounts; },
    reset() { revision++; render(); },
    select(title) { scenarioKey = title; Demo = modules.scenario(title).Demo; render(); },
    click(text) { h.act(() => Simulate.click(button(container, text))); },
    change(selector, value) { h.act(() => { const node = container.querySelector(selector); node.value = value; Simulate.change(node, { target: { value } }); }); } };
}
for (const title of ['Configure the actual web part Shell', 'Retain content through Shell presentation changes', 'Observe content remount between subtle and Card']) {
  test(`${title}: absent adapter explains native prerequisites without synthesizing Shell`, t => {
    const h = createHarness(); t.after(h.close);
    const { Demo } = pilotModules(h).scenario(title);
    const container = document.createElement('div'); document.body.appendChild(container);
    t.after(() => { h.act(() => ReactDOM.unmountComponentAtNode(container)); container.remove(); });
    h.act(() => ReactDOM.render(h.React.createElement(Demo), container));
    assert.match(container.textContent, /actual WebPartShell host/);
    assert.equal(container.querySelectorAll('[data-web-part-shell-surface]').length, 0);
  });
}
for (const title of ['Configure the actual web part Shell', 'Retain content through Shell presentation changes', 'Observe content remount between subtle and Card']) for (const mode of [1, 2]) test(`${title}: reset uses actual ${mode === 2 ? 'Edit' : 'Read'} mode and never writes on mount`, t => {
  const h = fixture(t, title, mode);
  assert.equal(h.instance.properties.title, 'Persisted title'); assert.equal(h.refreshes, 0);
  assert.match(h.container.textContent, new RegExp(`Mode: ${mode === 2 ? 'Edit' : 'Read'}`));
  assert.equal(h.container.querySelectorAll('[data-real-shell]').length, 1);
  assert.equal(h.container.querySelector('h2').textContent, 'Persisted title');
  h.reset();
  const writesTitle = title === 'Configure the actual web part Shell' && mode === 2;
  assert.equal(h.instance.properties.title, writesTitle ? 'SPFx React Toolkit' : 'Persisted title');
  assert.equal(h.refreshes, writesTitle ? 1 : 0);
  assert.equal(h.instance.properties.extra, 'keep');
  // A later provider mode update is real; mounting at a nonzero revision isn't reset.
  h.instance.displayMode = mode === 1 ? 2 : 1; h.render();
  assert.equal(h.refreshes, writesTitle ? 1 : 0);
});
test('real outer Shell retains content across Card variants, size, placement and host-owned visibility', t => {
  const h = fixture(t, 'Retain content through Shell presentation changes');
  h.click('Use Card baseline');
  h.change('input[aria-label="Project note"]', 'Retained note'); h.click('Increment counter');
  const input = h.container.querySelector('input[aria-label="Project note"]'), mounts = h.mounts;
  h.change('select[aria-label="Card appearance"]', 'filled-alternative');
  h.change('select[aria-label="Card size"]', 'large');
  h.change('select[aria-label="Title position"]', 'inside');
  assert.equal(h.container.querySelector('input[aria-label="Project note"]'), input);
  assert.equal(input.value, 'Retained note'); assert.match(h.container.textContent, /Child counter: 1/);
  button(h.container, 'Hide title').focus(); h.click('Hide title');
  assert.equal(h.instance.properties.showTitle, false); assert.equal(h.container.querySelector('h2'), null);
  h.setProperties({ showTitle: true }); assert.equal(h.mounts, mounts);
  assert.equal(h.container.querySelector('input[aria-label="Project note"]'), input);
  h.reset(); assert.equal(input.value, 'Intranet project review'); assert.match(h.container.textContent, /Child counter: 0/);
});
test('subtle/Card changes remount real body but preserve upper explorer selection and never loop a preset', t => {
  const h = fixture(t, 'Observe content remount between subtle and Card');
  const entry = h.modules.catalogue.executableExamples.find(item => item.title === 'Observe content remount between subtle and Card');
  h.act(() => { h.navigation.selectFeature(entry.featureKey); h.navigation.selectScenario(entry.featureKey, entry.id); });
  h.change('input[aria-label="Project note"]', 'Before Card'); h.click('Increment counter');
  const input = h.container.querySelector('input[aria-label="Project note"]');
  h.click('Use Card baseline');
  assert.notEqual(h.container.querySelector('input[aria-label="Project note"]'), input);
  assert.match(h.container.textContent, /Child counter: 0/);
  assert.equal(h.navigation.selectedFeatureKey, entry.featureKey);
  assert.equal(h.navigation.selectedScenarioByFeature[entry.featureKey], entry.id);
  const count = h.mounts; h.render(); assert.equal(h.mounts, count);
  h.click('Use subtle baseline'); assert.equal(h.mounts, count + 1);
  assert.equal(h.port.presentation.appearance, 'subtle');
  h.render(); assert.equal(h.mounts, count + 1);
});


test('Pagination reset and scenario departure cancel late completion without replacing or reconfiguring the actual outer Shell', async t => {
  const h = fixture(t, 'Browse controlled rows with pending completion');
  h.act(() => h.port.updatePresentation({ appearance: 'outline', size: 'large', titlePosition: 'inside' }));
  const root = h.container.querySelector('[data-real-shell]');
  const body = h.container.querySelector('[data-web-part-shell-body]');
  const rows = () => [...h.container.querySelectorAll('[data-order-id]')].map(node => Number(node.dataset.orderId));
  assert.deepEqual(rows(), [1, 2, 3, 4, 5]);
  h.click('Next page');
  h.act(() => { Simulate.click(button(h.container, 'Complete pending request')); h.reset(); });
  await require('react-dom/test-utils').act(async () => { await Promise.resolve(); });
  assert.deepEqual(rows(), [1, 2, 3, 4, 5]);
  assert.equal(h.container.querySelector('[data-real-shell]'), root);
  h.click('Next page');
  h.act(() => { Simulate.click(button(h.container, 'Complete pending request')); h.select('Show an empty controlled result'); });
  await require('react-dom/test-utils').act(async () => { await Promise.resolve(); });
  assert.deepEqual(rows(), []); assert.match(h.container.textContent, /0–0 of 0/);
  assert.equal(h.container.querySelector('[data-real-shell]'), root);
  assert.equal(h.container.querySelector('[data-web-part-shell-body]'), body);
  assert.deepEqual(h.port.presentation, { appearance: 'outline', size: 'large', titlePosition: 'inside', padding: 'auto' });
  assert.equal(h.instance.properties.title, 'Persisted title'); assert.equal(h.refreshes, 0);
});


test('initially hidden Shell mount, scenario selection and body remount preserve navigation focus', t => {
  const h = fixture(t, 'Configure the actual web part Shell', 1, { properties: { showTitle: false }, focusNavigation: true });
  assert.ok(document.activeElement === h.navigationControl, 'initially hidden mount must not move focus');
  for (const title of ['Retain content through Shell presentation changes', 'Observe content remount between subtle and Card', 'Configure the actual web part Shell']) {
    h.navigationControl.focus(); h.select(title);
    assert.ok(document.activeElement === h.navigationControl, 'hidden scenario selection preserves navigation focus');
    h.act(() => h.port.updatePresentation({ appearance: h.port.presentation.appearance === 'subtle' ? 'outline' : 'subtle' }));
    assert.ok(document.activeElement === h.navigationControl, 'already hidden body remount preserves navigation focus');
  }
});
test('configuration owns title commands and visible-to-hidden focus recovery without note/counter state', t => {
  const h = fixture(t, 'Configure the actual web part Shell', 2);
  assert.ok(h.container.querySelector('input[aria-label="Project note"]') === null, 'configuration has no content-retention state');
  h.click('Formatted title'); assert.match(h.instance.properties.titleHTML, /<strong>Project<\/strong>/);
  button(h.container, 'Hide title').focus(); h.click('Hide title');
  assert.equal(h.instance.properties.showTitle, false);
  assert.equal(document.activeElement, button(h.container, 'Show title'));
  h.click('Show title'); assert.equal(h.instance.properties.showTitle, true);
});
for (const title of ['Retain content through Shell presentation changes', 'Observe content remount between subtle and Card']) for (const mode of [1, 2]) test(`${title}: reset preserves persisted plain/rich/hidden properties in native mode ${mode}`, t => {
  const properties = { title: 'Saved', titleHTML: '<h2><strong>Saved</strong></h2>', showTitle: false };
  const h = fixture(t, title, mode, { properties });
  h.click('Use Card baseline'); h.change('input[aria-label="Project note"]', 'Changed'); h.click('Increment counter');
  h.reset();
  assert.deepEqual(h.instance.properties, { ...properties, extra: 'keep' });
  assert.equal(h.refreshes, 0);
  assert.equal(h.container.querySelector('input[aria-label="Project note"]').value, 'Intranet project review');
  assert.match(h.container.textContent, /Child counter: 0/);
  for (const label of ['Plain title', 'Formatted title', 'Empty title', 'Show title']) assert.equal([...h.container.querySelectorAll('button')].some(node => node.textContent === label), false, 'content scenarios do not own title commands');
  if (title === 'Retain content through Shell presentation changes') {
    assert.equal([...h.container.querySelector('select[aria-label="Card appearance"]').options].some(option => option.value === 'subtle'), false);
    assert.equal([...h.container.querySelectorAll('button')].some(node => node.textContent === 'Use subtle baseline'), false);
  } else assert.ok(h.container.querySelector('select') === null, 'remount case owns explicit appearance switching only');
});

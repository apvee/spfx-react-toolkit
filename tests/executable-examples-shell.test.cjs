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
    h.React.createElement(WebPartShellDemoHost, { instance, 'data-real-shell': '', ...options.consumerProps }, h.React.createElement(Content))), container)); }
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
  assert.deepEqual(h.port.presentation, { appearance: 'outline', size: 'large', titlePosition: 'inside', padding: 'auto', surface: undefined, dir: undefined, strings: undefined, titleEditorStrings: undefined, action: undefined });
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

test('explicit explorer Reset restores focus after actual Shell body remount without stealing later focus', async t => {
  const h = createHarness(); t.after(h.close);
  const modules = pilotModules(h);
  const { WebPartShellDemoHost } = modules.load('WebPartShellDemoHost.tsx');
  const { ExampleExplorer } = modules.load('ExampleExplorer.tsx');
  const groups = modules.catalogue.executableGroups.filter(group => group.features.some(feature => feature.title === 'WebPartShell'))
    .map(group => ({ ...group, features: group.features.filter(feature => feature.title === 'WebPartShell') }));
  const instance = h.instance('reset-focus'); instance.displayMode = 1;
  const renderer = createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  const container = document.createElement('div'), elsewhere = document.createElement('button');
  elsewhere.textContent = 'Later user focus'; document.body.append(container, elsewhere);
  t.after(() => { h.act(() => ReactDOM.unmountComponentAtNode(container)); container.remove(); elsewhere.remove(); });
  const flush = async () => { await require('react-dom/test-utils').act(async () => {
    await new Promise((resolve, reject) => {
      const observer = new window.MutationObserver(check);
      const deadline = setTimeout(() => { observer.disconnect(); reject(new Error('Shell scenario did not commit')); }, 5000);
      function check() { if (container.querySelector('select[aria-label="Appearance"]')) { clearTimeout(deadline); observer.disconnect(); resolve(); } }
      observer.observe(container, { childList: true, subtree: true }); check();
    });
  }); };
  elsewhere.focus();
  h.act(() => ReactDOM.render(h.React.createElement(RendererProvider, { renderer, targetDocument: document },
    h.React.createElement(WebPartShellDemoHost, { instance }, h.React.createElement(ExampleExplorer, { groups }))), container));
  // The host defaults to the full catalogue's first feature. Select the actual
  // Shell feature through the mobile selector, then observe its real lazy commit.
  h.act(() => Simulate.change(container.querySelector('select[aria-label="Feature"]'), { target: { value: groups[0].features[0].key } }));
  await flush(); assert.equal(document.activeElement, elsewhere, 'mount/selection never invents reset focus');
  const appearance = () => container.querySelector('select[aria-label="Appearance"]');
  h.act(() => Simulate.change(appearance(), { target: { value: 'filled' } })); await flush();
  const oldReset = button(container, 'Reset scenario'); oldReset.focus();
  h.act(() => Simulate.click(oldReset)); await flush();
  assert.equal(oldReset.isConnected, false, 'real subtle/Card body change destroys the old toolbar');
  assert.ok(document.activeElement === button(container, 'Reset scenario'), 'explicit Reset focus follows its replacement after real Shell remount');
  h.act(() => Simulate.change(appearance(), { target: { value: 'filled' } })); await flush();
  const nextReset = button(container, 'Reset scenario'); nextReset.focus();
  h.act(() => { Simulate.click(nextReset); elsewhere.focus(); }); await flush();
  assert.equal(nextReset.isConnected, false);
  assert.equal(document.activeElement, elsewhere, 'later user focus wins over pending reset recovery');
  h.act(() => Simulate.change(appearance(), { target: { value: 'filled' } })); await flush();
  const pointerReset = button(container, 'Reset scenario'); pointerReset.focus();
  h.act(() => { Simulate.click(pointerReset); elsewhere.dispatchEvent(new window.Event('pointerdown', { bubbles: true })); }); await flush();
  assert.equal(pointerReset.isConnected, false);
  assert.ok(document.activeElement !== button(container, 'Reset scenario'), 'later pointer action cancels recovery even before focus changes');
  // Track only the exact callbacks owned by this private focus intent. Native
  // Fluent/provider listeners remain real and outside this assertion.
  const add = document.addEventListener, remove = document.removeEventListener;
  const pending = new Map();
  document.addEventListener = function(type, listener, capture) {
    if (capture === true && ['focusin', 'pointerdown'].includes(type)) pending.set(listener, type);
    return add.call(this, type, listener, capture);
  };
  document.removeEventListener = function(type, listener, capture) {
    pending.delete(listener); return remove.call(this, type, listener, capture);
  };
  try {
    const reset = button(container, 'Reset scenario'); reset.focus(); h.act(() => Simulate.click(reset));
    assert.equal(pending.size, 2, 'a retained Reset owns two per-document cancellation listeners');
    h.act(() => Simulate.change(container.querySelector('select[aria-label="Scenario"]'), { target: { value: groups[0].features[0].scenarios[1].id } }));
    await require('react-dom/test-utils').act(async () => { await Promise.resolve(); });
    assert.equal(pending.size, 0, 'scenario navigation clears pending focus intent');
    const lastReset = button(container, 'Reset scenario'); lastReset.focus(); h.act(() => Simulate.click(lastReset));
    // This scenario reset can cross the Shell boundary and consume immediately;
    // a second same-boundary reset remains pending until unmount/navigation.
    await require('react-dom/test-utils').act(async () => { await Promise.resolve(); });
    const retainedReset = button(container, 'Reset scenario'); retainedReset.focus(); h.act(() => Simulate.click(retainedReset));
    assert.equal(pending.size, 2);
    h.act(() => ReactDOM.unmountComponentAtNode(container));
    assert.equal(pending.size, 0, 'provider unmount removes its per-document listeners');
  } finally { document.addEventListener = add; document.removeEventListener = remove; }
});


test('source Shell controls drive actual surface, direction, independent dictionaries and action slot without mount writes', t => {
  const h = fixture(t, 'Configure the actual web part Shell', 2);
  const root = () => h.container.querySelector('[data-real-shell]');
  assert.equal(h.refreshes, 0);
  h.change('select[aria-label="Surface"]', 'application');
  assert.equal(root().getAttribute('data-web-part-shell-surface'), 'application');
  h.change('select[aria-label="Direction"]', 'rtl');
  assert.equal(root().closest('[dir]').getAttribute('dir'), 'rtl');
  function toggle(marker, checked) {
    h.act(() => { const node = h.container.querySelector(`[data-shell-demo="${marker}"]`); node.checked = checked; Simulate.change(node); });
  }
  toggle('strings', true);
  assert.deepEqual(h.port.presentation.strings, { titleEditorLabel: 'Project section heading', titlePlaceholder: 'Add a project heading' });
  assert.equal(h.port.presentation.titleEditorStrings, undefined);
  toggle('title-editor-strings', true);
  assert.deepEqual(h.port.presentation.titleEditorStrings, { bold: 'Strong', formattingControls: 'Project heading controls', colorNames: { red: 'Alert red' } });
  toggle('strings', false);
  assert.equal(h.port.presentation.strings, undefined);
  assert.equal(h.port.presentation.titleEditorStrings.bold, 'Strong');
  assert.equal(h.refreshes, 0, 'presentation and dictionaries never publish title properties');
  h.change('select[aria-label="Header action"]', 'ordinary');
  assert.equal(h.container.querySelectorAll('[data-shell-demo="action-slot"]').length, 0);
  assert.equal(h.container.querySelectorAll('[data-shell-demo="action"]').length, 1);
  h.change('select[aria-label="Header action"]', 'slot');
  assert.equal(h.container.querySelectorAll('[data-shell-demo="action-slot"] > [data-shell-demo="action"]').length, 1);
  const action = h.port.presentation.action;
  h.select('Retain content through Shell presentation changes');
  assert.equal(h.port.presentation.action, action, 'configured stateless slot survives scenario departure');
  h.click('Use Card baseline');
  assert.equal(h.port.presentation.action, action, 'native body remount does not tear down configured presentation');
  h.click('Hide title');
  assert.equal(h.instance.properties.showTitle, false);
  assert.equal(h.container.querySelectorAll('[data-shell-demo="action"], [data-shell-demo="action-slot"]').length, 0);
  h.setProperties({ showTitle: true });
  assert.equal(h.container.querySelectorAll('[data-shell-demo="action-slot"] > [data-shell-demo="action"]').length, 1);
  h.select('Configure the actual web part Shell'); h.reset();
  for (const key of ['surface', 'dir', 'strings', 'titleEditorStrings', 'action']) assert.equal(h.port.presentation[key], undefined, key);
  assert.equal(h.container.querySelectorAll('[data-shell-demo="action"]').length, 1, 'Reset restores generic native action');
  assert.equal(h.container.querySelectorAll('[data-shell-demo="action-slot"]').length, 0);
});

test('consumer native Shell props take precedence and source undefined releases only its overrides', t => {
  const consumerProps = { surface: 'section', dir: 'ltr', strings: { titleEditorLabel: 'Consumer title' }, titleEditorStrings: { bold: 'Consumer bold' } };
  const h = fixture(t, 'Configure the actual web part Shell', 1, { consumerProps });
  h.act(() => h.port.updatePresentation({ surface: 'application', dir: 'rtl', strings: { titleEditorLabel: 'Source title' }, titleEditorStrings: { bold: 'Source bold' } }));
  assert.equal(h.port.presentation.surface, 'section'); assert.equal(h.port.presentation.dir, 'ltr');
  assert.equal(h.port.presentation.strings, consumerProps.strings); assert.equal(h.port.presentation.titleEditorStrings, consumerProps.titleEditorStrings);
  delete consumerProps.surface; delete consumerProps.dir; delete consumerProps.strings; delete consumerProps.titleEditorStrings; h.render();
  assert.equal(h.port.presentation.surface, 'application'); assert.equal(h.port.presentation.dir, 'rtl');
  h.act(() => h.port.updatePresentation({ surface: undefined, dir: undefined, strings: undefined, titleEditorStrings: undefined }));
  assert.equal(h.port.presentation.surface, undefined); assert.equal(h.port.presentation.dir, undefined);
  assert.equal(h.port.presentation.strings, undefined); assert.equal(h.port.presentation.titleEditorStrings, undefined);
  assert.equal(h.refreshes, 0);
  consumerProps.action = h.React.createElement('button', { 'data-consumer-action': '' }, 'Consumer action'); h.render();
  h.change('select[aria-label="Header action"]', 'slot');
  assert.equal(h.container.querySelectorAll('[data-consumer-action]').length, 1);
  assert.equal(h.container.querySelectorAll('[data-shell-demo="action-slot"]').length, 0);
  delete consumerProps.action; h.render();
  assert.equal(h.container.querySelectorAll('[data-shell-demo="action-slot"]').length, 1);
});

test('configured action callback is suppressed while hidden and actual host disposal unmounts it', t => {
  const h = fixture(t, 'Configure the actual web part Shell');
  let renders = 0, mounts = 0, cleanups = 0;
  function ActionProbe() {
    h.React.useEffect(() => { mounts++; return () => { cleanups++; }; }, []);
    return h.React.createElement('button', { 'data-action-probe': '' }, 'Probe');
  }
  h.act(() => h.port.updatePresentation({ action: { children: (Component, props) => {
    renders++; return h.React.createElement(Component, props, h.React.createElement(ActionProbe));
  } } }));
  assert.equal(mounts, 1);
  h.setProperties({ showTitle: false });
  assert.equal(cleanups, 1); const hiddenRenders = renders;
  h.act(() => h.port.updatePresentation({ size: 'large' })); h.render();
  assert.equal(renders, hiddenRenders, 'hidden header never invokes source slot');
  h.setProperties({ showTitle: true }); assert.equal(mounts, 2);
  h.act(() => ReactDOM.unmountComponentAtNode(h.container));
  assert.equal(cleanups, 2); assert.equal(h.container.querySelector('[data-action-probe]'), null);
});

test('metadata scenario inspects real plain-title metadata and fresh map/descriptor ownership', t => {
  const h = createHarness(); t.after(h.close);
  const { Demo } = pilotModules(h).scenario('Inspect searchable Shell title metadata');
  const host = document.createElement('div'); document.body.appendChild(host);
  t.after(() => { h.act(() => ReactDOM.unmountComponentAtNode(host)); host.remove(); });
  let revision = 4;
  const render = () => h.act(() => ReactDOM.render(h.React.createElement(Demo, { resetRevision: revision }), host));
  render(); assert.match(host.textContent, /Plain title searchable: yes/); assert.match(host.textContent, /No comparison requested/);
  h.act(() => Simulate.click(button(host, 'Create fresh metadata')));
  assert.match(host.textContent, /Fresh map and title descriptor: yes/);
  revision++; render(); assert.match(host.textContent, /No comparison requested/);
});


test('configuration diagnostics safely project a consumer JSX action with a live React owner', t => {
  const h = createHarness(); t.after(h.close);
  const modules = pilotModules(h);
  const { WebPartShellDemoHost } = modules.load('WebPartShellDemoHost.tsx');
  const { useExecutableShellPort } = modules.load('shared/useExecutableShellPort.ts');
  const { Demo } = modules.scenario('Configure the actual web part Shell');
  const instance = h.instance('owned-consumer-action');
  let port;
  function Content() { port = useExecutableShellPort(); return h.React.createElement(Demo, { shell: port }); }
  function Consumer() {
    return h.React.createElement(WebPartShellDemoHost, { instance,
      action: h.React.createElement('button', { 'data-consumer-owned-action': '' }, 'Consumer owned action') }, h.React.createElement(Content));
  }
  const renderer = createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  const host = document.createElement('div'); document.body.appendChild(host);
  t.after(() => { h.act(() => ReactDOM.unmountComponentAtNode(host)); host.remove(); });
  h.act(() => ReactDOM.render(h.React.createElement(RendererProvider, { renderer }, h.React.createElement(Consumer)), host));
  assert.throws(() => JSON.stringify(port.presentation.action), /circular/i, 'the fixture owns a real React Fiber rather than an owner-null element');
  assert.equal(host.querySelectorAll('[data-consumer-owned-action]').length, 1);
  const observations = JSON.parse(host.querySelector('[data-shell-demo="observations"]').textContent);
  assert.equal(observations.action, 'ordinary'); assert.equal(observations.appearance, 'subtle');
});

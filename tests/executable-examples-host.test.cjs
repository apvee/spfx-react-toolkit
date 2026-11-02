const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { button, base } = require('./fixtures/executable-examples/pilot-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const ReactDOM = require('react-dom');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const { RendererProvider } = require('@griffel/react');
const { createDOMRenderer } = require('@griffel/core');
const title = 'Read a user-selected list with observable results';

function fixture(t, { explorer = false, simulatedBrowserTransport = false } = {}) {
  const calls = [];
  let initializationError;
  const boundaries = {
    '../WebPartShellDemoHost': { useShellPresentation() { throw new Error('Ordinary read scenario must not use the Shell adapter'); } },
  };
  const h = createHarness(boundaries); t.after(h.close);
  // JSDOM 15 omits the browser crypto identity used by real Fluent focus handling.
  Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  // Real context hook/service, list hook/service, PnP selector, fluent builder,
  // request URL and parser. Only SPFx authentication identity and HTTP send are doubled.
  const pnp = h.load(require.resolve('@pnp/sp'));
  const { DefaultParse } = h.load(require.resolve('@pnp/queryable'));
  boundaries['@pnp/sp'] = { ...pnp, SPFx: () => {
    if (initializationError) throw initializationError;
    return instance => {
      instance.using(pnp.DefaultInit(), pnp.DefaultHeaders(), DefaultParse());
      instance.on.send.replace((url, init) => new Promise((resolve, reject) => {
        assert.equal(init.method, 'GET', 'read-only scenario must never send CRUD');
        calls.push({ url: new URL(String(url)), init, resolve, reject });
      }));
      return instance;
    };
  } };
  if (simulatedBrowserTransport) boundaries['../services/spfx-pnp-context.service'] = h.load(path.resolve(__dirname, 'fixtures/demo-presentation/pnp-read-transport.ts'));
  const toolkit = {
    ...h.load('hooks/useSPFxPnPContext.ts'), ...h.load('hooks/useSPFxPnPList.ts'),
    ...h.load('hooks/useStableCallback.ts'),
  };
  for (const specifier of ['@apvee/spfx-react-toolkit', '@apvee/spfx-react-toolkit/hooks/useSPFxPnPContext', '@apvee/spfx-react-toolkit/hooks/useSPFxPnPList', '@apvee/spfx-react-toolkit/hooks/useStableCallback']) boundaries[specifier] = toolkit;
  const catalogue = h.load(path.join(base, 'generated/jsdoc-examples/catalogue.ts'));
  const entry = catalogue.executableExamples.find(item => item.title === title);
  assert.ok(entry, `Source-owned executable scenario: ${title}`);
  const { Demo } = h.load(path.join(base, `generated/jsdoc-examples/${entry.id}.tsx`));
  let revision = 19, show = true, renders = 0;
  let navigation;
  const state = explorer ? h.load(path.join(base, 'shared/ExampleExplorerState.tsx')) : undefined;
  const Explorer = explorer ? h.load(path.join(base, 'ExampleExplorer.tsx')).ExampleExplorer : undefined;
  // Include only the real read and generic callback features: no unavailable
  // native Shell API or unrelated sample transport participates in this test.
  const groups = catalogue.executableGroups.map(group => ({ ...group, features: group.features.filter(feature => ['useSPFxPnPList', 'useStableCallback'].includes(feature.title)) })).filter(group => group.features.length);
  function Probe() { navigation = state.useExampleExplorerState(); return h.React.createElement(Explorer, { groups, legacyEntries: [] }); }
  // Griffel requires null for its legacy isomorphic renderer; JSDOM lacks container queries.
  const renderer = createDOMRenderer(null);
  function Root() {
    return h.React.createElement(RendererProvider, { renderer },
      h.React.createElement(FluentProvider, { theme: webLightTheme }, show && (explorer
        ? h.React.createElement(state.ExampleExplorerStateProvider, { initialGroups: groups }, h.React.createElement(Probe))
        : h.React.createElement(Demo, { resetRevision: revision, render: regions => {
          renders++; return h.React.createElement('div', {}, regions.controls, regions.commands, regions.status, regions.children, regions.diagnostics);
        } }))));
  }
  const view = h.mountComponent(Root);
  const flush = async (delay = 20) => h.act(async () => { await new Promise(resolve => setTimeout(resolve, delay)); });
  const rerender = () => h.act(() => { ReactDOM.render(h.React.createElement(Root), view.container); });
  const change = value => h.act(() => {
    const input = view.container.querySelector('input[aria-label="List title"]');
    assert.ok(input); Simulate.change(input, { target: { value } });
  });
  return { ...h, ...view, calls, catalogue, entry, flush, rerender, change,
    get renders() { return renders; }, get navigation() { return navigation; },
    click(text = 'Read selected list') { h.act(() => Simulate.click(button(view.container, text))); },
    reset() { revision++; rerender(); },
    depart() { show = false; rerender(); },
    setInitializationError(error) { initializationError = error; h.context.current = { ...h.context.current }; rerender(); },
    async succeed(index, items) { await h.act(async () => calls[index].resolve(new Response(JSON.stringify({ value: items }), { headers: { 'Content-Type': 'application/json' } }))); await flush(); },
    async fail(index, message) { await h.act(async () => calls[index].reject(new Error(message))); await flush(); },
  };
}

test('read-only host pilot mounts empty at a nonzero revision and reads exactly once only on command', async t => {
  const h = fixture(t); await h.flush();
  assert.equal(h.container.querySelector('input').value, '');
  assert.ok(button(h.container, 'Read selected list').disabled);
  h.click(); h.change('   '); h.click(); await h.flush(); assert.equal(h.calls.length, 0);
  h.change('Team A'); await h.flush(); assert.equal(h.calls.length, 0);
  assert.equal(button(h.container, 'Read selected list').disabled, false);
  h.click(); await h.flush();
  assert.equal(h.calls.length, 1);
  assert.match(decodeURIComponent(h.calls[0].url.pathname), /lists\/getByTitle\('Team A'\)\/items$/i);
  assert.equal(h.calls[0].url.searchParams.get('$select'), 'Id,Title');
  assert.equal(h.calls[0].url.searchParams.get('$top'), '5');
  assert.ok(button(h.container, 'Read selected list').disabled);
  assert.match(h.container.textContent, /Loading.*Team A/);
  h.click(); await h.flush(); assert.equal(h.calls.length, 1);
  await h.succeed(0, [{ Id: 12, Title: 'Visible row' }]);
  assert.match(h.container.textContent, /Results for Team A/);
  assert.match(h.container.textContent, /12.*Visible row/);
  h.change('Team B'); assert.doesNotMatch(h.container.textContent, /Visible row|Results for Team A/);
  assert.equal(h.calls.length, 1);
});

test('read-only pilot separates context initialization from attributed request errors and catches readiness changes', async t => {
  const h = fixture(t);
  h.change('Team A'); h.click(); await h.flush(); await h.fail(0, 'Read denied');
  assert.match(h.container.textContent, /Request error for Team A: Read denied/);
  h.setInitializationError(new Error('Authentication unavailable'));
  assert.match(h.container.textContent, /Context initialization error: Authentication unavailable/);
  assert.ok(button(h.container, 'Read selected list').disabled);
  h.click(); await h.flush(); assert.equal(h.calls.length, 1);
  h.reset();
  assert.match(h.container.textContent, /Context initialization error: Authentication unavailable/);
  assert.doesNotMatch(h.container.textContent, /Request error|Read denied/);
  assert.equal(h.container.querySelector('input').value, '');
  h.setInitializationError(undefined); h.change('Team B'); h.click(); await h.flush();
  await h.succeed(1, []); assert.match(h.container.textContent, /Results for Team B: 0/);
  assert.doesNotMatch(h.container.textContent, /Authentication unavailable|Read denied/);
});

for (const late of ['success', 'error']) test(`late A ${late} is ignored after B selection and after reset`, async t => {
  const h = fixture(t);
  h.change('Team A'); h.click(); await h.flush();
  h.change('Team B'); assert.equal(h.calls.length, 1);
  h.click(); await h.flush(); assert.equal(h.calls.length, 2);
  await h.succeed(1, [{ Id: 2, Title: 'Current B' }]);
  if (late === 'success') await h.succeed(0, [{ Id: 1, Title: 'Obsolete A' }]);
  else await h.fail(0, 'Obsolete A');
  assert.match(h.container.textContent, /Results for Team B/);
  assert.match(h.container.textContent, /Current B/); assert.doesNotMatch(h.container.textContent, /Obsolete A/);
  h.click(); await h.flush(); const count = h.calls.length;
  const input = h.container.querySelector('input'); h.reset();
  assert.equal(h.container.querySelector('input'), input, 'reset retains mounted component');
  assert.equal(input.value, ''); assert.equal(h.calls.length, count);
  if (late === 'success') await h.succeed(2, [{ Id: 3, Title: 'Obsolete reset' }]);
  else await h.fail(2, 'Obsolete reset');
  assert.doesNotMatch(h.container.textContent, /Current B|Obsolete reset|Results for Team B|Request error/);
});

test('context readiness loss during a pending read catches and ignores the late request failure', async t => {
  const h = fixture(t); h.change('Team A'); h.click(); await h.flush();
  h.setInitializationError(new Error('Context changed during read'));
  await h.fail(0, 'Obsolete request failure');
  assert.match(h.container.textContent, /Context initialization error: Context changed during read/);
  assert.doesNotMatch(h.container.textContent, /Obsolete request failure|Request error/);
  assert.ok(button(h.container, 'Read selected list').disabled);
  assert.equal(h.calls.length, 1);
});

test('pending read unmount causes no view update or additional I/O', async t => {
  const h = fixture(t); h.change('Team A'); h.click(); await h.flush(); h.depart();
  const rendered = h.renders, output = h.container.textContent;
  await h.succeed(0, [{ Id: 9, Title: 'Unmounted row' }]);
  assert.equal(h.renders, rendered); assert.equal(h.container.textContent, output); assert.equal(h.calls.length, 1);
});

test('generated Hook navigation, expansion and scenario selection never query the selected list', async t => {
  const h = fixture(t, { explorer: true }); await h.flush();
  assert.equal(h.catalogue.executableExamples.length, 10);
  assert.equal(h.catalogue.executableGroups.flatMap(group => group.features).length, 5);
  assert.deepEqual(h.entry.metadata, { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Choose an existing list and run the read explicitly; loading, results and errors identify the selected list.' });
  const generic = h.catalogue.executableExamples.find(entry => entry.featureTitle === 'useStableCallback');
  h.act(() => h.navigation.selectFeature(generic.featureKey)); await h.flush();
  assert.equal(h.container.querySelector('input[aria-label="List title"]'), null);
  h.act(() => h.navigation.toggleGroup('hooks')); h.act(() => h.navigation.toggleGroup('hooks'));
  h.act(() => h.navigation.selectFeature(h.entry.featureKey)); await h.flush();
  h.act(() => h.navigation.selectScenario(h.entry.featureKey, h.entry.id)); await h.flush();
  h.change('Team A'); await h.flush(); assert.equal(h.calls.length, 0);
  h.click(); await h.flush(); assert.equal(h.calls.length, 1);
  h.click('Ripristina scenario'); assert.equal(h.container.querySelector('input[aria-label="List title"]').value, '');
  await h.fail(0, 'Late navigation error'); assert.doesNotMatch(h.container.textContent, /Late navigation error/);
  assert.equal(h.calls.length, 1);
  h.change('Team B'); h.click(); await h.flush();
  h.act(() => h.navigation.selectFeature(generic.featureKey)); await h.flush();
  await h.succeed(1, [{ Id: 3, Title: 'Departed scenario result' }]);
  assert.doesNotMatch(h.container.textContent, /Departed scenario result/);
  assert.equal(h.calls.length, 2);
});


test('browser fixture identifies simulation and uses only its read transport with real source hooks', async t => {
  const h = fixture(t, { simulatedBrowserTransport: true });
  const transport = window.pnpReadFixture;
  assert.equal(transport.simulation, true);
  h.change('Chosen list'); await h.flush(); assert.equal(transport.calls.length, 0);
  h.click(); await h.flush(90);
  assert.equal(transport.calls.length, 1); assert.equal(transport.calls[0].method, 'GET');
  assert.match(h.container.textContent, /Results for Chosen list: 1/);
  assert.match(h.container.textContent, /Simulated list row \(no tenant request\)/);
  transport.configure({ failure: 'Read denied' }); h.click(); await h.flush(90);
  assert.match(h.container.textContent, /Request error for Chosen list: Simulated request error: Read denied/);
  transport.configure({ hold: true }); h.click(); await h.flush();
  h.reset(); transport.release(); await h.flush();
  assert.doesNotMatch(h.container.textContent, /Results for Chosen list|Request error/);
  assert.equal(transport.calls.length, 3); assert.equal(h.calls.length, 0);
});

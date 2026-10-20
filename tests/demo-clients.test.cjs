const assert = require('node:assert/strict');
const path = require('node:path');
const { test } = require('node:test');
const { Simulate } = require('react-dom/test-utils');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const { createHarness } = require('./services-test-harness.cjs');
const panel = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/ClientsPanel.tsx');

function fixture(t, invoke) {
  const calls = [];
  const boundaries = { '@microsoft/sp-http': { HttpClient: { configurations: { v1: {} } }, SPHttpClient: { configurations: { v1: {} } }, AadHttpClient: { configurations: { v1: {} } } } };
  const h = createHarness(boundaries);
  boundaries['../shared'] = Object.assign({}, ...['DemoCard', 'InfoGrid', 'JsonDetails', 'StatusBadge']
    .map(name => h.load(path.resolve(path.dirname(panel), '../shared', name + '.tsx'))));
  const client = {
    isReady: true, isLoading: false, isInitializing: false,
    invoke: invoke || (async () => { calls.push('invoke'); return { title: 'Visible response', completed: false, userId: 8, nested: { diagnosticOnly: 'raw nested payload' } }; }),
  };
  boundaries['@apvee/spfx-react-toolkit'] = {
    useSPFxHttpClient: () => client,
    useSPFxSPHttpClient: () => ({ ...client, baseUrl: 'https://tenant.test/sites/sample' }),
    useSPFxMSGraphClient: () => client,
    useSPFxAadHttpClient: () => ({ ...client, resourceUrl: undefined, setResourceUrl: () => {} }),
    useSPFxAadTokenProvider: () => client,
    useSPFxApiPermissionPrecheck: () => ({ configurationState: 'configured', isConfigured: true, isChecking: false, results: [], missing: [], warnings: [], check: () => {}, retryWithoutCache: () => {} }),
  };
  window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  t.after(h.close);
  const Panel = h.load(panel).default;
  const view = h.mountComponent(() => h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Panel)));
  return { ...h, ...view, calls, client, refresh: () => h.act(() => require('react-dom').render(h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Panel)), view.container)) };
}

test('client response metadata stays visible while raw payload diagnostics start collapsed', async t => {
  const f = fixture(t);
  assert.deepEqual(f.calls, [], 'mount must not invoke a service');
  assert.equal(f.container.querySelector('[aria-label$=" preview"]'), null);
  const section = f.container.querySelector('[data-demo-section][aria-label="HttpClient"]');
  const commands = section.querySelector('[aria-label="HttpClient commands"]');
  assert.equal(commands.querySelectorAll('input').length, 0);
  assert.ok(section.querySelector('[aria-label="HttpClient results"]'), 'current result region remains available before execution');
  assert.equal(section.querySelector('[aria-label="HttpClient diagnostics"]'), null, 'no raw response means no empty diagnostics');
  await f.act(async () => { Simulate.click(commands.querySelector('button')); await Promise.resolve(); });
  assert.deepEqual(f.calls, ['invoke']);
  const status = section.querySelector('[aria-label="HttpClient results"]');
  assert.match(status.textContent, /Visible response/);
  assert.match(status.textContent, /false/);
  assert.match(status.textContent, /8/);
  assert.ok(!status.textContent.includes('raw nested payload'));
  const diagnosticTrigger = [...section.querySelectorAll('button')].find(node => node.textContent === 'Advanced diagnostics');
  assert.equal(diagnosticTrigger.getAttribute('aria-expanded'), 'false');
  f.act(() => Simulate.click(diagnosticTrigger));
  assert.equal(diagnosticTrigger.getAttribute('aria-expanded'), 'true');
  assert.match(section.querySelector('[aria-label="HttpClient diagnostics"]').textContent, /raw nested payload/);
});

test('client selection mounts one current result and performs no invocations', t => {
  const f = fixture(t);
  for (const value of ['status', 'http', 'sharepoint', 'graph', 'aad', 'permissions']) {
    f.act(() => Simulate.change(f.container.querySelector('select[data-demo-variant="ClientsPanel"]'), { target: { value } }));
    assert.equal(f.container.querySelectorAll('[data-demo-section]').length, 1);
    assert.equal(f.container.querySelectorAll('[data-demo-configuration]').length, 1);
    assert.equal(f.container.querySelectorAll('[data-demo-result]').length, 1);
    assert.deepEqual(f.calls, []);
  }
});
test('client completion from an obsolete service identity cannot publish into the current result', async t => {
  let resolve;
  const pending = new Promise(done => { resolve = done; });
  const f = fixture(t, () => pending);
  f.act(() => Simulate.click(f.container.querySelector('[aria-label="HttpClient commands"] button')));
  f.client.invoke = async () => ({ title: 'Current identity' });
  f.refresh();
  await f.act(async () => { resolve({ title: 'Obsolete identity' }); await pending; });
  assert.ok(!f.container.textContent.includes('Obsolete identity'));
  await f.act(async () => { Simulate.click(f.container.querySelector('[aria-label="HttpClient commands"] button')); await Promise.resolve(); });
  assert.match(f.container.querySelector('[data-demo-result]').textContent, /Current identity/);
});

const assert = require('node:assert/strict');
const path = require('node:path');
const { test } = require('node:test');
const { Simulate } = require('react-dom/test-utils');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const { createHarness } = require('./services-test-harness.cjs');
const panel = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/TenantPanel.tsx');
function deferred() { let resolve; const promise = new Promise(done => { resolve = done; }); return { promise, resolve }; }
function fixture(t, get) {
  const calls = [], host = { client: { name: 'A' } }, mounted = { current: true };
  const discover = async () => 'https://tenant.test/catalog';
  const boundaries = {
    './useAppCatalogUrl.internal': { useAppCatalogUrl: () => ({ spHttpClient: host.client, discoverAppCatalogUrl: discover, isMountedRef: mounted }) },
    '../services/spfx-tenant-property.service': { createSPFxTenantPropertyService: client => ({ get: async key => { calls.push([client.name, key]); return get ? get(client, key) : { value: key.endsWith('version') ? client.name : 7 }; } }) },
  };
  const h = createHarness(boundaries);
  boundaries['../shared'] = Object.assign({}, ...['ActionResult', 'DemoCard', 'InfoGrid', 'JsonDetails', 'StatusBadge'].map(name => h.load(path.resolve(path.dirname(panel), '../shared', name + '.tsx'))));
  boundaries['@apvee/spfx-react-toolkit'] = { useSPFxTenantProperty: h.load('hooks/useSPFxTenantProperty.ts').useSPFxTenantProperty, useSPFxTenantKeyValueStore: () => { throw new Error('Inactive store must not mount'); } };
  window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  t.after(h.close);
  const Panel = h.load(panel).default;
  const render = () => h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Panel));
  const view = h.mountComponent(render);
  const execute = async () => h.act(async () => { Simulate.click([...view.container.querySelectorAll('button')].find(button => button.textContent === 'Run')); for (let i = 0; i < 8; i++) await Promise.resolve(); });
  return { ...h, ...view, host, calls, execute, refresh: () => h.act(() => { require('react-dom').render(h.React.createElement(render), view.container); }), result: () => view.container.querySelector('[data-demo-result]').textContent };
}
test('Tenant Properties completed execution is invalidated when real property hooks change host identity', async t => {
  const f = fixture(t);
  assert.deepEqual(f.calls, []);
  await f.execute();
  assert.match(f.result(), /spfx-toolkit-test-version:A/);
  f.host.client = { name: 'B' }; f.refresh();
  assert.match(f.result(), /Run to read/);
  assert.ok(!f.result().includes('(not set)'));
  assert.equal(f.calls.length, 2, 'changing identity never reads');
  await f.execute();
  assert.match(f.result(), /spfx-toolkit-test-version:B/);
});
test('Tenant Properties obsolete pending execution cannot mark a new host as read or replace its newer result', async t => {
  const pending = deferred();
  const f = fixture(t, (client, key) => client.name === 'A' ? pending.promise : { value: key.endsWith('version') ? 'Current B' : 9 });
  await f.execute();
  f.host.client = { name: 'B' }; f.refresh();
  assert.match(f.result(), /Run to read/);
  assert.ok(!f.result().includes('(not set)'));
  await f.execute();
  await f.act(async () => { pending.resolve({ value: 'Obsolete A' }); await pending.promise; });
  assert.match(f.result(), /Current B/);
  assert.ok(!f.result().includes('Obsolete A'));
});

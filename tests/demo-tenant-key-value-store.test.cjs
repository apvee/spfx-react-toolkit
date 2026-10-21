const assert = require('node:assert/strict');
const path = require('node:path');
const { test } = require('node:test');
const { Simulate } = require('react-dom/test-utils');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const { createHarness } = require('./services-test-harness.cjs');
const panel = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/TenantPanel.tsx');

function fixture(t, overrides = {}) {
  const calls = [], values = new Map(), client = {}, isMountedRef = { current: true };
  const discoverAppCatalogUrl = async () => 'https://tenant.test/catalog';
  const checkWritePermission = async () => true;
  const service = {
    ensureListReady: async () => { throw new Error('Demo reads must never provision'); },
    get: async key => values.get(key), list: async () => [...values.values()],
    save: async (key, value, catalog, description) => {
      calls.push({ key, value, catalog, description });
      values.set(key, { key, value: typeof value === 'bigint' ? String(value) : value, description, id: 1 });
    }, remove: async key => { values.delete(key); }, ...overrides
  };
  const boundaries = {
    './useAppCatalogUrl.internal': { useAppCatalogUrl: () => ({ spHttpClient: client, discoverAppCatalogUrl, checkWritePermission, isMountedRef }) },
    '../services/spfx-tenant-key-value-store.service': { createSPFxTenantKeyValueStoreService: () => service }
  };
  const h = createHarness(boundaries);
  boundaries['../shared'] = Object.assign({}, ...['ActionResult', 'DemoCard', 'InfoGrid', 'JsonDetails', 'StatusBadge']
    .map(name => h.load(path.resolve(path.dirname(panel), '../shared', name + '.tsx'))));
  boundaries['@apvee/spfx-react-toolkit'] = {
    useSPFxTenantKeyValueStore: h.load('hooks/useSPFxTenantKeyValueStore.ts').useSPFxTenantKeyValueStore,
    useSPFxTenantProperty: () => ({ load: async () => {}, isLoading: false })
  };
  t.after(h.close);
  window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  t.mock.method(window, 'setTimeout', (callback, delay) => setTimeout(callback, delay));
  const Panel = h.load(panel).default;
  const view = h.mountComponent(() => h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Panel)));
  function select(selector, value) {
    const element = view.container.querySelector(selector); assert.ok(element, selector);
    h.act(() => Simulate.change(element, { target: { value } }));
  }
  select('[data-demo-variant="TenantPanel"]', 'store');
  const operation = value => select('[data-demo-variant="TenantStoreScenario"]', value);
  const input = (index, value) => h.act(() => Simulate.change(view.container.querySelectorAll('input')[index], { target: { value } }));
  const run = async () => h.act(async () => {
    const button = [...view.container.querySelectorAll('button')].find(element => element.textContent === 'Run');
    assert.equal(button.disabled, false); Simulate.click(button);
    for (let i = 0; i < 12; i++) await Promise.resolve();
  });
  return { ...view, calls, values, operation, input, run, select };
}

test('tenant demo saves empty text, explicitly clears description and displays the read value type', async t => {
  const f = fixture(t); assert.deepEqual(f.calls, []);
  f.operation('save'); f.input(2, ''); await f.run();
  assert.equal(f.calls[0].value, ''); assert.equal(f.calls[0].description, '');
  f.operation('get'); await f.run();
  assert.match(f.container.textContent, /Value type.*string/);
});

test('tenant demo passes bigint without precision loss and makes its string readback observable', async t => {
  const f = fixture(t); f.operation('save');
  f.select('[data-tenant-store-format]', 'bigint'); f.input(1, '9007199254740993'); await f.run();
  assert.equal(f.calls[0].value, 9007199254740993n);
  f.operation('get'); await f.run();
  assert.match(f.container.textContent, /9007199254740993/); assert.match(f.container.textContent, /Value type.*string/);
});

test('tenant demo parses JSON and Date inputs and reports invalid input before writing', async t => {
  const f = fixture(t); f.operation('save');
  f.select('[data-tenant-store-format]', 'json'); f.input(1, '{"enabled":true}'); await f.run();
  assert.deepEqual(f.calls[0].value, { enabled: true });
  f.select('[data-tenant-store-format]', 'date'); f.input(1, '2026-01-02T03:04:05Z'); await f.run();
  assert.equal(f.calls[1].value.toISOString(), '2026-01-02T03:04:05.000Z');
  f.select('[data-tenant-store-format]', 'json'); f.input(1, '{invalid'); await f.run();
  assert.equal(f.calls.length, 2); assert.doesNotMatch(f.container.textContent, /Saved spfx/);
});

for (const operation of ['get', 'list', 'save', 'remove']) {
  test(`tenant demo ${operation} displays failure without successful fallback messaging`, async t => {
    const f = fixture(t, { [operation]: async () => { throw new Error('Access denied'); } });
    f.operation(operation); await f.run();
    assert.match(f.container.textContent, /Access denied/);
    assert.doesNotMatch(f.container.textContent, /not found\.|Loaded 0 tenant|Saved spfx|Removed spfx/);
  });
}

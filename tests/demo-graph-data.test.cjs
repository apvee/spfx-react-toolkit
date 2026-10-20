const assert = require('node:assert/strict');
const path = require('node:path');
const { test } = require('node:test');
const { Simulate } = require('react-dom/test-utils');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const { createHarness } = require('./services-test-harness.cjs');
const panel = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/GraphDataPanel.tsx');
function fixture(t, options = {}) {
  const calls = [];
  const graph = { client: {}, isReady: true, isInitializing: false, ...options.graph };
  let settle;
  const boundaries = {
    './useSPFxMSGraphClient': { useSPFxMSGraphClient: () => graph },
    '../services/spfx-onedrive-app-data.service': { createSPFxOneDriveAppDataService: client => options.serviceFactory ? options.serviceFactory(client, calls) : ({ read: async (file, folder) => { calls.push(['read', file, folder]); return { isNotFound: true }; }, write: async (...args) => { calls.push(['write', ...args]); await new Promise(resolve => { settle = resolve; }); } }) },
  };
  const h = createHarness(boundaries);
  boundaries['../shared'] = Object.assign({}, ...['DemoCard', 'InfoGrid', 'JsonDetails'].map(name => h.load(path.resolve(path.dirname(panel), '../shared', name + '.tsx'))));
  boundaries['@apvee/spfx-react-toolkit'] = { useSPFxMSGraphClient: () => graph, useSPFxOneDriveAppData: h.load('hooks/useSPFxOneDriveAppData.ts').useSPFxOneDriveAppData, useSPFxUserInfo: () => ({ displayName: 'Sample user' }), useSPFxUserPhoto: () => ({ isReady: true, reload: async () => {} }) };
  window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  t.after(h.close);
  const Panel = h.load(panel).default;
  const Component = () => h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Panel));
  const view = h.mountComponent(Component);
  const choose = (marker, value) => h.act(() => Simulate.change(view.container.querySelector(`select[data-demo-variant="${marker}"]`), { target: { value } }));
  const execute = async () => h.act(async () => { Simulate.click([...view.container.querySelectorAll('button')].find(button => button.textContent === 'Run')); for (let i = 0; i < 8; i++) await Promise.resolve(); });
  return { ...h, ...view, graph, calls, choose, execute, refresh: () => h.act(() => { require('react-dom').render(h.React.createElement(Component), view.container); }), settle: async () => h.act(async () => { settle(); await Promise.resolve(); }) };
}
test('OneDrive missing-file read does not create; selector changes never invoke services and explicit write locks its target', async t => {
  const f = fixture(t);
  assert.deepEqual(f.calls, []);
  f.choose('GraphDataPanel', 'onedrive');
  assert.deepEqual(f.calls, [], 'selecting OneDrive must not read or write');
  assert.equal(f.container.querySelectorAll('[data-demo-section]').length, 1);
  await f.execute();
  assert.deepEqual(f.calls, [['read', 'test-data.json', 'spfx-react-toolkit-demo']]);
  assert.match(f.container.querySelector('[data-demo-result]').textContent, /File not found/);
  f.choose('OneDriveScenario', 'write');
  assert.equal(f.calls.length, 1);
  assert.ok(!f.container.querySelector('[data-demo-result]').textContent.includes('File not found'));
  await f.execute();
  assert.equal(f.calls.length, 2);
  assert.equal(f.calls[1][0], 'write');
  assert.equal(f.container.querySelector('select[data-demo-variant="GraphDataPanel"]').disabled, true);
  assert.equal(f.container.querySelector('select[data-demo-variant="OneDriveScenario"]').disabled, true);
  assert.equal(f.container.querySelector('input').disabled, true);
  await f.settle();
  assert.equal(f.container.querySelector('select[data-demo-variant="GraphDataPanel"]').disabled, false);
});

function deferred() { let resolve, reject; const promise = new Promise((a, b) => { resolve = a; reject = b; }); return { promise, resolve, reject }; }
for (const isInitializing of [false, true]) {
  test(`unavailable OneDrive client (${isInitializing ? 'initializing' : 'missing'}) cannot execute or claim a persisted default`, async t => {
    const f = fixture(t, { graph: { client: undefined, isReady: false, isInitializing } });
    f.choose('GraphDataPanel', 'onedrive');
    const button = [...f.container.querySelectorAll('button')].find(button => button.textContent === 'Run');
    assert.equal(button.disabled, true);
    await f.execute(); // Simulated dispatch also tests the defensive callback guard.
    assert.deepEqual(f.calls, []);
    const result = f.container.querySelector('[data-demo-result]').textContent;
    assert.match(result, /Graph unavailable/);
    assert.ok(!result.includes('Saved message'));
    assert.ok(!result.includes('Counter'));
  });
}
test('obsolete OneDrive write rejection cannot replace a newer read on another real-hook service identity', async t => {
  const pending = deferred();
  const f = fixture(t, { serviceFactory: (client, calls) => ({
    read: async () => { calls.push(['read', client]); return { isNotFound: false, data: { message: 'Current service read', counter: 4, timestamp: 1 } }; },
    write: async () => { calls.push(['write', client]); return pending.promise; },
  }) });
  f.choose('GraphDataPanel', 'onedrive'); f.choose('OneDriveScenario', 'write'); await f.execute();
  const oldClient = f.graph.client;
  f.graph.client = {}; f.refresh();
  f.choose('OneDriveScenario', 'read'); await f.execute();
  assert.match(f.container.querySelector('[data-demo-result]').textContent, /Current service read/);
  await f.act(async () => { pending.reject(new Error('Obsolete write failed')); await pending.promise.catch(() => {}); });
  assert.ok(!f.container.textContent.includes('Obsolete write failed'));
  assert.match(f.container.querySelector('[data-demo-result]').textContent, /Current service read/);
  assert.equal(f.calls[0][1], oldClient);
  assert.notEqual(f.calls[1][1], oldClient);
});
test('obsolete successful OneDrive write cannot erase a draft edited after the Graph identity changes', async t => {
  const pending = deferred();
  const f = fixture(t, { serviceFactory: () => ({ read: async () => ({ isNotFound: true }), write: () => pending.promise }) });
  f.choose('GraphDataPanel', 'onedrive'); f.choose('OneDriveScenario', 'write'); await f.execute();
  f.graph.client = {}; f.refresh();
  f.act(() => Simulate.change(f.container.querySelector('input'), { target: { value: 'Draft for the new service' } }));
  await f.act(async () => { pending.resolve(); await pending.promise; });
  assert.equal(f.container.querySelector('input').value, 'Draft for the new service');
  assert.ok(!f.container.querySelector('[data-demo-result]').textContent.includes('Saved message'));
});

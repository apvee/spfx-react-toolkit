const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const panel = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/RuntimePanel.tsx');
function fixture(t, time = async () => ({ result: 'Complete', durationMs: 1 })) {
  const boundaries = { '../SpFxReactToolkitTest.module.scss': { default: {} } };
  const h = createHarness(boundaries);
  boundaries['../shared'] = Object.assign({}, ...['ActionResult', 'DemoCard', 'InfoGrid'].map(name => h.load(path.resolve(path.dirname(panel), '../shared', `${name}.tsx`))));
  const storage = { value: '', setValue: () => {}, remove: () => {} };
  boundaries['@apvee/spfx-react-toolkit'] = {
    useSPFxProperties: () => ({ properties: {}, setProperties: () => {}, updateProperties: () => {} }),
    useSPFxSessionStorage: () => storage, useSPFxLocalStorage: () => storage,
    useSPFxLogger: () => ({ info: () => {}, warn: () => {}, error: () => {} }),
    useSPFxPerformance: () => ({ time }),
  };
  window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  const timers = new Map(); let next = 100000;
  const originalSet = window.setTimeout.bind(window), originalClear = window.clearTimeout.bind(window);
  t.mock.method(window, 'setTimeout', (callback, delay, ...args) => {
    if (delay !== 3000) return originalSet(callback, delay, ...args);
    const id = ++next; timers.set(id, callback); return id;
  });
  t.mock.method(window, 'clearTimeout', id => { if (timers.has(id)) timers.delete(id); else originalClear(id); });
  const Component = h.load(panel).default;
  const view = h.mountComponent(() => h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Component)));
  t.after(h.close);
  const click = text => { const button = [...view.container.querySelectorAll('button')].find(node => node.textContent === text); assert.ok(button); h.act(() => Simulate.click(button)); };
  return { ...h, ...view, timers, click };
}
test('Runtime replaces the notification timeout and cancels it when its tab unmounts', t => {
  const f = fixture(t);
  f.click('setProperties');
  const first = [...f.timers.keys()][0]; assert.equal(f.timers.size, 1);
  f.click('updateProperties');
  assert.equal(f.timers.size, 1, 'an old expiry must not clear the new message');
  assert.ok(!f.timers.has(first), 'replaced result clears its old timeout');
  assert.match(f.container.textContent, /Properties updated with updateProperties/);
  f.unmount();
  assert.equal(f.timers.size, 0, 'tab unmount cancels notification publication');
});
test('Runtime pending performance completion cannot publish after tab unmount', async t => {
  let resolve;
  const pending = new Promise(done => { resolve = done; });
  const f = fixture(t, () => pending);
  const errors = t.mock.method(console, 'error', () => {});
  f.act(() => Simulate.change(f.container.querySelector('select'), { target: { value: 'performance' } }));
  f.click('Run timed operation'); f.unmount();
  await f.act(async () => { resolve({ result: 'Deferred complete', durationMs: 12 }); await pending; });
  assert.equal(errors.mock.callCount(), 0, 'completion after unmount emits no React state warning');
});

test('Runtime feedback stays in its owning configuration card and moves without duplicate messages', t => {
  const f = fixture(t);
  const properties = f.container.querySelector('[data-demo-section][aria-label="Properties"] [data-demo-configuration]');
  f.click('setProperties');
  assert.match(properties.textContent, /Properties updated with setProperties/);
  f.act(() => Simulate.change(f.container.querySelector('select'), { target: { value: 'storage' } }));
  const storage = f.container.querySelector('[data-demo-section][aria-label="Scoped Storage"] [data-demo-configuration]');
  assert.equal(f.container.querySelectorAll('[data-demo-section]').length, 1);
  const input = f.container.querySelector('input[aria-label="Session value"]');
  f.act(() => Simulate.change(input, { target: { value: 'session demo' } }));
  f.click('Save Session');
  assert.match(storage.textContent, /Session storage value saved/);
  assert.equal(properties.isConnected, false);
  assert.equal(f.container.querySelectorAll('.fui-MessageBar').length, 1, 'only the current owner displays feedback');
  const expiry = [...f.timers.values()][0];
  f.act(() => expiry());
  assert.ok(!storage.textContent.includes('Session storage value saved'));
});

test('Logger replaces the visible operation result and keeps history in initially collapsed diagnostics', t => {
  const f = fixture(t);
  f.act(() => Simulate.change(f.container.querySelector('select[data-demo-variant="RuntimePanel"]'), { target: { value: 'logger' } }));
  f.click('Info');
  assert.match(f.container.querySelector('[data-demo-result]').textContent, /demo info/);
  f.click('Warning');
  const result = f.container.querySelector('[data-demo-result]').textContent;
  assert.match(result, /demo warning/);
  assert.ok(!result.includes('demo info'));
  const diagnostics = [...f.container.querySelectorAll('button')].find(button => button.textContent === 'Advanced diagnostics');
  assert.equal(diagnostics.getAttribute('aria-expanded'), 'false');
  f.act(() => Simulate.click(diagnostics));
  const history = f.container.querySelector('[aria-label="Logger diagnostics"]').textContent;
  assert.match(history, /demo info/);
  assert.match(history, /demo warning/);
});

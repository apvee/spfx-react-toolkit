const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const demo = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/demos/PnPListDemo.tsx');
function fixture(t, remove = async () => {}) {
  const calls = [];
  const previousConfirm = Object.getOwnPropertyDescriptor(global, 'confirm');
  Object.defineProperty(global, 'confirm', { configurable: true, value: () => false });
  t.after(() => { if (previousConfirm) Object.defineProperty(global, 'confirm', previousConfirm); else delete global.confirm; });
  let publishRows;
  const boundaries = { '@apvee/spfx-react-toolkit': { useSPFxPnPList: () => {
    const [items, setItems] = h.React.useState([{ Id: 42, Title: 'Intended item' }, { Id: 73, Title: 'Other item' }]);
    publishRows = id => setItems(current => current.filter(item => item.Id !== id));
    return { items, loading: false, isEmpty: false, hasMore: false, query: async () => {}, remove: id => { calls.push(id); return remove(id); } };
  } } };
  const h = createHarness(boundaries);
  for (const name of ['Node', 'NodeFilter', 'HTMLElement', 'Element']) {
    const previous = Object.getOwnPropertyDescriptor(global, name);
    Object.defineProperty(global, name, { configurable: true, value: window[name] });
    t.after(() => { if (previous) Object.defineProperty(global, name, previous); else delete global[name]; });
  }
  Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  const Component = h.load(demo).PnPListDemo;
  const view = h.mountComponent(() => h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Component)));
  t.after(h.close);
  const buttons = name => [...document.querySelectorAll('button')].filter(node => node.textContent === name);
  const click = button => h.act(() => Simulate.click(button));
  const open = () => { const trigger = buttons('Delete')[0]; trigger.focus(); click(trigger); return trigger; };
  return { ...h, ...view, calls, buttons, click, open, publishRows: id => h.act(() => publishRows(id)) };
}
test('PnP delete opens Fluent consent and cancellation restores the intended trigger without removal', t => {
  const f = fixture(t); const trigger = f.open();
  assert.ok(document.querySelector('[role="dialog"],[role="alertdialog"]'), 'Fluent confirmation surface must open');
  assert.deepEqual(f.calls, []);
  f.click(f.buttons('Cancel')[0]);
  assert.deepEqual(f.calls, []);
  assert.equal(document.activeElement, trigger);
});
test('PnP consent captures intended ID and approves once even before removal settles', async t => {
  let finish; const pending = new Promise(resolve => { finish = resolve; });
  const f = fixture(t, () => pending); f.open();
  assert.deepEqual(f.calls, []);
  const approve = f.buttons('Delete item')[0]; assert.ok(approve, 'explicit approval action');
  f.act(() => { Simulate.click(approve); Simulate.click(approve); });
  assert.deepEqual(f.calls, [42]);
  await f.act(async () => { finish(); await pending; });
  assert.deepEqual(f.calls, [42]);
});
test('PnP deletion retains original error reporting', async t => {
  const failure = new Error('Mock removal failed');
  const errors = t.mock.method(console, 'error', () => {});
  const f = fixture(t, async () => { throw failure; }); f.open();
  assert.ok(f.buttons('Delete item')[0], 'explicit approval action');
  await f.act(async () => { f.click(f.buttons('Delete item')[0]); });
  assert.deepEqual(f.calls, [42]);
  assert.ok(errors.mock.calls.some(call => call.arguments[0] === 'Delete error:' && call.arguments[1] === failure));
});


test('PnP approval retains stable focus when removal resolves before later refetch row publication', async t => {
  const f = fixture(t);
  const trigger = f.open();
  await f.act(async () => { f.click(f.buttons('Delete item')[0]); });
  assert.deepEqual(f.calls, [42]);
  assert.equal(trigger.isConnected, true, 'real remove resolves while row is still rendered');
  f.publishRows(42);
  assert.equal(trigger.isConnected, false, 'later debounced query publication removes the row');
  assert.equal(document.activeElement === f.container.querySelector('input[placeholder="e.g., Site Pages"]'), true, 'approval focus remains on a stable list control after deferred row removal');
});

test('PnP late removal and refetch preserve the focus chosen by subsequent user navigation', async t => {
  let finish; const pending = new Promise(resolve => { finish = resolve; });
  const f = fixture(t, () => pending); f.open();
  f.click(f.buttons('Delete item')[0]);
  const destination = f.buttons('Edit')[1]; destination.focus();
  await f.act(async () => { finish(); await pending; });
  assert.equal(document.activeElement === destination, true, 'removal completion must not steal deliberate focus');
  f.publishRows(42);
  assert.equal(document.activeElement === destination, true, 'later refetch must not steal deliberate focus');
  assert.deepEqual(f.calls, [42]);
});

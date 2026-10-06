const assert = require('node:assert/strict');
const path = require('node:path');
const { test } = require('node:test');
const { Simulate } = require('react-dom/test-utils');
const { createHarness } = require('./services-test-harness.cjs');
require('@fluentui/react').registerIcons({ icons: { TableGroup: 'table', Info: 'info', InfoSolid: 'info', Success: 'success', ErrorBadge: 'error', Completed: 'done', CheckMark: 'check', Cancel: 'cancel', BulletedList: 'list', Search: 'search', Save: 'save', Delete: 'delete' } });
const panel = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/SitePanel.tsx');
function deferred() { let resolve; const promise = new Promise(r => { resolve = r; }); return { promise, resolve }; }
function fixture(t, overrides = {}) {
  const calls = [];
  const service = {
    canCurrentUserWrite: async () => false,
    get: async (...args) => { calls.push(['get', ...args]); return undefined; },
    list: async (...args) => { calls.push(['list', ...args]); return []; },
    save: async (...args) => { calls.push(['save', ...args]); },
    remove: async (...args) => { calls.push(['remove', ...args]); },
    ...overrides,
  };
  const pageContext = { site: { absoluteUrl: 'https://tenant.test/sites/root' }, web: { absoluteUrl: 'https://tenant.test/sites/root/sub' } };
  const boundaries = {
    './useSPFxPageContext': { useSPFxPageContext: () => pageContext },
    './useSPFxSPHttpClient': { useSPFxSPHttpClient: () => ({ client }) },
    '../services/spfx-site-key-value-store.service': { createSPFxSiteKeyValueStoreService: () => service },
    '../SpFxReactToolkitTest.module.scss': { default: {} },
  };
  const client = {};
  const h = createHarness(boundaries);
  // Resolve the existing shared components explicitly: this harness does not resolve directory indexes.
  boundaries['../shared'] = Object.assign({}, ...['ActionResult', 'DemoCard', 'InfoGrid', 'JsonDetails', 'StatusBadge']
    .map(name => h.load(path.resolve(path.dirname(panel), '../shared', name + '.tsx'))));
  boundaries['@apvee/spfx-react-toolkit'] = {
    useSPFxSiteKeyValueStore: h.load('hooks/useSPFxSiteKeyValueStore.ts').useSPFxSiteKeyValueStore,
    useSPFxPageContext: () => pageContext,
  };
  t.after(h.close);
  // Fluent DelayedRender uses window.setTimeout and global clearTimeout; share IDs as browsers do.
  t.mock.method(window, 'setTimeout', (callback, delay) => setTimeout(callback, delay));
  const Component = h.load(panel).default;
  const view = h.mountComponent(Component);
  const flush = async () => h.act(async () => { for (let i = 0; i < 8; i++) await Promise.resolve(); await new Promise(resolve => setTimeout(resolve, 0)); });
  function click(text) {
    const button = Array.from(view.container.querySelectorAll('button')).find(b => b.querySelector('.ms-Button-label')?.textContent === text);
    assert.ok(button, text + ' action exists');
    assert.equal(button.disabled, false, text + ' action is available');
    h.act(() => Simulate.click(button));
  }
  function rerender() { h.act(() => { require('react-dom').render(h.React.createElement(Component), view.container); }); }
  return { ...h, ...view, calls, pageContext, flush, click, rerender };
}
for (const [action, method] of [['Save', 'save'], ['Remove', 'remove']]) {
  test(action + ' catches rejected writes and displays failure without success', async t => {
    const f = fixture(t, { [method]: async () => { throw new Error('Write denied'); } });
    f.click(action); await f.flush();
    assert.ok(f.container.textContent.includes('Write denied'));
    assert.ok(!f.container.textContent.includes(action === 'Save' ? 'Saved ' : 'Removed '));
  });
}
for (const [action, method, success] of [['Get', 'get', 'not found'], ['List', 'list', 'Loaded 0']]) {
  test(action + ' fallback reports current read error and never later becomes success', async t => {
    let fail = true;
    const f = fixture(t, { [method]: async () => { if (fail) throw new Error('Read denied'); return method === 'list' ? [] : undefined; } });
    f.click(action); await f.flush();
    assert.ok(f.container.textContent.includes('Read denied'));
    assert.ok(!f.container.textContent.includes(success));
    fail = false;
    f.click(action); await f.flush();
    assert.ok(!f.container.textContent.includes('Read denied'));
    assert.ok(f.container.textContent.includes(success));
  });
}
test('mount performs no CRUD and manual Save accepts an empty string at the collection root', async t => {
  const f = fixture(t); await f.flush();
  assert.deepEqual(f.calls, []);
  f.click('Save'); await f.flush();
  assert.equal(f.calls.length, 1);
  assert.match(f.calls[0][1], /^spfx-toolkit-site-demo-/);
  assert.equal(f.calls[0][2], '');
  assert.equal(f.calls[0][3], f.pageContext.site.absoluteUrl);
  assert.ok(f.container.textContent.includes('Saved '));
});
test('pending action cannot publish into a changed collection', async t => {
  const pending = deferred();
  const f = fixture(t, { get: () => pending.promise });
  f.click('Get');
  f.pageContext.site.absoluteUrl = 'https://tenant.test/sites/other'; f.rerender();
  await f.act(async () => { pending.resolve({ key: 'obsolete', value: 'obsolete', id: 1 }); await pending.promise; });
  assert.ok(!f.container.textContent.includes('Found '));
  assert.ok(!f.container.textContent.includes('obsolete'));
});
test('reader Get and List display actual public-hook data from the collection root', async t => {
  const item = { key: 'sample-result', value: 'persisted-value', description: 'saved metadata', id: 12 };
  const f = fixture(t, {
    get: async (...args) => { f.calls.push(['get', ...args]); return item; },
    list: async (...args) => { f.calls.push(['list', ...args]); return [item]; },
  });
  f.click('Get'); await f.flush();
  assert.ok(f.container.textContent.includes('Found '));
  assert.ok(f.container.textContent.includes('persisted-value'));
  assert.equal(f.calls[0][2], f.pageContext.site.absoluteUrl);
  f.click('List'); await f.flush();
  assert.ok(f.container.textContent.includes('Loaded 1'));
  assert.ok(f.container.textContent.includes('persisted-value'));
  assert.equal(f.calls[1][1], f.pageContext.site.absoluteUrl);
});
test('unmount ignores pending action completion without React warnings', async t => {
  const pending = deferred();
  const f = fixture(t, { list: () => pending.promise });
  f.click('List'); f.unmount();
  const errors = t.mock.method(console, 'error', () => {});
  await f.act(async () => { pending.resolve([]); await pending.promise; });
  assert.equal(errors.mock.callCount(), 0);
});

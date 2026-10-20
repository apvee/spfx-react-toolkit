const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { loadHook } = require('./async-test-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const fluent = require('@fluentui/react-components');
const { RendererProvider } = require('@griffel/react');
const { createDOMRenderer } = require('@griffel/core');
const { Permission } = require('./fixtures/access-trim-harness.cjs');
// JSDOM 15 has no Web Crypto; supply the native crypto API used by Tabster.
Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
const keys = { page: {}, sp: {}, graph: {}, theme: {} };
const constants = Object.assign(Permission, { viewListItems: new Permission({ Low: 1 }), editListItems: new Permission({ Low: 4 }), addListItems: new Permission({ Low: 2 }), manageWeb: new Permission({ Low: 0x40000000 }), manageLists: new Permission({ Low: 0x800 }) });
function fixture(t, name) {
  const toolkit = {}; let useContext, usePage;
  const pending = []; const calls = [];
  const sp = { get: (url, config, options) => new Promise(resolve => { calls.push({ url, config, options }); pending.push(() => resolve({ ok: true, json: async () => mode === 'error' ? { malformed: true } : url.includes('currentuser') ? { Id: 17 } : url.includes('/users') ? { value: mode === 'denied' ? [] : [{ Id: 17 }] } : url.includes('sitegroups') ? { value: [{ Id: 2, Title: decodeURIComponent(url).includes('Changed group') ? 'Changed group' : 'Project Readers' }] } : { High: '0', Low: mode === 'denied' ? '0' : '7' } })); }) };
  let mode = 'allowed';
  const groupId = 'bbbbbbbb-0000-0000-0000-000000000001';
  const graph = { api: url => ({ get: () => new Promise((resolve, reject) => { calls.push({ url }); pending.push(() => mode === 'error' ? reject(new Error('Graph unavailable')) : resolve(url.startsWith('/groups?') ? { value: [{ id: groupId, displayName: decodeURIComponent(url).includes('Changed group') ? 'Changed group' : 'Project Readers', groupTypes: ['Unified'], securityEnabled: false, mailEnabled: true, onPremisesSyncEnabled: false }] } : { value: [] })); }), post: body => new Promise((resolve, reject) => { calls.push({ url, body }); pending.push(() => mode === 'error' ? reject(new Error('Graph unavailable')) : resolve({ value: mode === 'denied' ? [] : body.groupIds })); }) }) };
  const page = { web: { absoluteUrl: 'https://tenant.test/sites/current' }, user: { loginName: 'first', email: 'first@test' } };
  const h = createHarness({ '@apvee/spfx-react-toolkit': toolkit,
    '@microsoft/sp-http': { SPHttpClient: { serviceKey: keys.sp, configurations: { v1: {} } }, MSGraphClientFactory: { serviceKey: keys.graph } },
    '@microsoft/sp-page-context': { SPPermission: constants, PageContext: { serviceKey: keys.page } },
    '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 } }, '@microsoft/sp-component-base': { ThemeProvider: { serviceKey: keys.theme } },
    './useSPFxContext': { useSPFxContext: () => useContext() }, './useSPFxPageContext': { useSPFxPageContext: () => usePage() },
    './useAsyncInvoke.internal': { useAsyncInvoke: loadHook('useAsyncInvoke.internal', 'useAsyncInvoke') },
  });
  useContext = h.load('core/context.internal.tsx').useSPFxContext; usePage = h.load('hooks/useSPFxPageContext.ts').useSPFxPageContext;
  Object.assign(toolkit, h.load('hooks/useSPFxCrossSitePermissions.ts'));
  for (const [family, slug] of [['SharePointPermissionTrim', 'sharepoint-permission-trim'], ['SharePointGroupTrim', 'sharepoint-group-trim'], ['EntraGroupTrim', 'entra-group-trim']]) Object.assign(toolkit, h.load(`controls/${slug}/${slug}.tsx`));
  const Provider = h.load('core/provider-base.internal.tsx').SPFxProviderBase;
  const theme = { tryGetTheme: () => ({}), themeChangedEvent: { add() {}, remove() {} } };
  const factory = { getClient: async () => graph };
  const instance = { context: { instanceId: 'demo', serviceScope: { whenFinished: fn => fn(), consume: key => key === keys.page ? page : key === keys.sp ? sp : key === keys.graph ? factory : theme } }, properties: {}, displayMode: 1, domElement: document.createElement('div'), render() {} };
  const Component = h.load(path.resolve(__dirname, `../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/${name}Panel.tsx`)).default;
  const renderer = createDOMRenderer(document);
  const view = h.mountComponent(() => h.React.createElement(RendererProvider, { renderer }, h.React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme }, h.React.createElement(Provider, { instance }, h.React.createElement(Component)))));
  t.after(() => { view.unmount(); h.close(); });
  const select = (selector, value) => { const node = view.container.querySelector(selector); assert.ok(node, selector); h.act(() => Simulate.change(node, { target: { value } })); };
  const click = selector => { const node = view.container.querySelector(selector); assert.ok(node, selector); h.act(() => Simulate.click(node)); };
  const flush = async () => { await h.act(async () => { for (let i = 0; i < 50; i++) await Promise.resolve(); }); };
  const finish = async () => { for (let i = 0; i < 12; i++) { await flush(); while (pending.length) { await h.act(async () => pending.shift()()); } } };
  return { h, view, calls, select, click, flush, finish, mode: value => { mode = value; } };
}
for (const name of ['SharePointPermissionTrim', 'SharePointGroupTrim', 'EntraGroupTrim']) test(`${name} actual panel stages native reads and protects component lifetimes`, async t => {
  const f = fixture(t, name); await f.flush();
  assert.equal(f.view.container.querySelectorAll('[data-demo-section]').length, 1);
  assert.equal(f.calls.length, 0); assert.ok(f.view.container.querySelector('[data-trim-demo=idle]'));
  f.click('[data-trim-demo=run]'); await f.flush();
  assert.equal(f.view.container.querySelector('[data-trim-demo=protected]'), null);
  assert.equal(f.view.container.querySelector('[data-trim-demo=state]').getAttribute('data-trim-state'), 'loading');
  await f.finish(); assert.ok(f.view.container.querySelector('[data-trim-demo=protected]'));
  f.click('[data-trim-demo=increment]'); assert.match(f.view.container.textContent, /Protected counter: 1/); assert.match(f.view.container.textContent, /Protected mounts: 1; active: true/);
  const applied = f.view.container.querySelector('[data-trim-demo=applied-target]').textContent;
  if (name === 'SharePointPermissionTrim') f.select('[data-trim-demo=web-url]', 'https://tenant.test/sites/next'); else f.select('[data-trim-demo=groups]', 'Changed group');
  assert.equal(f.view.container.querySelector('[data-trim-demo=applied-target]').textContent, applied); assert.match(f.view.container.textContent, /Protected counter: 1/);
  const before = f.calls.length;
  f.mode('denied');
  f.click('[data-trim-demo=run]'); await f.flush();
  assert.equal(f.view.container.querySelector('[data-trim-demo=protected]'), null); await f.finish();
  assert.equal(f.view.container.querySelector('[data-trim-demo=state]').getAttribute('data-trim-state'), 'denied');
  assert.ok(f.calls.length > before); assert.match(f.view.container.textContent, /active: false/);
  if (name !== 'SharePointPermissionTrim') f.select('[data-trim-demo=groups]', 'Project Readers');
  else f.select('[data-trim-demo=web-url]', 'invalid URL');
  f.mode('error'); f.click('[data-trim-demo=run]'); await f.finish();
  assert.equal(f.view.container.querySelector('[data-trim-demo=state]').getAttribute('data-trim-state'), 'error');
});
test('permission sample nested AND checks mount inner Entra group only after permission grant', async t => {
  const f = fixture(t, 'SharePointPermissionTrim'); await f.flush();
  f.select('[data-demo-variant=sharepoint-permission-trim]', 'and'); f.mode('denied'); f.click('[data-trim-demo=run]'); await f.finish();
  assert.equal(f.calls.some(call => call.url.startsWith('/groups')), false); assert.equal(f.view.container.querySelector('[data-trim-demo=protected]'), null);
  f.mode('allowed'); f.click('[data-trim-demo=run]'); await f.finish(); assert.ok(f.calls.some(call => call.url.startsWith('/groups'))); assert.ok(f.view.container.querySelector('[data-trim-demo=protected]'));
});

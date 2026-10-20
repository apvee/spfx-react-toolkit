const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const { deferred, loadHook } = require('./async-test-harness.cjs');
async function flush(h, fn = () => {}) { await h.act(async () => { fn(); for (let i = 0; i < 16; i++) await Promise.resolve(); }); }
class Permission { constructor(mask) { this.mask = mask; } hasPermission(p) { return (this.mask.Low & p.mask.Low) === p.mask.Low; } }
const sdk = { SPHttpClient: { configurations: { v1: {} } }, MSGraphClientFactory: { serviceKey: {} } };
function setup() {
  const provider = {}; const page = { web: { absoluteUrl: 'https://tenant.test/sites/current' }, user: { loginName: 'a' } };
  const calls = []; const requests = []; let respond = async () => ({ ok: true, json: async () => ({ High: '0', Low: '1' }) });
  let client = { get: (url, config, options) => { calls.push(url); requests.push({ config, options }); return respond(url); } };
  const h = createHarness({ '@microsoft/sp-http': sdk, '@microsoft/sp-page-context': { SPPermission: Permission }, './useSPFxContext': { useSPFxContext: () => provider }, './useSPFxPageContext': { useSPFxPageContext: () => page }, './useSPFxSPHttpClient': { useSPFxSPHttpClient: () => ({ client, invoke: fn => fn(client) }) } });
  const hook = h.load('hooks/useSPFxCrossSitePermissions.ts').useSPFxCrossSitePermissions;
  return { h, hook, calls, requests, page, replaceClient() { client = { get: (url, config, options) => { calls.push(url); requests.push({ config, options }); return respond(url); } }; }, respond(fn) { respond = fn; } };
}
test('scoped list requests only selected resource, scalar-equivalent options do not refetch', async t => {
  const f = setup(); t.after(() => f.h.close());
  const options = { scope: 'list', listTitle: "Project's Files" };
  const m = f.h.mount(p => f.hook(p), options); await flush(f.h);
  assert.equal(f.calls.length, 1); assert.match(f.calls[0], /getByTitle\('Project%27%27s%20Files'\)\/effectivebasepermissions$/);
  assert.equal(m.current.hasPermission(new Permission({ Low: 1 })), true);
  m.render({ ...options }); await flush(f.h); assert.equal(f.calls.length, 1);
});
test('current web native mask avoids REST, absent mask fetches web only', async t => {
  const f = setup(); t.after(() => f.h.close()); f.page.web.permissions = new Permission({ Low: 1 });
  const m = f.h.mount(() => f.hook({})); await flush(f.h); assert.equal(f.calls.length, 0); assert.equal(m.current.permissions, f.page.web.permissions);
  f.page.web.permissions = undefined; m.render(); await flush(f.h); assert.equal(f.calls.length, 1); assert.match(f.calls[0], /_api\/web\/effectivebasepermissions$/);
});
test('legacy undefined stays idle and old selected grants clear before effects', async t => {
  const f = setup(); t.after(() => f.h.close()); const observed = [];
  const m = f.h.mount(p => { const info = f.hook(p.value); observed.push(info); return info; }, { value: undefined });
  await flush(f.h); assert.equal(f.calls.length, 0); assert.equal(m.current.isLoading, false);
  m.render({ value: { scope: 'list', listTitle: 'A' } }); await flush(f.h); assert.ok(m.current.permissions);
  const d = deferred(); f.respond(() => d.promise); const first = observed.length;
  m.render({ value: { scope: 'list', listTitle: 'B' } }); assert.equal(observed[first].permissions, undefined);
  await flush(f.h, () => d.resolve({ ok: true, json: async () => ({ Low: 0, High: 0 }) }));
  assert.equal(m.current.hasPermission(new Permission({ Low: 1 })), false);
});
test('Graph factory replacement discards ready client during render and obsolete initialization', async t => {
  const ds = [deferred(), deferred(), deferred()]; let index = 0; const factories = ds.map(d => ({ getClient: () => d.promise }));
  const consumers = factories.map(factory => () => factory); const observations = [];
  const h = createHarness({ '@microsoft/sp-http': sdk, './useSPFxServiceScope': { useSPFxServiceScope: () => ({ consume: consumers[index] }) }, './useAsyncInvoke.internal': { useAsyncInvoke: loadHook('useAsyncInvoke.internal', 'useAsyncInvoke') } }); t.after(() => h.close());
  const hook = h.load('hooks/useSPFxMSGraphClient.ts').useSPFxMSGraphClient;
  const m = h.mount(() => { const value = hook(); observations.push(value); return value; });
  const a = {}; await flush(h, () => ds[0].resolve(a)); assert.equal(m.current.client, a);
  index = 1; const first = observations.length; m.render(); assert.equal(observations[first].client, undefined); assert.equal(m.current.isInitializing, true);
  index = 2; m.render(); const c = {}; await flush(h, () => ds[2].resolve(c)); await flush(h, () => ds[1].resolve({})); assert.equal(m.current.client, c);
});
for (const value of [
  { scope: 'list' }, { scope: 'list', listId: 'bad' }, { scope: 'list', listTitle: '' },
  { scope: 'list', listTitle: 'A', listPath: 'Lists/A' }, { scope: 'web', listTitle: 'A' },
  { scope: 'item', listTitle: 'A', itemId: 0 }, { scope: 'item', listTitle: 'A', itemId: 1.5 },
  { scope: 'list', listTitle: 'A', itemId: 1 }, { scope: 'list', listTitle: 'A', folderPath: '/A' },
  { scope: 'folder', listTitle: 'A', folderPath: 'A' }, { scope: 'folder', listTitle: 'A', folderPath: '/A/../B' },
  { scope: 'list', listPath: '/Lists/A' }, { scope: 'list', listPath: 'Lists/A?x' },
  { scope: 'web', webUrl: '/sites/relative' }, { scope: 'web', webUrl: 'https://tenant.test/?x' },
  { scope: 'unknown' },
]) test(`invalid scoped target blocks all HTTP: ${JSON.stringify(value)}`, async t => {
  const f = setup(); t.after(() => f.h.close()); const m = f.h.mount(() => f.hook(value)); await flush(f.h);
  assert.equal(f.calls.length, 0); assert.ok(m.current.error); assert.equal(m.current.permissions, undefined); assert.equal(m.current.isLoading, false);
});
for (const payload of [undefined, {}, { Low: 'bad', High: 0 }, { Low: -1, High: 0 }, { Low: 1, High: 4294967296 }, { Low: 1.5, High: 0 }, { Low: null, High: 0 }]) {
  test(`invalid permission response fails closed: ${JSON.stringify(payload)}`, async t => {
    const f = setup(); t.after(() => f.h.close()); f.respond(async () => ({ ok: true, json: async () => payload }));
    const m = f.h.mount(() => f.hook({})); await flush(f.h); assert.ok(m.current.error); assert.equal(m.current.permissions, undefined);
  });
}
test('HTTP failure is an error even when payload would grant permissions', async t => {
  const f = setup(); t.after(() => f.h.close()); let decoded = false;
  f.respond(async () => ({ ok: false, status: 403, json: async () => { decoded = true; return { Low: 1, High: 0 }; } }));
  const m = f.h.mount(() => f.hook({})); await flush(f.h); assert.match(m.current.error.message, /403/); assert.equal(decoded, false); assert.equal(m.current.hasPermission(new Permission({ Low: 1 })), false);
});
test('item selects only item mask, ID normalization and decoded list path use selected web', async t => {
  const f = setup(); t.after(() => f.h.close());
  const m = f.h.mount(() => f.hook({ scope: 'item', webUrl: 'https://tenant.test/sites/other/', listId: '{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}', itemId: 42 }));
  await flush(f.h); assert.deepEqual(f.calls, ["https://tenant.test/sites/other/_api/web/lists(guid'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee')/items(42)/effectivebasepermissions"]);
  m.unmount(); f.h.mount(() => f.hook({ scope: 'list', listPath: 'Shared Documents/50% # files' })); await flush(f.h);
  assert.match(f.calls[1], /GetListUsingPath\(DecodedUrl='%2Fsites%2Fcurrent%2FShared%20Documents%2F50%25%20%23%20files'\)/);
});
for (const mismatch of [false, true]) test(`folder checks authoritative selected list identity: mismatch ${mismatch}`, async t => {
  const f = setup(); t.after(() => f.h.close()); const id = '11111111-2222-3333-4444-555555555555';
  f.respond(async url => ({ ok: true, json: async () => url.includes('GetFolder') ? {
    EffectiveBasePermissions: { Low: '1', High: '0' }, ParentList: { Id: mismatch ? 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' : id },
  } : { Id: id } }));
  const m = f.h.mount(() => f.hook({ scope: 'folder', listTitle: 'Docs', folderPath: "/sites/current/Docs/A's 50% #" })); await flush(f.h);
  assert.equal(f.calls.length, 2); assert.equal(f.calls.some(url => /_api\/(site|web)\/effectivebasepermissions/.test(url)), false);
  assert.match(f.calls[1], /A%27%27s%2050%25%20%23/);
  assert.equal(Boolean(m.current.error), mismatch); assert.equal(m.current.hasPermission(new Permission({ Low: 1 })), !mismatch);
});
test('legacy three-set shape accepts top-level, named and verbose masks', async t => {
  const f = setup(); t.after(() => f.h.close()); f.respond(async url => ({ ok: true, json: async () => url.includes('/site/') ? { d: { EffectiveBasePermissions: { Low: '1', High: '0' } } } : { EffectiveBasePermissions: { Low: 1, High: 0 } } }));
  const m = f.h.mount(() => f.hook('https://tenant.test/sites/other', { listId: '11111111-2222-3333-4444-555555555555' })); await flush(f.h);
  assert.equal(f.calls.length, 3); for (const method of ['hasWebPermission', 'hasSitePermission', 'hasListPermission']) assert.equal(m.current[method](new Permission({ Low: 1 })), true);
});
for (const identity of ['client', 'user']) test(`identity ${identity} revokes during render and ignores obsolete completion`, async t => {
  const f = setup(); t.after(() => f.h.close()); const ds = [deferred(), deferred(), deferred()]; let n = 0; const observations = [];
  f.respond(() => ds[n++].promise);
  const m = f.h.mount(() => { const value = f.hook({ scope: 'list', listTitle: 'A' }); observations.push(value); return value; }); await flush(f.h);
  await flush(f.h, () => ds[0].resolve({ ok: true, json: async () => ({ Low: 1, High: 0 }) })); assert.ok(m.current.permissions);
  if (identity === 'client') f.replaceClient(); else f.page.user.loginName = 'b';
  const first = observations.length; m.render(); assert.equal(observations[first].permissions, undefined); await flush(f.h);
  if (identity === 'client') f.replaceClient(); else f.page.user.loginName = 'c'; m.render(); await flush(f.h);
  await flush(f.h, () => ds[2].resolve({ ok: true, json: async () => ({ Low: 0, High: 0 }) }));
  await flush(f.h, () => ds[1].reject(new Error('obsolete'))); assert.equal(m.current.error, undefined); assert.equal(m.current.hasPermission(new Permission({ Low: 1 })), false);
});
test('inflight permission reads share per provider only, settlement retains no grant', async t => {
  const f = setup(); t.after(() => f.h.close()); const d = deferred(); f.respond(() => d.promise);
  const a = f.h.mount(() => f.hook({ scope: 'list', listTitle: 'A' })); const b = f.h.mount(() => f.hook({ scope: 'list', listTitle: 'A' })); await flush(f.h); assert.equal(f.calls.length, 1);
  a.unmount(); await flush(f.h, () => d.resolve({ ok: true, json: async () => ({ Low: 1, High: 0 }) })); assert.ok(b.current.permissions);
  f.h.mount(() => f.hook({ scope: 'list', listTitle: 'A' })); await flush(f.h); assert.equal(f.calls.length, 2);
  const other = setup(); t.after(() => other.h.close()); other.h.mount(() => other.hook({ scope: 'list', listTitle: 'A' })); await flush(other.h); assert.equal(other.calls.length, 1);
});
test('shared access lifecycle guards provider identity, suppresses unmounted state and normalizes failures', async t => {
  const h = createHarness(); t.after(() => h.close()); const { useAccessRequest } = h.load('hooks/useAccessRequest.internal.ts');
  const ds = [deferred(), deferred()]; const providers = [{}, {}]; const observations = []; let index = 0;
  const m = h.mount(() => { const state = useAccessRequest({ provider: providers[index], identity: [providers[index]], requestKey: 'test', request: () => ds[index].promise }); observations.push(state); return state; }); await flush(h);
  index = 1; m.render(); await flush(h); await flush(h, () => ds[1].resolve(true)); await flush(h, () => ds[0].resolve(false)); assert.equal(m.current.data, true);
  const pending = deferred(); const end = h.mount(() => useAccessRequest({ provider: {}, identity: [], request: () => pending.promise })); await flush(h); end.unmount(); await flush(h, () => pending.reject('unmounted')); assert.equal(end.current.error, undefined);
});
test('shared access disable and reenable starts pending without a former decision', async t => {
  const h = createHarness(); t.after(() => h.close()); const { useAccessRequest } = h.load('hooks/useAccessRequest.internal.ts'); const provider = {}; const old = deferred(), next = deferred(); let active = true, n = 0; const seen = [];
  const m = h.mount(() => { const value = useAccessRequest({ provider, identity: ['same'], request: active ? () => ++n === 1 ? old.promise : next.promise : undefined }); seen.push(value); return value; });
  await flush(h, () => old.resolve(true)); assert.equal(m.current.data, true);
  active = false; m.render(); assert.equal(m.current.data, undefined);
  active = true; const first = seen.length; m.render(); assert.equal(seen[first].data, undefined); assert.equal(seen[first].isLoading, true);
  await flush(h, () => next.resolve(false)); assert.equal(m.current.data, false);
});
test('legacy site set uses collection root web effective permissions', async t => {
  const f = setup(); t.after(() => f.h.close()); f.h.mount(() => f.hook('https://tenant.test/sites/other/subweb')); await flush(f.h);
  assert.equal(f.calls[0], 'https://tenant.test/sites/other/subweb/_api/site/rootweb/effectivebasepermissions');
});
test('Graph factory replacement hides invocation error before effects', async t => {
  const a = deferred(), b = deferred(); let index = 0; const consumers = [() => ({ getClient: () => a.promise }), () => ({ getClient: () => b.promise })]; const seen = [];
  const h = createHarness({ '@microsoft/sp-http': sdk, './useSPFxServiceScope': { useSPFxServiceScope: () => ({ consume: consumers[index] }) }, './useAsyncInvoke.internal': { useAsyncInvoke: loadHook('useAsyncInvoke.internal', 'useAsyncInvoke') } }); t.after(() => h.close());
  const { useSPFxMSGraphClient } = h.load('hooks/useSPFxMSGraphClient.ts'); const m = h.mount(() => { const value = useSPFxMSGraphClient(); seen.push(value); return value; });
  await flush(h, () => a.resolve({})); await h.act(async () => { await m.current.invoke(async () => { throw new Error('old client'); }).catch(() => {}); }); assert.ok(m.current.error);
  index = 1; const first = seen.length; m.render(); assert.equal(seen[first].error, undefined); assert.equal(seen[first].isLoading, false);
});
for (const value of [{ scope: null }, { webUrl: null }, { scope: 'web', webUrl: 1 }]) test(`invalid explicit absence/type is rejected: ${JSON.stringify(value)}`, async t => {
  const f = setup(); t.after(() => f.h.close()); const m = f.h.mount(() => f.hook(value)); await flush(f.h);
  assert.ok(m.current.error); assert.equal(f.calls.length, 0);
});
test('native permission requests use v4-compatible metadata header with v1 configuration', async t => {
  const f = setup(); t.after(() => f.h.close()); f.h.mount(() => f.hook({})); await flush(f.h);
  assert.equal(f.requests[0].config, sdk.SPHttpClient.configurations.v1);
  assert.equal(f.requests[0].options.headers.Accept, 'application/json;odata.metadata=none');
});
test('Graph initialization failures normalize, replacement recovers, unmount suppresses publication', async t => {
  const ds = [deferred(), deferred(), deferred()]; let index = 0; const consumers = ds.map(d => () => ({ getClient: () => d.promise }));
  const h = createHarness({ '@microsoft/sp-http': sdk, './useSPFxServiceScope': { useSPFxServiceScope: () => ({ consume: consumers[index] }) }, './useAsyncInvoke.internal': { useAsyncInvoke: loadHook('useAsyncInvoke.internal', 'useAsyncInvoke') } }); t.after(() => h.close());
  const { useSPFxMSGraphClient } = h.load('hooks/useSPFxMSGraphClient.ts'); const m = h.mount(() => useSPFxMSGraphClient());
  await flush(h, () => ds[0].reject('initialization denied')); assert.equal(m.current.initError.message, 'initialization denied'); assert.equal(m.current.isReady, false);
  index = 1; m.render(); assert.equal(m.current.initError, undefined); const ready = {}; await flush(h, () => ds[1].resolve(ready)); assert.equal(m.current.client, ready); assert.equal(m.current.isReady, true);
  index = 2; m.render(); await flush(h); m.unmount(); await flush(h, () => ds[2].reject(new Error('after unmount'))); assert.equal(m.current.initError, undefined);
});
test('shared access normalizes current failures and clears error for changed target', async t => {
  const h = createHarness(); t.after(() => h.close()); const { useAccessRequest } = h.load('hooks/useAccessRequest.internal.ts'); const provider = {}; let target = 'a'; const d = deferred();
  const m = h.mount(() => useAccessRequest({ provider, identity: [target], request: () => target === 'a' ? Promise.reject('current denial') : d.promise })); await flush(h);
  assert.equal(m.current.error.message, 'current denial'); assert.equal(m.current.data, undefined); assert.equal(m.current.isLoading, false);
  target = 'b'; m.render(); assert.equal(m.current.error, undefined); assert.equal(m.current.isLoading, true); await flush(h, () => d.resolve(false)); assert.equal(m.current.data, false);
});
for (const nativeMask of [false, true]) {
  const invalidContainers = [
    ['array', []], ['array with scoped fields', Object.assign([], { scope: 'web' })],
    ['Date', new Date(0)], ['Map', new Map()], ['Set', new Set()], ['RegExp', /web/],
    ['boxed string', new String('web')], ['boxed number', new Number(1)],
  ];
  for (const [label, options] of invalidContainers) test(`invalid scoped container ${label} is rejected with native mask ${nativeMask}`, async t => {
    const f = setup(); t.after(() => f.h.close()); if (nativeMask) f.page.web.permissions = new Permission({ Low: 1 });
    const m = f.h.mount(() => f.hook(options)); await flush(f.h);
    assert.ok(m.current.error); assert.equal(m.current.permissions, undefined);
    assert.equal(m.current.hasPermission(new Permission({ Low: 1 })), false);
    assert.equal(m.current.isLoading, false); assert.equal(f.calls.length, 0);
  });
  test(`valid default becomes invalid container during render with native mask ${nativeMask}`, async t => {
    const f = setup(); t.after(() => f.h.close()); if (nativeMask) f.page.web.permissions = new Permission({ Low: 1 }); const observations = [];
    const m = f.h.mount(p => { const value = f.hook(p.target); observations.push(value); return value; }, { target: {} }); await flush(f.h);
    assert.equal(m.current.hasPermission(new Permission({ Low: 1 })), true); const initialCalls = f.calls.length;
    const first = observations.length; m.render({ target: [] }); assert.equal(observations[first].permissions, undefined);
    assert.ok(m.current.error); await flush(f.h); assert.equal(f.calls.length, initialCalls);
    m.render({ target: {} }); await flush(f.h); assert.equal(m.current.error, undefined); assert.equal(m.current.hasPermission(new Permission({ Low: 1 })), true);
  });
  test(`legitimate structural option objects remain supported with native mask ${nativeMask}`, async t => {
    const f = setup(); t.after(() => f.h.close()); if (nativeMask) f.page.web.permissions = new Permission({ Low: 1 });
    class WebOptions { constructor() { this.scope = 'web'; } }
    for (const target of [{}, Object.freeze({ scope: 'web' }), Object.create(null), new WebOptions()]) {
      const m = f.h.mount(() => f.hook(target)); await flush(f.h);
      assert.equal(m.current.error, undefined); assert.equal(m.current.hasPermission(new Permission({ Low: 1 })), true); m.unmount();
    }
    assert.equal(f.calls.length, nativeMask ? 0 : 4);
  });
}

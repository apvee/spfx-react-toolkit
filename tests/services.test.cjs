const assert = require('node:assert/strict');
const { test } = require('node:test');
const { createHarness } = require('./services-test-harness.cjs');

// These use the installed PnP caching/queryable pipeline. Only HTTP send is replaced.
for (const storage of ['session', 'local']) {
  test('PnP cache expires after custom timeout in ' + storage + ' storage', async t => {
    const h = createHarness(); t.after(h.close);
    t.mock.timers.enable({ apis: ['Date'], now: 1700000000000 });
    const service = h.load('services/spfx-pnp-context.service.ts').createSPFxPnPContextService(h.context.current, h.pageContext);
    const sp = service.createSPFI(undefined, { cache: { enabled: true, storage, timeout: 1250 } });
    const calls = h.withTransport(sp);
    assert.deepEqual(await sp.web(), { Title: 'response-1' });
    t.mock.timers.tick(1249);
    assert.deepEqual(await sp.web(), { Title: 'response-1' });
    assert.equal(calls.length, 1);
    t.mock.timers.tick(2);
    assert.deepEqual(await sp.web(), { Title: 'response-2' });
    assert.equal(calls.length, 2);
  });
}
for (const timeout of [undefined, 0]) {
  test('PnP cache retains legacy five minute fallback for timeout ' + timeout, async t => {
    const h = createHarness(); t.after(h.close);
    t.mock.timers.enable({ apis: ['Date'], now: 1700000000000 });
    const service = h.load('services/spfx-pnp-context.service.ts').createSPFxPnPContextService(h.context.current, h.pageContext);
    const sp = service.createSPFI(undefined, { cache: { enabled: true, timeout } });
    const calls = h.withTransport(sp);
    await sp.web();
    t.mock.timers.tick(299999);
    assert.deepEqual(await sp.web(), { Title: 'response-1' });
    t.mock.timers.tick(2);
    assert.deepEqual(await sp.web(), { Title: 'response-2' });
    assert.equal(calls.length, 2);
  });
}

test('PnP hook changes cache key factory and retains equal inline configuration', async t => {
  const h = createHarness(); t.after(h.close);
  const useHook = h.load('hooks/useSPFxPnPContext.ts').useSPFxPnPContext;
  const firstFactory = () => 'first';
  const secondFactory = () => 'second';
  const fixture = h.mount(p => useHook(p.siteUrl, p.config), { config: { cache: { enabled: true, keyFactory: firstFactory } } });
  const first = fixture.current;
  const calls = h.withTransport(first.sp);
  await first.sp.web();
  assert.ok(sessionStorage.getItem('first'));
  fixture.render({ config: { cache: { enabled: true, keyFactory: firstFactory } } });
  assert.equal(fixture.current, first, 'equal inline config should retain instance and return value');
  fixture.render({ config: { cache: { enabled: true, keyFactory: secondFactory } } });
  assert.notEqual(fixture.current.sp, first.sp, 'different function identity must rebuild SPFI');
  h.withTransport(fixture.current.sp);
  await fixture.current.sp.web();
  assert.ok(sessionStorage.getItem('second'));
  assert.equal(calls.length, 1);
});

function deferred() { let resolve, reject; const promise = new Promise((a, b) => { resolve = a; reject = b; }); return { promise, resolve, reject }; }
function permissionFixture(t) {
  const requests = [];
  const client = { get(url) { const request = { url, ...deferred() }; requests.push(request); return request.promise; } };
  const invoke = fn => fn(client);
  class Permission { constructor(mask) { this.mask = mask; } hasPermission(permission) { return this.mask.Low === permission.mask.Low; } }
  const h = createHarness({ './useSPFxSPHttpClient': { useSPFxSPHttpClient: () => ({ invoke }) }, '@microsoft/sp-page-context': { SPPermission: Permission }, '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } } });
  t.after(h.close);
  const hook = h.load('hooks/useSPFxCrossSitePermissions.ts').useSPFxCrossSitePermissions;
  const rendered = h.mount(p => hook(p.url, p.options), { url: 'https://tenant.test/A' });
  async function finish(indices, Low) { await h.act(async () => { for (const i of indices) requests[i].resolve({ ok: true, json: async () => ({ EffectiveBasePermissions: { High: 0, Low } }) }); await Promise.all(indices.map(i => requests[i].promise)); }); }
  return { ...h, ...rendered, get current() { return rendered.current; }, requests, finish };
}
test('cross-site permissions retain current site when old requests finish last', async t => {
  const f = permissionFixture(t);
  f.render({ url: 'https://tenant.test/B' });
  await f.finish([2, 3], 2);
  await f.finish([0, 1], 1);
  assert.equal(f.current.webPermissions.mask.Low, 2);
  assert.equal(f.current.sitePermissions.mask.Low, 2);
  assert.equal(f.current.isLoading, false);
  f.render({ url: 'https://tenant.test/B', options: {} });
  assert.equal(f.requests.length, 4, 'same scalar identity must not fetch again');
});
test('clearing cross-site URL invalidates pending permissions and stays idle', async t => {
  const f = permissionFixture(t);
  f.render({ url: '' });
  await f.finish([0, 1], 1);
  assert.equal(f.current.webPermissions, undefined);
  assert.equal(f.current.sitePermissions, undefined);
  assert.equal(f.current.isLoading, false);
});
test('cross-site permissions clear previous grants while target is loading', async t => {
  const f = permissionFixture(t);
  await f.finish([0, 1], 1);
  f.render({ url: 'https://tenant.test/B' });
  assert.equal(f.current.webPermissions, undefined);
  assert.equal(f.current.sitePermissions, undefined);
  assert.equal(f.current.isLoading, true);
  await f.finish([2, 3], 2);
});
function photoFixture(t) {
  const requests = [];
  const client = { api(endpoint) { return { get() { const request = { endpoint, ...deferred() }; requests.push(request); return request.promise; } }; } };
  const h = createHarness({ './useSPFxMSGraphClient': { useSPFxMSGraphClient: () => ({ client }) } });
  t.after(h.close);
  const hook = h.load('hooks/useSPFxUserPhoto.ts').useSPFxUserPhoto;
  const rendered = h.mount(p => hook(p), { userId: 'A' });
  async function finish(index, blob) { await h.act(async () => { requests[index].resolve(blob); await requests[index].promise; }); }
  return { ...h, ...rendered, get current() { return rendered.current; }, requests, finish };
}
test('user photo retains current user blob when old user finishes last', async t => {
  const f = photoFixture(t);
  f.render({ userId: 'B' });
  const b = new Blob(['B']);
  await f.finish(1, b);
  const url = f.current.photoUrl;
  await f.finish(0, new Blob(['A']));
  assert.equal(f.current.photoBlob, b);
  assert.equal(f.current.photoUrl, url);
  f.render({ userId: 'B' });
  assert.equal(f.requests.length, 2);
});
test('user photo overlapping reload keeps latest blob when the first completes last', async t => {
  const f = photoFixture(t);
  let reload;
  f.act(() => { reload = f.current.reload(); });
  const latest = new Blob(['latest']);
  await f.finish(1, latest);
  await f.finish(0, new Blob(['old']));
  await reload;
  assert.equal(f.current.photoBlob, latest);
});
test('user photo clears previous user while a different user photo is pending', async t => {
  const f = photoFixture(t);
  await f.finish(0, new Blob(['A']));
  f.render({ userId: 'B' });
  assert.equal(f.current.photoBlob, undefined);
  assert.equal(f.current.photoUrl, undefined);
  await f.finish(1, new Blob(['B']));
});
test('same-name overlapping time reports each operation duration independently', async t => {
  const h = createHarness({ './useSPFxInstanceInfo': { useSPFxInstanceInfo: () => ({ id: 'one', kind: 'WebPart' }) }, './useSPFxCorrelationInfo': { useSPFxCorrelationInfo: () => ({ correlationId: 'one' }) } });
  t.after(h.close);
  const hook = h.load('hooks/useSPFxPerformance.ts').useSPFxPerformance;
  const f = h.mount(hook);
  t.mock.timers.enable({ apis: ['Date'], now: 1000 });
  const descriptor = Object.getOwnPropertyDescriptor(global, 'performance');
  const marks = new Map(); const entries = [];
  Object.defineProperty(global, 'performance', { configurable: true, value: { mark: name => marks.set(name, Date.now()), clearMarks: name => marks.delete(name), measure: (name, start) => entries.push({ name, duration: Date.now() - marks.get(start) }), getEntriesByName: name => entries.filter(e => e.name === name) } });
  t.after(() => Object.defineProperty(global, 'performance', descriptor));
  const a = deferred(); const b = deferred();
  const ar = f.current.time('same', () => a.promise); await Promise.resolve();
  t.mock.timers.tick(100);
  const br = f.current.time('same', () => b.promise); await Promise.resolve();
  t.mock.timers.tick(50);
  a.resolve('A'); const first = await ar;
  b.resolve('B'); const second = await br;
  assert.equal(first.durationMs, 150);
  assert.equal(second.durationMs, 50);
  assert.equal(first.name, 'same'); assert.equal(first.result, 'A');
  assert.equal(second.name, 'same'); assert.equal(second.result, 'B');
  assert.equal(marks.size, 0, 'time must release its own marks');
  f.current.mark('manual'); t.mock.timers.tick(25);
  assert.equal(f.current.measure('manual metric', 'manual').durationMs, 25);
  const failure = new Error('operation failed');
  await assert.rejects(f.current.time('failure', () => { throw failure; }), error => error === failure);
  assert.deepEqual(Array.from(marks.keys()), ['manual'], 'time must clean up only its own marks');
});

test('PnP context resolves site, injects headers and rebuilds changed scalar config', async t => {
  const h = createHarness(); t.after(h.close);
  const hook = h.load('hooks/useSPFxPnPContext.ts').useSPFxPnPContext;
  const f = h.mount(p => hook(p.siteUrl, p.config), { siteUrl: '/sites/other/', config: { headers: { 'X-Test': 'one' } } });
  assert.equal(f.current.siteUrl, 'https://tenant.test/sites/other');
  const first = f.current.sp;
  const calls = h.withTransport(first); await first.web();
  assert.equal(calls[0].init.headers['X-Test'], 'one');
  f.render({ siteUrl: '/sites/other/', config: { headers: { 'X-Test': 'two' } } });
  assert.notEqual(f.current.sp, first);
  const changed = h.withTransport(f.current.sp); await f.current.sp.web();
  assert.equal(changed[0].init.headers['X-Test'], 'two');
  assert.equal(sessionStorage.length, 0, 'disabled caching must not write storage');
});
test('PnP hook exposes initialization failure and recovers when context becomes available', t => {
  const h = createHarness(); t.after(h.close);
  const context = h.context.current; h.context.current = undefined;
  const hook = h.load('hooks/useSPFxPnPContext.ts').useSPFxPnPContext;
  const f = h.mount(() => hook());
  assert.equal(f.current.isInitialized, false); assert.ok(f.current.error instanceof Error); assert.equal(f.current.sp, undefined);
  h.context.current = context; f.render();
  assert.equal(f.current.isInitialized, true); assert.equal(f.current.error, undefined);
});
test('cross-site old rejection cannot replace new success and unmount ignores deferred updates', async t => {
  const f = permissionFixture(t);
  f.render({ url: 'https://tenant.test/B' });
  await f.finish([2, 3], 2);
  await f.act(async () => { f.requests[0].reject(new Error('old denied')); f.requests[1].resolve({ json: async () => ({}) }); });
  assert.equal(f.current.error, undefined); assert.equal(f.current.webPermissions.mask.Low, 2);
  f.render({ url: 'https://tenant.test/C' }); f.unmount();
  const errors = t.mock.method(console, 'error', () => {});
  await f.finish([4, 5], 3);
  assert.equal(errors.mock.callCount(), 0, 'no React unmounted-state warning');
});
test('current cross-site network rejection surfaces error and clears loading', async t => {
  const f = permissionFixture(t);
  await f.act(async () => { f.requests[0].reject(new Error('denied')); f.requests[1].resolve({ json: async () => ({}) }); });
  assert.equal(f.current.error.message, 'denied'); assert.equal(f.current.isLoading, false); assert.equal(f.current.webPermissions, undefined);
});
test('photo older success cannot finish newer loading and stale error cannot replace newer success', async t => {
  const f = photoFixture(t); t.mock.method(console, 'error', () => {});
  f.render({ userId: 'B' });
  await f.finish(0, new Blob(['A']));
  assert.equal(f.current.isLoading, true); assert.equal(f.current.photoBlob, undefined);
  let newer; f.act(() => { newer = f.current.reload(); });
  const b = new Blob(['B']); await f.finish(2, b); await newer;
  await f.act(async () => { f.requests[1].reject({ statusCode: 403, message: 'Forbidden' }); });
  assert.equal(f.current.error, undefined); assert.equal(f.current.photoBlob, b); assert.equal(f.current.isLoading, false);
});
test('photo current 403 maps error and unmount releases its blob URL without creating a stale one', async t => {
  const f = photoFixture(t); t.mock.method(console, 'error', () => {});
  await f.act(async () => { f.requests[0].reject({ statusCode: 403, message: 'Forbidden' }); });
  assert.match(f.current.error.message, /Insufficient permissions/); assert.equal(f.current.isLoading, false);
  let reload; f.act(() => { reload = f.current.reload(); });
  await f.finish(1, new Blob(['A'])); await reload;
  const url = f.current.photoUrl;
  const revokeOriginal = URL.revokeObjectURL;
  const revokes = t.mock.method(URL, 'revokeObjectURL', value => revokeOriginal(value));
  const createOriginal = URL.createObjectURL;
  const creates = t.mock.method(URL, 'createObjectURL', blob => createOriginal(blob));
  f.render({ userId: 'B' }); f.unmount();
  await f.finish(2, new Blob(['B']));
  assert.ok(revokes.mock.calls.some(call => call.arguments[0] === url));
  assert.equal(creates.mock.callCount(), 0);
});
function teamsFixture(t, sdk) {
  const h = createHarness(); t.after(h.close);
  h.context.current = { sdks: { microsoftTeams: sdk } };
  const state = h.load('core/state.internal.tsx');
  const store = h.load('core/runtime-store.internal.ts').createSPFxRuntimeStore();
  function Wrapper({ children }) { return h.React.createElement(state.SPFxRuntimeStoreContext.Provider, { value: store }, children); }
  const hook = h.load('hooks/useSPFxTeams.ts').useSPFxTeams;
  const rendered = h.mount(hook, {}, Wrapper);
  async function flush() { await h.act(async () => { for (let i = 0; i < 10; i++) await Promise.resolve(); }); }
  return { ...h, ...rendered, get current() { return rendered.current; }, store, flush };
}
for (const kind of ['wrapper-v2', 'direct-v2', 'direct-v1']) {
  test('Teams supports ' + kind + ' SDK shape and normalizes theme', async t => {
    const context = kind === 'direct-v1' ? { theme: 'contrast' } : { app: { theme: 'dark' } };
    const direct = kind === 'direct-v1' ? { getContext: callback => callback(context) } : { app: { getContext: async () => context } };
    const sdk = kind === 'wrapper-v2' ? { teamsJs: direct, context } : direct;
    const f = teamsFixture(t, sdk); await f.flush();
    assert.equal(f.current.supported, true); assert.equal(f.current.context, context);
    assert.equal(f.current.theme, kind === 'direct-v1' ? 'highContrast' : 'dark');
  });
}
test('Teams initialized context replacement acquires new SDK context', async t => {
  const first = { app: { theme: 'dark' } }; const next = { app: { theme: 'default' } };
  const f = teamsFixture(t, { app: { getContext: async () => first } }); await f.flush();
  f.context.current = { sdks: { microsoftTeams: { app: { getContext: async () => next } } } }; f.render(); await f.flush();
  assert.equal(f.current.context, next); assert.equal(f.current.theme, 'default');
});
test('Teams ignores a superseded deferred context and completion after unmount', async t => {
  const first = deferred(); const next = { app: { theme: 'dark' } };
  const f = teamsFixture(t, { app: { getContext: () => first.promise } });
  f.context.current = { sdks: { microsoftTeams: { app: { getContext: async () => next } } } }; f.render(); await f.flush();
  await f.act(async () => { first.resolve({ app: { theme: 'default' } }); await first.promise; });
  assert.equal(f.current.context, next);
  const after = deferred(); f.context.current = { sdks: { microsoftTeams: { app: { getContext: () => after.promise } } } }; f.render(); f.unmount();
  const snapshot = f.store.getState().teams;
  await f.act(async () => { after.resolve({ app: { theme: 'default' } }); await after.promise; });
  assert.equal(f.store.getState().teams, snapshot);
});
const tokenFor = scope => 'e30.' + Buffer.from(JSON.stringify({ aud: 'https://graph.microsoft.com', scp: scope, exp: 9999999999 })).toString('base64url') + '.signature';
function precheckFixture(t, provider) {
  const holder = { current: provider }; const config = { graph: ['User.Read'] };
  const h = createHarness({ './useSPFxAadTokenProvider': { useSPFxAadTokenProvider: () => ({ tokenProvider: holder.current, isInitializing: false, isReady: true }) }, './useSPFxInstanceInfo': { useSPFxInstanceInfo: () => ({ id: 'one', kind: 'WebPart' }) } }); t.after(h.close);
  const hook = h.load('hooks/useSPFxApiPermissionPrecheck.ts').useSPFxApiPermissionPrecheck;
  const rendered = h.mount(p => hook(config, p.options));
  async function flush() { await h.act(async () => { for (let i = 0; i < 20; i++) await Promise.resolve(); }); }
  return { ...h, ...rendered, get current() { return rendered.current; }, holder, flush };
}
test('API precheck invalidates available results and checks replacement provider', async t => {
  const f = precheckFixture(t, { getToken: async () => tokenFor('User.Read') }); await f.flush();
  assert.equal(f.current.results[0].status, 'available');
  const next = deferred(); let calls = 0;
  f.holder.current = { getToken: () => { calls++; return next.promise; } }; f.render();
  assert.equal(f.current.results.length, 0); assert.equal(f.current.isChecking, true); assert.equal(calls, 1);
  await f.act(async () => { next.resolve(tokenFor('Other.Scope')); await next.promise; }); await f.flush();
  assert.equal(f.current.results[0].status, 'missingScope');
});
test('API precheck replacement invalidates old in-flight check and retains new result', async t => {
  const old = deferred(); const f = precheckFixture(t, { getToken: () => old.promise });
  f.holder.current = { getToken: async () => tokenFor('Other.Scope') }; f.render(); await f.flush();
  await f.act(async () => { old.resolve(tokenFor('User.Read')); await old.promise; }); await f.flush();
  assert.equal(f.current.results[0].status, 'missingScope'); assert.equal(f.current.isChecking, false);
});
test('API precheck passive events, retry cache options and unmount preserve contracts', async t => {
  const popups = new Set(); const redirects = new Set(); const options = []; const observers = [];
  const event = handlers => ({ add(observer, handler) { observers.push(observer); handlers.add(handler); }, remove(_observer, handler) { handlers.delete(handler); } });
  let popupCanceled = 0; let redirectCanceled = 0;
  const provider = { popupEvent: event(popups), onBeforeRedirectEvent: event(redirects), getToken: async (_url, option) => {
    options.push(option.useCachedToken);
    for (const handler of popups) handler({ cancel() { popupCanceled++; } });
    for (const handler of redirects) handler({ cancel() { redirectCanceled++; } });
    return tokenFor('User.Read');
  } };
  const f = precheckFixture(t, provider); await f.flush();
  assert.equal(f.current.results[0].status, 'available');
  await f.act(async () => { await f.current.retry(); await f.current.retryWithoutCache(); });
  assert.deepEqual(options, [true, true, false]);
  assert.equal(popupCanceled, 3); assert.equal(redirectCanceled, 3);
  assert.equal(popups.size, 0); assert.equal(redirects.size, 0);
  assert.ok(observers.every(observer => observer.isDisposed));
  const pending = deferred(); f.holder.current = { getToken: () => pending.promise }; f.render(); f.unmount();
  const errors = t.mock.method(console, 'error', () => {});
  await f.act(async () => { pending.resolve(tokenFor('User.Read')); await pending.promise; }); await f.flush();
  assert.equal(errors.mock.callCount(), 0);
});

test('API precheck groups token acquisition and maps timeout without losing scope results', async t => {
  const h = createHarness(); t.after(h.close);
  let calls = 0;
  const service = h.load('services/spfx-api-permission-precheck.service.ts').createSPFxApiPermissionPrecheckService({ getToken: async () => { calls++; return tokenFor('User.Read Other.Scope'); } });
  const results = await service.check({ graph: ['User.Read', 'Other.Scope'] });
  assert.deepEqual(results.map(result => result.status), ['available', 'available']);
  assert.equal(calls, 1);
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const pending = deferred();
  const timeoutService = h.load('services/spfx-api-permission-precheck.service.ts').createSPFxApiPermissionPrecheckService({ getToken: () => pending.promise });
  const work = timeoutService.check({ graph: ['User.Read', 'Other.Scope'] }, { timeoutMs: 10 });
  t.mock.timers.tick(10);
  assert.deepEqual((await work).map(result => result.status), ['timeout', 'timeout']);
  pending.resolve(tokenFor('User.Read'));
});

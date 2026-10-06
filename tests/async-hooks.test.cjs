const test = require('node:test');
const assert = require('node:assert/strict');
const { loadHook, mount, deferred, start, settle, flush, catalogBoundary, driveBoundary, pnpBoundary, listPage, searchPage } = require('./async-test-harness.cjs');
const cleanup = (t, h) => { t.after(() => h.unmount()); return h; };
for (const first of [0, 1])
    test(`AsyncInvoke loading covers overlapping calls, completion ${first} first`, async (t) => {
        const ds = [deferred(), deferred()];
        const hook = loadHook('useAsyncInvoke.internal', 'useAsyncInvoke');
        const client = {};
        const h = cleanup(t, mount(() => hook(client), {}));
        const ps = ds.map(d => start(() => h.current.invoke(() => d.promise)));
        assert.equal(h.current.isLoading, true);
        await settle(async () => { ds[first].resolve(first); await ps[first]; });
        assert.equal(h.current.isLoading, true);
        await settle(async () => { ds[1 - first].resolve(1 - first); await ps[1 - first]; });
        assert.equal(h.current.isLoading, false);
    });
test('AsyncInvoke ignores former client error and preserves rejection after unmount', async (t) => {
    const hook = loadHook('useAsyncInvoke.internal', 'useAsyncInvoke');
    const h = cleanup(t, mount(p => hook(p.client), { client: {} }));
    const d = deferred();
    const p = start(() => h.current.invoke(() => d.promise));
    const rejected = p.catch(e => e);
    h.render({ client: {} });
    await settle(async () => { d.reject(new Error('old client')); await rejected; });
    assert.equal(h.current.error, undefined);
    assert.equal(h.current.isLoading, false);
    const end = deferred();
    const ep = start(() => h.current.invoke(() => end.promise)).catch(e => e);
    h.unmount();
    await settle(async () => { end.reject(new Error('unmounted')); assert.equal((await ep).message, 'unmounted'); });
});
for (const op of ['get', 'list'])
    test(`TenantKV ${op} awaits service and captures rejection as fallback`, async (t) => {
        const d = deferred();
        const service = { ensureListReady: async () => { }, get: () => d.promise, list: () => d.promise };
        const hook = loadHook('useSPFxTenantKeyValueStore', 'useSPFxTenantKeyValueStore', catalogBoundary(service));
        const h = cleanup(t, mount(() => hook(), {}));
        await flush();
        const p = start(() => h.current[op]('key'));
        const handled = p.catch(e => e);
        await flush();
        assert.equal(h.current.isLoading, true);
        await settle(async () => { d.reject(new Error('network')); const value = await handled; assert.deepEqual(value, op === 'get' ? undefined : []); });
        assert.equal(h.current.error.message, 'network');
        assert.equal(h.current.isLoading, false);
    });
test('TenantKV reads and writes track all pending operations independently', async (t) => {
    const ds = [deferred(), deferred(), deferred(), deferred()];
    const service = { ensureListReady: async () => { }, get: () => ds[0].promise, list: () => ds[1].promise, save: () => ds[2].promise, remove: () => ds[3].promise };
    const hook = loadHook('useSPFxTenantKeyValueStore', 'useSPFxTenantKeyValueStore', catalogBoundary(service));
    const h = cleanup(t, mount(() => hook(), {}));
    await flush();
    const ps = [start(() => h.current.get('x')), start(() => h.current.list()), start(() => h.current.save('x', 'v')), start(() => h.current.remove('x'))];
    await flush();
    await settle(async () => { ds[0].resolve({ key: 'x', value: 'v' }); ds[2].resolve(); await Promise.all([ps[0], ps[2]]); });
    assert.equal(h.current.isLoading, true);
    assert.equal(h.current.isWriting, true);
    await settle(async () => { ds[1].resolve([]); ds[3].resolve(); await Promise.all([ps[1], ps[3]]); });
    assert.equal(h.current.isLoading, false);
    assert.equal(h.current.isWriting, false);
});
test('TenantProperty autoFetch change discards reverse completion and stale rejection', async (t) => {
    const ds = { a: deferred(), b: deferred(), c: deferred() };
    const service = { get: k => ds[k].promise };
    const hook = loadHook('useSPFxTenantProperty', 'useSPFxTenantProperty', catalogBoundary(service));
    const h = cleanup(t, mount(p => hook(p.propertyKey, p.auto), { propertyKey: 'a', auto: true }));
    await flush();
    h.render({ propertyKey: 'b', auto: true });
    await flush();
    await settle(async () => { ds.b.resolve({ value: 'B', description: 'beta' }); });
    await settle(async () => { ds.a.resolve({ value: 'A', description: 'alpha' }); });
    assert.equal(h.current.data, 'B');
    assert.equal(h.current.description, 'beta');
    h.render({ propertyKey: 'c', auto: false });
    assert.equal(h.current.data, undefined);
    const p = start(() => h.current.load());
    await flush();
    h.render({ propertyKey: 'b', auto: false });
    await settle(async () => { ds.c.reject(new Error('old')); await p; });
    assert.equal(h.current.error, undefined);
});
test('OneDrive old 404 cannot create default in the current file', async (t) => {
    const ds = { a: deferred(), b: deferred() };
    const writes = [];
    const service = { read: k => ds[k].promise, write: async (k, v) => writes.push([k, v]) };
    const hook = loadHook('useSPFxOneDriveAppData', 'useSPFxOneDriveAppData', driveBoundary(service));
    const h = cleanup(t, mount(p => hook(p.file, { autoFetch: true, createIfMissing: true, defaultValue: 'default' }), { file: 'a' }));
    await flush();
    h.render({ file: 'b' });
    await flush();
    await settle(async () => { ds.b.resolve({ data: 'B', isNotFound: false }); });
    await settle(async () => { ds.a.resolve({ data: undefined, isNotFound: true }); });
    await flush();
    assert.equal(h.current.data, 'B');
    assert.equal(h.current.isNotFound, false);
    assert.deepEqual(writes, []);
});
test('OneDrive identity clears old missing state before createIfMissing effect', async (t) => {
    const writes = [];
    const service = { read: async () => ({ isNotFound: true, data: undefined }), write: async (k, v) => writes.push([k, v]) };
    const hook = loadHook('useSPFxOneDriveAppData', 'useSPFxOneDriveAppData', driveBoundary(service));
    const h = cleanup(t, mount(p => hook(p.file, { autoFetch: p.auto, createIfMissing: p.create, defaultValue: 'default' }), { file: 'a', auto: true, create: false }));
    await flush();
    assert.equal(h.current.isNotFound, true);
    h.render({ file: 'b', auto: false, create: true });
    await flush();
    assert.deepEqual(writes, []);
    assert.equal(h.current.data, 'default');
});
test('OneDrive newer write wins over older read and reverse write completion', async (t) => {
    const read = deferred(), w1 = deferred(), w2 = deferred();
    const service = { read: () => read.promise, write: (_k, v) => v === 'first' ? w1.promise : w2.promise };
    const hook = loadHook('useSPFxOneDriveAppData', 'useSPFxOneDriveAppData', driveBoundary(service));
    const h = cleanup(t, mount(() => hook('file', { autoFetch: false }), {}));
    const rp = start(() => h.current.load());
    const p1 = start(() => h.current.write('first'));
    const p2 = start(() => h.current.write('second'));
    await settle(async () => { w2.resolve(); await p2; });
    assert.equal(h.current.isWriting, true);
    await settle(async () => { w1.resolve(); read.resolve({ data: 'old read', isNotFound: false }); await Promise.all([p1, rp]); });
    assert.equal(h.current.data, 'second');
    assert.equal(h.current.isWriting, false);
});
test('PnPList identity and latest query prevent stale state/error', async (t) => {
    const a = deferred(), b = deferred(), c = deferred();
    let n = 0;
    const hook = loadHook('useSPFxPnPList', 'useSPFxPnPList', pnpBoundary((_sp, title) => ({ query: () => title === 'a' ? a.promise : (n++ === 0 ? b.promise : c.promise) }), 'list'));
    const h = cleanup(t, mount(p => hook(p.title), { title: 'a' }));
    const pa = start(() => h.current.query());
    h.render({ title: 'b' });
    const pb = start(() => h.current.query());
    await settle(async () => { b.resolve(listPage(['B'])); await pb; });
    await settle(async () => { a.resolve(listPage(['A'])); await pa; });
    assert.deepEqual(h.current.items, ['B']);
    const pc = start(() => h.current.query()).catch(e => e);
    h.render({ title: 'c' });
    assert.deepEqual(h.current.items, []);
    await settle(async () => { c.reject(new Error('old')); await pc; });
    assert.equal(h.current.error, undefined);
});
test('PnPList rejects duplicate page dispatch and ignores page from superseded query', async (t) => {
    const page = deferred();
    let pages = 0;
    const service = { query: async (builder) => listPage([builder ? 'new' : 'first'], 1, true), loadMore: () => { pages++; return page.promise; } };
    const hook = loadHook('useSPFxPnPList', 'useSPFxPnPList', pnpBoundary(() => service, 'list'));
    const h = cleanup(t, mount(() => hook('list', { pageSize: 1 }), {}));
    await settle(() => h.current.query());
    const a = start(() => h.current.loadMore()), b = start(() => h.current.loadMore());
    assert.equal(pages, 1);
    await settle(() => h.current.query(x => x));
    await settle(async () => { page.resolve(listPage(['stale'], 2, false)); await Promise.all([a, b]); });
    assert.deepEqual(h.current.items, ['new']);
});
test('PnPSearch clears refiners on new search but keeps them on refetch', async (t) => {
    const calls = [];
    const service = { search: async (q, o) => { calls.push({ q, filters: [...o.refinementFilters] }); return searchPage([]); } };
    const hook = loadHook('useSPFxPnPSearch', 'useSPFxPnPSearch', pnpBoundary(() => service, 'search'));
    const h = cleanup(t, mount(() => hook(), {}));
    await settle(() => h.current.search('first'));
    await settle(() => h.current.applyRefiner('color', 'red'));
    await settle(() => h.current.refetch());
    assert.deepEqual(calls[2].filters, [['color', ['red']]]);
    await settle(() => h.current.search('second'));
    assert.deepEqual(calls[3].filters, []);
});
test('PnPSearch latest query ignores reversed responses and stale errors', async (t) => {
    const ds = { a: deferred(), b: deferred(), c: deferred() };
    const service = { search: q => ds[q].promise };
    const hook = loadHook('useSPFxPnPSearch', 'useSPFxPnPSearch', pnpBoundary(() => service, 'search'));
    const h = cleanup(t, mount(() => hook(), {}));
    const a = start(() => h.current.search('a'));
    const b = start(() => h.current.search('b'));
    await settle(async () => { ds.b.resolve(searchPage(['B'])); await b; });
    await settle(async () => { ds.a.resolve(searchPage(['A'])); await a; });
    assert.equal(h.current.results[0].id, 'B');
    const c = start(() => h.current.search('c')).catch(e => e);
    await settle(() => h.current.search('b'));
    await settle(async () => { ds.c.reject(new Error('old')); await c; });
    assert.equal(h.current.error, undefined);
});
test('PnPSearch same-tick loadMore dispatches one page and superseded page cannot append', async (t) => {
    const page = deferred();
    let pages = 0;
    const service = { search: (q, o) => o.startRow ? (pages++, page.promise) : Promise.resolve(searchPage([q], 3)) };
    const hook = loadHook('useSPFxPnPSearch', 'useSPFxPnPSearch', pnpBoundary(() => service, 'search'));
    const h = cleanup(t, mount(() => hook({ pageSize: 1 }), {}));
    await settle(() => h.current.search('first'));
    const a = start(() => h.current.loadMore()), b = start(() => h.current.loadMore());
    assert.equal(pages, 1);
    await settle(() => h.current.search('new'));
    await settle(async () => { page.resolve(searchPage(['stale'], 3)); await Promise.all([a, b]); });
    assert.deepEqual(h.current.results.map(x => x.id), ['new']);
});
test('AppCatalog cached URL invalidates with pageContext/client service identity', async (t) => {
    let client = {}, context = { tenant: 'a' };
    const ds = { a: deferred(), b: deferred() };
    const hook = loadHook('useAppCatalogUrl.internal', 'useAppCatalogUrl', { './useSPFxSPHttpClient': { useSPFxSPHttpClient: () => ({ client }) }, './useSPFxPageContext': { useSPFxPageContext: () => context }, '../services/spfx-app-catalog.service': { createSPFxAppCatalogService: (_c, p) => ({ discoverUrl: () => ds[p.tenant].promise, canCurrentUserWrite: async () => true }) } });
    const h = cleanup(t, mount(() => hook(), {}));
    const a = start(() => h.current.discoverAppCatalogUrl());
    client = {};
    context = { tenant: 'b' };
    h.render({});
    const b = start(() => h.current.discoverAppCatalogUrl());
    await settle(async () => { ds.b.resolve('catalog-b'); await b; });
    await settle(async () => { ds.a.resolve('catalog-a'); await a; });
    assert.equal(h.current.appCatalogUrl, 'catalog-b');
    assert.equal(await h.current.discoverAppCatalogUrl(), 'catalog-b');
});
test('OneDrive inline object default does not restart a successful autoFetch', async (t) => {
    const first = deferred(), next = deferred();
    let reads = 0;
    const service = { read: () => ++reads === 1 ? first.promise : next.promise };
    const hook = loadHook('useSPFxOneDriveAppData', 'useSPFxOneDriveAppData', driveBoundary(service));
    const h = cleanup(t, mount(() => hook('file', { defaultValue: { theme: 'light' } }), {}));
    await flush();
    assert.equal(reads, 1);
    await settle(async () => { first.resolve({ data: { theme: 'dark' }, isNotFound: false }); });
    assert.equal(reads, 1);
    assert.equal(h.current.data.theme, 'dark');
});
test('AsyncInvoke newer success is not replaced by an older failure and clearError works', async (t) => {
    const client = {};
    const hook = loadHook('useAsyncInvoke.internal', 'useAsyncInvoke');
    const h = cleanup(t, mount(() => hook(client), {}));
    const a = deferred(), b = deferred();
    const pa = start(() => h.current.invoke(() => a.promise)).catch(e => e);
    const pb = start(() => h.current.invoke(() => b.promise));
    await settle(async () => { b.resolve('ok'); assert.equal(await pb, 'ok'); });
    await settle(async () => { a.reject(new Error('old')); await pa; });
    assert.equal(h.current.error, undefined);
    await settle(async () => { await h.current.invoke(async () => { throw 'current'; }).catch(e => assert.equal(e.message, 'current')); });
    assert.equal(h.current.error.message, 'current');
    start(() => h.current.clearError());
    assert.equal(h.current.error, undefined);
});
test('TenantKV provisioning read fallback remains successful and write rejection propagates', async (t) => {
    let ready = false;
    const fail = new Error('write denied');
    const service = { ensureListReady: async () => {
            if (!ready)
                throw new Error('not provisioned');
        }, get: async () => ({ key: 'x', value: 1 }), list: async () => [], save: async () => { throw fail; }, remove: async () => { throw fail; } };
    const hook = loadHook('useSPFxTenantKeyValueStore', 'useSPFxTenantKeyValueStore', catalogBoundary(service));
    const h = cleanup(t, mount(() => hook(), {}));
    await flush();
    await settle(async () => { assert.equal(await h.current.get('x'), undefined); assert.deepEqual(await h.current.list(), []); });
    assert.equal(h.current.error, undefined);
    ready = true;
    await settle(async () => { assert.equal((await h.current.get('x')).value, 1); assert.equal(await h.current.save('x', 1).catch(e => e), fail); });
    assert.equal(h.current.writeError, fail);
    assert.equal(h.current.isWriting, false);
});
test('PnPSearch old suggestion rejection cannot replace newer successful search error state', async (t) => {
    const suggestion = deferred();
    const service = { suggest: () => suggestion.promise, search: async () => searchPage(['fresh']) };
    const hook = loadHook('useSPFxPnPSearch', 'useSPFxPnPSearch', pnpBoundary(() => service, 'search'));
    const h = cleanup(t, mount(() => hook(), {}));
    const p = start(() => h.current.suggest('old')).catch(e => e);
    await settle(() => h.current.search('fresh'));
    await settle(async () => { suggestion.reject(new Error('old suggestion')); await p; });
    assert.equal(h.current.error, undefined);
});
for (const kind of ['property', 'drive', 'list', 'search', 'store'])
    test(`${kind} deferred failure after unmount cannot update React state`, async (t) => {
        const d = deferred();
        let hook, invoke;
        if (kind === 'property') {
            hook = loadHook('useSPFxTenantProperty', 'useSPFxTenantProperty', catalogBoundary({ get: () => d.promise }));
            invoke = h => h.load();
        }
        if (kind === 'drive') {
            hook = loadHook('useSPFxOneDriveAppData', 'useSPFxOneDriveAppData', driveBoundary({ read: () => d.promise }));
            invoke = h => h.load();
        }
        if (kind === 'list') {
            hook = loadHook('useSPFxPnPList', 'useSPFxPnPList', pnpBoundary(() => ({ query: () => d.promise }), 'list'));
            invoke = h => h.query();
        }
        if (kind === 'search') {
            hook = loadHook('useSPFxPnPSearch', 'useSPFxPnPSearch', pnpBoundary(() => ({ search: () => d.promise }), 'search'));
            invoke = h => h.search('x');
        }
        if (kind === 'store') {
            hook = loadHook('useSPFxTenantKeyValueStore', 'useSPFxTenantKeyValueStore', catalogBoundary({ ensureListReady: async () => { }, get: () => d.promise }));
            invoke = h => h.get('x');
        }
        const h = cleanup(t, mount(() => kind === 'property' ? hook('x', false) : kind === 'drive' ? hook('file', { autoFetch: false }) : kind === 'list' ? hook('list') : hook(), {}));
        const errors = [];
        const original = console.error;
        console.error = (...args) => errors.push(args.map(String).join(' '));
        t.after(() => { console.error = original; });
        const p = start(() => invoke(h.current)).catch(e => e);
        await flush();
        h.unmount();
        await settle(async () => { d.reject(new Error('late')); await p; });
        assert.deepEqual(errors, []);
    });
test('PnPSearch query pagination uses accumulated count for hasMore', async (t) => {
    const service = { search: async (_q, o) => searchPage([String(o.startRow)], 3) };
    const hook = loadHook('useSPFxPnPSearch', 'useSPFxPnPSearch', pnpBoundary(() => service, 'search'));
    const h = cleanup(t, mount(() => hook({ pageSize: 1 }), {}));
    await settle(() => h.current.search('x'));
    assert.equal(h.current.hasMore, true);
    await settle(() => h.current.loadMore());
    assert.equal(h.current.hasMore, true);
    await settle(() => h.current.loadMore());
    assert.equal(h.current.hasMore, false);
    assert.deepEqual(h.current.results.map(x => x.id), ['0', '1', '2']);
});
test('PnPList latest query wins within one identity and failed page retries the same offset', async (t) => {
    const a = deferred(), b = deferred();
    let queries = 0, pages = 0;
    const offsets = [];
    const service = { query: () => ++queries === 1 ? a.promise : b.promise, loadMore: async (_q, _size, skip) => {
            offsets.push(skip);
            if (++pages === 1)
                throw new Error('page failed');
            return listPage(['page'], 2, false);
        } };
    const hook = loadHook('useSPFxPnPList', 'useSPFxPnPList', pnpBoundary(() => service, 'list'));
    const h = cleanup(t, mount(() => hook('list', { pageSize: 1 }), {}));
    const pa = start(() => h.current.query());
    const pb = start(() => h.current.query());
    await settle(async () => { b.resolve(listPage(['new'], 1, true)); await pb; });
    await settle(async () => { a.resolve(listPage(['old'], 1, true)); await pa; });
    assert.deepEqual(h.current.items, ['new']);
    await settle(async () => { await h.current.loadMore().catch(e => assert.equal(e.message, 'page failed')); });
    assert.equal(h.current.loadingMore, false);
    await settle(() => h.current.loadMore());
    assert.deepEqual(offsets, [1, 1]);
    assert.deepEqual(h.current.items, ['new', 'page']);
    assert.equal(h.current.error, undefined);
});
test('PnPSearch failed page keeps offset and inline options remain stable', async (t) => {
    let pages = 0;
    const rows = [];
    const service = { search: async (_q, o) => {
            rows.push(o.startRow);
            if (o.startRow === 1 && ++pages === 1)
                throw new Error('failed');
            return searchPage([String(o.startRow)], 2);
        } };
    const hook = loadHook('useSPFxPnPSearch', 'useSPFxPnPSearch', pnpBoundary(() => service, 'search'));
    const h = cleanup(t, mount(() => hook({ pageSize: 1, selectProperties: ['Title'] }), {}));
    await settle(() => h.current.search('x'));
    await settle(async () => { await h.current.loadMore().catch(e => assert.equal(e.message, 'failed')); });
    assert.equal(h.current.loadingMore, false);
    await settle(() => h.current.loadMore());
    assert.deepEqual(rows, [0, 1, 1]);
    assert.equal(h.current.hasMore, false);
    assert.equal(h.current.error, undefined);
});
test('OneDrive write errors propagate and successful missing creation writes once', async (t) => {
    let fail = true;
    const writes = [];
    const service = { read: async () => ({ data: undefined, isNotFound: true }), write: async (k, v) => {
            writes.push([k, v]);
            if (fail)
                throw new Error('denied');
        } };
    const hook = loadHook('useSPFxOneDriveAppData', 'useSPFxOneDriveAppData', driveBoundary(service));
    const h = cleanup(t, mount(() => hook('file', { autoFetch: false, defaultValue: 'default', createIfMissing: true }), {}));
    await settle(async () => { await h.current.write('value').catch(e => assert.equal(e.message, 'denied')); });
    assert.equal(h.current.writeError.message, 'denied');
    assert.equal(h.current.isWriting, false);
    fail = false;
    await settle(() => h.current.load());
    await flush();
    assert.deepEqual(writes, [['file', 'value'], ['file', 'default']]);
    assert.equal(h.current.data, 'default');
    assert.equal(h.current.isNotFound, false);
    assert.equal(h.current.writeError, undefined);
});
test('AppCatalog resolved cache is cleared on service change before discovery', async (t) => {
    let context = { tenant: 'a' };
    const client = {};
    const hook = loadHook('useAppCatalogUrl.internal', 'useAppCatalogUrl', { './useSPFxSPHttpClient': { useSPFxSPHttpClient: () => ({ client }) }, './useSPFxPageContext': { useSPFxPageContext: () => context }, '../services/spfx-app-catalog.service': { createSPFxAppCatalogService: (_c, p) => ({ discoverUrl: async () => 'catalog-' + p.tenant, canCurrentUserWrite: async () => true }) } });
    const h = cleanup(t, mount(() => hook(), {}));
    await settle(() => h.current.discoverAppCatalogUrl());
    assert.equal(h.current.appCatalogUrl, 'catalog-a');
    context = { tenant: 'b' };
    h.render({});
    assert.equal(h.current.appCatalogUrl, undefined);
    await settle(async () => { assert.equal(await h.current.discoverAppCatalogUrl(), 'catalog-b'); });
    assert.equal(h.current.appCatalogUrl, 'catalog-b');
});
test('OneDrive lazy default is the current identity seed without starting a read', async (t) => {
    let reads = 0;
    const service = { read: async () => { reads++; return { data: 'remote', isNotFound: false }; } };
    const hook = loadHook('useSPFxOneDriveAppData', 'useSPFxOneDriveAppData', driveBoundary(service));
    const h = cleanup(t, mount(p => hook(p.file, { autoFetch: false, defaultValue: { theme: p.theme } }), { file: 'a', theme: 'light' }));
    assert.deepEqual(h.current.data, { theme: 'light' });
    assert.equal(h.current.isReady, true);
    h.render({ file: 'b', theme: 'dark' });
    assert.deepEqual(h.current.data, { theme: 'dark' });
    assert.equal(h.current.isNotFound, false);
    assert.equal(reads, 0);
});
for (const operation of ['create', 'update', 'remove', 'createBatch', 'updateBatch', 'removeBatch']) {
    test(`PnPList pending ${operation} refetches the current query after replacement`, async (t) => {
        const mutation = deferred();
        const calls = [];
        const A = x => x;
        const B = x => x;
        const service = {
            query: async (builder) => {
                const name = builder === A ? 'A' : 'B';
                calls.push(name);
                return listPage([name], 1, true);
            },
            [operation]: () => mutation.promise
        };
        const hook = loadHook('useSPFxPnPList', 'useSPFxPnPList', pnpBoundary(() => service, 'list'));
        const h = cleanup(t, mount(() => hook('list', { pageSize: 1 }), {}));
        await settle(() => h.current.query(A));
        const args = operation === 'create' ? [{ Title: 'new' }] : operation === 'update' ? [1, { Title: 'new' }] : operation === 'remove' ? [1] : operation === 'createBatch' ? [[{ Title: 'new' }]] : operation === 'updateBatch' ? [[{ id: 1, item: { Title: 'new' } }]] : [[1]];
        const pending = start(() => h.current[operation](...args));
        await settle(() => h.current.query(B));
        await settle(async () => {
            mutation.resolve(operation.endsWith('Batch') ? { value: operation === 'createBatch' ? [1] : undefined, errors: [], summaryError: undefined } : operation === 'create' ? 1 : undefined);
            await pending;
        });
        await settle(() => new Promise(resolve => setTimeout(resolve, 120)));
        assert.deepEqual(calls, ['A', 'B', 'B']);
        assert.deepEqual(h.current.items, ['B']);
        assert.equal(h.current.loading, false);
    });
}
test('PnPList failed refetch preserves the successful next-page offset', async (t) => {
    let queries = 0;
    const offsets = [];
    const failure = new Error('refetch failed');
    const service = {
        query: async () => {
            if (++queries > 1)
                throw failure;
            return { items: ['0', '1'], nextSkip: 2, hasMore: true, effectivePageSize: 2 };
        },
        loadMore: async (_builder, _size, skip) => {
            offsets.push(skip);
            return { items: ['2', '3'], nextSkip: 4, hasMore: false, effectivePageSize: 2 };
        }
    };
    const hook = loadHook('useSPFxPnPList', 'useSPFxPnPList', pnpBoundary(() => service, 'list'));
    const h = cleanup(t, mount(() => hook('list', { pageSize: 2 }), {}));
    await settle(() => h.current.query());
    await settle(async () => { assert.equal(await h.current.refetch().catch(e => e), failure); });
    assert.deepEqual(h.current.items, ['0', '1']);
    assert.equal(h.current.hasMore, true);
    assert.equal(h.current.error, failure);
    await settle(() => h.current.loadMore());
    assert.deepEqual(offsets, [2]);
    assert.deepEqual(h.current.items, ['0', '1', '2', '3']);
    assert.equal(h.current.error, undefined);
});
test('PnPList CRUD refresh uses the replacement query even while its first response is pending', async (t) => {
    const mutation = deferred(), replacement = deferred();
    const calls = [];
    const A = x => x, B = x => x;
    let replacementCalls = 0;
    const service = {
        query: builder => {
            calls.push(builder === A ? 'A' : 'B');
            return builder === B && ++replacementCalls === 1 ? replacement.promise : Promise.resolve(listPage([builder === A ? 'A' : 'fresh B'], 1, true));
        },
        create: () => mutation.promise
    };
    const hook = loadHook('useSPFxPnPList', 'useSPFxPnPList', pnpBoundary(() => service, 'list'));
    const h = cleanup(t, mount(() => hook('list', { pageSize: 1 }), {}));
    await settle(() => h.current.query(A));
    const write = start(() => h.current.create({ Title: 'new' }));
    const read = start(() => h.current.query(B));
    await settle(async () => { mutation.resolve(1); await write; });
    await settle(() => new Promise(resolve => setTimeout(resolve, 120)));
    assert.deepEqual(calls, ['A', 'B', 'B']);
    assert.deepEqual(h.current.items, ['fresh B']);
    await settle(async () => { replacement.resolve(listPage(['stale B'], 1, true)); await read; });
    assert.deepEqual(h.current.items, ['fresh B']);
});
test('PnPList a newer query invalidates an already scheduled CRUD refresh', async (t) => {
    const calls = [];
    const A = x => x, B = x => x;
    const service = { query: async (builder) => { calls.push(builder === A ? 'A' : 'B'); return listPage([builder === A ? 'A' : 'B']); }, create: async () => 1 };
    const hook = loadHook('useSPFxPnPList', 'useSPFxPnPList', pnpBoundary(() => service, 'list'));
    const h = cleanup(t, mount(() => hook('list'), {}));
    await settle(() => h.current.query(A));
    await settle(() => h.current.create({ Title: 'new' }));
    await settle(() => h.current.query(B));
    await settle(() => new Promise(resolve => setTimeout(resolve, 120)));
    assert.deepEqual(calls, ['A', 'B']);
    assert.deepEqual(h.current.items, ['B']);
});
test('PnPList former-service mutation cannot refresh the replacement list', async (t) => {
    const mutation = deferred();
    const calls = [];
    const hook = loadHook('useSPFxPnPList', 'useSPFxPnPList', pnpBoundary((_sp, title) => ({ query: async () => { calls.push(title); return listPage([title]); }, create: () => mutation.promise }), 'list'));
    const h = cleanup(t, mount(p => hook(p.title), { title: 'a' }));
    await settle(() => h.current.query());
    const write = start(() => h.current.create({ Title: 'new' }));
    h.render({ title: 'b' });
    await settle(() => h.current.query());
    await settle(async () => { mutation.resolve(1); await write; });
    await settle(() => new Promise(resolve => setTimeout(resolve, 120)));
    assert.deepEqual(calls, ['a', 'b']);
    assert.deepEqual(h.current.items, ['b']);
});
test('PnPList unmount cancels scheduled mutation refresh and ignores later mutation completion', async (t) => {
    const mutation = deferred();
    let queries = 0, writes = 0;
    const service = { query: async () => { queries++; return listPage(['data']); }, create: () => ++writes === 1 ? Promise.resolve(1) : mutation.promise };
    const hook = loadHook('useSPFxPnPList', 'useSPFxPnPList', pnpBoundary(() => service, 'list'));
    const h = cleanup(t, mount(() => hook('list'), {}));
    await settle(() => h.current.query());
    await settle(() => h.current.create({ Title: 'first' }));
    const second = start(() => h.current.create({ Title: 'second' }));
    h.unmount();
    await settle(async () => { mutation.resolve(2); await second; });
    await settle(() => new Promise(resolve => setTimeout(resolve, 120)));
    assert.equal(queries, 1);
});

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadHook, mount, deferred, start, settle, pnpBoundary, listPage } = require('./async-test-harness.cjs');
const { createHarness } = require('./services-test-harness.cjs');
const { createTransport } = require('./fixtures/pnp-list-transport.cjs');
const wrappers = [
    ['Id', 'id', 'id', '11111111-2222-3333-4444-555555555555', 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'],
    ['Url', 'url', 'serverRelativeUrl', '/sites/projects/Lists/Tasks', '/sites/projects/Lists/Other'],
    ['Path', 'path', 'webRelativePath', 'Lists/Tasks', 'Lists/Other']
];
const cleanup = (t, h) => { t.after(() => h.unmount()); return h; };
const waitRefresh = () => settle(() => new Promise(resolve => setTimeout(resolve, 120)));
const mutationArgs = {
    create: [{}], update: [1, {}], remove: [1],
    createBatch: [[{}]], updateBatch: [[{ id: 1, item: {} }]], removeBatch: [[1]]
};
const mutationResult = op => op.endsWith('Batch') ? { value: op === 'createBatch' ? [1] : undefined, errors: [] } : op === 'create' ? 1 : undefined;
for (const [suffix, kind, field, a, b] of wrappers) {
    const name = 'useSPFxPnPListBy' + suffix;
    const load = factory => loadHook(name, name, pnpBoundary(factory, 'list'));
    test(`${name} preserves equal primitive identity and resets target, client, initialization and pageSize`, async t => {
        const calls = [];
        const hook = load((sp, target, size) => {
            calls.push({ sp, target, size });
            return { query: async () => listPage([target[field]], 1, true) };
        });
        const context = { sp: {}, isInitialized: true };
        const props = { value: a, size: 1, context };
        const h = cleanup(t, mount(p => hook(p.value, { pageSize: p.size }, p.context), props));
        assert.equal(calls.length, 1);
        assert.deepEqual(calls[0].target, { kind, [field]: a });
        assert.equal(h.current.loading, false);
        assert.deepEqual(h.current.items, []);
        await settle(() => h.current.query());
        const query = h.current.query;
        h.render({ ...props, context: { ...context } });
        assert.equal(calls.length, 1);
        assert.equal(h.current.query, query);
        assert.deepEqual(h.current.items, [a]);
        for (const next of [
            { ...props, value: b },
            { ...props, value: b, size: 2 },
            { ...props, value: b, size: 2, context: { sp: {}, isInitialized: true } }
        ]) {
            const before = calls.length;
            h.render(next);
            assert.equal(calls.length, before + 1);
            assert.deepEqual(h.current.items, []);
            assert.equal(h.current.hasMore, false);
            await settle(async () => { await assert.rejects(h.current.refetch(), /No previous query/); });
            await settle(() => h.current.query());
        }
        h.render({ ...props, context: { ...context, isInitialized: false } });
        assert.deepEqual(h.current.items, []);
        h.render(props);
        assert.deepEqual(h.current.items, []);
        assert.equal(calls.length, 5);
    });
    for (const reject of [false, true]) {
        test(`${name} suppresses stale query ${reject ? 'error' : 'success'} across targets`, async t => {
            const old = deferred();
            const hook = load((_sp, target) => ({ query: () => target[field] === a ? old.promise : Promise.resolve(listPage([b])) }));
            const h = cleanup(t, mount(p => hook(p.value), { value: a }));
            const pending = start(() => h.current.query());
            const outcome = pending.catch(error => error);
            h.render({ value: b });
            await settle(() => h.current.query());
            await settle(async () => { reject ? old.reject(new Error('old')) : old.resolve(listPage([a])); await outcome; });
            assert.deepEqual(h.current.items, [b]);
            assert.equal(h.current.error, undefined);
            assert.equal(h.current.loading, false);
        });
    }
    test(`${name} isolates stale pages and retains page dispatch locks`, async t => {
        const page = deferred();
        let pages = 0;
        const hook = load((_sp, target) => ({ query: async () => listPage([target[field]], 1, true), loadMore: () => { pages++; return page.promise; } }));
        const h = cleanup(t, mount(p => hook(p.value, { pageSize: 1 }), { value: a }));
        await settle(() => h.current.query());
        const pending = start(() => h.current.loadMore());
        assert.deepEqual(await h.current.loadMore(), []);
        assert.equal(pages, 1);
        h.render({ value: b });
        await settle(() => h.current.query());
        await settle(async () => { page.resolve(listPage(['old page'], 2)); await pending; });
        assert.deepEqual(h.current.items, [b]);
        assert.equal(h.current.loadingMore, false);
        assert.equal(h.current.hasMore, true);
    });
    for (const op of Object.keys(mutationArgs)) {
        test(`${name} stale ${op} cannot refresh or set errors on the new target`, async t => {
            const mutation = deferred();
            const calls = [];
            const hook = load((_sp, target) => ({
                query: async () => { calls.push(target[field]); return listPage([target[field]]); },
                [op]: () => mutation.promise
            }));
            const h = cleanup(t, mount(p => hook(p.value), { value: a }));
            await settle(() => h.current.query());
            const pending = start(() => h.current[op](...mutationArgs[op]));
            h.render({ value: b });
            await settle(() => h.current.query());
            await settle(async () => { mutation.resolve(mutationResult(op)); await pending; });
            await waitRefresh();
            assert.deepEqual(calls, [a, b]);
            assert.deepEqual(h.current.items, [b]);
            assert.equal(h.current.error, undefined);
        });
    }
    for (const op of ['getById', ...Object.keys(mutationArgs)]) {
        test(`${name} suppresses former-target ${op} rejection`, async t => {
            const failed = deferred();
            const hook = load((_sp, target) => ({ query: async () => listPage([target[field]]), [op]: () => failed.promise }));
            const h = cleanup(t, mount(p => hook(p.value), { value: a }));
            const pending = start(() => h.current[op](...(mutationArgs[op] || [1])));
            const caught = pending.catch(error => error);
            h.render({ value: b });
            await settle(() => h.current.query());
            const error = new Error('former target failed');
            await settle(async () => { failed.reject(error); assert.equal(await caught, op === 'getById' ? undefined : error); });
            assert.equal(h.current.error, undefined);
            assert.deepEqual(h.current.items, [b]);
        });
    }
    test(`${name} suppresses stale page errors after client replacement`, async t => {
        const page = deferred();
        const hook = load(() => ({ query: async () => listPage(['fresh'], 1, true), loadMore: () => page.promise }));
        const h = cleanup(t, mount(p => hook(a, { pageSize: 1 }, p.context), { context: { sp: {}, isInitialized: true } }));
        await settle(() => h.current.query());
        const pending = start(() => h.current.loadMore()).catch(error => error);
        h.render({ context: { sp: {}, isInitialized: true } });
        await settle(() => h.current.query());
        const error = new Error('former client page failed');
        await settle(async () => { page.reject(error); assert.equal(await pending, error); });
        assert.equal(h.current.error, undefined);
        assert.equal(h.current.loadingMore, false);
        assert.deepEqual(h.current.items, ['fresh']);
    });
    test(`${name} cancels scheduled debounce on identity change and unmount`, async t => {
        const calls = [];
        const late = deferred();
        let writes = 0;
        const hook = load((_sp, target) => ({ query: async () => { calls.push(target[field]); return listPage([target[field]]); }, create: () => ++writes === 3 ? late.promise : Promise.resolve(1) }));
        const h = cleanup(t, mount(p => hook(p.value), { value: a }));
        await settle(() => h.current.query());
        await settle(() => h.current.create({}));
        h.render({ value: b });
        await settle(() => h.current.query());
        await waitRefresh();
        assert.deepEqual(calls, [a, b]);
        await settle(() => h.current.create({}));
        const pending = start(() => h.current.create({}));
        h.unmount();
        await settle(async () => { late.resolve(1); await pending; });
        await waitRefresh();
        assert.deepEqual(calls, [a, b]);
    });
    test(`${name} simultaneous instances retain independent state`, async t => {
        const pending = deferred();
        let factories = 0;
        const hook = load(() => { const n = ++factories; return { query: () => n === 1 ? pending.promise : Promise.resolve(listPage(['second'])) }; });
        const first = cleanup(t, mount(() => hook(a), {}));
        const second = cleanup(t, mount(() => hook(a), {}));
        const request = start(() => first.current.query());
        await settle(() => second.current.query());
        assert.equal(first.current.loading, true);
        assert.deepEqual(second.current.items, ['second']);
        first.unmount();
        await settle(async () => { pending.resolve(listPage(['first'])); await request; });
        assert.deepEqual(second.current.items, ['second']);
    });
    test(`${name} invalid editable input mounts without HTTP and rejects only on request`, async t => {
        const services = createHarness();
        const { createSPFxPnPListService } = services.load('services/spfx-pnp-list.service.ts');
        const transport = createTransport(services, undefined, [{ value: [{ Id: 1 }] }]);
        const hook = loadHook(name, name, pnpBoundary(createSPFxPnPListService, 'list'));
        const context = { sp: transport.sp, isInitialized: true };
        const h = cleanup(t, mount(p => hook(p.value, undefined, context), { value: '' }));
        assert.equal(transport.calls.length, 0);
        assert.equal(h.current.error, undefined);
        await settle(async () => { await assert.rejects(h.current.query()); });
        assert.ok(h.current.error instanceof Error);
        assert.equal(transport.calls.length, 0);
        h.render({ value: a });
        assert.equal(h.current.error, undefined);
        await settle(() => h.current.query());
        assert.deepEqual(h.current.items, [{ Id: 1 }]);
        assert.equal(transport.calls.length, 1);
    });
}
test('legacy title wrapper still forwards raw strings without selector inference', async t => {
    const targets = [];
    const hook = loadHook('useSPFxPnPList', 'useSPFxPnPList', pnpBoundary((_sp, target) => { targets.push(target); return { query: async () => listPage([target]) }; }, 'list'));
    const h = cleanup(t, mount(p => hook(p.title), { title: '/Lists/Tasks' }));
    await settle(() => h.current.query());
    h.render({ title: '11111111-2222-3333-4444-555555555555' });
    assert.deepEqual(targets, ['/Lists/Tasks', '11111111-2222-3333-4444-555555555555']);
});

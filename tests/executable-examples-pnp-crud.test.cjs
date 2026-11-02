const assert = require('node:assert/strict');
const { test } = require('node:test');
const path = require('node:path');
// Only this Dialog suite needs native DOM Node; restore its original descriptor at suite end.
const originalNode = Object.getOwnPropertyDescriptor(global, 'Node');
const { createClientsFixture } = require('./fixtures/executable-examples/clients-scenarios.cjs');
const { Simulate } = require('react-dom/test-utils');
const { flushSync } = require('react-dom');
Object.defineProperty(global, 'Node', { configurable: true, value: window.Node });
require('node:test').after(() => { if (originalNode)
    Object.defineProperty(global, 'Node', originalNode);
else
    delete global.Node; });
const title = 'Manage list items with paging and captured delete consent';
const owner = 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPList.ts#useSPFxPnPList';
function authored(value, match) {
    if (Array.isArray(value)) {
        for (const child of value) {
            const found = authored(child, match);
            if (found)
                return found;
        }
        return;
    }
    if (!value || typeof value !== 'object')
        return;
    if (match(value.props ?? {}))
        return value.props;
    return authored(value.props?.children, match);
}
function fixture(t, options = {}) {
    const prior = Object.getOwnPropertyDescriptor(window, 'ResizeObserver');
    Object.defineProperty(window, 'ResizeObserver', { configurable: true, value: class {
            observe() { }
            unobserve() { }
            disconnect() { }
        } });
    t.after(() => { if (prior)
        Object.defineProperty(window, 'ResizeObserver', prior);
    else
        delete window.ResizeObserver; });
    const f = createClientsFixture(t, title, { ...options, pnpContext: true, pnpCrud: true, ownerKey: owner, compiled: options.compiled ?? process.env.PNP_CRUD_IMPLEMENTATION === 'compiled' });
    f.props = predicate => { for (const region of Object.values(f.regions ?? {})) {
        const found = authored(region, predicate);
        if (found)
            return found;
    } };
    f.action = name => { const props = f.props(p => p.children === name && typeof p.onClick === 'function'); assert.ok(props, name); return props.onClick; };
    f.input = (label, value) => { const props = f.props(p => p['aria-label'] === label); assert.ok(props, label); f.h.act(() => props.onChange({}, { value })); };
    return f;
}
const json = body => new Response(JSON.stringify(body), { headers: { 'Content-Type': 'application/json' } });
async function drain(f) { await f.flush(); await new Promise(resolve => setTimeout(resolve, 20)); await f.flush(); }
async function finish(f, index, response) { await f.flush(() => f.calls[index].resolve(response)); await drain(f); }
const rows = [{ Id: 42, Title: 'Intended item' }, { Id: 73, Title: 'Other item' }];
test('CRUD checkpoint emitted Load reaches exact native title GET rather than a no-op', async (t) => {
    const f = fixture(t, { revision: 17 });
    await drain(f);
    assert.equal(f.calls.length, 0);
    f.input('List Title', "O'Brien tasks");
    f.click('Load');
    await drain(f);
    assert.equal(f.calls.length, 1, 'explicit Load must reach native terminal');
    const url = new URL(f.calls[0].url);
    assert.equal(url.pathname, "/sites/clients-one/_api/web/lists/getByTitle('O''Brien%20tasks')/items");
    assert.deepEqual([...url.searchParams].sort(), [['$orderby', 'Id desc'], ['$select', 'Id,Title'], ['$top', '10']]);
    await finish(f, 0, json({ value: rows }));
    assert.match(f.container.textContent, /Intended item/);
});
async function load(f, count = rows) { f.input('List Title', 'Tasks'); f.click('Load'); await drain(f); await finish(f, f.calls.length - 1, json({ value: count })); }
function clickRow(f, name, index = 0) { const node = [...f.container.querySelectorAll('button')].filter(n => n.textContent === name)[index]; assert.ok(node, name); f.h.act(() => Simulate.click(node)); }
function business(f) { return f.calls.filter(r => r.method === 'POST' && !r.url.endsWith('/contextinfo')); }
async function digest(f) {
    await drain(f);
    const index = f.calls.length - 1;
    assert.ok(f.calls[index].url.endsWith('/_api/contextinfo'));
    assert.equal(business(f).length, 0);
    assert.equal(new Headers(f.calls[index].init.headers).get('X-PnPjs-NoDigest'), '1');
    await finish(f, index, json({ FormDigestValue: 'fixture-pnp-crud-digest', FormDigestTimeoutSeconds: 1800 }));
}
test('CRUD checkpoint Create81 MERGE42 DELETE42/73 use native digest parsers and delayed actual refresh', async (t) => {
    const f = fixture(t);
    await load(f);
    const mode = process.env.PNP_CRUD_IMPLEMENTATION === 'compiled' ? 'lib' : 'src';
    assert.equal(f.toolkit.SPFxWebPartProvider, f.h.load(path.resolve('packages/spfx-react-toolkit', mode, 'core/provider-webpart.' + (mode === 'src' ? 'tsx' : 'js'))).SPFxWebPartProvider, 'actual matching provider in this module cache');
    for (const file of ['hooks/useSPFxPnPList', 'hooks/useSPFxPnPList.internal', 'services/spfx-pnp-list.service', 'hooks/useSPFxPnPContext', 'services/spfx-pnp-context.service'])
        assert.ok(f.pnpFiles.includes(path.resolve('packages/spfx-react-toolkit', mode, file + (mode === 'src' ? '.ts' : '.js'))), file);
    f.input('New item title', 'Created raw title');
    const pending = f.action('Create')();
    await digest(f);
    const create = f.calls.at(-1);
    assert.equal(create.method, 'POST');
    assert.equal(new URL(create.url).pathname, "/sites/clients-one/_api/web/lists/getByTitle('Tasks')/items");
    assert.equal(new Headers(create.init.headers).get('X-RequestDigest'), 'fixture-pnp-crud-digest');
    assert.deepEqual(JSON.parse(create.init.body), { Title: 'Created raw title' });
    const before = f.calls.length;
    // Observe settlement in its microtask, not after finish() yields to later wall-clock timers.
    const atCreateSettlement = pending.then(() => ({ calls: f.calls.length, text: f.container.textContent }));
    await finish(f, before - 1, json({ Id: 81 }));
    await pending;
    assert.match(f.container.textContent, /Created item #81/);
    assert.equal((await atCreateSettlement).calls, before, 'public promise settles before refresh dispatch');
    assert.equal((await atCreateSettlement).text.includes('Refreshing…'), false, 'mutation settlement does not start an immediate query');
    await new Promise(r => setTimeout(r, 110));
    await drain(f);
    assert.equal(f.calls.length, before + 1);
    await finish(f, before, json({ value: rows }));
    clickRow(f, 'Edit');
    f.input('Edit title', 'Saved raw title');
    const save = f.action('Save')();
    await drain(f);
    const update = f.calls.at(-1);
    assert.ok(update.url.endsWith('/items(42)'));
    assert.equal(update.method, 'POST');
    assert.deepEqual(JSON.parse(update.init.body), { Title: 'Saved raw title' });
    assert.equal(new Headers(update.init.headers).get('X-HTTP-Method'), 'MERGE');
    assert.equal(new Headers(update.init.headers).get('IF-Match'), '*');
    await finish(f, f.calls.length - 1, new Response(undefined, { status: 204, headers: { ETag: '"v2"' } }));
    await save;
    await new Promise(r => setTimeout(r, 110));
    await drain(f);
    await finish(f, f.calls.length - 1, json({ value: rows }));
    for (const [id, index] of [[42, 0], [73, 1]]) {
        clickRow(f, 'Delete', index);
        assert.ok(document.querySelector('[role="dialog"]'));
        const approve = f.action('Delete item');
        let removal;
        f.h.act(() => { removal = approve(); approve(); });
        await drain(f);
        const request = f.calls.at(-1);
        assert.ok(request.url.endsWith(`/items(${id})`));
        assert.equal(new Headers(request.init.headers).get('X-HTTP-Method'), 'DELETE');
        assert.equal(new Headers(request.init.headers).get('IF-Match'), '*');
        const n = f.calls.length;
        const atDeleteSettlement = removal.then(() => ({ calls: f.calls.length, text: f.container.textContent }));
        await finish(f, n - 1, new Response(undefined, { status: 204 }));
        await removal;
        assert.equal((await atDeleteSettlement).calls, n, 'deletion settles before refresh dispatch');
        assert.equal((await atDeleteSettlement).text.includes('Refreshing…'), false);
        assert.match(f.container.textContent, /Intended item/);
        await new Promise(r => setTimeout(r, 110));
        await drain(f);
        await finish(f, n, json({ value: rows }));
    }
    assert.equal(business(f).length, 4);
    assert.equal(f.calls.filter(r => r.url.endsWith('/contextinfo')).length, 1);
});
test('CRUD checkpoint presentation helper maps to matching actual factory and holds digest then business Response', async (t) => {
    const f = fixture(t);
    const descriptor = Object.getOwnPropertyDescriptor(window, 'fetch');
    const native = f.h.load(path.resolve('tests/fixtures/demo-presentation/trim-transports.ts'));
    t.after(() => { if (descriptor)
        Object.defineProperty(window, 'fetch', descriptor);
    else
        delete window.fetch; });
    t.mock.method(global, 'fetch', window.fetch.bind(window));
    const helper = f.h.load(path.resolve('tests/fixtures/demo-presentation/pnp-read-transport.ts'));
    const helperMode = process.env.PNP_CRUD_IMPLEMENTATION === 'compiled' ? 'lib' : 'src';
    assert.equal(f.pnpFactory.createSPFxPnPContextService, f.h.load(path.resolve('packages/spfx-react-toolkit', helperMode, 'services/spfx-pnp-context.service.' + (helperMode === 'src' ? 'ts' : 'js'))).createSPFxPnPContextService, 'helper dependency is the exact matching factory in this cache');
    const sp = helper.createSPFxPnPContextService(f.instance.context, f.instance.context.pageContext).createSPFI('https://simulated.invalid/sites/crud-helper');
    const mode = process.env.PNP_CRUD_IMPLEMENTATION === 'compiled' ? 'lib' : 'src';
    const service = f.h.load(path.resolve('packages/spfx-react-toolkit', mode, 'services/spfx-pnp-list.service.' + (mode === 'src' ? 'ts' : 'js'))).createSPFxPnPListService(sp, 'Tasks', 10);
    native.clientsNative.configure({ holdRequests: true, pnpCrud: true });
    let done = false;
    const pending = service.create({ Title: 'Helper create' }).then(id => { done = true; return id; });
    await drain(f);
    assert.equal(native.clientsNative.requests.length, 1);
    assert.equal(native.clientsNative.requests[0].kind, 'pnp-digest');
    assert.equal(done, false);
    native.clientsNative.requests[0].resolve(json({ FormDigestValue: 'helper-digest', FormDigestTimeoutSeconds: 1800 }));
    await drain(f);
    assert.equal(native.clientsNative.requests.length, 2);
    const request = native.clientsNative.requests[1];
    assert.equal(request.kind, 'pnp-item-write');
    assert.equal(done, false);
    assert.equal(new Headers(request.settings.headers).get('X-RequestDigest'), 'helper-digest');
    request.resolve(json({ Id: 81 }));
    assert.equal(await pending, 81);
    assert.equal(window.pnpReadFixture.calls.length, 0);
});
test('CRUD fixture admission defaults OFF, enables only finite native protocol and keeps top5 pilot independent', async (t) => {
    const f = fixture(t);
    const descriptor = Object.getOwnPropertyDescriptor(window, 'fetch');
    const native = f.h.load(path.resolve('tests/fixtures/demo-presentation/trim-transports.ts'));
    t.after(() => { if (descriptor)
        Object.defineProperty(window, 'fetch', descriptor);
    else
        delete window.fetch; });
    t.mock.method(global, 'fetch', window.fetch.bind(window));
    const base = "https://simulated.invalid/sites/finite/_api/web/lists/getByTitle('Tasks')/items";
    await assert.rejects(window.fetch(base, { method: 'POST', headers: { 'X-RequestDigest': 'known' }, body: '{"Title":"New"}' }), /Unsupported native PnP fixture route/);
    assert.equal(native.clientsNative.requests.length, 0);
    native.clientsNative.configure({ pnpCrud: true });
    for (const [url, init] of [[base, { method: 'PATCH' }], [base + '?extra=1', { method: 'POST', headers: { 'X-RequestDigest': 'known' }, body: '{"Title":"New"}' }], [base, { method: 'POST', body: '{"Title":"New"}' }], [base.replace('simulated.invalid', 'external.invalid'), { method: 'POST' }], [base + '?$select=Id,Title&$orderby=Id%20desc&$top=11', { method: 'GET' }], [base, { method: 'POST', headers: { 'X-RequestDigest': 'known' }, body: '{"Other":"New"}' }], [base + "/evil('x')/items", { method: 'GET' }], [base + '(42)', { method: 'POST', headers: { 'X-RequestDigest': 'known', 'X-HTTP-Method': 'DELETE', 'IF-Match': '*' }, body: '{}' }]])
        await assert.rejects(window.fetch(url, init), /Unsupported native PnP/);
    assert.equal(native.clientsNative.requests.length, 0);
    const helper = f.h.load(path.resolve('tests/fixtures/demo-presentation/pnp-read-transport.ts'));
    const sp = helper.createSPFxPnPContextService(f.instance.context, f.instance.context.pageContext).createSPFI('https://simulated.invalid/sites/finite');
    const mode = process.env.PNP_CRUD_IMPLEMENTATION === 'compiled' ? 'lib' : 'src';
    const service = f.h.load(path.resolve('packages/spfx-react-toolkit', mode, 'services/spfx-pnp-list.service.' + (mode === 'src' ? 'ts' : 'js'))).createSPFxPnPListService(sp, 'Pilot');
    await assert.rejects(f.h.load(require.resolve('@pnp/sp')).SPQueryable(sp.web, 'arbitrary')(), /Unsupported native PnP CRUD fixture route/);
    assert.deepEqual((await service.query(q => q.select('Id', 'Title').top(5))).items, [{ Id: 1, Title: 'Simulated list row (no tenant request)' }]);
    assert.equal(window.pnpReadFixture.calls.length, 1);
    assert.equal(native.clientsNative.requests.length, 0);
    native.clientsNative.configure({ pnpCrud: false });
    await assert.rejects(sp.web.lists.getByTitle('Pilot').items.add({ Title: 'Forbidden' }), /Simulated read fixture forbids CRUD/);
    assert.equal(window.pnpReadFixture.calls.at(-1).method, 'POST');
    assert.equal(native.clientsNative.requests.filter(r => r.kind === 'pnp-item-write').length, 0);
});
test('CRUD retained approval cannot consume a later consent and retained target input cannot bypass a dispatched write lock', async (t) => {
    const f = fixture(t);
    await load(f);
    clickRow(f, 'Delete');
    const oldApprove = f.action('Delete item');
    f.h.act(() => f.action('Cancel')());
    clickRow(f, 'Delete', 1);
    await f.flush(() => oldApprove());
    await drain(f);
    assert.equal(business(f).length, 0, 'retired approval cannot dispatch later consent');
    f.h.act(() => f.action('Cancel')());
    const oldInput = f.props(p => p['aria-label'] === 'List Title').onChange;
    f.input('New item title', 'Held create');
    const promise = f.action('Create')();
    await digest(f);
    await f.flush(() => oldInput({}, { value: 'Forbidden retarget' }));
    assert.equal(f.container.querySelector('[aria-label="List Title"]').value, 'Tasks');
    await finish(f, f.calls.length - 1, json({ Id: 81 }));
    await promise;
});
function replaceClient(f) { const next = f.host(f.instance.instanceId); next.context.serviceScope = f.instance.context.serviceScope; f.replace(next); }
async function startCreate(f) { f.input('New item title', 'Original draft'); let promise; f.h.act(() => { promise = f.action('Create')(); }); await digest(f); return { promise, index: f.calls.length - 1 }; }
for (const change of ['reset', 'client'])
    for (const outcome of ['success', 'failure'])
        test(`CRUD ${change} retains original physical promise while hiding old ${outcome} and preserving new drafts`, async (t) => {
            const f = fixture(t);
            await load(f);
            const oldCreate = f.action('Create');
            const oldInput = f.props(p => p['aria-label'] === 'List Title').onChange;
            const { promise, index } = await startCreate(f);
            if (change === 'reset') {
                const element = f.container.querySelector('[aria-label="List Title"]');
                f.reset();
                assert.equal(f.container.querySelector('[aria-label="List Title"]') === element, true, 'Reset retains the mounted input');
                assert.equal(element.value, '');
                f.input('List Title', 'Tasks');
            }
            else
                replaceClient(f);
            f.input('New item title', 'New session draft');
            const count = f.calls.length;
            await f.flush(() => { oldCreate(); oldInput({}, { value: 'Obsolete target' }); f.action('Create')(); f.action('Load')(); });
            await drain(f);
            assert.equal(f.calls.length, count);
            assert.equal(f.container.querySelector('[aria-label="List Title"]').value, 'Tasks');
            assert.ok(!f.container.textContent.includes('Intended item'));
            // A native settlement is not permission to dispatch in the same tick before the original hook promise settles.
            await f.flush(() => { f.calls[index].resolve(outcome === 'success' ? json({ Id: 81 }) : new Response('Obsolete denial', { status: 400 })); f.action('Create')(); });
            await promise;
            await drain(f);
            assert.equal(f.calls.length, count);
            assert.equal(f.container.querySelector('[aria-label="New item title"]').value, 'New session draft');
            assert.ok(!f.container.textContent.includes('Obsolete denial'));
            assert.ok(!f.container.textContent.includes('Created item #81'));
            await new Promise(r => setTimeout(r, 110));
            await drain(f);
            assert.equal(f.calls.length, count, 'old service cannot refresh after committed blank/client replacement');
            let next;
            f.h.act(() => { next = f.action('Create')(); });
            await drain(f);
            assert.equal(f.calls.length, count + 1);
            assert.deepEqual(JSON.parse(f.calls.at(-1).init.body), { Title: 'New session draft' });
            await finish(f, count, json({ Id: 82 }));
            await next;
        });
test('CRUD scheduled refresh Reset commits blank and same-title reentry waits for fresh explicit Load', async (t) => {
    const f = fixture(t);
    await load(f);
    const { promise, index } = await startCreate(f);
    await finish(f, index, json({ Id: 81 }));
    await promise;
    const n = f.calls.length;
    f.reset();
    assert.equal(f.container.querySelector('[aria-label="List Title"]').value, '');
    f.input('List Title', 'Tasks');
    await new Promise(r => setTimeout(r, 120));
    await drain(f);
    assert.equal(f.calls.length, n);
    assert.ok(!f.container.textContent.includes('Intended item'));
    f.click('Load');
    await drain(f);
    assert.equal(f.calls.length, n + 1);
    await finish(f, n, json({ value: [] }));
    assert.match(f.container.textContent, /No items found/);
});
test('CRUD paging keeps builder and offset across failed next page then reaches an empty terminal page', async (t) => {
    const f = fixture(t);
    const ten = Array.from({ length: 10 }, (_, i) => ({ Id: 42 - i, Title: 'First ' + i }));
    await load(f, ten);
    const page = f.action('Load More');
    await f.flush(() => { page(); page(); page(); });
    await drain(f);
    assert.equal(f.calls.length, 2);
    let u = new URL(f.calls[1].url);
    assert.equal(u.searchParams.get('$skiptoken'), 'Paged=TRUE&p_ID=10');
    assert.equal(u.searchParams.get('$top'), '10');
    assert.equal(u.searchParams.get('$orderby'), 'Id desc');
    await finish(f, 1, new Response('Page denied', { status: 403 }));
    assert.match(f.container.textContent, /Page denied/);
    f.click('Dismiss');
    f.click('Load More');
    await drain(f);
    assert.equal(new URL(f.calls[2].url).searchParams.get('$skiptoken'), 'Paged=TRUE&p_ID=10');
    await finish(f, 2, json({ value: ten.map((r, i) => ({ Id: 22 - i, Title: 'Next ' + i })) }));
    assert.match(f.container.textContent, /Items \(20\)/);
    const n = f.calls.length;
    await f.flush(() => page());
    await drain(f);
    assert.equal(f.calls.length, n, 'retained old page callback cannot replay an old offset');
    f.click('Load More');
    await drain(f);
    assert.equal(new URL(f.calls.at(-1).url).searchParams.get('$skiptoken'), 'Paged=TRUE&p_ID=20');
    await finish(f, n, json({ value: [] }));
    assert.equal([...f.container.querySelectorAll('button')].some(n => n.textContent === 'Load More'), false);
});
for (const kind of ['query', 'page'])
    for (const change of ['reset', 'client', 'departure', 'unmount'])
        test(`CRUD pending ${kind} ${change} consumes stale failure without publication`, async (t) => {
            const f = fixture(t);
            if (kind === 'page')
                await load(f, Array.from({ length: 10 }, (_, i) => ({ Id: 42 - i, Title: 'Row ' + i })));
            else
                f.input('List Title', 'Tasks');
            const old = f.action(kind === 'page' ? 'Load More' : 'Load');
            let promise;
            f.h.act(() => { promise = old(); });
            await drain(f);
            const index = f.calls.length - 1;
            if (change === 'reset')
                f.reset();
            if (change === 'client')
                replaceClient(f);
            if (change === 'departure')
                f.select('Inspect the current native host identity', 'packages/spfx-react-toolkit/src/core/context.internal.tsx#useSPFxContext');
            if (change === 'unmount')
                f.unmount();
            const n = f.calls.length;
            await f.flush(() => old());
            assert.equal(f.calls.length, n);
            await finish(f, index, new Response('Stale read error', { status: 403 }));
            await promise;
            assert.ok(!f.container.textContent.includes('Stale read error'));
        });
for (const change of ['reset', 'client', 'departure', 'unmount'])
    test(`CRUD unused consent ${change} invalidates retained approval with zero writes`, async (t) => {
        const f = fixture(t);
        await load(f);
        clickRow(f, 'Delete');
        const approve = f.action('Delete item');
        if (change === 'reset')
            f.reset();
        if (change === 'client')
            replaceClient(f);
        if (change === 'departure')
            f.select('Inspect the current native host identity', 'packages/spfx-react-toolkit/src/core/context.internal.tsx#useSPFxContext');
        if (change === 'unmount')
            f.unmount();
        await f.flush(() => { approve(); approve(); });
        await drain(f);
        assert.equal(business(f).length, 0);
        assert.equal(f.calls.length, 1);
    });
test('CRUD Edit Cancel Save lock conflicting page/draft/retained commands and consume failures without clearing newer values', async (t) => {
    const f = fixture(t);
    await load(f);
    clickRow(f, 'Edit');
    f.input('Edit title', 'Cancelled draft');
    const oldSave = f.action('Save');
    f.h.act(() => f.action('Cancel')());
    await f.flush(() => oldSave());
    assert.equal(business(f).length, 0);
    clickRow(f, 'Edit');
    f.input('Edit title', 'Actual save');
    let save;
    f.h.act(() => { const action = f.action('Save'); save = action(); action(); });
    await digest(f);
    assert.equal(business(f).length, 1);
    assert.equal(f.container.querySelector('[aria-label="Edit title"]').readOnly, true);
    assert.equal([...f.container.querySelectorAll('button')].find(n => n.textContent === 'Cancel').disabled, true);
    await finish(f, f.calls.length - 1, new Response('Save denied', { status: 400 }));
    await save;
    assert.match(f.container.textContent, /Save denied/);
    assert.equal(f.container.querySelector('[aria-label="Edit title"]').value, 'Actual save');
    f.click('Dismiss');
    assert.ok(!f.container.textContent.includes('Save denied'));
});
test('CRUD successful mutation remains success when later real refresh fails', async (t) => {
    const f = fixture(t);
    await load(f);
    const { promise, index } = await startCreate(f);
    await finish(f, index, json({ Id: 81 }));
    await promise;
    await new Promise(r => setTimeout(r, 110));
    await drain(f);
    await finish(f, f.calls.length - 1, new Response('Refresh denied', { status: 403 }));
    assert.match(f.container.textContent, /Created item #81/);
    assert.match(f.container.textContent, /Refresh denied/);
    assert.match(f.container.textContent, /Intended item/);
});
function negativeFixture(t, transform) { const fs = require('node:fs'), os = require('node:os'); const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs'); const selected = scenario(title, owner), read = fs.readFileSync; const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pnp-crud-negative-')), copy = path.join(dir, 'scenario.tsx'); const altered = transform(selected.source); assert.notEqual(altered, selected.source); fs.writeFileSync(copy, altered); t.after(() => fs.rmSync(dir, { recursive: true, force: true })); const mock = t.mock.method(fs, 'readFileSync', (file, ...args) => read(String(file) === selected.file ? copy : file, ...args)); try {
    return fixture(t);
}
finally {
    mock.mock.restore();
} }
test('CRUD negative control detects missing physical admission lock before original promise settles', async (t) => {
    const f = negativeFixture(t, s => s.replace('physical.current !== undefined ||', ''));
    await load(f);
    f.input('New item title', 'Duplicate');
    const create = f.action('Create');
    await f.flush(() => { create(); create(); });
    await drain(f);
    assert.throws(() => assert.equal(f.calls.filter(r => r.url.endsWith('/contextinfo')).length, 1, 'one cold dispatch'), /one cold dispatch/);
});
test('CRUD negative control detects stale retained Load admission after Reset', async (t) => {
    const f = negativeFixture(t, s => s.replace('mounted.current && boundary.current.token === token', 'mounted.current'));
    await load(f);
    const read = f.action('Load');
    f.reset();
    await f.flush(() => read());
    await drain(f);
    assert.throws(() => assert.equal(f.calls.length, 1, 'retained load stays retired'), /retained load stays retired/);
});
test('CRUD two actual providers and standalone fallback keep requests and outcomes isolated', async (t) => {
    const f = fixture(t, { directProvider: true, fallback: true });
    const peer = f.mountPeer(f.host('crud-peer'));
    t.after(peer.unmount);
    const field = peer.container.querySelector('[aria-label="List Title"]');
    f.h.act(() => { field.value = 'Peer'; Simulate.change(field, { target: { value: 'Peer' } }); });
    f.change('List Title', 'Primary');
    f.click('Load');
    peer.click('Load');
    await drain(f);
    assert.equal(f.calls.length, 2);
    assert.ok(f.calls.some(r => r.url.includes('/sites/crud-peer/')));
    await finish(f, 1, json({ value: [{ Id: 73, Title: 'Peer only' }] }));
    await finish(f, 0, json({ value: [{ Id: 42, Title: 'Primary only' }] }));
    assert.match(peer.container.textContent, /Peer only/);
    assert.ok(!peer.container.textContent.includes('Primary only'));
    assert.match(f.container.textContent, /Primary only/);
    assert.ok(!f.container.textContent.includes('Peer only'));
});
test('CRUD actual Explorer reset preserves input identity and departure invalidates pending explicit read', async (t) => {
    const f = fixture(t, { explorer: true });
    const feature = [...f.container.querySelectorAll('[data-example-feature]')].find(n => n.textContent === 'useSPFxPnPList');
    f.h.act(() => Simulate.click(feature));
    const select = f.container.querySelector('[aria-label="Scenario"]');
    assert.ok(select);
    // Explorer defaults to the preserved read pilot; select the independently named CRUD case.
    const option = [...select.options].find(n => n.textContent === title);
    assert.ok(option);
    f.h.act(() => Simulate.change(select, { target: { value: option.value } }));
    await drain(f);
    assert.equal(f.calls.length, 0);
    f.change('List Title', 'Tasks');
    f.click('Load');
    await drain(f);
    const input = f.container.querySelector('[aria-label="List Title"]');
    f.click('Ripristina scenario');
    assert.equal(f.container.querySelector('[aria-label="List Title"]') === input, true, 'Explorer Reset retains the mounted input');
    assert.equal(input.value, '');
    const context = [...f.container.querySelectorAll('[data-example-feature]')].find(n => n.textContent === 'useSPFxContext');
    f.h.act(() => Simulate.click(context));
    await finish(f, 0, new Response('Departed read', { status: 403 }));
    assert.ok(!f.container.textContent.includes('Departed read'));
});
for (const change of ['departure', 'unmount'])
    for (const outcome of ['success', 'failure'])
        test(`CRUD dispatched mutation ${change} settles without publishing obsolete ${outcome}`, async (t) => {
            const f = fixture(t);
            await load(f);
            const { promise, index } = await startCreate(f);
            const create = f.action('Create');
            if (change === 'departure')
                f.select('Inspect the current native host identity', 'packages/spfx-react-toolkit/src/core/context.internal.tsx#useSPFxContext');
            else
                f.unmount();
            const n = f.calls.length;
            await f.flush(() => create());
            assert.equal(f.calls.length, n);
            await finish(f, index, outcome === 'success' ? json({ Id: 81 }) : new Response('Old write denial', { status: 400 }));
            await promise;
            await new Promise(r => setTimeout(r, 110));
            await drain(f);
            assert.equal(f.calls.length, n);
            assert.ok(!f.container.textContent.includes('Created item #81'));
            assert.ok(!f.container.textContent.includes('Old write denial'));
        });
test('CRUD retained Dismiss cannot clear a newer target error and canceled old consent cannot retarget', async (t) => {
    const f = fixture(t);
    await load(f);
    clickRow(f, 'Delete');
    const approve = f.action('Delete item');
    f.h.act(() => f.action('Cancel')());
    f.input('List Title', 'Other');
    await f.flush(() => approve());
    assert.equal(business(f).length, 0);
    f.click('Load');
    await drain(f);
    await finish(f, f.calls.length - 1, new Response('First denial', { status: 403 }));
    const dismiss = f.action('Dismiss');
    f.reset();
    f.input('List Title', 'Current');
    f.click('Load');
    await drain(f);
    await finish(f, f.calls.length - 1, new Response('Current denial', { status: 403 }));
    await f.flush(() => dismiss());
    assert.match(f.container.textContent, /Current denial/);
    f.click('Dismiss');
    assert.ok(!f.container.textContent.includes('Current denial'));
});
test('CRUD approval consumes physical admission before focus can synchronously activate a conflicting create', async (t) => {
    const f = fixture(t);
    await load(f);
    f.input('New item title', 'Reentrant create');
    const create = f.action('Create');
    clickRow(f, 'Delete');
    f.container.querySelector('[aria-label="List Title"]').addEventListener('focus', () => { create(); }, { once: true });
    let remove;
    f.h.act(() => { remove = f.action('Delete item')(); });
    await digest(f);
    assert.equal(new Headers(f.calls.at(-1).init.headers).get('X-HTTP-Method'), 'DELETE', 'only captured deletion dispatches after pre-dispatch focus');
    assert.ok(f.calls.at(-1).url.endsWith('/items(42)'));
    await finish(f, f.calls.length - 1, new Response(undefined, { status: 204 }));
    await remove;
});
test('CRUD focus-triggered Reset before native dispatch invalidates consumed approval without a cleanup write', async (t) => {
    const f = fixture(t);
    await load(f);
    clickRow(f, 'Delete');
    f.container.querySelector('[aria-label="List Title"]').addEventListener('focus', () => {
        // A queued host update is not a committed revision; force the ownership change before dispatch.
        flushSync(() => f.reset());
    }, { once: true });
    await f.flush(() => f.action('Delete item')());
    await drain(f);
    assert.equal(f.calls.length, 1);
    assert.equal(business(f).length, 0);
    assert.equal(f.container.querySelector('[aria-label="List Title"]').value, '');
});
test('CRUD same-value creation draft keeps captured Create valid and creation before Load schedules no query', async t => {
    const f = fixture(t); f.input('List Title', 'Tasks'); f.input('New item title', 'Unchanged');
    const create = f.action('Create'); f.input('New item title', 'Unchanged');
    let pending; f.h.act(() => { pending = create(); }); await drain(f);
    assert.equal(f.calls.length, 1, 'same-value draft must not retire valid Create');
    await digest(f); await finish(f, f.calls.length - 1, json({ Id: 81 })); await pending;
    await new Promise(resolve => setTimeout(resolve, 110)); await drain(f);
    assert.equal(f.calls.length, 2, 'no active query exists before explicit Load');
    assert.match(f.container.textContent, /Created item #81/);
});
test('CRUD same-value edit keeps captured Save valid', async t => {
    const f = fixture(t); await load(f); clickRow(f, 'Edit');
    const save = f.action('Save'); f.input('Edit title', 'Intended item');
    let pending; f.h.act(() => { pending = save(); }); await drain(f);
    assert.equal(f.calls.length, 2, 'same-value edit must not retire valid Save');
    await digest(f); assert.deepEqual(JSON.parse(f.calls.at(-1).init.body), { Title: 'Intended item' });
    await finish(f, f.calls.length - 1, new Response(undefined, { status: 204 })); await pending;
});
for (const legacyTitle of [undefined, null, '']) test('CRUD legacy wire title ' + String(legacyTitle) + ' displays untitled and opens an empty edit draft', async t => {
    const f = fixture(t);
    // Legacy SharePoint JSON may contain null; no new source state uses null.
    await load(f, [{ Id: 42, Title: legacyTitle }]);
    assert.match(f.container.textContent, /\(untitled\)/);
    assert.doesNotThrow(() => clickRow(f, 'Edit'));
    assert.equal(f.container.querySelector('[aria-label="Edit title"]').value, '');
    assert.equal([...f.container.querySelectorAll('button')].find(node => node.textContent === 'Save').disabled, true);
    assert.equal(f.calls.length, 1);
});
test('CRUD target no-op retains Load while accepted ABA never revives old target callbacks', async t => {
    const f = fixture(t); f.input('List Title', 'Tasks'); const oldLoad = f.action('Load');
    f.input('List Title', 'Tasks'); let pending; f.h.act(() => { pending = oldLoad(); }); await drain(f);
    assert.equal(f.calls.length, 1); await finish(f, 0, json({ value: rows })); await pending;
    const input = f.props(props => props['aria-label'] === 'List Title').onChange;
    f.h.act(() => { input({}, { value: 'Other' }); oldLoad(); input({}, { value: 'Tasks' }); });
    assert.equal(f.container.querySelector('[aria-label="List Title"]').value, 'Other');
    f.input('List Title', 'Tasks'); await f.flush(() => oldLoad()); await drain(f); assert.equal(f.calls.length, 1);
    f.click('Load'); await drain(f); assert.equal(f.calls.length, 2); await finish(f, 1, json({ value: [] }));
});
test('CRUD accepted creation ABA retires old commands, commits fresh handlers, and old input cannot overwrite success baseline', async t => {
    const f = fixture(t); f.input('List Title', 'Tasks'); f.input('New item title', 'A');
    const oldCreate = f.action('Create'), input = f.props(props => props['aria-label'] === 'New item title').onChange;
    f.h.act(() => { input({}, { value: 'B' }); input({}, { value: 'A' }); oldCreate(); });
    await drain(f); assert.equal(f.calls.length, 0); assert.equal(f.container.querySelector('[aria-label="New item title"]').value, 'A');
    f.h.act(() => input({}, { value: 'Stale rendered input' })); assert.equal(f.container.querySelector('[aria-label="New item title"]').value, 'A');
    f.click('Create'); await digest(f);
    assert.deepEqual(JSON.parse(f.calls.at(-1).init.body), { Title: 'A' }); await finish(f, f.calls.length - 1, json({ Id: 81 }));
    f.h.act(() => input({}, { value: 'Obsolete' })); assert.equal(f.container.querySelector('[aria-label="New item title"]').value, '');
});
test('CRUD accepted edit ABA commits fresh Save and a retired edit input cannot overwrite Cancel/reentry baseline', async t => {
    const f = fixture(t); await load(f); clickRow(f, 'Edit');
    const oldSave = f.action('Save'), input = f.props(props => props['aria-label'] === 'Edit title').onChange;
    f.h.act(() => { input({}, { value: 'B' }); input({}, { value: 'Intended item' }); oldSave(); });
    await drain(f); assert.equal(f.calls.length, 1); assert.equal(f.container.querySelector('[aria-label="Edit title"]').value, 'Intended item');
    f.h.act(() => input({}, { value: 'Stale rendered input' })); assert.equal(f.container.querySelector('[aria-label="Edit title"]').value, 'Intended item');
    f.h.act(() => f.action('Cancel')()); clickRow(f, 'Edit'); f.h.act(() => input({}, { value: 'Obsolete' }));
    assert.equal(f.container.querySelector('[aria-label="Edit title"]').value, 'Intended item');
    f.click('Save'); await digest(f);
    assert.deepEqual(JSON.parse(f.calls.at(-1).init.body), { Title: 'Intended item' }); await finish(f, f.calls.length - 1, new Response(undefined, { status: 204 }));
    assert.match(f.container.textContent, /Updated item #42/);
});
test('CRUD actual construction fault invalidates unused consent, stays separate from request errors and recovers', async t => {
    const f = fixture(t); await load(f); clickRow(f, 'Delete'); const approve = f.action('Delete item');
    f.setPnPInitializationError(new Error('Controlled CRUD construction fault')); replaceClient(f);
    await f.flush(() => approve()); await drain(f); assert.equal(f.calls.length, 1); assert.equal(business(f).length, 0);
    assert.match(f.container.textContent, /Context initialization error: Controlled CRUD construction fault/);
    f.reset(); assert.match(f.container.textContent, /Context initialization error/);
    f.setPnPInitializationError(undefined); replaceClient(f); f.input('List Title', 'Tasks'); f.click('Load'); await drain(f);
    await finish(f, 1, json({ value: rows })); assert.match(f.container.textContent, /Intended item/);
});
test('CRUD approval during the separate background refresh retains usable consent until read settlement', async t => {
    const f = fixture(t); await load(f); const { promise, index } = await startCreate(f);
    await finish(f, index, json({ Id: 81 })); await promise; clickRow(f, 'Delete');
    await new Promise(resolve => setTimeout(resolve, 110)); await drain(f); const refresh = f.calls.length - 1;
    assert.equal(f.calls[refresh].method, 'GET'); const approve = f.action('Delete item');
    await f.flush(() => approve()); assert.equal(business(f).length, 1, 'pending read blocks a second write');
    await finish(f, refresh, json({ value: rows })); let deletion; f.h.act(() => { deletion = approve(); }); await drain(f);
    assert.equal(business(f).length, 2, 'blocked approval must retain its captured consent');
    assert.equal(new Headers(f.calls.at(-1).init.headers).get('X-HTTP-Method'), 'DELETE');
    await finish(f, f.calls.length - 1, new Response(undefined, { status: 204 })); await deletion;
});
for (const action of ['Cancel', 'Escape']) for (const origin of ['disabled', 'removed']) test(`CRUD ${action} during background refresh restores stable focus with ${origin} origin`, async t => {
    const f = fixture(t); await load(f); const { promise, index } = await startCreate(f);
    await finish(f, index, json({ Id: 81 })); await promise; clickRow(f, 'Delete');
    await new Promise(resolve => setTimeout(resolve, 110)); await drain(f); const refresh = f.calls.length - 1;
    await f.flush(() => f.action('Delete item')());
    if (origin === 'removed') await finish(f, refresh, json({ value: [rows[1]] }));
    const dialog = document.querySelector('[role="dialog"]'); assert.ok(dialog);
    f.h.act(() => {
        if (action === 'Escape') Simulate.keyDown(dialog, { key: 'Escape' });
        else Simulate.click([...dialog.querySelectorAll('button')].find(node => node.textContent === 'Cancel'));
    });
    assert.equal(document.querySelector('[role="dialog"]') === null, true, 'Cancel/Escape remain usable after blocked approval');
    assert.equal(document.activeElement === f.container.querySelector('[aria-label="List Title"]'), true);
    if (origin === 'disabled') await finish(f, refresh, json({ value: rows }));
    assert.equal(business(f).length, 1);
    assert.equal(document.activeElement === f.container.querySelector('[aria-label="List Title"]'), true, 'later read publication leaves focus alone');
});

const assert = require('node:assert/strict');
const { test } = require('node:test');
const path = require('node:path');
const { createClientsFixture } = require('./fixtures/executable-examples/clients-scenarios.cjs');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const cases = {
    context: ['spfx-pnp-context.service', 'createSPFxPnPContextService', 'Explore PnP context caching and request headers'],
    generic: ['spfx-pnp.service', 'createSPFxPnPService', 'Read through a supplied PnP service client'],
    list: ['spfx-pnp-list.service', 'createSPFxPnPListService', 'Read and page an explicitly selected list service'],
};
const compiled = process.env.PNP_SERVICES_IMPLEMENTATION === 'compiled';
function fixture(t, kind = 'context', options = {}) {
    const previous = Object.getOwnPropertyDescriptor(window, 'ResizeObserver');
    Object.defineProperty(window, 'ResizeObserver', { configurable: true, value: class {
            observe() { }
            unobserve() { }
            disconnect() { }
        } });
    t.after(() => { if (previous)
        Object.defineProperty(window, 'ResizeObserver', previous);
    else
        delete window.ResizeObserver; });
    const [file, name, title] = cases[kind];
    return createClientsFixture(t, title, { ...options, pnpContext: true, pnpServices: true, ownerKey: `packages/spfx-react-toolkit/src/services/${file}.ts#${name}`, compiled });
}
async function drain(f) { await f.flush(); await new Promise(resolve => setTimeout(resolve, 25)); await f.flush(); }
const json = value => new Response(JSON.stringify(value), { headers: { 'Content-Type': 'application/json' } });
async function finish(f, index, response) { await f.flush(() => f.calls[index].resolve(response)); await drain(f); }
function props(tree, predicate) {
    if (!tree || typeof tree !== 'object')
        return undefined;
    if (predicate(tree.props ?? {}))
        return tree.props;
    for (const child of Array.isArray(tree) ? tree : [tree.props?.children]) {
        const found = props(child, predicate);
        if (found)
            return found;
    }
}
function action(f, name) {
    for (const region of Object.values(f.regions)) {
        const p = props(region, p => p.children === name && typeof p.onClick === 'function');
        if (p)
            return p.onClick;
    }
    assert.fail('Missing source command ' + name);
}
function input(f, label, value) { const p = props(f.regions.controls, p => p['aria-label'] === label); assert.ok(p, label); f.h.act(() => p.onChange({}, { value, target: { value } })); }
// This fails on a bounded native route before source examples are added, rather than a missing owner.
test('native read-only service mode admits exact title page GET without enabling CRUD', async (t) => {
    const f = createClientsFixture(t, 'Read current or cross-site web information', { pnpContext: true, compiled, ownerKey: 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPContext.ts#useSPFxPnPContext' });
    const descriptor = Object.getOwnPropertyDescriptor(window, 'fetch');
    const native = f.h.load(path.resolve('tests/fixtures/demo-presentation/trim-transports.ts'));
    t.after(() => { if (descriptor)
        Object.defineProperty(window, 'fetch', descriptor);
    else
        delete window.fetch; });
    native.clientsNative.configure({ pnpServices: true });
    const url = "https://simulated.invalid/sites/read/_api/web/lists/getByTitle('Tasks')/items?$select=Id,Title&$orderby=Id%20desc&$top=10&$skiptoken=Paged%3DTRUE%26p_ID%3D10";
    const response = await window.fetch(url);
    assert.ok(response instanceof Response);
    assert.equal(native.clientsNative.requests.at(-1).kind, 'pnp-service-items');
    await assert.rejects(window.fetch(url.split('?')[0], { method: 'POST', headers: { 'X-RequestDigest': 'known' }, body: '{"Title":"Forbidden"}' }), /Unsupported native PnP/);
    assert.equal(native.isPnPCrudEnabled(), false);
});
for (const [file, name, title] of Object.values(cases))
    test('service availability: ' + name, () => {
        const selected = scenario(title, `packages/spfx-react-toolkit/src/services/${file}.ts#${name}`);
        assert.equal(selected.entry.metadata.kind, 'read');
    });
for (const kind of Object.keys(cases))
    test('checkpoint actual emitted ' + kind + ' uses matching provider/service origin with zero construction I/O', async (t) => {
        const f = fixture(t, kind, { revision: 17 });
        await drain(f);
        assert.equal(f.calls.length + f.initializations.length + f.tokens.length + f.downloads.length, 0);
        const mode = compiled ? 'lib' : 'src', other = compiled ? 'src' : 'lib';
        for (const stem of ['spfx-pnp-context.service', ...(kind === 'context' ? [] : [cases[kind][0]])]) {
            assert.ok(f.pnpFiles.includes(path.resolve('packages/spfx-react-toolkit', mode, 'services/' + stem + (compiled ? '.js' : '.ts'))), stem);
            assert.ok(!f.pnpFiles.some(file => file.includes('/' + other + '/services/' + stem + '.')));
        }
        assert.equal(f.toolkit.createSPFxPnPContextService, f.pnpFactory.createSPFxPnPContextService, 'public subject and helper witness use same actual module-cache factory');
        f.reset();
        await drain(f);
        assert.equal(f.calls.length, 0);
    });
function ownedStorage(t) {
    const records = new Map(), restorations = [];
    const store = { get length() { return records.size; }, key(index) { return [...records.keys()][index] ?? null; }, getItem(key) { return records.get(key) ?? null; }, setItem(key, value) { records.set(String(key), String(value)); }, removeItem(key) { records.delete(key); }, clear() { records.clear(); } };
    for (const target of [global, window])
        for (const name of ['sessionStorage', 'localStorage']) {
            const previous = Object.getOwnPropertyDescriptor(target, name);
            const value = name === 'sessionStorage' ? store : { ...store, getItem: () => null, setItem: () => assert.fail('session caching must not write local storage'), clear() { } };
            Object.defineProperty(target, name, { configurable: true, value });
            restorations.push(() => { if (previous)
                Object.defineProperty(target, name, previous);
            else
                delete target[name]; });
        }
    return { records, restore() { restorations.reverse().forEach(fn => fn()); } };
}
test('checkpoint context reports header and installed session cache replays59999 then expires60000', async (t) => {
    const storage = ownedStorage(t), records = storage.records;
    t.mock.timers.enable({ apis: ['Date'], now: 1700000000000 });
    const f = fixture(t);
    t.after(storage.restore);
    await drain(f);
    f.click('Read web');
    await drain(f);
    assert.equal(f.calls.length, 1);
    assert.equal(f.calls[0].url, 'https://tenant.test/sites/projects/_api/web?%24select=Title%2CUrl%2CDescription');
    assert.equal(new Headers(f.calls[0].init.headers).get('X-Application'), 'reports');
    await finish(f, 0, json({ Title: 'First cached web', Url: 'https://tenant.test/sites/projects', Description: 'Stored' }));
    assert.match(f.container.textContent, /First cached web/);
    assert.match(f.container.textContent, /serialization only.*\{"cache":\{"enabled":false\}\}/);
    const entries = [...records].filter(([key]) => key.startsWith('pnp-cache-'));
    assert.equal(entries.length, 1);
    assert.equal(new Date(JSON.parse(entries[0][1]).expiration).getTime(), 1700000060000);
    t.mock.timers.tick(59999);
    f.click('Read web');
    await drain(f);
    assert.equal(f.calls.length, 1);
    assert.match(f.container.textContent, /First cached web/);
    t.mock.timers.tick(1);
    f.click('Read web');
    await drain(f);
    assert.equal(f.calls.length, 2);
    await finish(f, 1, json({ Title: 'Expired replacement', Url: 'https://tenant.test/sites/projects' }));
    assert.match(f.container.textContent, /Expired replacement/);
    assert.match(f.container.textContent, /\(no description\)/);
});
function batchResponse() {
    return new Response(['--batchresponse_services', 'Content-Type: application/http', 'Content-Transfer-Encoding: binary', '', 'HTTP/1.1 200 OK', 'Content-Type: application/json', '', '{"Title":"Batched service title"}', '--batchresponse_services', 'Content-Type: application/http', 'Content-Transfer-Encoding: binary', '', 'HTTP/1.1 200 OK', 'Content-Type: application/json', '', '{"Url":"https://tenant.test/batched-result"}', '--batchresponse_services--', ''].join('\r\n'), { headers: { 'Content-Type': 'multipart/mixed; boundary=batchresponse_services' } });
}
test('checkpoint supplied client invoke and exact two synchronously queued batch GETs reach installed parser', async (t) => {
    const f = fixture(t, 'generic');
    assert.match(f.container.textContent, /Exact supplied client retained: true/);
    f.click('Invoke web read');
    await drain(f);
    await finish(f, 0, json({ Title: 'Invoke service title', Url: 'https://tenant.test/invoke' }));
    assert.match(f.container.textContent, /Invoke service title/);
    f.click('Batch web reads');
    await drain(f);
    let index = 1;
    if (f.calls[index].url.endsWith('/_api/contextinfo')) {
        assert.equal(new Headers(f.calls[index].init.headers).get('X-PnPjs-NoDigest'), '1');
        await finish(f, index++, json({ FormDigestValue: 'service-digest', FormDigestTimeoutSeconds: 1800 }));
    }
    const request = f.calls[index];
    assert.equal(request.method, 'POST');
    assert.ok(new Headers(request.init.headers).get('X-RequestDigest'));
    assert.deepEqual([...request.init.body.matchAll(/^GET (.+) HTTP\/1\.1$/gm)].map(m => m[1]), ['https://tenant.test/sites/clients-one/_api/web?%24select=Title', 'https://tenant.test/sites/clients-one/_api/web?%24select=Url']);
    assert.equal(request.init.body.includes('changeset'), false);
    await finish(f, index, batchResponse());
    assert.match(f.container.textContent, /Batched service title/);
    assert.match(f.container.textContent, /https:\/\/tenant.test\/batched-result/);
});
test('checkpoint list captured service/builder/page size and count offset produce10 then2 rows', async (t) => {
    const f = fixture(t, 'list');
    input(f, 'List title', 'Tasks');
    f.click('Run');
    await drain(f);
    assert.equal(f.calls.length, 1);
    assert.match(f.calls[0].url, /getByTitle\('Tasks'\)\/items/);
    const ten = Array.from({ length: 10 }, (_, i) => ({ Id: i + 1, Title: 'Row ' + (i + 1) }));
    await finish(f, 0, json({ value: ten }));
    assert.match(f.container.textContent, /Items \(10\)/);
    assert.deepEqual([...f.container.querySelectorAll('dd')].map(n => n.textContent), ['10', '10', 'true', '10']);
    f.click('Next page');
    await drain(f);
    const params = new URL(f.calls[1].url).searchParams;
    assert.deepEqual([...params].sort(), [['$orderby', 'Id desc'], ['$select', 'Id,Title'], ['$skiptoken', 'Paged=TRUE&p_ID=10'], ['$top', '10']]);
    await finish(f, 1, json({ value: [{ Id: 11, Title: 'Next11' }, { Id: 12, Title: 'Next12' }] }));
    assert.match(f.container.textContent, /Items \(12\)/);
    assert.deepEqual([...f.container.querySelectorAll('dd')].map(n => n.textContent), ['2', '10', 'false', '12']);
});
function replaceClient(f) { const next = f.host(f.instance.instanceId); next.context.serviceScope = f.instance.context.serviceScope; f.replace(next); }
const command = kind => kind === 'context' ? 'Read web' : kind === 'generic' ? 'Invoke web read' : 'Run';
function prepare(f, kind) { if (kind === 'list')
    input(f, 'List title', 'Tasks'); }
function result(kind, name) { return json(kind === 'list' ? { value: [{ Id: 42, Title: name }] } : { Title: name, Url: 'https://tenant.test/' + name }); }
const denial = text => new Response(text, { status: 403 });
for (const kind of Object.keys(cases)) {
    test(kind + ' current error Dismiss retry and synchronous duplicate lock', async (t) => {
        const f = fixture(t, kind);
        prepare(f, kind);
        const run = action(f, command(kind));
        await f.flush(() => { run(); run(); run(); });
        await drain(f);
        assert.equal(f.calls.length, 1);
        await finish(f, 0, denial('Current direct denial'));
        assert.match(f.container.textContent, /Current direct denial/);
        const dismiss = action(f, 'Dismiss');
        f.click('Dismiss');
        assert.equal(f.container.querySelector('[role="alert"]'), null);
        f.click(command(kind));
        await drain(f);
        await finish(f, 1, result(kind, 'Current recovered'));
        assert.match(f.container.textContent, /Current recovered/);
        await f.flush(() => dismiss());
        assert.match(f.container.textContent, /Current recovered/);
    });
    test(kind + ' construction fault is separate and recovers through actual context/client replacement', async (t) => {
        const f = fixture(t, kind);
        prepare(f, kind);
        const old = action(f, command(kind));
        f.setPnPInitializationError(new Error('Controlled direct construction fault'));
        replaceClient(f);
        await drain(f);
        assert.match(f.container.textContent, /Client initialization error: Controlled direct construction fault/);
        await f.flush(() => old());
        assert.equal(f.calls.length, 0);
        f.setPnPInitializationError(undefined);
        replaceClient(f);
        await drain(f);
        f.click(command(kind));
        await drain(f);
        await finish(f, 0, result(kind, 'Initialization recovered'));
        assert.match(f.container.textContent, /Initialization recovered/);
    });
    for (const change of ['reset', 'client', 'webscalar', 'departure', 'unmount'])
        for (const outcome of ['success', 'error'])
            test(`${kind} ${change} denies retained commands and old ${outcome} without unlocking new work`, async (t) => {
                const f = fixture(t, kind);
                prepare(f, kind);
                const old = action(f, command(kind));
                f.click(command(kind));
                await drain(f);
                assert.equal(f.calls.length, 1);
                if (change === 'reset') {
                    f.reset();
                    prepare(f, kind);
                }
                if (change === 'client')
                    replaceClient(f);
                if (change === 'webscalar') {
                    f.instance.context.pageContext.web.absoluteUrl = 'https://tenant.test/sites/scalar-new';
                    f.render();
                }
                if (change === 'departure')
                    f.select('Inspect the current native host identity', 'packages/spfx-react-toolkit/src/core/context.internal.tsx#useSPFxContext');
                if (change === 'unmount')
                    f.unmount();
                await f.flush(() => old());
                await drain(f);
                assert.equal(f.calls.length, 1, 'retained old command denied before service invocation');
                if (change !== 'departure' && change !== 'unmount') {
                    f.click(command(kind));
                    await drain(f);
                    assert.equal(f.calls.length, 2);
                }
                await finish(f, 0, outcome === 'error' ? denial('Obsolete direct denial') : result(kind, 'Obsolete direct result'));
                assert.equal(f.container.textContent.includes('Obsolete direct'), false);
                if (change !== 'departure' && change !== 'unmount') {
                    const current = action(f, command(kind));
                    await f.flush(() => current());
                    await drain(f);
                    assert.equal(f.calls.length, 2, 'old finally must not release current lock');
                    await finish(f, 1, result(kind, 'Latest direct result'));
                    assert.match(f.container.textContent, /Latest direct result/);
                }
            });
}
test('context raw draft no-op stays usable while accepted ABA and retained input/Read are retired synchronously', async (t) => {
    const f = fixture(t);
    let p = props(f.regions.controls, p => p['aria-label'] === 'Site URL'), run = action(f, 'Read web');
    await f.flush(() => p.onChange({}, { value: '/sites/projects' }));
    await f.flush(() => run());
    await drain(f);
    assert.equal(f.calls.length, 1);
    await finish(f, 0, result('context', 'No-op usable'));
    p = props(f.regions.controls, p => p['aria-label'] === 'Site URL');
    run = action(f, 'Read web');
    await f.flush(() => { p.onChange({}, { value: '/sites/new' }); run(); p.onChange({}, { value: '/sites/old-handler' }); });
    await drain(f);
    assert.equal(f.calls.length, 1);
    assert.equal(f.container.querySelector('input').value, '/sites/new');
    input(f, 'Site URL', '/sites/projects');
    await f.flush(() => run());
    await drain(f);
    assert.equal(f.calls.length, 1, 'ABA does not revive original command');
    f.click('Read web');
    await drain(f);
    assert.match(f.container.textContent, /No-op usable/, 'same URL legitimately replays installed cache');
    assert.equal(f.calls.length, 1);
});
test('list no-op mode preserves its raw draft and retained next-page cannot replay an old offset', async (t) => {
    const f = fixture(t, 'list');
    input(f, 'List title', 'Tasks');
    const mode = props(f.regions.controls, p => p['aria-label'] === 'Selector mode');
    await f.flush(() => mode.onChange({}, { value: 'title' }));
    assert.equal(f.container.querySelector('[aria-label="List title"]').value, 'Tasks');
    f.click('Run');
    await drain(f);
    const ten = Array.from({ length: 10 }, (_, i) => ({ Id: i + 1, Title: 'First ' + i }));
    await finish(f, 0, json({ value: ten }));
    const more = action(f, 'Next page');
    f.click('Next page');
    await drain(f);
    await finish(f, 1, json({ value: ten.map(r => ({ ...r, Id: r.Id + 10 })) }));
    const count = f.calls.length;
    await f.flush(() => more());
    await drain(f);
    assert.equal(f.calls.length, count, 'captured old offset cannot run again');
});
test('list failed next page retains rows and retries captured offset without synthetic success', async (t) => {
    const f = fixture(t, 'list');
    prepare(f, 'list');
    f.click('Run');
    await drain(f);
    await finish(f, 0, json({ value: Array.from({ length: 10 }, (_, i) => ({ Id: i + 1, Title: 'Kept' + i })) }));
    f.click('Next page');
    await drain(f);
    await finish(f, 1, denial('Current page denial'));
    assert.match(f.container.textContent, /Items \(10\)/);
    assert.match(f.container.textContent, /Current page denial/);
    f.click('Dismiss');
    f.click('Next page');
    await drain(f);
    assert.equal(new URL(f.calls[2].url).searchParams.get('$skiptoken'), 'Paged=TRUE&p_ID=10');
    await finish(f, 2, json({ value: [] }));
    assert.match(f.container.textContent, /Items \(10\)/);
    assert.deepEqual([...f.container.querySelectorAll('dd')].map(n => n.textContent), ['0', '10', 'false', '10']);
});
const selectors = [
    ['title', 'List title', "  O'Brien Tasks  ", "/lists/getByTitle('%20%20O''Brien%20Tasks%20%20')"],
    ['id', 'List GUID', ' {ABCDEF01-2345-6789-ABCD-EF0123456789} ', "/lists('abcdef01-2345-6789-abcd-ef0123456789')"],
    ['url', 'Server-relative list URL', "/sites/elsewhere/Lists/O'Brien café % # %20/", "/getList('%2Fsites%2Felsewhere%2FLists%2FO''Brien%20caf%C3%A9%20%25%20%23%20%2520%2F')"],
    ['path', 'Web-relative list path', 'Lists/Tasks/', "/getList('%2Fsites%2Fclients-one%2FLists%2FTasks%2F')"],
];
for (const [mode, label, value, endpoint] of selectors)
    test('direct list selector raw ' + mode + ' routes on the configured physical web without I/O before Run', async (t) => {
        const f = fixture(t, 'list');
        if (mode !== 'title')
            input(f, 'Selector mode', mode);
        input(f, label, value);
        await drain(f);
        assert.equal(f.calls.length, 0);
        f.click('Run');
        await drain(f);
        assert.equal(f.calls[0].url.split('?')[0], 'https://tenant.test/sites/clients-one/_api/web' + endpoint + '/items');
        await finish(f, 0, json({ value: [] }));
        assert.match(f.container.textContent, /No items loaded/);
    });
for (const [mode, label, value] of [['id', 'List GUID', 'not-a-guid'], ['url', 'Server-relative list URL', 'Lists/Invalid'], ['path', 'Web-relative list path', '/Invalid']])
    test('invalid ' + mode + ' remains lazy and rejects before native dispatch', async (t) => {
        const f = fixture(t, 'list');
        input(f, 'Selector mode', mode);
        input(f, label, value);
        assert.equal(f.calls.length, 0);
        f.click('Run');
        await drain(f);
        assert.equal(f.calls.length, 0);
        assert.match(f.container.textContent, /Request error:.*createSPFxPnPListService/);
    });
test('same-cache providers have independent factory clients and UI while identical URLs legitimately share cached data', async (t) => {
    const storage = ownedStorage(t);
    const f = fixture(t);
    t.after(storage.restore);
    const peer = f.mountPeer();
    await drain(f);
    assert.equal(f.calls.length, 0);
    f.click('Read web');
    await drain(f);
    await finish(f, 0, json({ Title: 'Shared URL cached response', Url: 'https://tenant.test/sites/projects' }));
    peer.click('Read web');
    await drain(f);
    assert.equal(f.calls.length, 1, 'default URL-only key replays data for another provider');
    assert.match(peer.container.textContent, /Shared URL cached response/);
    f.reset();
    await drain(f);
    assert.equal(f.container.textContent.includes('Shared URL cached response'), false);
    assert.match(peer.container.textContent, /Shared URL cached response/);
    const a = f.toolkit.createSPFxPnPContextService(f.instance.context, f.instance.context.pageContext), b = f.toolkit.createSPFxPnPContextService(peer.instance.context, peer.instance.context.pageContext);
    assert.notEqual(a, b);
    assert.notEqual(a.createSPFI(), b.createSPFI());
    peer.unmount();
});
for (const kind of ['generic', 'list'])
    test(kind + ' two provider reverse completions publish independently', async (t) => {
        const f = fixture(t, kind);
        prepare(f, kind);
        const peer = f.mountPeer();
        if (kind === 'list') {
            const node = peer.container.querySelector('input');
            const { Simulate } = require('react-dom/test-utils');
            f.h.act(() => { node.value = 'Other'; Simulate.change(node, { target: { value: 'Other' } }); });
        }
        f.click(command(kind));
        peer.click(command(kind));
        await drain(f);
        assert.equal(f.calls.length, 2);
        assert.notEqual(f.calls[0].url, f.calls[1].url);
        await finish(f, 1, result(kind, 'Peer independent'));
        assert.match(peer.container.textContent, /Peer independent/);
        assert.equal(f.container.textContent.includes('Peer independent'), false);
        await finish(f, 0, result(kind, 'Primary independent'));
        assert.match(f.container.textContent, /Primary independent/);
        assert.match(peer.container.textContent, /Peer independent/);
        peer.unmount();
    });
for (const kind of Object.keys(cases))
    test(kind + ' newer result stays visible after obsolete failure and success', async (t) => {
        const f = fixture(t, kind);
        prepare(f, kind);
        f.click(command(kind));
        await drain(f);
        f.reset();
        prepare(f, kind);
        f.click(command(kind));
        await drain(f);
        await finish(f, 1, result(kind, 'Winning direct result'));
        await finish(f, 0, denial('Losing direct failure'));
        assert.match(f.container.textContent, /Winning direct result/);
        assert.equal(f.container.textContent.includes('Losing direct failure'), false);
    });
for (const kind of ['context', 'list'])
    test(kind + ' retained draft and Dismiss cannot mutate a reset session', async (t) => {
        const f = fixture(t, kind);
        prepare(f, kind);
        const label = kind === 'context' ? 'Site URL' : 'List title';
        const oldInput = props(f.regions.controls, p => p['aria-label'] === label).onChange;
        f.click(command(kind));
        await drain(f);
        await finish(f, 0, denial('Former error'));
        const dismiss = action(f, 'Dismiss');
        f.reset();
        prepare(f, kind);
        f.click(command(kind));
        await drain(f);
        await finish(f, 1, denial('Current reset error'));
        await f.flush(() => { oldInput({}, { value: 'Obsolete draft' }); dismiss(); });
        assert.match(f.container.textContent, /Current reset error/);
        assert.equal(f.container.querySelector('input').value, kind === 'context' ? '/sites/projects' : 'Tasks');
    });
test('direct-service native admission is separate default OFF and rejects extra query, origin, method and arbitrary batch shapes', async (t) => {
    const f = fixture(t, 'generic'), descriptor = Object.getOwnPropertyDescriptor(window, 'fetch');
    const native = f.h.load(path.resolve('tests/fixtures/demo-presentation/trim-transports.ts'));
    t.after(() => { if (descriptor)
        Object.defineProperty(window, 'fetch', descriptor);
    else
        delete window.fetch; });
    const root = 'https://simulated.invalid/sites/finite';
    const title = root + "/_api/web/lists/getByTitle('Tasks')/items?$select=Id,Title&$orderby=Id%20desc&$top=10";
    const selector = root + "/_api/web/lists('11111111-2222-3333-4444-555555555555')/items?$select=Id,Title&$orderby=Id%20desc&$top=10&$skiptoken=Paged%3DTRUE%26p_ID%3D10";
    for (const url of [title, selector, root + '/_api/web?$select=Url'])
        await assert.rejects(window.fetch(url), /Unsupported native PnP/);
    native.clientsNative.configure({ pnpServices: true });
    for (const url of [title + '&extra=1', title + '&$top=10', selector.replace('p_ID%3D10', 'p_ID%3Dbad'), title.replace('simulated.invalid', 'external.invalid'), root + '/_api/web?$select=Url&extra=1'])
        await assert.rejects(window.fetch(url), /Unsupported native PnP/);
    for (const method of ['POST', 'PATCH', 'DELETE'])
        await assert.rejects(window.fetch(title.split('?')[0], { method, headers: { 'X-RequestDigest': 'known' }, body: '{"Title":"Forbidden"}' }), /Unsupported native PnP/);
    const malformed = '--batch_bad\nGET ' + root + '/_api/web?%24select=Url HTTP/1.1\nGET ' + root + '/_api/web?%24select=Title HTTP/1.1\n--batch_bad--\n';
    await assert.rejects(window.fetch(root + '/_api/$batch', { method: 'POST', headers: { 'Content-Type': 'multipart/mixed; boundary=batch_bad', 'X-RequestDigest': 'known' }, body: malformed }), /Unsupported native PnP read batch/);
    assert.equal(native.clientsNative.requests.length, 0);
    assert.equal(native.isPnPCrudEnabled(), false);
    const two = '--batch_good\nGET ' + root + '/_api/web?%24select=Title HTTP/1.1\nGET ' + root + '/_api/web?%24select=Url HTTP/1.1\n--batch_good--\n';
    await window.fetch(root + '/_api/$batch', { method: 'POST', headers: { 'Content-Type': 'multipart/mixed; boundary=batch_good', 'X-RequestDigest': 'known' }, body: two });
    assert.equal(native.clientsNative.requests.at(-1).kind, 'pnp-batch');
    native.clientsNative.configure({ pnpServices: false });
    await assert.rejects(window.fetch(title), /Unsupported native PnP/);
});
function transformedFixture(t, kind, transform) {
    const fs = require('node:fs'), os = require('node:os');
    const [file, name, title] = cases[kind], selected = scenario(title, `packages/spfx-react-toolkit/src/services/${file}.ts#${name}`);
    const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'pnp-service-negative-')), copy = path.join(directory, 'scenario.tsx'), original = fs.readFileSync;
    const altered = transform(selected.source);
    assert.notEqual(altered, selected.source);
    fs.writeFileSync(copy, altered);
    t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
    const mock = t.mock.method(fs, 'readFileSync', (file, ...args) => original(String(file) === selected.file ? copy : file, ...args));
    try {
        return fixture(t, kind);
    }
    finally {
        mock.mock.restore();
    }
}
test('negative control independently detects a removed pre-dispatch currentness guard', async (t) => {
    const f = transformedFixture(t, 'context', s => s.replace('if (!current() || !sp || dispatch.current?.token === token) return;', 'if (!sp || dispatch.current?.token === token) return;'));
    const old = action(f, 'Read web');
    f.reset();
    await f.flush(() => old());
    await drain(f);
    assert.throws(() => assert.equal(f.calls.length, 0, 'retired command must issue zero HTTP'), /retired command/);
    await finish(f, 0, result('context', 'Broken guard request consumed'));
});
test('negative control independently detects obsolete publication after reset', async (t) => {
    const f = transformedFixture(t, 'generic', s => s.replace('const owns = (): boolean => current() && sequence.current === request;', 'const owns = (): boolean => true;').replace('const visible = current() && outcome?.token === token ? outcome : undefined;', 'const visible = outcome;'));
    f.click('Invoke web read');
    await drain(f);
    f.reset();
    await finish(f, 0, result('generic', 'Injected obsolete result'));
    assert.throws(() => assert.equal(f.container.textContent.includes('Injected obsolete result'), false, 'obsolete result hidden'), /obsolete result hidden/);
});
test('negative control independently detects lost captured paging generation guard', async (t) => {
    const f = transformedFixture(t, 'list', s => s.replace(' || visible?.request !== sequence.current', ''));
    prepare(f, 'list');
    f.click('Run');
    await drain(f);
    const ten = Array.from({ length: 10 }, (_, i) => ({ Id: i + 1, Title: 'Before' + i }));
    await finish(f, 0, json({ value: ten }));
    const more = action(f, 'Next page');
    f.click('Next page');
    await drain(f);
    await finish(f, 1, json({ value: ten.map(r => ({ ...r, Id: r.Id + 10 })) }));
    await f.flush(() => more());
    await drain(f);
    assert.throws(() => assert.equal(f.calls.length, 2, 'old offset must not dispatch again'), /old offset/);
    await finish(f, 2, json({ value: [] }));
});
test('negative control independently detects omitted configured caching on repeat SDK read', async (t) => {
    const storage = ownedStorage(t);
    const f = transformedFixture(t, 'context', s => s.replace("cache: { enabled: true, storage: 'session', timeout: 60000 }", "cache: { enabled: false, storage: 'session', timeout: 60000 }"));
    t.after(storage.restore);
    f.click('Read web');
    await drain(f);
    await finish(f, 0, result('context', 'Uncached first'));
    f.click('Read web');
    await drain(f);
    assert.throws(() => assert.equal(f.calls.length, 1, 'repeat read before expiry must replay cached data'), /repeat read/);
    await finish(f, 1, result('context', 'Uncached second'));
});
for (const kind of Object.keys(cases))
    test(kind + ' complete portable fallback uses actual public context and explicit command', async (t) => {
        const f = fixture(t, kind, { fallback: true, directProvider: true });
        if (kind === 'list')
            f.change('List title', 'Tasks');
        assert.equal(f.calls.length, 0);
        f.click(command(kind));
        await drain(f);
        await finish(f, 0, result(kind, 'Portable result'));
        assert.match(f.container.textContent, /Portable result/);
    });
test('negative control independently detects obsolete finally releasing a newer operation lock', async (t) => {
    const f = transformedFixture(t, 'generic', s => s.replace('if (owns() && dispatch.current?.request === request && dispatch.current.token === token) dispatch.current = undefined;', 'dispatch.current = undefined;'));
    f.click('Invoke web read');
    await drain(f);
    f.reset();
    f.click('Invoke web read');
    await drain(f);
    await finish(f, 0, result('generic', 'Old settled'));
    const current = action(f, 'Invoke web read');
    await f.flush(() => current());
    await drain(f);
    assert.throws(() => assert.equal(f.calls.length, 2, 'obsolete finally must preserve newer lock'), /obsolete finally/);
    await finish(f, 1, result('generic', 'Intermediate'));
    await finish(f, 2, result('generic', 'Latest'));
});
test('negative control detects a batch callback that does not queue its second request', async (t) => {
    const f = transformedFixture(t, 'generic', s => s.replace("const second = client.web.select('Url')();", "const second = Promise.resolve({ Url: 'Never requested' });"));
    f.click('Batch web reads');
    await drain(f);
    let index = 0;
    if (f.calls[0].url.endsWith('/_api/contextinfo'))
        await finish(f, index++, json({ FormDigestValue: 'negative-digest', FormDigestTimeoutSeconds: 1800 }));
    assert.match(f.container.textContent, /Request error:.*exact two-read batch targets/);
    assert.equal(f.calls.some(request => request.url.endsWith('/_api/$batch')), false, 'broken queue never reaches native transport');
});
for (const description of [undefined, null, '']) {
    test('context legacy missing text fallback: ' + String(description), async t => {
        const f = fixture(t);
        f.click('Read web'); await drain(f);
        await finish(f, 0, json({ Title: 'Legacy context', Url: 'https://tenant.test/legacy', Description: description }));
        assert.match(f.container.textContent, /\(no description\)/);
    });
    test('generic normalizes legacy missing description: ' + String(description), async t => {
        const f = fixture(t, 'generic');
        f.click('Invoke web read'); await drain(f);
        await finish(f, 0, json({ Title: 'Legacy generic', Url: 'https://tenant.test/legacy', Description: description }));
        assert.match(f.container.textContent, /Legacy generic/);
        assert.equal([...f.container.querySelectorAll('dt')].some(node => node.textContent === 'Description'), false);
    });
}
test('Node read-only terminal rejects malformed digest and batch queries before recording a request', t => {
    const f = fixture(t, 'generic');
    assert.throws(() => global.fetch('https://tenant.test/sites/current/_api/contextinfo?extra=1', { method: 'POST', headers: { 'X-PnPjs-NoDigest': '1' } }), /digest query/);
    assert.throws(() => global.fetch('https://tenant.test/sites/current/_api/$batch', { method: 'POST', headers: { 'Content-Type': 'multipart/mixed; boundary=batch_bad', 'X-RequestDigest': 'known' }, body: '--batch_bad\nGET https://tenant.test/sites/current/_api/web?%24select=Url HTTP/1.1\n--batch_bad--\n' }), /exact two-read/);
    assert.equal(f.calls.length, 0);
});

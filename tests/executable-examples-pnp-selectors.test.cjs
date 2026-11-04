const assert = require('node:assert/strict');
const { test } = require('node:test');
const path = require('node:path');
const { createClientsFixture } = require('./fixtures/executable-examples/clients-scenarios.cjs');
const cases = [
  { suffix: 'Id', title: 'Read items by list GUID', label: 'List GUID', action: 'Query by GUID', value: '11111111-2222-3333-4444-555555555555', endpoint: "/sites/clients-one/_api/web/lists('11111111-2222-3333-4444-555555555555')/items" },
  { suffix: 'Url', title: 'Read items by decoded server-relative URL', label: 'Server-relative list URL', action: 'Query by URL', value: '/sites/projects/Lists/Tasks', endpoint: "/sites/clients-one/_api/web/getList('%2Fsites%2Fprojects%2FLists%2FTasks')/items" },
  { suffix: 'Path', title: 'Read items by decoded web-relative path', label: 'Web-relative list path', action: 'Query by path', value: 'Lists/Tasks', endpoint: "/sites/clients-one/_api/web/getList('%2Fsites%2Fclients-one%2FLists%2FTasks')/items" },
];
const owner = c => 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPListBy' + c.suffix + '.ts#useSPFxPnPListBy' + c.suffix;
function fixture(t, c, options = {}) {
  const previous = Object.getOwnPropertyDescriptor(window, 'ResizeObserver');
  Object.defineProperty(window, 'ResizeObserver', { configurable: true, value: class { observe() {} unobserve() {} disconnect() {} } });
  t.after(() => { if (previous) Object.defineProperty(window, 'ResizeObserver', previous); else delete window.ResizeObserver; });
  const f = createClientsFixture(t, c.title, { ...options, pnpContext: true, pnpSelectors: true, ownerKey: owner(c), compiled: options.compiled ?? process.env.PNP_SELECTORS_IMPLEMENTATION === 'compiled' });
  f.action = name => {
    for (const region of Object.values(f.regions ?? {})) { const props = authored(region, p => p.children === name && typeof p.onClick === 'function'); if (props) return props.onClick; }
    assert.fail('Missing authored action: ' + name);
  };
  return f;
}
async function drain(f) { await f.flush(); await new Promise(resolve => setTimeout(resolve, 25)); await f.flush(); }
const json = items => new Response(JSON.stringify({ value: items }), { headers: { 'Content-Type': 'application/json' } });
async function finish(f, index, response) { await f.flush(() => f.calls[index].resolve(response)); await drain(f); }
for (const c of cases) test(c.suffix + ' checkpoint real emitted source/lib/provider selector read reaches native item transport', async t => {
  const f = fixture(t, c, { revision: 17 }); await drain(f);
  assert.equal(f.calls.length, 0);
  const mode = process.env.PNP_SELECTORS_IMPLEMENTATION === 'compiled' ? 'lib' : 'src';
  for (const file of ['hooks/useSPFxPnPListBy' + c.suffix, 'hooks/useSPFxPnPList.internal', 'services/spfx-pnp-list.service', 'hooks/useSPFxPnPContext', 'services/spfx-pnp-context.service'])
    assert.ok(f.pnpFiles.includes(path.resolve('packages/spfx-react-toolkit', mode, file + (mode === 'src' ? '.ts' : '.js'))), file);
  assert.ok(!f.pnpFiles.some(file => file.includes('/' + (mode === 'src' ? 'lib' : 'src') + '/services/spfx-pnp-')));
  f.change(c.label, c.value); await drain(f); assert.equal(f.calls.length, 0);
  f.click(c.action); await drain(f); assert.equal(f.calls.length, 1, 'explicit selector command must dispatch');
  const url = new URL(f.calls[0].url);
  assert.equal(url.origin, 'https://tenant.test'); assert.equal(url.pathname, c.endpoint);
  assert.deepEqual([...url.searchParams].sort(), [['$orderby', 'Id desc'], ['$select', 'Id,Title'], ['$top', '10']]);
  assert.equal(f.calls[0].method, 'GET');
  await finish(f, 0, json([{ Id: 42, Title: 'Result title' }]));
  assert.match(f.container.textContent, /#42 - Result title/);
});

test('presentation helper forwards selector item GETs to native queue while original title pilot stays independent', async t => {
  const f = fixture(t, cases[0], { compiled: true });
  const descriptor = Object.getOwnPropertyDescriptor(window, 'fetch');
  const native = f.h.load(path.resolve('tests/fixtures/demo-presentation/trim-transports.ts'));
  t.after(() => { if (descriptor) Object.defineProperty(window, 'fetch', descriptor); else delete window.fetch; });
  t.mock.method(global, 'fetch', window.fetch.bind(window));
  const helper = f.h.load(path.resolve('tests/fixtures/demo-presentation/pnp-read-transport.ts'));
  const context = helper.createSPFxPnPContextService(f.instance.context, f.instance.context.pageContext);
  const sp = context.createSPFI('https://simulated.invalid/sites/projects');
  const create = f.h.load(path.resolve('packages/spfx-react-toolkit/lib/services/spfx-pnp-list.service.js')).createSPFxPnPListService;
  native.clientsNative.configure({ holdRequests: true });
  const pending = create(sp, { kind: 'id', id: '11111111-2222-3333-4444-555555555555' }, 10).query(q => q.select('Id', 'Title').orderBy('Id', false));
  await drain(f);
  assert.equal(native.clientsNative.requests.length, 1, 'selector goes to native terminal, not pilot fallback');
  assert.equal(window.pnpReadFixture.calls.length, 0);
  assert.equal(native.clientsNative.requests[0].kind, 'pnp-items');
  native.clientsNative.requests[0].resolve(json([{ Id: 73, Title: 'Native selector' }]));
  assert.deepEqual((await pending).items, [{ Id: 73, Title: 'Native selector' }]);
  for (const [target, endpoint] of [
    [{ kind: 'url', serverRelativeUrl: '/sites/elsewhere/Lists/Tasks' }, "/sites/projects/_api/web/getList('%2Fsites%2Felsewhere%2FLists%2FTasks')/items"],
    [{ kind: 'path', webRelativePath: 'Lists/Tasks' }, "/sites/projects/_api/web/getList('%2Fsites%2Fprojects%2FLists%2FTasks')/items"]
  ]) {
    const index = native.clientsNative.requests.length;
    const read = create(sp, target, 10).query(q => q.select('Id', 'Title').orderBy('Id', false)); await drain(f);
    assert.equal(native.clientsNative.requests.length, index + 1); assert.equal(new URL(native.clientsNative.requests[index].target).pathname, endpoint);
    native.clientsNative.requests[index].resolve(json([{ Id: 42, Title: 'Other selector' }])); assert.deepEqual((await read).items, [{ Id: 42, Title: 'Other selector' }]);
  }
  const beforeRejected = native.clientsNative.requests.length;
  const valid = "https://simulated.invalid/sites/projects/_api/web/lists('11111111-2222-3333-4444-555555555555')/items?$select=Id,Title&$orderby=Id%20desc&$top=10";
  for (const request of [
    [valid, { method: 'POST' }], [valid.replace('$top=10', '$top=11'), { method: 'GET' }],
    [valid.replace('simulated.invalid', 'external.invalid'), { method: 'GET' }], [valid + '&unexpected=1', { method: 'GET' }],
    ['https://simulated.invalid/sites/projects/_api/web/arbitrary', { method: 'GET' }],
  ]) await assert.rejects(window.fetch(...request), /Unsupported native PnP fixture route/);
  assert.equal(native.clientsNative.requests.length, beforeRejected); assert.equal(window.pnpReadFixture.calls.length, 0);
  const pilot = await create(sp, 'Pilot').query(q => q.select('Id', 'Title').top(5));
  assert.deepEqual(pilot.items, [{ Id: 1, Title: 'Simulated list row (no tenant request)' }]);
  assert.equal(window.pnpReadFixture.calls.length, 1); assert.equal(window.pnpReadFixture.calls[0].method, 'GET');
  assert.equal(native.clientsNative.requests.length, 3);
});

function authored(value, match) {
  if (Array.isArray(value)) { for (const child of value) { const found = authored(child, match); if (found) return found; } return undefined; }
  if (!value || typeof value !== 'object') return undefined;
  if (match(value.props ?? {})) return value.props;
  return authored(value.props?.children, match);
}
function inputHandler(f, c) { const props = authored(f.regions.controls, p => p['aria-label'] === c.label); assert.ok(props); return value => props.onChange({}, { value }); }
const denied = message => new Response(message, { status: 403, statusText: 'Forbidden' });
function replaceClient(f) {
  const next = f.host(f.instance.instanceId);
  // Keep the finished scope: this replaces the real native context/SPFI without provider unmount.
  next.context.serviceScope = f.instance.context.serviceScope;
  f.replace(next);
}
function alternate(c) { return c.suffix === 'Id' ? 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' : c.suffix === 'Url' ? '/sites/projects/Lists/Other' : 'Lists/Other'; }
for (const c of cases) {
  test(c.suffix + ' blank/whitespace drafts, invalid selector, first revision, reset and confirmed empty are distinct', async t => {
    const f = fixture(t, c, { revision: 19 }); await drain(f);
    const button = () => [...f.container.querySelectorAll('button')].find(n => n.textContent === c.action);
    assert.equal(button().disabled, true); assert.ok(!f.container.textContent.includes('No items loaded.'));
    f.change(c.label, '   '); assert.equal(button().disabled, true);
    await f.flush(() => f.action(c.action)()); await drain(f); assert.equal(f.calls.length, 0);
    f.change(c.label, c.suffix === 'Id' ? 'invalid' : c.suffix === 'Url' ? 'Lists/Tasks' : '/Lists/Tasks');
    assert.equal(button().disabled, false); assert.equal(f.calls.length, 0);
    f.click(c.action); await drain(f); assert.equal(f.calls.length, 0);
    assert.match(f.container.textContent, /Request error:.*Invalid/);
    f.click('Dismiss'); assert.ok(!f.container.textContent.includes('Request error:'));
    f.change(c.label, c.value); f.click(c.action); await drain(f);
    await finish(f, 0, json([])); assert.match(f.container.textContent, /No items loaded\./);
    const input = f.container.querySelector('input'); f.reset(); await drain(f);
    assert.equal(f.container.querySelector('input'), input); assert.equal(input.value, '');
    assert.ok(!f.container.textContent.includes('No items loaded.')); assert.equal(f.calls.length, 1);
  });
  test(c.suffix + ' read errors retry and legacy optional/null/empty titles preserve visible fallback and count', async t => {
    const f = fixture(t, c); f.change(c.label, c.value); f.click(c.action); await drain(f);
    await finish(f, 0, denied('Read denied')); assert.match(f.container.textContent, /Read denied/);
    const dismiss = f.action('Dismiss'); f.click(c.action); await drain(f);
    await f.flush(() => dismiss()); assert.match(f.container.textContent, /Loading/);
    await finish(f, 1, json([{ Id: 42, Title: 'Result title' }, { Id: 43, Title: null }, { Id: 44 }, { Id: 45, Title: '' }]));
    assert.match(f.container.textContent, /Items \(4\)/); assert.match(f.container.textContent, /#42 - Result title/);
    for (const id of [43, 44, 45]) assert.ok(f.container.textContent.includes('#' + id + ' - (untitled)'));
    assert.ok(!f.container.textContent.includes('Read denied'));
  });
  for (const change of ['target', 'client', 'initialization', 'reset', 'web']) for (const settlement of ['success', 'rejection'])
    test(c.suffix + ' ' + change + ' blocks retained commands and late ' + settlement + ' cannot unlock or erase new pending request', async t => {
      const f = fixture(t, c); f.change(c.label, c.value); const oldRun = f.action(c.action), oldInput = inputHandler(f, c);
      f.click(c.action); await drain(f); assert.equal(f.calls.length, 1);
      if (change === 'target') f.change(c.label, alternate(c));
      if (change === 'client') replaceClient(f);
      if (change === 'initialization') {
        f.setPnPInitializationError(new Error('Controlled selector construction fault')); replaceClient(f); await drain(f);
        assert.match(f.container.textContent, /Context initialization error: Controlled selector construction fault/);
        assert.equal([...f.container.querySelectorAll('button')].find(n => n.textContent === c.action).disabled, true);
        await f.flush(() => oldRun()); assert.equal(f.calls.length, 1);
        f.setPnPInitializationError(undefined); replaceClient(f);
      }
      if (change === 'reset') { const node = f.container.querySelector('input'); f.reset(); assert.equal(f.container.querySelector('input'), node); f.change(c.label, c.value); }
      if (change === 'web') { f.instance.context.pageContext.web.absoluteUrl = 'https://tenant.test/sites/new-web'; f.render(); }
      await f.flush(() => { oldRun(); oldInput('obsolete input'); }); await drain(f);
      assert.equal(f.calls.length, 1); assert.notEqual(f.container.querySelector('input').value, 'obsolete input');
      const run = f.action(c.action); await f.flush(() => { run(); run(); }); await drain(f); assert.equal(f.calls.length, 2, 'current session dispatches once');
      await finish(f, 0, settlement === 'success' ? json([{ Id: 1, Title: 'Obsolete result' }]) : denied('Obsolete denial'));
      assert.match(f.container.textContent, /Loading/); assert.ok(!f.container.textContent.includes('Obsolete'));
      await f.flush(() => run()); await drain(f); assert.equal(f.calls.length, 2, 'old finally did not unlock newer command');
      await finish(f, 1, json([{ Id: 73, Title: 'Current result' }])); assert.match(f.container.textContent, /Current result/);
    });
  test(c.suffix + ' synchronous duplicate admission, no-op input and accepted ABA preserve current ownership', async t => {
    const f = fixture(t, c); f.change(c.label, c.value); const oldRun = f.action(c.action), edit = inputHandler(f, c);
    await f.flush(() => { edit(c.value); oldRun(); oldRun(); }); await drain(f); assert.equal(f.calls.length, 1);
    await finish(f, 0, denied('Old error')); const dismiss = f.action('Dismiss');
    await f.flush(() => { edit(alternate(c)); oldRun(); dismiss(); }); await drain(f); assert.equal(f.calls.length, 1);
    f.change(c.label, c.value); await f.flush(() => { oldRun(); edit('obsolete'); dismiss(); }); await drain(f); assert.equal(f.calls.length, 1);
    assert.equal(f.container.querySelector('input').value, c.value);
    f.click(c.action); await drain(f); await finish(f, 1, json([{ Id: 2, Title: 'ABA current' }])); assert.match(f.container.textContent, /ABA current/);
  });
  for (const stale of ['success', 'rejection']) test(c.suffix + ' reverse target completion retains newer ' + stale + ' and blocks obsolete Dismiss', async t => {
    const f = fixture(t, c); f.change(c.label, c.value); f.click(c.action); await drain(f);
    f.change(c.label, alternate(c)); f.click(c.action); await drain(f);
    await finish(f, 1, stale === 'success' ? json([{ Id: 2, Title: 'Latest success' }]) : denied('Latest denial'));
    const dismiss = stale === 'rejection' ? f.action('Dismiss') : undefined;
    await finish(f, 0, stale === 'success' ? denied('Older denial') : json([{ Id: 1, Title: 'Older success' }]));
    assert.match(f.container.textContent, stale === 'success' ? /Latest success/ : /Latest denial/); assert.ok(!f.container.textContent.includes('Older'));
    if (dismiss) { f.change(c.label, c.value); f.click(c.action); await drain(f); await finish(f, 2, denied('Current denial')); await f.flush(() => dismiss()); assert.match(f.container.textContent, /Current denial/); f.click('Dismiss'); assert.ok(!f.container.textContent.includes('Current denial')); }
  });
  test(c.suffix + ' same-target reset masks hook publication and keeps old retained actions inert', async t => {
    const f = fixture(t, c); f.change(c.label, c.value); const run = f.action(c.action); f.click(c.action); await drain(f);
    f.reset(); f.change(c.label, c.value); f.click(c.action); await drain(f);
    await finish(f, 1, json([{ Id: 2, Title: 'After reset' }])); await finish(f, 0, denied('Before reset'));
    await f.flush(() => run()); assert.equal(f.calls.length, 2); assert.match(f.container.textContent, /After reset/); assert.ok(!f.container.textContent.includes('Before reset'));
  });
  test(c.suffix + ' real providers share one module cache but isolate physical webs, failures and peer lifecycle', async t => {
    const f = fixture(t, c); f.change(c.label, c.value); f.click(c.action); await drain(f);
    const peerHost = f.host('selectors-peer'), peer = f.mountPeer(peerHost);
    const { Simulate } = require('react-dom/test-utils');
    f.h.act(() => Simulate.change(peer.container.querySelector('input'), { target: { value: c.value } })); peer.click(c.action); await drain(f);
    assert.equal(f.calls.length, 2); assert.ok(f.calls[0].url.includes('/sites/clients-one/_api/')); assert.ok(f.calls[1].url.includes('/sites/selectors-peer/_api/'));
    await finish(f, 1, json([{ Id: 73, Title: 'Peer result' }])); assert.match(peer.container.textContent, /Peer result/); assert.match(f.container.textContent, /Loading/);
    await finish(f, 0, denied('First provider denial')); assert.match(f.container.textContent, /First provider denial/); assert.ok(!peer.container.textContent.includes('First provider denial'));
    f.reset(); assert.match(peer.container.textContent, /Peer result/); peer.unmount(); f.change(c.label, c.value); f.click(c.action); await drain(f);
    await finish(f, 2, json([{ Id: 42, Title: 'First recovered' }])); assert.match(f.container.textContent, /First recovered/);
  });
  test(c.suffix + ' standalone fallback uses the same real request and actual Explorer reset/departure invalidates work', async t => {
    const f = fixture(t, c, { fallback: true }); f.change(c.label, c.value); f.click(c.action); await drain(f);
    await finish(f, 0, json([{ Id: 42, Title: 'Standalone result' }])); assert.match(f.container.textContent, /Standalone result/);
    const explorer = fixture(t, c, { explorer: true }); await drain(explorer);
    const { Simulate } = require('react-dom/test-utils');
    const feature = [...explorer.container.querySelectorAll('[data-example-feature]')].find(node => node.textContent === 'useSPFxPnPListBy' + c.suffix); assert.ok(feature);
    explorer.h.act(() => Simulate.click(feature)); await drain(explorer); assert.equal(explorer.calls.length, 0); explorer.change(c.label, c.value); explorer.click(c.action); await drain(explorer);
    const node = explorer.container.querySelector('input'); explorer.click('Reset scenario'); assert.equal(explorer.container.querySelector('input'), node);
    await finish(explorer, 0, json([{ Id: 1, Title: 'Before Explorer reset' }])); assert.ok(!explorer.container.textContent.includes('Before Explorer reset'));
    explorer.change(c.label, c.value); explorer.click(c.action); await drain(explorer);
    const nav = [...explorer.container.querySelectorAll('nav button')].find(button => button.textContent === 'useSPFxContext'); assert.ok(nav);
    explorer.h.act(() => Simulate.click(nav)); await drain(explorer);
    await finish(explorer, 1, denied('Departed selector denial')); assert.ok(!explorer.container.textContent.includes('Departed selector denial')); assert.match(explorer.container.textContent, /Native context available/);
  });
  test(c.suffix + ' unmount consumes late rejection and retained Query performs no further request', async t => {
    const f = fixture(t, c); f.change(c.label, c.value); const run = f.action(c.action); f.click(c.action); await drain(f); f.unmount();
    await finish(f, 0, denied('Unmounted rejection')); await f.flush(() => run()); assert.equal(f.calls.length, 1);
  });
}

for (const [c, raw, webUrl, expected] of [
  [cases[0], ' \t{ABCDEF01-2345-6789-ABCD-EF0123456789}\n', undefined, "/sites/clients-one/_api/web/lists('abcdef01-2345-6789-abcd-ef0123456789')/items"],
  [cases[1], "/sites/elsewhere/Lists/O'Brien café % # %20/", undefined, "/sites/clients-one/_api/web/getList('%2Fsites%2Felsewhere%2FLists%2FO''Brien%20caf%C3%A9%20%25%20%23%20%2520%2F')/items"],
  [cases[2], "Lists/O'Brien café % # %20/", 'https://tenant.test/sites/Project%20One/sub%25%23', "/sites/Project%20One/sub%25%23/_api/web/getList('%2Fsites%2FProject%20One%2Fsub%25%23%2FLists%2FO''Brien%20caf%C3%A9%20%25%20%23%20%2520%2F')/items"],
  [cases[2], 'Lists/%20', 'https://tenant.test/sites/Literal%2520', "/sites/Literal%2520/_api/web/getList('%2Fsites%2FLiteral%2520%2FLists%2F%2520')/items"],
]) test(c.suffix + ' actual emitted raw routing/escaping oracle ' + JSON.stringify(raw), async t => {
  const f = fixture(t, c);
  if (webUrl) { f.instance.context.pageContext.web.absoluteUrl = webUrl; f.render(); }
  f.change(c.label, raw); assert.equal(f.calls.length, 0); f.click(c.action); await drain(f);
  assert.equal(f.calls.length, 1); assert.equal(new URL(f.calls[0].url).pathname, expected);
  await finish(f, 0, json([{ Id: 42, Title: 'Literal route result' }])); assert.match(f.container.textContent, /Literal route result/);
});
for (const c of cases) test(c.suffix + ' real scope departure resets baseline, denies old commands and late results', async t => {
  const f = fixture(t, c); f.change(c.label, c.value); const old = f.action(c.action); f.click(c.action); await drain(f);
  f.replace(f.host(f.instance.instanceId)); await drain(f);
  assert.equal(f.container.querySelector('input').value, '', 'provider scope change remounts its child');
  await f.flush(() => old()); assert.equal(f.calls.length, 1);
  f.change(c.label, c.value); f.click(c.action); await drain(f); await finish(f, 1, json([{ Id: 73, Title: 'Scope result' }]));
  await finish(f, 0, denied('Old scope')); assert.match(f.container.textContent, /Scope result/); assert.ok(!f.container.textContent.includes('Old scope'));
});
function negativeFixture(t, c, transform) {
  const fs = require('node:fs'), os = require('node:os');
  const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
  const selected = scenario(c.title, owner(c)), read = fs.readFileSync;
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'pnp-selectors-negative-')), copy = path.join(directory, 'scenario.tsx');
  const altered = transform(selected.source); assert.notEqual(altered, selected.source, 'negative control changes the intended guard'); fs.writeFileSync(copy, altered);
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const mock = t.mock.method(fs, 'readFileSync', (file, ...args) => read(String(file) === selected.file ? copy : file, ...args));
  try { return fixture(t, c); } finally { mock.mock.restore(); }
}
for (const c of cases) {
  test(c.suffix + ' negative control catches removed pre-dispatch ownership on retained Query after reset', async t => {
    const f = negativeFixture(t, c, source => source.replace('mounted.current && boundary.current.token === token', 'mounted.current'));
    f.change(c.label, c.value); const run = f.action(c.action); f.reset(); await f.flush(() => run()); await drain(f);
    assert.throws(() => assert.equal(f.calls.length, 0, 'obsolete Query cannot dispatch'), /obsolete Query cannot dispatch/);
    await finish(f, 0, json([{ Id: 1, Title: 'Faulty dispatch' }]));
  });
  test(c.suffix + ' negative control catches missing synchronous duplicate dispatch lock', async t => {
    const f = negativeFixture(t, c, source => source.replace(' || dispatch.current?.token === token', ''));
    f.change(c.label, c.value); const run = f.action(c.action); await f.flush(() => { run(); run(); }); await drain(f);
    assert.throws(() => assert.equal(f.calls.length, 1, 'one same-tick command'), /one same-tick command/);
    await finish(f, 0, json([])); await finish(f, 1, json([]));
  });
  test(c.suffix + ' negative control catches stale publication erasing newer settled outcome', async t => {
    const f = negativeFixture(t, c, source => source.replace('const owns = (): boolean => current() && sequence.current === request;', 'const owns = (): boolean => true;'));
    f.change(c.label, c.value); f.click(c.action); await drain(f); f.change(c.label, alternate(c)); f.click(c.action); await drain(f);
    await finish(f, 1, denied('Latest protected error')); await finish(f, 0, json([{ Id: 1, Title: 'Obsolete result' }]));
    assert.throws(() => assert.match(f.container.textContent, /Latest protected error/), /Latest protected error/);
  });
}

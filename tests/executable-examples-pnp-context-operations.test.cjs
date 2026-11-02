const assert = require('node:assert/strict');
const { test } = require('node:test');
const path = require('node:path');
const { createClientsFixture } = require('./fixtures/executable-examples/clients-scenarios.cjs');
const title = 'Read current or cross-site web information';
const ownerKey = 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPContext.ts#useSPFxPnPContext';
function fixture(t, options = {}) {
  const previous = Object.getOwnPropertyDescriptor(window, 'ResizeObserver');
  Object.defineProperty(window, 'ResizeObserver', { configurable: true, value: class { observe() {} unobserve() {} disconnect() {} } });
  t.after(() => { if (previous) Object.defineProperty(window, 'ResizeObserver', previous); else delete window.ResizeObserver; });
  return createClientsFixture(t, title, { ...options, pnpContext: true, ownerKey, compiled: options.compiled ?? process.env.PNP_CONTEXT_IMPLEMENTATION === 'compiled' });
}
async function drain(f) { await f.flush(); await new Promise(resolve => setTimeout(resolve, 25)); await f.flush(); }
function web(Title = 'Actual web result', Description = 'Native description') { return new Response(JSON.stringify({ Title, Url: 'https://tenant.test/result', Description }), { headers: { 'Content-Type': 'application/json' } }); }
test('checkpoint actual emitted provider constructs matching-mode hook and factory without requests', async t => {
  const f = fixture(t, { revision: 17 }); await drain(f);
  assert.equal(f.calls.length, 0); assert.equal(f.tokens.length + f.initializations.length + f.downloads.length, 0);
  const mode = process.env.PNP_CONTEXT_IMPLEMENTATION === 'compiled' ? 'lib' : 'src';
  const other = mode === 'src' ? 'lib' : 'src';
  const stem = path.resolve('packages/spfx-react-toolkit', mode);
  assert.ok(f.pnpFiles.includes(path.join(stem, 'hooks/useSPFxPnPContext.' + (mode === 'src' ? 'ts' : 'js'))));
  assert.ok(f.pnpFiles.includes(path.join(stem, 'services/spfx-pnp-context.service.' + (mode === 'src' ? 'ts' : 'js'))), 'actual factory read by the actual hook in the same module cache');
  assert.ok(!f.pnpFiles.some(file => file.includes('/' + other + '/services/spfx-pnp-context.service.')));
  f.change('Site URL', '/sites/draft'); f.reset(); await drain(f);
  assert.equal(f.calls.length, 0); assert.equal(f.container.querySelector('input').value, '');
});
test('checkpoint explicit Run reaches real PnP terminal fetch and renders a web object', async t => {
  const f = fixture(t); f.click('Run'); await drain(f);
  assert.equal(f.calls.length, 1);
  const url = new URL(f.calls[0].url);
  assert.equal(url.pathname, '/sites/clients-one/_api/web');
  assert.equal(url.searchParams.get('$select'), 'Title,Url,Description');
  assert.equal(f.calls[0].method, 'GET'); assert.equal(f.calls[0].kind, 'pnp');
  await f.flush(() => f.calls[0].resolve(web())); await drain(f);
  assert.match(f.container.textContent, /Actual web result/); assert.match(f.container.textContent, /Native description/);
  assert.match(f.container.textContent, /https:\/\/tenant.test\/result/);
});
function authored(value, match) {
  if (!value || typeof value !== 'object') return undefined;
  if (match(value.props ?? {})) return value.props;
  for (const child of Array.isArray(value) ? value : [value.props?.children]) { const found = authored(child, match); if (found) return found; }
  return undefined;
}
function handler(f, name) {
  for (const region of Object.values(f.regions ?? {})) { const found = authored(region, props => props.children === name && typeof props.onClick === 'function'); if (found) return found.onClick; }
  const node = [...f.container.querySelectorAll('button')].find(node => node.textContent === name);
  assert.ok(node, name);
  const key = Object.keys(node).find(key => key.startsWith('__reactFiber$'));
  let fiber = node[key], callback;
  while (fiber) { if (fiber.memoizedProps?.onClick) callback = fiber.memoizedProps.onClick; fiber = fiber.return; }
  assert.ok(callback); return () => callback({ currentTarget: node, target: node, preventDefault() {}, stopPropagation() {} });
}
function changeHandler(f) { const props = authored(f.regions.controls, props => props['aria-label'] === 'Site URL'); assert.ok(props); return value => props.onChange({}, { value }); }
async function finish(f, index, result) { await f.flush(() => f.calls[index].resolve(result)); await drain(f); }
const denied = message => new Response(message, { status: 403, statusText: 'Forbidden' });
for (const [raw, expected] of [['/sites/other/', 'https://tenant.test/sites/other'], ['https://elsewhere.invalid/sites/absolute/', 'https://elsewhere.invalid/sites/absolute'], ['', 'https://tenant.test/sites/clients-one']])
  test(`raw target ${JSON.stringify(raw)} resolves publicly and Run reads that exact web`, async t => {
    const f = fixture(t); f.change('Site URL', raw); await drain(f); assert.equal(f.calls.length, 0);
    assert.ok(f.container.textContent.includes('Effective web: ' + expected + '; initialized: Yes'));
    f.click('Run'); await drain(f); assert.equal(f.calls.length, 1);
    assert.equal(f.calls[0].url.split('/_api/')[0], expected);
    await finish(f, 0, web('Chosen target')); assert.match(f.container.textContent, /Chosen target/);
  });
test('whitespace remains a raw public target and never silently means the current web', async t => {
  const f = fixture(t); f.change('Site URL', '   '); await drain(f);
  assert.equal(f.calls.length, 0); assert.ok(f.container.textContent.includes('Effective web:    ; initialized: Yes'));
  f.click('Run'); await drain(f); assert.equal(f.calls.length, 1);
  assert.notEqual(f.calls[0].url, 'https://tenant.test/sites/clients-one/_api/web?%24select=Title%2CUrl%2CDescription');
  await finish(f, 0, web('Raw target'));
});
for (const description of [undefined, '', null]) test(`SharePoint legacy Description ${JSON.stringify(description)} displays the explicit fallback`, async t => {
  const f = fixture(t); f.click('Run'); await drain(f); await finish(f, 0, web('Description case', description));
  // Explicit undefined uses web helper's default; construct actual omitted field separately.
  if (description === undefined) { f.click('Run'); await drain(f); await finish(f, 1, new Response(JSON.stringify({ Title: 'Missing', Url: 'https://tenant.test' }), { headers: { 'Content-Type': 'application/json' } })); }
  assert.match(f.container.textContent, /\(no description\)/);
});
test('current request rejection is consumed, Dismiss clears only request error and Run retries', async t => {
  const f = fixture(t); f.click('Run'); await drain(f); await finish(f, 0, denied('Read denied'));
  assert.match(f.container.textContent, /Request error:.*Read denied/); assert.match(f.container.textContent, /initialized: Yes/);
  f.click('Dismiss'); assert.ok(!f.container.textContent.includes('Request error:'));
  f.click('Run'); await drain(f); await finish(f, 1, web('Retry succeeded')); assert.match(f.container.textContent, /Retry succeeded/);
});
test('controlled construction-dependency fault presents initialization separately with actual hook and factory loaded', async t => {
  const f = fixture(t); f.setPnPInitializationError(new Error('Controlled spfi construction fault'));
  f.replace(f.host(f.instance.instanceId)); await drain(f);
  assert.match(f.container.textContent, /Context initialization error: Controlled spfi construction fault/);
  assert.equal([...f.container.querySelectorAll('button')].find(n => n.textContent === 'Run').disabled, true);
  await f.flush(() => handler(f, 'Run')()); await drain(f); assert.equal(f.calls.length, 0);
  f.reset(); assert.match(f.container.textContent, /Context initialization error:/); assert.equal(f.calls.length, 0);
  f.setPnPInitializationError(undefined); f.replace(f.host(f.instance.instanceId)); await drain(f);
  assert.ok(!f.container.textContent.includes('Context initialization error:')); f.click('Run'); await drain(f);
  await finish(f, 0, web('Recovered construction')); assert.match(f.container.textContent, /Recovered construction/);
});
for (const change of ['target', 'client', 'reset', 'host']) for (const stale of ['success', 'rejection'])
  test(`${change} invalidates retained Run and late ${stale} without erasing a newer pending result`, async t => {
    const f = fixture(t); const original = handler(f, 'Run'); f.click('Run'); await drain(f); assert.equal(f.calls.length, 1);
    if (change === 'target') f.change('Site URL', '/sites/new-target');
    if (change === 'client') f.replace(f.host(f.instance.instanceId));
    if (change === 'reset') { const input = f.container.querySelector('input'); f.reset(); assert.equal(f.container.querySelector('input'), input, 'Reset keeps Demo identity'); }
    if (change === 'host') f.replace(f.host('new-host'));
    await f.flush(() => original()); await drain(f); assert.equal(f.calls.length, 1, 'retained old command does not dispatch');
    f.click('Run'); await drain(f); assert.equal(f.calls.length, 2);
    await finish(f, 0, stale === 'success' ? web('Obsolete success') : denied('Obsolete rejection'));
    assert.match(f.container.textContent, /Loading selected web/); assert.ok(!f.container.textContent.includes('Obsolete'));
    await finish(f, 1, web('Current result')); assert.match(f.container.textContent, /Current result/); assert.ok(!f.container.textContent.includes('Obsolete'));
  });
for (const last of ['success', 'rejection']) test(`same-context reverse completion preserves latest ${last} after older completion`, async t => {
  const f = fixture(t); const run = handler(f, 'Run'); await f.flush(() => { run(); run(); }); await drain(f); assert.equal(f.calls.length, 2);
  await finish(f, 1, last === 'success' ? web('Latest success') : denied('Latest denial'));
  await finish(f, 0, last === 'success' ? denied('Old denial') : web('Old success'));
  assert.match(f.container.textContent, last === 'success' ? /Latest success/ : /Latest denial/); assert.ok(!f.container.textContent.includes('Old denial') && !f.container.textContent.includes('Old success'));
});
test('authored draft admission retires old Run and Dismiss synchronously, while same-value edits keep commands usable', async t => {
  const f = fixture(t); const old = handler(f, 'Run'), edit = changeHandler(f);
  await f.flush(() => { edit(''); old(); }); await drain(f); assert.equal(f.calls.length, 1, 'same raw draft does not retire Run');
  await finish(f, 0, denied('Old error')); const dismiss = handler(f, 'Dismiss'), run = handler(f, 'Run'), next = changeHandler(f);
  await f.flush(() => { next('/sites/B'); run(); dismiss(); }); await drain(f); assert.equal(f.calls.length, 1);
  f.change('Site URL', ''); await f.flush(() => { old(); run(); }); await drain(f); assert.equal(f.calls.length, 1, 'accepted A-B-A cannot resurrect a retained command');
  f.click('Run'); await drain(f); await finish(f, 1, denied('New error')); await f.flush(() => dismiss());
  assert.match(f.container.textContent, /New error/);
});
test('retained Dismiss from an earlier request cannot clear the newer same-target failure', async t => {
  const f = fixture(t); f.click('Run'); await drain(f); await finish(f, 0, denied('First failure')); const oldDismiss = handler(f, 'Dismiss');
  f.click('Run'); await drain(f); await finish(f, 1, denied('Second failure')); await f.flush(() => oldDismiss());
  assert.match(f.container.textContent, /Second failure/); f.click('Dismiss'); assert.ok(!f.container.textContent.includes('Request error:'));
});
for (const stale of ['success', 'rejection']) test(`departure and unmount consume late ${stale} without publication or retained dispatch`, async t => {
  const f = fixture(t); const run = handler(f, 'Run'); f.click('Run'); await drain(f);
  f.select('Inspect the current native host identity', 'packages/spfx-react-toolkit/src/core/context.internal.tsx#useSPFxContext');
  await f.flush(() => run()); await drain(f); assert.equal(f.calls.length, 1);
  await finish(f, 0, stale === 'success' ? web('Departed result') : denied('Departed error')); assert.ok(!f.container.textContent.includes('Departed'));
  f.select(title, ownerKey); f.click('Run'); await drain(f); f.unmount(); await finish(f, 1, stale === 'success' ? web('Unmounted') : denied('Unmounted error'));
});
test('two real providers in one module cache keep client targets and UI results isolated', async t => {
  const f = fixture(t); const peer = f.mountPeer(); t.after(peer.unmount); await drain(f); assert.equal(f.calls.length, 0);
  f.click('Run'); peer.click('Run'); await drain(f);
  assert.deepEqual(f.calls.map(r => new URL(r.url).pathname).sort(), ['/sites/clients-one/_api/web', '/sites/clients-peer/_api/web']);
  await finish(f, 1, web('Peer web')); await finish(f, 0, web('Primary web'));
  assert.match(peer.container.textContent, /Peer web/); assert.ok(!peer.container.textContent.includes('Primary web'));
  assert.match(f.container.textContent, /Primary web/); assert.ok(!f.container.textContent.includes('Peer web'));
});
test('portable standalone fallback uses actual provider and native request results', async t => {
  const f = fixture(t, { directProvider: true, fallback: true }); f.click('Run'); await drain(f); await finish(f, 0, web('Standalone result')); assert.match(f.container.textContent, /Standalone result/);
});

test('actual Explorer departure keeps navigation available and retires the dispatched context case', async t => {
  const { Simulate } = require('react-dom/test-utils');
  const f = fixture(t, { explorer: true }); await drain(f);
  const select = name => { const node = [...f.container.querySelectorAll('[data-example-feature]')].find(node => node.textContent === name); assert.ok(node, name); f.h.act(() => Simulate.click(node)); };
  select('useSPFxPnPContext'); await drain(f); assert.equal(f.calls.length, 0);
  const original = handler(f, 'Run'); f.click('Run'); await drain(f); assert.equal(f.calls.length, 1);
  const reset = [...f.container.querySelectorAll('button')].find(node => node.textContent === 'Ripristina scenario'); assert.ok(reset);
  f.h.act(() => Simulate.click(reset)); await f.flush(() => original()); await drain(f); assert.equal(f.calls.length, 1);
  select('useSPFxContext'); await drain(f); await finish(f, 0, web('Departed Explorer result'));
  assert.ok(!f.container.textContent.includes('Departed Explorer result')); assert.match(f.container.textContent, /Native context available/);
});
function negativeFixture(t, transform) {
  const fs = require('node:fs'), os = require('node:os');
  const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
  const selected = scenario(title, ownerKey), original = fs.readFileSync;
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'pnp-context-negative-'));
  const copy = path.join(directory, 'scenario.tsx');
  const altered = transform(selected.source); assert.notEqual(altered, selected.source, 'negative control changes the intended guard'); fs.writeFileSync(copy, altered);
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const mocked = t.mock.method(fs, 'readFileSync', (file, ...args) => original(String(file) === selected.file ? copy : file, ...args));
  try { return fixture(t); } finally { mocked.mock.restore(); }
}
test('negative control: missing current token guard is detected by retained Run after Reset', async t => {
  const f = negativeFixture(t, source => source.replace('mounted.current && boundary.current.token === token', 'mounted.current'));
  const old = handler(f, 'Run'); f.reset(); await f.flush(() => old()); await drain(f);
  assert.throws(() => assert.equal(f.calls.length, 0, 'retained command must not dispatch'), /retained command must not dispatch/);
  await finish(f, 0, web('Faulty dispatch'));
});
test('negative control: missing request ownership is detected by reverse same-context settlement', async t => {
  const f = negativeFixture(t, source => source.replace('current() && sequence.current === request', 'current()'));
  const run = handler(f, 'Run'); await f.flush(() => { run(); run(); }); await drain(f);
  await finish(f, 1, denied('Latest denial')); await finish(f, 0, web('Old success'));
  assert.throws(() => assert.match(f.container.textContent, /Latest denial/), /Latest denial/);
  assert.match(f.container.textContent, /Old success/);
});
test('existing compiled presentation helper routes web objects through the controlled native queue and preserves pilot API/list rows', async t => {
  const f = fixture(t, { compiled: true });
  const descriptor = Object.getOwnPropertyDescriptor(window, 'fetch');
  const native = f.h.load(path.resolve('tests/fixtures/demo-presentation/trim-transports.ts'));
  t.after(() => { if (descriptor) Object.defineProperty(window, 'fetch', descriptor); else delete window.fetch; });
  t.mock.method(global, 'fetch', window.fetch.bind(window));
  const helper = f.h.load(path.resolve('tests/fixtures/demo-presentation/pnp-read-transport.ts'));
  const service = helper.createSPFxPnPContextService(f.instance.context, f.instance.context.pageContext);
  const client = service.createSPFI('https://simulated.invalid/sites/context-native');
  native.clientsNative.configure({ holdRequests: true });
  let value; const pending = client.web.select('Title', 'Url', 'Description')().then(result => { value = result; }); await drain(f);
  assert.equal(window.pnpReadFixture.calls.length, 0, 'new web protocol does not consume pilot list ledger');
  assert.equal(native.clientsNative.requests.length, 1); const request = native.clientsNative.requests[0];
  assert.equal(request.kind, 'pnp-web'); assert.equal(new URL(request.target).searchParams.get('$select'), 'Title,Url,Description');
  request.resolve(web('Native protocol result')); await pending; assert.equal(value.Title, 'Native protocol result');
  assert.deepEqual(Object.keys(window.pnpReadFixture).sort(), ['calls', 'configure', 'release', 'simulation']);
  f.h.load(require.resolve('@pnp/sp/lists')); f.h.load(require.resolve('@pnp/sp/items'));
  const rows = await client.web.lists.getByTitle('Pilot').items.select('Id', 'Title').top(5)();
  assert.deepEqual(rows, [{ Id: 1, Title: 'Simulated list row (no tenant request)' }]);
  assert.equal(window.pnpReadFixture.calls.length, 1); assert.equal(window.pnpReadFixture.calls[0].method, 'GET');
  const before = native.clientsNative.requests.length;
  await assert.rejects(window.fetch('https://external.invalid/_api/web?$select=Title,Url,Description'), /Unsupported native PnP fixture route/);
  await assert.rejects(window.fetch('https://simulated.invalid/_api/web?$select=Title,Url,Description', { method: 'POST' }), /Unsupported native PnP fixture route/);
  assert.equal(native.clientsNative.requests.length, before);
});

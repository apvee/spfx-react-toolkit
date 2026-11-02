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

// These regressions exercise real .ts generic arrows through the same loader as TSX cases.
test('services harness executes the actual generic source hook without TSX misclassification', async t => {
  const { createHarness } = require('./services-test-harness.cjs');
  const h = createHarness(); t.after(h.close);
  const usePnP = h.load('hooks/useSPFxPnP.ts').useSPFxPnP;
  const mounted = h.mount(() => usePnP());
  const calls = h.withTransport(mounted.current.sp);
  let result;
  await h.act(async () => { result = await mounted.current.invoke(client => client.web.select('Title')()); });
  assert.deepEqual(result, { Title: 'response-1' });
  assert.equal(calls.length, 1); assert.equal(mounted.current.isLoading, false);
});
test('services harness executes the actual generic source service and preserves supplied client identity', async t => {
  const { createHarness } = require('./services-test-harness.cjs');
  const h = createHarness(); t.after(h.close);
  const create = h.load('services/spfx-pnp.service.ts').createSPFxPnPService;
  const context = h.load('services/spfx-pnp-context.service.ts').createSPFxPnPContextService(h.context.current, h.pageContext);
  const sp = context.createSPFI(); h.withTransport(sp);
  const service = create(sp);
  assert.equal(service.sp, sp);
  assert.deepEqual(await service.invoke(client => client.web.select('Title')()), { Title: 'response-1' });
});

const operationsTitle = 'Compare explicit invoke and batch reads';
const operationsOwner = 'packages/spfx-react-toolkit/src/hooks/useSPFxPnP.ts#useSPFxPnP';
function operationsFixture(t, options = {}) {
  const previous = Object.getOwnPropertyDescriptor(window, 'ResizeObserver');
  Object.defineProperty(window, 'ResizeObserver', { configurable: true, value: class { observe() {} unobserve() {} disconnect() {} } });
  t.after(() => { if (previous) Object.defineProperty(window, 'ResizeObserver', previous); else delete window.ResizeObserver; });
  const f = createClientsFixture(t, operationsTitle, { ...options, pnpContext: true, pnpOperations: true, ownerKey: operationsOwner, compiled: options.compiled ?? process.env.PNP_CONTEXT_IMPLEMENTATION === 'compiled' });
  return f;
}
const json = body => new Response(JSON.stringify(body), { headers: { 'Content-Type': 'application/json' } });
function requestMode(f, value) {
  const props = authored(f.regions.controls, p => p['aria-label'] === 'Request'); assert.ok(props);
  f.h.act(() => props.onChange({}, { value }));
}
function batchReply(status = 200) {
  const values = [{ value: [{ Title: 'Operations tasks' }, { Title: 'Operations news' }] }, { Title: 'Operations reader' }, { Title: 'Operations web' }];
  const lines = values.flatMap((value, index) => ['--batchresponse_operations', 'Content-Type: application/http', 'Content-Transfer-Encoding: binary', '', `HTTP/1.1 ${index === 1 ? status : 200} ${index === 1 && status !== 200 ? 'Forbidden' : 'OK'}`, 'Content-Type: application/json', '', JSON.stringify(index === 1 && status !== 200 ? { error: { message: 'Inner current user denied' } } : value)]);
  return new Response([...lines, '--batchresponse_operations--', ''].join('\r\n'), { headers: { 'Content-Type': 'multipart/mixed; boundary=batchresponse_operations' } });
}
test('operations checkpoint actual emitted provider invokes exact nonhidden top10 read with matching subjects', async t => {
  const f = operationsFixture(t, { revision: 17 }); await drain(f);
  const mode = process.env.PNP_CONTEXT_IMPLEMENTATION === 'compiled' ? 'lib' : 'src';
  for (const file of ['hooks/useSPFxPnP', 'services/spfx-pnp.service', 'hooks/useSPFxPnPContext', 'services/spfx-pnp-context.service'])
    assert.ok(f.pnpFiles.includes(path.resolve('packages/spfx-react-toolkit', mode, file + (mode === 'src' ? '.ts' : '.js'))), file);
  assert.equal(f.calls.length, 0); requestMode(f, 'batch'); requestMode(f, 'invoke'); f.reset(); await drain(f); assert.equal(f.calls.length, 0);
  f.click('Run'); await drain(f); assert.equal(f.calls.length, 1, 'explicit invoke must reach native transport');
  const url = new URL(f.calls[0].url); assert.equal(url.pathname, '/sites/clients-one/_api/web/lists');
  assert.deepEqual([...url.searchParams].sort(), [['$filter', 'Hidden eq false'], ['$select', 'Title,ItemCount'], ['$top', '10']]);
  await finish(f, 0, json({ value: [{ Title: 'Operations tasks', ItemCount: 17 }] }));
  assert.match(f.container.textContent, /Operations tasks.*17 items/);
});
test('operations checkpoint cold batch obtains digest and dispatches three real inner GETs', async t => {
  const f = operationsFixture(t); requestMode(f, 'batch'); f.click('Run'); await drain(f);
  assert.equal(f.calls.length, 1, 'cold digest request before batch');
  assert.equal(f.calls[0].method, 'POST'); assert.equal(new URL(f.calls[0].url).pathname, '/sites/clients-one/_api/contextinfo');
  assert.equal(f.calls[0].init.headers['X-PnPjs-NoDigest'], '1');
  await finish(f, 0, json({ FormDigestValue: 'fixture-pnp-operations-digest', FormDigestTimeoutSeconds: 1800 }));
  assert.equal(f.calls.length, 2); const envelope = f.calls[1];
  assert.equal(new URL(envelope.url).pathname, '/sites/clients-one/_api/$batch'); assert.equal(envelope.method, 'POST');
  assert.equal(envelope.init.headers['X-RequestDigest'], 'fixture-pnp-operations-digest');
  assert.match(envelope.init.headers['Content-Type'], /^multipart\/mixed; boundary=batch_/);
  const inner = [...envelope.init.body.matchAll(/^GET (.+) HTTP\/1\.1$/gm)].map(match => new URL(match[1]));
  assert.equal(inner.length, 3); assert.ok(!envelope.init.body.includes('changeset'));
  assert.deepEqual(inner.map(url => [url.origin + url.pathname, [...url.searchParams]]), [
    ['https://tenant.test/sites/clients-one/_api/web/lists', [['$select', 'Title'], ['$top', '5']]],
    ['https://tenant.test/sites/clients-one/_api/web/currentUser', [['$select', 'Title']]],
    ['https://tenant.test/sites/clients-one/_api/web', [['$select', 'Title']]],
  ]);
  await finish(f, 1, batchReply());
  for (const expected of ['Operations tasks', 'Operations news', 'Operations reader', 'Operations web']) assert.ok(f.container.textContent.includes(expected), expected);
});
test('operations presentation helper and native queue preserve real cold digest and multipart parsing', async t => {
  const f = operationsFixture(t, { compiled: true });
  const descriptor = Object.getOwnPropertyDescriptor(window, 'fetch');
  const native = f.h.load(path.resolve('tests/fixtures/demo-presentation/trim-transports.ts'));
  t.after(() => { if (descriptor) Object.defineProperty(window, 'fetch', descriptor); else delete window.fetch; });
  t.mock.method(global, 'fetch', window.fetch.bind(window));
  const helper = f.h.load(path.resolve('tests/fixtures/demo-presentation/pnp-read-transport.ts'));
  const context = helper.createSPFxPnPContextService(f.instance.context, f.instance.context.pageContext);
  const client = context.createSPFI('https://simulated.invalid/sites/operations-native');
  const create = f.h.load(path.resolve('packages/spfx-react-toolkit/lib/services/spfx-pnp.service.js')).createSPFxPnPService;
  const service = create(client); native.clientsNative.configure({ holdRequests: true });
  let result, error;
  const pending = service.batch(sp => Promise.all([sp.web.lists.select('Title').top(5)(), sp.web.currentUser.select('Title')(), sp.web.select('Title')()])).then(value => { result = value; }, failure => { error = failure; });
  await drain(f); assert.equal(error, undefined); assert.equal(native.clientsNative.requests.length, 1);
  assert.equal(native.clientsNative.requests[0].kind, 'pnp-digest');
  native.clientsNative.requests[0].resolve(json({ FormDigestValue: 'fixture-pnp-operations-digest', FormDigestTimeoutSeconds: 1800 })); await drain(f);
  assert.equal(native.clientsNative.requests.length, 2); const envelope = native.clientsNative.requests[1];
  assert.equal(envelope.kind, 'pnp-batch'); assert.equal(envelope.settings.method, 'POST');
  assert.equal(new Headers(envelope.settings.headers).get('X-RequestDigest'), 'fixture-pnp-operations-digest');
  assert.equal([...envelope.settings.body.matchAll(/^GET /gm)].length, 3);
  envelope.resolve(batchReply()); await pending; assert.equal(error, undefined);
  assert.deepEqual(result, [[{ Title: 'Operations tasks' }, { Title: 'Operations news' }], { Title: 'Operations reader' }, { Title: 'Operations web' }]);
  assert.equal(window.pnpReadFixture.calls.length, 0, 'read batch never consumes pilot ledger');
});
async function startOperation(f, mode) {
  requestMode(f, mode); const before = f.calls.length; f.click('Run'); await drain(f);
  if (f.calls[before]?.url.endsWith('/_api/contextinfo')) await finish(f, before, json({ FormDigestValue: 'fixture-pnp-operations-digest', FormDigestTimeoutSeconds: 1800 }));
  assert.equal(f.calls.length, before + (f.calls[before]?.url.endsWith('/_api/contextinfo') ? 2 : 1));
  return f.calls.length - 1;
}
const operationResult = (mode, title) => mode === 'invoke' ? json({ value: [{ Title: title, ItemCount: 17 }] }) : batchReply();
const operationFailure = mode => mode === 'invoke' ? denied('Old invoke denied') : batchReply(403);
for (const firstMode of ['invoke', 'batch']) for (const stale of ['success', 'rejection']) for (const order of ['pending', 'completed'])
  test(`operations ${firstMode} departure to other mode masks old ${stale} after newer ${order}`, async t => {
    const f = operationsFixture(t); const old = await startOperation(f, firstMode); const oldRun = handler(f, 'Run');
    const secondMode = firstMode === 'invoke' ? 'batch' : 'invoke'; const current = await startOperation(f, secondMode);
    const nextRun = handler(f, 'Run');
    if (order === 'completed') await finish(f, current, operationResult(secondMode, 'Current invoke'));
    await finish(f, old, stale === 'success' ? operationResult(firstMode, 'Obsolete invoke') : operationFailure(firstMode));
    assert.ok(!f.container.textContent.includes('Obsolete') && !f.container.textContent.includes('Old invoke denied') && !f.container.textContent.includes('Inner current user denied'));
    if (order === 'pending') {
      assert.match(f.container.textContent, /Loading selected request/);
      const count = f.calls.length; await f.flush(() => { oldRun(); nextRun(); nextRun(); }); await drain(f); assert.equal(f.calls.length, count, 'old finally cannot unlock newer command');
      await finish(f, current, operationResult(secondMode, 'Current invoke'));
    }
    assert.match(f.container.textContent, secondMode === 'invoke' ? /Current invoke/ : /Operations reader/);
  });
test('operations synchronous duplicate Run dispatches once and reset retains Demo while invalidating old actions', async t => {
  const f = operationsFixture(t, { revision: 29 }); const old = handler(f, 'Run');
  await f.flush(() => { old(); old(); old(); }); await drain(f); assert.equal(f.calls.length, 1);
  const selector = f.container.querySelector('[aria-label="Request"]'); f.reset(); assert.equal(f.container.querySelector('[aria-label="Request"]'), selector);
  await f.flush(() => old()); await drain(f); assert.equal(f.calls.length, 1);
  await finish(f, 0, denied('Reset obsolete')); assert.ok(!f.container.textContent.includes('Reset obsolete'));
  const next = await startOperation(f, 'invoke'); await finish(f, next, json({ value: [] }));
  assert.match(f.container.textContent, /No lists returned/); assert.match(f.container.textContent, /Selected request completed/);
});
test('operations mode edits retire retained Run and input immediately, no-op edits remain usable and ABA stays retired', async t => {
  const f = operationsFixture(t); const original = handler(f, 'Run');
  const change = authored(f.regions.controls, props => props['aria-label'] === 'Request').onChange;
  await f.flush(() => { change({}, { value: 'invoke' }); original(); }); await drain(f); assert.equal(f.calls.length, 1);
  await finish(f, 0, denied('Earlier denial')); const dismiss = handler(f, 'Dismiss'), run = handler(f, 'Run');
  await f.flush(() => { change({}, { value: 'batch' }); run(); dismiss(); change({}, { value: 'invoke' }); }); await drain(f);
  assert.equal(f.calls.length, 1); assert.equal(f.container.querySelector('[aria-label="Request"]').value, 'batch', 'retained input cannot retarget after retirement');
  requestMode(f, 'invoke'); await f.flush(() => { original(); run(); dismiss(); }); await drain(f); assert.equal(f.calls.length, 1);
  f.click('Run'); await drain(f); await finish(f, 1, denied('Current denial'));
  await f.flush(() => dismiss()); assert.match(f.container.textContent, /Current denial/);
  f.click('Dismiss'); assert.ok(!f.container.textContent.includes('Request error:'));
});
for (const change of ['client', 'reset', 'initialization', 'departure', 'unmount']) for (const stale of ['success', 'rejection'])
  test(`operations ${change} consumes stale ${stale} and denies retained Run`, async t => {
    const f = operationsFixture(t); const oldRun = handler(f, 'Run'); const old = await startOperation(f, 'invoke');
    if (change === 'client') f.replace(f.host(f.instance.instanceId));
    if (change === 'reset') f.reset();
    if (change === 'initialization') { f.setPnPInitializationError(new Error('Controlled operations construction fault')); f.replace(f.host(f.instance.instanceId)); }
    if (change === 'departure') f.select('Inspect the current native host identity', 'packages/spfx-react-toolkit/src/core/context.internal.tsx#useSPFxContext');
    if (change === 'unmount') f.unmount();
    await f.flush(() => oldRun()); await drain(f); assert.equal(f.calls.length, 1);
    let next;
    if (change === 'client' || change === 'reset') next = await startOperation(f, 'invoke');
    await finish(f, old, stale === 'success' ? json({ value: [{ Title: 'Obsolete client', ItemCount: 2 }] }) : denied('Obsolete failure'));
    assert.ok(!f.container.textContent.includes('Obsolete'));
    if (next !== undefined) { assert.match(f.container.textContent, /Loading selected request/); await finish(f, next, json({ value: [{ Title: 'New client', ItemCount: 17 }] })); assert.match(f.container.textContent, /New client/); }
    if (change === 'initialization') {
      assert.match(f.container.textContent, /Context initialization error: Controlled operations construction fault/); f.reset(); assert.match(f.container.textContent, /Context initialization error/);
      f.setPnPInitializationError(undefined); f.replace(f.host(f.instance.instanceId)); const retry = await startOperation(f, 'invoke'); await finish(f, retry, json({ value: [{ Title: 'Recovered operations', ItemCount: 17 }] })); assert.match(f.container.textContent, /Recovered operations/);
    }
  });
test('operations combined inner read failure is distinct from successful results and retries without another digest', async t => {
  const f = operationsFixture(t); const first = await startOperation(f, 'batch'); await finish(f, first, batchReply(403));
  assert.match(f.container.textContent, /Request error:.*Inner current user denied/); assert.ok(!f.container.querySelector('[aria-label="Batch results"]'));
  const dismiss = handler(f, 'Dismiss'); const second = await startOperation(f, 'batch'); assert.equal(f.calls.filter(r => r.url.endsWith('/contextinfo')).length, 1);
  await finish(f, second, batchReply(403)); await f.flush(() => dismiss()); assert.match(f.container.textContent, /Request error:/);
  f.click('Dismiss'); assert.ok(!f.container.textContent.includes('Request error:')); const retry = await startOperation(f, 'batch'); await finish(f, retry, batchReply()); assert.match(f.container.textContent, /Operations reader/);
});
test('operations outer nonretryable batch rejection settles and a later command can retry', async t => {
  const f = operationsFixture(t); const first = await startOperation(f, 'batch'); await finish(f, first, new Response('Outer bad request', { status: 400, statusText: 'Bad Request' }));
  assert.match(f.container.textContent, /Request error:.*Outer bad request/); assert.equal(f.calls.length, 2);
  const next = await startOperation(f, 'invoke'); await finish(f, next, json({ value: [{ Title: 'After envelope failure', ItemCount: 17 }] })); assert.match(f.container.textContent, /After envelope failure/);
});
test('operations terminal network rejection follows installed three-attempt retry and publishes only current failure', async t => {
  const f = operationsFixture(t); f.click('Run'); await drain(f);
  for (let index = 0; index < 3; index++) { assert.equal(f.calls.length, index + 1); await f.flush(() => f.calls[index].reject(new Error('Offline fixture'))); await drain(f); }
  assert.equal(f.calls.length, 3); assert.match(f.container.textContent, /Request error: Offline fixture/);
  const next = await startOperation(f, 'invoke'); await finish(f, next, json({ value: [] })); assert.match(f.container.textContent, /No lists returned/);
});
test('operations actual providers stay isolated and portable fallback runs actual invoke', async t => {
  const f = operationsFixture(t, { directProvider: true, fallback: true }); const peer = f.mountPeer(); t.after(peer.unmount);
  f.click('Run'); peer.click('Run'); await drain(f); assert.deepEqual(f.calls.map(r => new URL(r.url).pathname).sort(), ['/sites/clients-one/_api/web/lists', '/sites/clients-peer/_api/web/lists']);
  await finish(f, 1, json({ value: [{ Title: 'Peer operations', ItemCount: 73 }] })); await finish(f, 0, json({ value: [{ Title: 'Primary operations', ItemCount: 42 }] }));
  assert.match(f.container.textContent, /Primary operations.*42 items/); assert.ok(!f.container.textContent.includes('Peer operations'));
  assert.match(peer.container.textContent, /Peer operations.*73 items/); assert.ok(!peer.container.textContent.includes('Primary operations'));
});
test('operations real Explorer selection, stable Reset and departure do no automatic requests', async t => {
  const { Simulate } = require('react-dom/test-utils'); const f = operationsFixture(t, { explorer: true }); await drain(f);
  const select = name => { const button = [...f.container.querySelectorAll('[data-example-feature]')].find(n => n.textContent === name); assert.ok(button); f.h.act(() => Simulate.click(button)); };
  select('useSPFxPnP'); await drain(f); assert.equal(f.calls.length, 0); const run = handler(f, 'Run'); f.click('Run'); await drain(f);
  const selector = f.container.querySelector('[aria-label="Request"]'); f.click('Ripristina scenario'); assert.equal(f.container.querySelector('[aria-label="Request"]'), selector);
  await f.flush(() => run()); await drain(f); assert.equal(f.calls.length, 1);
  select('useSPFxContext'); await finish(f, 0, denied('Departed operation')); assert.ok(!f.container.textContent.includes('Departed operation'));
});
function operationsNegativeFixture(t, transform) {
  const fs = require('node:fs'), os = require('node:os'); const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
  const selected = scenario(operationsTitle, operationsOwner), read = fs.readFileSync;
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'pnp-operations-negative-')); const copy = path.join(directory, 'scenario.tsx');
  const altered = transform(selected.source); assert.notEqual(altered, selected.source); fs.writeFileSync(copy, altered);
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const mocked = t.mock.method(fs, 'readFileSync', (file, ...args) => read(String(file) === selected.file ? copy : file, ...args));
  try { return operationsFixture(t); } finally { mocked.mock.restore(); }
}
test('operations negative control detects obsolete token admission after reset', async t => {
  const f = operationsNegativeFixture(t, source => source.replace('mounted.current && boundary.current.token === token', 'mounted.current'));
  const run = handler(f, 'Run'); f.reset(); await f.flush(() => run()); await drain(f);
  assert.throws(() => assert.equal(f.calls.length, 0, 'obsolete command must not dispatch'), /obsolete command must not dispatch/);
  await finish(f, 0, json({ value: [] }));
});
test('operations negative control detects duplicate dispatch when synchronous lock is bypassed', async t => {
  const f = operationsNegativeFixture(t, source => source.replace('dispatch.current?.token === token', 'false'));
  const run = handler(f, 'Run'); await f.flush(() => { run(); run(); }); await drain(f);
  assert.throws(() => assert.equal(f.calls.length, 1, 'one accepted dispatch'), /one accepted dispatch/);
  await finish(f, 0, json({ value: [] })); await finish(f, 1, json({ value: [] }));
});
test('operations negative control detects missing publication guard after a newer completed mode', async t => {
  const f = operationsNegativeFixture(t, source => source.replaceAll('if (owns()) setOutcome', 'if (true) setOutcome'));
  const old = await startOperation(f, 'invoke'); const current = await startOperation(f, 'batch'); await finish(f, current, batchReply());
  await finish(f, old, json({ value: [{ Title: 'Obsolete publication', ItemCount: 17 }] }));
  assert.throws(() => assert.match(f.container.textContent, /Operations reader/), /Operations reader/);
});
test('operations terminal batch rejection retries envelope three times and settles without claiming child cancellation', async t => {
  const f = operationsFixture(t); const first = await startOperation(f, 'batch'); assert.equal(first, 1);
  for (let index = 1; index <= 3; index++) { assert.equal(f.calls.length, index + 1); await f.flush(() => f.calls[index].reject(new Error('Batch offline fixture'))); await drain(f); }
  assert.equal(f.calls.length, 4); assert.match(f.container.textContent, /Request error: Batch offline fixture/);
  assert.equal(f.calls.filter(request => request.url.endsWith('/contextinfo')).length, 1);
});
test('operations native presentation fallback parses finite read responses and rejects mutation envelopes before queuing', async t => {
  const f = operationsFixture(t, { compiled: true }); const descriptor = Object.getOwnPropertyDescriptor(window, 'fetch');
  const native = f.h.load(path.resolve('tests/fixtures/demo-presentation/trim-transports.ts'));
  t.after(() => { if (descriptor) Object.defineProperty(window, 'fetch', descriptor); else delete window.fetch; }); t.mock.method(global, 'fetch', window.fetch.bind(window));
  const helper = f.h.load(path.resolve('tests/fixtures/demo-presentation/pnp-read-transport.ts'));
  const sp = helper.createSPFxPnPContextService(f.instance.context, f.instance.context.pageContext).createSPFI('https://simulated.invalid/sites/native-fallback');
  const service = f.h.load(path.resolve('packages/spfx-react-toolkit/lib/services/spfx-pnp.service.js')).createSPFxPnPService(sp);
  assert.deepEqual(await service.invoke(client => client.web.lists.select('Title', 'ItemCount').filter('Hidden eq false').top(10)()), [{ Title: 'Operations tasks', ItemCount: 17 }]);
  assert.deepEqual(await service.batch(client => Promise.all([client.web.lists.select('Title').top(5)(), client.web.currentUser.select('Title')(), client.web.select('Title')()])), [[{ Title: 'Operations tasks' }, { Title: 'Operations news' }], { Title: 'Operations reader' }, { Title: 'Operations web' }]);
  const before = native.clientsNative.requests.length;
  await assert.rejects(window.fetch('https://simulated.invalid/sites/native-fallback/_api/$batch', { method: 'POST', headers: { 'Content-Type': 'multipart/mixed; boundary=batch_bad', 'X-RequestDigest': 'fixture' }, body: '--batch_bad\nPOST https://simulated.invalid/sites/native-fallback/_api/web/lists HTTP/1.1\n\n--batch_bad--\n' }), /Unsupported native PnP read batch/);
  await assert.rejects(window.fetch('https://external.invalid/_api/contextinfo', { method: 'POST' }), /Unsupported native PnP fixture route/);
  await assert.rejects(window.fetch('https://simulated.invalid/_api/web/lists', { method: 'POST' }), /Unsupported native PnP fixture route/);
  assert.equal(native.clientsNative.requests.length, before); assert.equal(window.pnpReadFixture.calls.length, 0);
});
test('remaining PnP panel default constructs real selectors and services under actual provider with zero native traffic', async t => {
  const f = operationsFixture(t); const mode = process.env.PNP_CONTEXT_IMPLEMENTATION === 'compiled' ? 'lib' : 'src';
  for (const name of ['useSPFxPnPList', 'useSPFxPnPListById', 'useSPFxPnPListByUrl', 'useSPFxPnPListByPath']) Object.assign(f.toolkit, f.h.load(path.resolve('packages/spfx-react-toolkit', mode, 'hooks', name + (mode === 'src' ? '.ts' : '.js'))));
  const Panel = f.h.load(path.resolve('apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/PnPPanel.tsx')).default;
  const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
  const view = f.h.mountComponent(() => f.h.React.createElement(FluentProvider, { theme: webLightTheme }, f.h.React.createElement(f.toolkit.SPFxWebPartProvider, { instance: f.host('remaining-pnp') }, f.h.React.createElement(Panel))));
  t.after(view.unmount); await drain(f);
  assert.equal(view.container.querySelector('select[data-demo-variant="PnPPanel"]').value, 'listAccess');
  assert.equal([...view.container.querySelectorAll('button')].find(node => node.textContent === 'Query by GUID').disabled, true);
  assert.equal(view.container.querySelector('input').value, ''); assert.equal(f.calls.length, 0);
});

test('operations diagnostic plumbing exposes a labeled synthetic source-owned probe to an outside observer', t => {
  // Plumbing proof only: this is not a real source-owned unmount fault or tenant diagnostic.
  const message = "Synthetic diagnostic plumbing probe: Can't perform a React state update on an unmounted component";
  const observed = [];
  const error = console.error;
  t.mock.method(console, 'error', (...args) => {
    if (args[0] === message) observed.push(args);
    else error(...args);
  });
  // The observer must sit outside the fixture so a fixture interceptor cannot conceal this probe.
  const f = operationsFixture(t);
  const stack = `Synthetic source-owned Demo stack for ${f.entry.id}`;
  console.error(message, stack);
  assert.deepEqual(observed, [[message, stack]], 'outside observer must receive the labeled probe');
});

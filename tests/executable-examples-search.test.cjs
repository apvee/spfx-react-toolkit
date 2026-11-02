const assert = require('node:assert/strict');
const { test } = require('node:test');
const path = require('node:path');
const { createClientsFixture } = require('./fixtures/executable-examples/clients-scenarios.cjs');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const compiled = process.env.PNP_SEARCH_IMPLEMENTATION === 'compiled';
const owner = name => `packages/spfx-react-toolkit/src/hooks/useSPFxPnPSearch.ts#${name}`;
const titles = { text: 'Search SharePoint text with explicit execution', vertical: 'Search selected verticals with builder paging' };
const json = value => new Response(JSON.stringify(value), { headers: { 'Content-Type': 'application/json' } });
function page(title = 'Native search row', count = 1, total = count) { return { PrimaryQueryResult: { RelevantResults: { RowCount: count, TotalRows: total, TotalRowsIncludingDuplicates: total, Table: { Rows: Array.from({ length: count }, (_, i) => ({ Cells: [{ Key: 'DocId', Value: title + i }, { Key: 'Title', Value: title }, { Key: 'Path', Value: 'https://tenant.test/result' + i }, { Key: 'Author', Value: 'Ada' }, { Key: 'FileType', Value: 'pdf' }] })) } } } }; }
async function drain(f) { await f.flush(); await new Promise(resolve => setTimeout(resolve, 25)); await f.flush(); }
async function finish(f, index, value) { await f.flush(() => f.calls[index].resolve(value)); await drain(f); }
function props(value, predicate) {
  if (Array.isArray(value)) { for (const child of value) { const found = props(child, predicate); if (found) return found; } return; }
  if (!value || typeof value !== 'object') return;
  if (predicate(value.props ?? {})) return value.props;
  return props(value.props?.children, predicate);
}
function fixture(t, kind = 'text', options = {}) {
  const prior = Object.getOwnPropertyDescriptor(window, 'ResizeObserver');
  Object.defineProperty(window, 'ResizeObserver', { configurable: true, value: class { observe() {} unobserve() {} disconnect() {} } });
  t.after(() => { if (prior) Object.defineProperty(window, 'ResizeObserver', prior); else delete window.ResizeObserver; });
  const f = createClientsFixture(t, titles[kind], { ...options, compiled, pnpContext: true, pnpSearch: true, ownerKey: owner(kind === 'text' ? 'useSPFxPnPSearch' : 'SearchVerticals') });
  f.props = predicate => { for (const region of Object.values(f.regions ?? {})) { const found = props(region, predicate); if (found) return found; } };
  f.action = name => { const found = f.props(p => p.children === name && typeof p.onClick === 'function'); assert.ok(found, name); return found.onClick; };
  f.edit = value => { const found = f.props(p => p['aria-label'] === 'Search SharePoint...'); f.h.act(() => found.onChange({}, { value })); };
  return f;
}
async function queryIndex(f, index = 0) { await drain(f); if (f.calls[index]?.url.endsWith('/_api/contextinfo')) { await finish(f, index++, json({ FormDigestValue: 'search-digest', FormDigestTimeoutSeconds: 1800 })); } assert.ok(f.calls[index]?.url.endsWith('/_api/search/postquery')); return index; }
function request(f, index) { return JSON.parse(f.calls[index].init.body).request; }
test('native Search mode admits exact postquery and rejects Suggest and writes', async t => {
  const f = createClientsFixture(t, 'Read current or cross-site web information', { pnpContext: true, compiled, ownerKey: 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPContext.ts#useSPFxPnPContext' });
  const descriptor = Object.getOwnPropertyDescriptor(window, 'fetch');
  const native = f.h.load(path.resolve('tests/fixtures/demo-presentation/trim-transports.ts'));
  t.after(() => { if (descriptor) Object.defineProperty(window, 'fetch', descriptor); else delete window.fetch; });
  native.clientsNative.configure({ pnpSearch: true });
  const body = JSON.stringify({ request: { Querytext: 'training', RowLimit: 10, HitHighlightedProperties: [], Properties: [], RefinementFilters: [], ReorderingRules: [], SelectProperties: [], SortList: [] } });
  const response = await window.fetch('https://simulated.invalid/sites/demo/_api/search/postquery', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-RequestDigest': 'known' }, body });
  assert.ok(response instanceof Response);
  assert.equal(native.clientsNative.requests.at(-1).kind, 'pnp-search');
  await assert.rejects(window.fetch("https://simulated.invalid/sites/demo/_api/search/suggest?querytext='tr'"), /Unsupported native PnP/);
  await assert.rejects(window.fetch("https://simulated.invalid/sites/demo/_api/web/lists/getByTitle('Tasks')/items", { method: 'POST', body: '{"Title":"Forbidden"}' }), /Unsupported native PnP/);
  assert.equal(native.isPnPCrudEnabled(), false);
});
for (const [kind, name] of [['text', 'useSPFxPnPSearch'], ['vertical', 'SearchVerticals']]) test('Search availability ' + kind, () => assert.equal(scenario(titles[kind], owner(name)).entry.metadata.kind, 'read'));
test('checkpoint emitted text reaches actual provider PnP POST/parser', async t => {
  const f = fixture(t); await drain(f); assert.equal(f.calls.length, 0);
  f.change('Search SharePoint...', ' training '); assert.equal(f.calls.length, 0); f.click('Search');
  const i = await queryIndex(f);
  assert.deepEqual(request(f, i), { Querytext: ' training ', RowLimit: 10, HitHighlightedProperties: [], Properties: [], RefinementFilters: [], ReorderingRules: [], SelectProperties: [], SortList: [] });
  await finish(f, i, json(page())); assert.match(f.container.textContent, /Native search row/);
});
test('checkpoint emitted vertical reaches actual GUID sort page5 append', async t => {
  const f = fixture(t, 'vertical'); await drain(f); assert.equal(f.calls.length, 0); f.click('Documents'); assert.equal(f.calls.length, 0); f.click('Search All');
  const i = await queryIndex(f);
  assert.deepEqual(request(f, i), { Querytext: '*', SourceId: '8413cd39-2156-4e00-b54d-11efd9abdb89', RowLimit: 5, HitHighlightedProperties: [], Properties: [], RefinementFilters: [], ReorderingRules: [], SelectProperties: ['Title', 'Path', 'Author', 'LastModifiedTime', 'FileType'], SortList: [{ Property: 'LastModifiedTime', Direction: 1 }] });
  await finish(f, i, json(page('First', 5, 7))); assert.match(f.container.textContent, /Found 7 results in "Documents"/); f.click('Load More'); const next = await queryIndex(f, i + 1); assert.equal(request(f, next).StartRow, 5); await finish(f, next, json(page('Second', 2, 7))); assert.match(f.container.textContent, /First/); assert.match(f.container.textContent, /Second/); assert.ok(![...f.container.querySelectorAll('button')].some(b => b.textContent === 'Load More'));
});
for (const kind of ['text', 'vertical']) test('actual matching-mode origins and reset/fallback silence ' + kind, async t => {
  const f = fixture(t, kind, { revision: 37, fallback: true }); await drain(f);
  assert.equal(f.calls.length + f.initializations.length + f.tokens.length, 0);
  const mode = compiled ? 'lib' : 'src', other = compiled ? 'src' : 'lib';
  for (const stem of ['hooks/useSPFxPnPSearch', 'services/spfx-pnp-search.service', 'services/spfx-pnp-context.service']) {
    assert.ok(f.pnpFiles.includes(path.resolve('packages/spfx-react-toolkit', mode, stem + (compiled ? '.js' : '.ts'))), stem);
    assert.ok(!f.pnpFiles.some(file => file.includes('/' + other + '/' + stem + '.')));
  }
  assert.equal(f.toolkit.useSPFxPnPSearch, f.h.load(path.resolve('packages/spfx-react-toolkit', mode, 'hooks/useSPFxPnPSearch' + (compiled ? '.js' : '.ts'))).useSPFxPnPSearch);
  f.reset(); await drain(f); assert.equal(f.calls.length, 0);
});
test('text no-op draft and Enter issue one exact current query, preserve native input and render complete result states', async t => {
  const f = fixture(t); await drain(f); const input = f.container.querySelector('input'); input.focus();
  f.edit('   '); const enter = f.props(p => p['aria-label'] === 'Search SharePoint...').onKeyDown;
  f.h.act(() => { enter({ key: 'Enter' }); }); await drain(f); assert.equal(f.calls.length, 0);
  f.edit('report'); const run = f.action('Search'); f.edit('report');
  f.h.act(() => { void run(); void run(); f.props(p => p['aria-label'] === 'Search SharePoint...').onKeyDown({ key: 'Enter' }); });
  const i = await queryIndex(f); assert.equal(f.calls.filter(c => c.url.endsWith('/postquery')).length, 1);
  const response = page('', 1, 1); response.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells = [{ Key: 'DocId', Value: 'blank' }, { Key: 'Path', Value: 'https://tenant.test/blank.pdf' }, { Key: 'FileType', Value: 'pdf' }];
  await finish(f, i, json(response)); assert.match(f.container.textContent, /Found 1 result\b/); assert.match(f.container.textContent, /\(No Title\)/); assert.match(f.container.textContent, /Type: pdf/);
  const link = f.container.querySelector('a[href="https://tenant.test/blank.pdf"]'); assert.equal(link.target, '_blank'); assert.equal(link.rel, 'noopener noreferrer');
  assert.equal(f.container.querySelector('input'), input); assert.equal(document.activeElement, input);
  f.edit('empty'); assert.doesNotMatch(f.container.textContent, /Found 1 result/); f.click('Search'); const j = await queryIndex(f, i + 1); await finish(f, j, json(page('unused', 0, 0))); assert.match(f.container.textContent, /No results found\. Try a different query\./);
  f.reset(); await drain(f); assert.equal(f.container.querySelector('input'), input); assert.equal(input.value, ''); assert.doesNotMatch(f.container.textContent, /No results found/);
});
for (const kind of ['text', 'vertical']) for (const settle of ['success', 'failure']) test('reset and fresh command ignore obsolete ' + settle + ' and finally ' + kind, async t => {
  const f = fixture(t, kind); if (kind === 'text') f.edit('old'); f.click(kind === 'text' ? 'Search' : 'Search All'); const old = await queryIndex(f);
  const retained = f.action(kind === 'text' ? 'Search' : 'Search All'); f.reset(); await drain(f);
  f.h.act(() => { void retained(); }); await drain(f); assert.equal(f.calls.length, old + 1);
  if (kind === 'text') f.edit('new'); f.click(kind === 'text' ? 'Search' : 'Search Documents'); const fresh = await queryIndex(f, old + 1);
  if (settle === 'success') await finish(f, old, json(page('Obsolete'))); else { await finish(f, old, new Response('Obsolete error', { status: 403, statusText: 'Forbidden' })); }
  assert.doesNotMatch(f.container.textContent, /Obsolete/); assert.match(f.container.textContent, /Loading/);
  const latest = f.action(kind === 'text' ? 'Search' : 'Search All'); f.h.act(() => { void latest(); }); await drain(f); assert.equal(f.calls.length, fresh + 1, 'old finally must not release newer lock');
  await finish(f, fresh, json(page('Current'))); assert.match(f.container.textContent, /Current/);
});
for (const kind of ['text', 'vertical']) test('same-value no-op and accepted A-B-A retire commands and inputs ' + kind, async t => {
  const f = fixture(t, kind); if (kind === 'text') f.edit('A'); else f.click('Documents');
  const old = f.action(kind === 'text' ? 'Search' : 'Search All'); const edit = kind === 'text' ? f.props(p => p['aria-label'] === 'Search SharePoint...').onChange : f.action('Pages');
  if (kind === 'text') { f.edit('A'); f.edit('B'); f.edit('A'); } else { f.click('Documents'); f.click('Pages'); f.click('Documents'); }
  f.h.act(() => { void old(); if (kind === 'text') edit({}, { value: 'Obsolete input' }); else void edit(); }); await drain(f); assert.equal(f.calls.length, 0);
  if (kind === 'text') assert.equal(f.container.querySelector('input').value, 'A');
  f.click(kind === 'text' ? 'Search' : 'Search All'); const i = await queryIndex(f); await finish(f, i, json(page('Accepted'))); assert.match(f.container.textContent, /Accepted/);
});
for (const [vertical, guid] of [['All', undefined], ['Documents', '8413cd39-2156-4e00-b54d-11efd9abdb89'], ['Pages', '5e34578e-4d68-4783-8c79-1f07d10bed4f'], ['People', 'b09a7990-05ea-4af9-81ef-edfab16c4e31'], ['Videos', '38403c8c-3975-41a8-826e-717f2d41568a']]) test('actual vertical GUID and both query commands ' + vertical, async t => {
  const f = fixture(t, 'vertical'); assert.deepEqual([...f.container.querySelectorAll('button')].map(b => b.textContent).filter(v => ['All', 'Documents', 'Pages', 'People', 'Videos'].includes(v)), ['All', 'Documents', 'Pages', 'People', 'Videos']);
  f.click(vertical); await drain(f); assert.equal(f.calls.length, 0); f.click('Search Documents'); const i = await queryIndex(f); const q = request(f, i);
  assert.equal(q.Querytext, 'filetype:docx OR filetype:pdf'); assert.equal(q.SourceId, guid); assert.equal(Object.hasOwn(q, 'SourceId'), guid !== undefined); assert.equal(q.RowLimit, 5); assert.deepEqual(q.SortList, [{ Property: 'LastModifiedTime', Direction: 1 }]);
  const response = page('Selected');
  if (vertical === 'People') response.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells.push({ Key: 'LastModifiedTime', Value: '2026-01-05T12:00:00Z' });
  await finish(f, i, json(response)); assert.match(f.container.textContent, new RegExp('in "' + vertical + '"'));
  if (vertical === 'People') { assert.match(f.container.textContent, /Ada - [^<]*2026/); assert.doesNotMatch(f.container.querySelector('article').textContent, /N\/A/); }
  else assert.match(f.container.textContent, /Ada - N\/A/);
});
test('paging retry keeps rows, captured source and offset; retired page callbacks cannot repeat offsets', async t => {
  const f = fixture(t, 'vertical'); f.click('People'); f.click('Search Documents'); const i = await queryIndex(f); await finish(f, i, json(page('First', 5, 12)));
  const oldMore = f.action('Load More'); f.click('Load More'); const j = await queryIndex(f, i + 1); await finish(f, j, new Response('Page failed', { status: 403, statusText: 'Forbidden' }));
  assert.match(f.container.textContent, /First/); assert.match(f.container.textContent, /Page failed/); const oldDismiss = f.action('Dismiss'); f.click('Dismiss'); assert.doesNotMatch(f.container.textContent, /Page failed/);
  f.click('Load More'); const k = await queryIndex(f, j + 1); assert.equal(request(f, k).StartRow, 5); assert.equal(request(f, k).SourceId, 'b09a7990-05ea-4af9-81ef-edfab16c4e31'); assert.equal(request(f, k).Querytext, 'filetype:docx OR filetype:pdf');
  await finish(f, k, json(page('Second', 5, 12))); f.h.act(() => { void oldMore(); oldDismiss(); }); await drain(f); assert.equal(f.calls.length, k + 1); f.click('Load More'); const last = await queryIndex(f, k + 1); assert.equal(request(f, last).StartRow, 10); await finish(f, last, json(page('Third', 2, 12)));
});
for (const kind of ['text', 'vertical']) test('actual failed first request is dismissible and retry executes ' + kind, async t => {
  const f = fixture(t, kind); if (kind === 'text') f.edit('retry'); f.click(kind === 'text' ? 'Search' : 'Search All'); const i = await queryIndex(f);
  await finish(f, i, new Response('Denied Search', { status: 403, statusText: 'Forbidden' })); assert.match(f.container.textContent, /403/); const dismiss = f.action('Dismiss'); f.click('Dismiss'); assert.doesNotMatch(f.container.textContent, /403/);
  f.click(kind === 'text' ? 'Search' : 'Search All'); const j = await queryIndex(f, i + 1); f.h.act(() => dismiss()); assert.match(f.container.textContent, /Loading/); await finish(f, j, json(page('Retried'))); assert.match(f.container.textContent, /Retried/);
});
test('Search native admission is independent default OFF, strict before ledger, and never enables service or CRUD routes', async t => {
  const f = fixture(t), descriptor = Object.getOwnPropertyDescriptor(window, 'fetch');
  const native = f.h.load(path.resolve('tests/fixtures/demo-presentation/trim-transports.ts'));
  t.after(() => { if (descriptor) Object.defineProperty(window, 'fetch', descriptor); else delete window.fetch; });
  const root = 'https://simulated.invalid/sites/search', url = root + '/_api/search/postquery';
  const query = { Querytext: 'training', RowLimit: 10, HitHighlightedProperties: [], Properties: [], RefinementFilters: [], ReorderingRules: [], SelectProperties: [], SortList: [] };
  const init = { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-RequestDigest': 'known' }, body: JSON.stringify({ request: query }) };
  await assert.rejects(window.fetch(url, init), /Unsupported native PnP/);
  native.clientsNative.configure({ pnpServices: true }); await assert.rejects(window.fetch(url, init), /Unsupported native PnP/);
  native.clientsNative.configure({ pnpServices: false, pnpCrud: true }); await assert.rejects(window.fetch(url, init), /Unsupported native PnP/);
  native.clientsNative.configure({ pnpCrud: false, pnpSearch: true });
  const invalid = [
    [url + '?extra=1', init], [url.replace('simulated.invalid', 'external.invalid'), init], [url, { ...init, method: 'GET' }], [url, { ...init, method: 'DELETE' }],
    [url, { ...init, headers: { 'Content-Type': 'application/json' } }], [url, { ...init, body: '{bad' }], [url, { ...init, body: JSON.stringify({ request: { ...query, SourceId: 'simulated-documents' } }) }],
    [url, { ...init, body: JSON.stringify({ request: { ...query, RefinementFilters: ['FileType:pdf'] } }) }], [url, { ...init, body: JSON.stringify({ request: { ...query, RowLimit: 99 } }) }],
    [url, { ...init, body: JSON.stringify({ request: { ...query, Querytext: '*', RowLimit: 5, SourceId: ['8413cd39-2156-4e00-b54d-11efd9abdb89'], SelectProperties: ['Title', 'Path', 'Author', 'LastModifiedTime', 'FileType'], SortList: [{ Property: 'LastModifiedTime', Direction: 1 }] } }) }],
    [url, { ...init, body: JSON.stringify({ request: { ...query, Extra: true } }) }], [url, { ...init, body: JSON.stringify({ request: { ...query, StartRow: 5 } }) }],
    [root + "/_api/search/suggest?querytext='tr'", { method: 'GET' }], [root + '/_api/web?$select=Url', {}], [root + "/_api/web/lists/getByTitle('Tasks')/items?$select=Id,Title&$orderby=Id%20desc&$top=10", {}],
    [root + '/_api/contextinfo?extra=1', { method: 'POST', headers: { 'X-PnPjs-NoDigest': '1' } }],
  ];
  for (const [target, options] of invalid) await assert.rejects(window.fetch(target, options), /Unsupported native PnP/);
  assert.equal(native.clientsNative.requests.length, 0, 'invalid inputs are rejected before native ledger insertion');
  await window.fetch(url, init); assert.equal(native.clientsNative.requests.length, 1);
  native.clientsNative.configure({ pnpSearch: false }); await assert.rejects(window.fetch(url, init), /Unsupported native PnP/); assert.equal(native.clientsNative.requests.length, 1);
});
test('native Search helper forwards before CRUD unknown rejection and leaves the original pilot ledger untouched', async t => {
  const f = fixture(t); const descriptor = Object.getOwnPropertyDescriptor(window, 'fetch');
  const native = f.h.load(path.resolve('tests/fixtures/demo-presentation/trim-transports.ts'));
  t.after(() => { if (descriptor) Object.defineProperty(window, 'fetch', descriptor); else delete window.fetch; });
  native.clientsNative.configure({ pnpSearch: true, pnpCrud: true });
  t.mock.method(global, 'fetch', (input, init) => window.fetch(input, init));
  const helper = f.h.load(path.resolve('tests/fixtures/demo-presentation/pnp-read-transport.ts'));
  assert.equal(f.pnpFactory.createSPFxPnPContextService, f.h.load(path.resolve('packages/spfx-react-toolkit', compiled ? 'lib' : 'src', 'services/spfx-pnp-context.service' + (compiled ? '.js' : '.ts'))).createSPFxPnPContextService);
  const context = { pageContext: { web: { absoluteUrl: 'https://simulated.invalid/sites/search' } } };
  const sp = helper.createSPFxPnPContextService(context, context.pageContext).createSPFI();
  const service = f.h.load(path.resolve('packages/spfx-react-toolkit', compiled ? 'lib' : 'src', 'services/spfx-pnp-search.service' + (compiled ? '.js' : '.ts'))).createSPFxPnPSearchService(sp, { pageSize: 10 });
  const response = await service.search('training'); assert.equal(response.results[0].data.Title, 'Fixture Search result 1'); assert.equal(window.pnpReadFixture.calls.length, 0);
  await assert.rejects(service.suggest('tr'), /Unsupported native PnP Search/); assert.equal(window.pnpReadFixture.calls.length, 0);
});
for (const kind of ['text', 'vertical']) test('same provider context replacement retires pending commands without replacing native controls ' + kind, async t => {
  const f = fixture(t, kind); if (kind === 'text') f.edit('before'); const original = f.container.querySelector(kind === 'text' ? 'input' : 'button');
  f.click(kind === 'text' ? 'Search' : 'Search All'); const i = await queryIndex(f); const retained = f.action(kind === 'text' ? 'Search' : 'Search All');
  f.replace({ ...f.instance, context: { ...f.instance.context } }); await drain(f); assert.equal(f.container.querySelector(kind === 'text' ? 'input' : 'button'), original);
  f.h.act(() => { void retained(); }); await drain(f); assert.equal(f.calls.length, i + 1);
  f.click(kind === 'text' ? 'Search' : 'Search All'); const j = await queryIndex(f, i + 1); await finish(f, j, json(page('New client'))); await finish(f, i, json(page('Old client'))); assert.match(f.container.textContent, /New client/); assert.doesNotMatch(f.container.textContent, /Old client/);
});
for (const kind of ['text', 'vertical']) test('same-cache two providers publish independently and default Search has no installed data cache ' + kind, async t => {
  const f = fixture(t, kind); const peer = f.mountPeer(); if (kind === 'text') { f.edit('primary'); const { Simulate } = require('react-dom/test-utils'); const node = peer.container.querySelector('input'); f.h.act(() => { node.value = 'peer'; Simulate.change(node, { target: { value: 'peer' } }); }); }
  f.click(kind === 'text' ? 'Search' : 'Search All'); const a = await queryIndex(f); peer.click(kind === 'text' ? 'Search' : 'Search All'); const b = await queryIndex(f, a + 1);
  assert.notEqual(f.calls[a].url, f.calls[b].url); await finish(f, b, json(page('Peer winner'))); await finish(f, a, json(page('Primary winner'))); assert.match(peer.container.textContent, /Peer winner/); assert.doesNotMatch(peer.container.textContent, /Primary winner/); assert.match(f.container.textContent, /Primary winner/);
  f.click(kind === 'text' ? 'Search' : 'Search All'); const next = await queryIndex(f, b + 1); await finish(f, next, json(page('Fresh repeated request'))); assert.match(f.container.textContent, /Fresh repeated request/); peer.unmount();
});
for (const kind of ['text', 'vertical']) test('departure and unmount deny retained commands and consume late rejection ' + kind, async t => {
  const f = fixture(t, kind); if (kind === 'text') f.edit('depart'); f.click(kind === 'text' ? 'Search' : 'Search All'); const i = await queryIndex(f); const retained = f.action(kind === 'text' ? 'Search' : 'Search All');
  const other = kind === 'text' ? 'vertical' : 'text'; f.select(titles[other], owner(other === 'text' ? 'useSPFxPnPSearch' : 'SearchVerticals'));
  await finish(f, i, new Response('Retired rejection', { status: 403 })); assert.doesNotMatch(f.container.textContent, /Retired rejection/); f.h.act(() => { void retained(); }); await drain(f); assert.equal(f.calls.length, i + 1);
  if (other === 'text') f.edit('unmount'); f.click(other === 'text' ? 'Search' : 'Search All'); const j = await queryIndex(f, i + 1); const action = f.action(other === 'text' ? 'Search' : 'Search All'); f.unmount(); f.h.act(() => { void action(); }); await finish(f, j, new Response('Unmounted rejection', { status: 403 })); assert.equal(f.calls.length, j + 1);
});
test('remaining legacy Refiners and Suggestions mount and select with actual provider Search/native zero I/O', async t => {
  const f = fixture(t); const Panel = f.h.load(path.resolve('apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/SearchPanel.tsx')).default;
  const peer = f.h.mountComponent(() => f.h.React.createElement(f.toolkit.SPFxWebPartProvider, { instance: f.host('remaining-search') }, f.h.React.createElement(Panel)));
  await drain(f); assert.equal(f.calls.length, 0); const select = peer.container.querySelector('select'); const { Simulate } = require('react-dom/test-utils');
  for (const value of ['refiners', 'suggestions', 'refiners']) { f.h.act(() => { select.value = value; Simulate.change(select, { target: { value } }); }); await drain(f); assert.equal(f.calls.length, 0, value); }
  peer.unmount();
});
for (const kind of ['text', 'vertical']) for (const failure of [false, true]) test('newer target completion wins before obsolete ' + (failure ? 'error ' : 'success ') + kind, async t => {
  const f = fixture(t, kind); if (kind === 'text') f.edit('old'); f.click(kind === 'text' ? 'Search' : 'Search All'); const i = await queryIndex(f);
  if (kind === 'text') f.edit('new'); else f.click('Pages'); f.click(kind === 'text' ? 'Search' : 'Search Documents'); const j = await queryIndex(f, i + 1);
  await finish(f, j, json(page('Newest target'))); await finish(f, i, failure ? new Response('Obsolete target denial', { status: 403 }) : json(page('Obsolete target')));
  assert.match(f.container.textContent, /Newest target/); assert.doesNotMatch(f.container.textContent, /Obsolete target/);
});
for (const change of ['vertical', 'reset']) test('old pending page cannot append into a new ' + change + ' query', async t => {
  const f = fixture(t, 'vertical'); f.click('Search All'); const i = await queryIndex(f); await finish(f, i, json(page('First page', 5, 12))); f.click('Load More'); const j = await queryIndex(f, i + 1);
  if (change === 'reset') f.reset(); else f.click('Documents'); f.click('Search Documents'); const k = await queryIndex(f, j + 1); await finish(f, k, json(page('New baseline'))); await finish(f, j, json(page('Old page')));
  assert.match(f.container.textContent, /New baseline/); assert.doesNotMatch(f.container.textContent, /First page|Old page/);
});
for (const kind of ['text', 'vertical']) test('construction dependency error is separate from explicit request failure and recovers ' + kind, async t => {
  const f = fixture(t, kind); if (kind === 'text') f.edit('ready'); f.setPnPInitializationError(new Error('Controlled construction fault')); f.replace({ ...f.instance, context: { ...f.instance.context } }); await drain(f);
  assert.match(f.container.textContent, /Client initialization error: Controlled construction fault/); assert.equal(f.calls.length, 0);
  f.setPnPInitializationError(undefined); f.replace({ ...f.instance, context: { ...f.instance.context } }); await drain(f); assert.doesNotMatch(f.container.textContent, /Controlled construction fault/);
  f.click(kind === 'text' ? 'Search' : 'Search All'); const i = await queryIndex(f); await finish(f, i, json(page('Recovered context'))); assert.match(f.container.textContent, /Recovered context/);
});
function alteredFixture(t, kind, transform) {
  const fs = require('node:fs'), os = require('node:os'); const selected = scenario(titles[kind], owner(kind === 'text' ? 'useSPFxPnPSearch' : 'SearchVerticals'));
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'search-negative-')), copy = path.join(directory, 'scenario.tsx'); const original = fs.readFileSync;
  const changed = transform(selected.source); assert.notEqual(changed, selected.source); fs.writeFileSync(copy, changed); t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const mock = t.mock.method(fs, 'readFileSync', (file, ...args) => original(String(file) === selected.file ? copy : file, ...args));
  try { return fixture(t, kind); } finally { mock.mock.restore(); }
}
test('negative control detects removed pre-dispatch session guard', async t => {
  const f = alteredFixture(t, 'text', s => s.replace('if (!current() || !context.isInitialized || !query.trim()', 'if (!context.isInitialized || !query.trim()'));
  f.edit('retired'); const old = f.action('Search'); f.reset(); f.h.act(() => { void old(); }); const i = await queryIndex(f);
  assert.throws(() => assert.equal(f.calls.filter(c => c.url.endsWith('/postquery')).length, 0, 'retired action must send zero Search requests'), /retired action/); await finish(f, i, json(page('Consumed faulty dispatch')));
});
test('negative control detects lost draft generation through A-B-A', async t => {
  const f = alteredFixture(t, 'text', s => s.replace(' && session.version === version', ''));
  f.edit('A'); const old = f.action('Search'); f.edit('B'); f.edit('A'); f.h.act(() => { void old(); }); const i = await queryIndex(f);
  assert.throws(() => assert.equal(f.calls.filter(c => c.url.endsWith('/postquery')).length, 0, 'ABA must retire old action'), /ABA/); await finish(f, i, json(page('Consumed faulty ABA')));
});
test('negative control detects obsolete publication hiding a newer result', async t => {
  const f = alteredFixture(t, 'text', s => s.replace('const owns = (): boolean => current() && sequence.current === request;', 'const owns = (): boolean => true;'));
  f.edit('old'); f.click('Search'); const i = await queryIndex(f); f.reset(); f.edit('new'); f.click('Search'); const j = await queryIndex(f, i + 1); await finish(f, j, json(page('Newest guarded result'))); await finish(f, i, json(page('Faulty old publication')));
  assert.throws(() => assert.match(f.container.textContent, /Newest guarded result/, 'newest result must remain visible'), /newest result/);
});
test('negative control detects obsolete finally releasing newer lock', async t => {
  const f = alteredFixture(t, 'text', s => s.replace('if (owns() && dispatch.current?.token === token && dispatch.current.request === request) dispatch.current = undefined;', 'dispatch.current = undefined;'));
  f.edit('old'); f.click('Search'); const i = await queryIndex(f); f.reset(); f.edit('new'); f.click('Search'); const j = await queryIndex(f, i + 1); await finish(f, i, json(page('Obsolete')));
  const retained = f.action('Search'); f.h.act(() => { void retained(); }); const k = await queryIndex(f, j + 1);
  assert.throws(() => assert.equal(f.calls.length, j + 1, 'old finally must not release a newer lock'), /old finally/); await finish(f, j, json(page('Current'))); await finish(f, k, json(page('Consumed duplicate')));
});
test('negative control detects retained paging callback after a later page generation', async t => {
  const f = alteredFixture(t, 'vertical', s => s.replace(' || session.page !== pageVersion', ''));
  f.click('Search All'); const i = await queryIndex(f); await finish(f, i, json(page('First', 5, 20))); const old = f.action('Load More'); f.click('Load More'); const j = await queryIndex(f, i + 1); await finish(f, j, json(page('Second', 5, 20))); f.h.act(() => { void old(); }); const k = await queryIndex(f, j + 1);
  assert.throws(() => assert.equal(f.calls.length, j + 1, 'old page callback must not dispatch'), /old page callback/); await finish(f, k, json(page('Consumed old callback', 1, 20)));
});
test('negative control detects real People GUID substituted for Documents', async t => {
  const f = alteredFixture(t, 'vertical', s => s.replace('Documents: SearchVerticals.Documents', 'Documents: SearchVerticals.People'));
  f.click('Documents'); f.click('Search All'); const i = await queryIndex(f); assert.throws(() => assert.equal(request(f, i).SourceId, '8413cd39-2156-4e00-b54d-11efd9abdb89', 'Documents wire GUID'), /Documents wire GUID/); await finish(f, i, json(page('Consumed wrong source')));
});
for (const [kind, feature] of [['text', 'useSPFxPnPSearch'], ['vertical', 'SearchVerticals']]) test('real Explorer Reset retains mounted controls and Reset reference while pending ' + kind, async t => {
  const f = fixture(t, kind, { explorer: true }); const { Simulate } = require('react-dom/test-utils'); await drain(f);
  const select = name => { const node = [...f.container.querySelectorAll('[data-example-feature]')].find(node => node.textContent === name); assert.ok(node, name); f.h.act(() => Simulate.click(node)); };
  select(feature); await drain(f); assert.equal(f.calls.length, 0);
  const original = kind === 'text' ? f.container.querySelector('input') : [...f.container.querySelectorAll('button')].find(node => node.textContent === 'Documents');
  const reset = [...f.container.querySelectorAll('button')].find(node => node.textContent === 'Ripristina scenario'); assert.ok(reset);
  if (kind === 'text') f.change('Search SharePoint...', 'explorer'); f.click(kind === 'text' ? 'Search' : 'Search All'); const i = await queryIndex(f);
  reset.focus(); f.click('Ripristina scenario'); await drain(f); assert.equal([...f.container.querySelectorAll('button')].find(node => node.textContent === 'Ripristina scenario'), reset);
  assert.equal(kind === 'text' ? f.container.querySelector('input') : [...f.container.querySelectorAll('button')].find(node => node.textContent === 'Documents'), original); assert.equal(document.activeElement, reset);
  await finish(f, i, json(page('Retired Explorer result'))); assert.doesNotMatch(f.container.textContent, /Retired Explorer result/);
  select(kind === 'text' ? 'SearchVerticals' : 'useSPFxPnPSearch'); await drain(f); assert.equal(f.calls.length, i + 1);
});
for (const kind of ['text', 'vertical']) for (const missing of [undefined, '', '   ']) test('absent or blank Path creates no empty interactive result ' + kind + ':' + JSON.stringify(missing), async t => {
  const f = fixture(t, kind); if (kind === 'text') f.edit('missing path'); f.click(kind === 'text' ? 'Search' : 'Search All'); const i = await queryIndex(f);
  const response = page('Pathless result'); response.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells = [{ Key: 'DocId', Value: 'pathless' }, { Key: 'Title', Value: 'Pathless result' }, ...(missing === undefined ? [] : [{ Key: 'Path', Value: missing }])];
  await finish(f, i, json(response)); const article = f.container.querySelector('article'); assert.ok(article); assert.match(article.textContent, /Pathless result/); assert.equal(article.querySelector('a,button') === null, true, 'an absent or blank Path must not create an empty interactive Link');
  if (kind === 'text') assert.doesNotMatch(article.textContent, /Type:/);
});
for (const kind of ['text', 'vertical']) test('nonblank Path preserves its raw href and safe new-tab attributes ' + kind, async t => {
  const f = fixture(t, kind); if (kind === 'text') f.edit('raw path'); f.click(kind === 'text' ? 'Search' : 'Search All'); const i = await queryIndex(f);
  const rawPath = ' https://tenant.test/raw-result.pdf ';
  const response = page('Raw Path result'); response.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells = [{ Key: 'DocId', Value: 'raw-path' }, { Key: 'Title', Value: 'Raw Path result' }, { Key: 'Path', Value: rawPath }];
  await finish(f, i, json(response)); const link = f.container.querySelector('article a'); assert.ok(link); assert.equal(link.getAttribute('href'), rawPath); assert.equal(link.textContent, rawPath); assert.equal(link.target, '_blank'); assert.equal(link.rel, 'noopener noreferrer');
});
for (const kind of ['text', 'vertical']) test('raw target change retires Run and Enter before React commits ' + kind, async t => {
  const f = fixture(t, kind); if (kind === 'text') f.edit('A'); else f.click('Documents');
  const run = f.action(kind === 'text' ? 'Search' : 'Search All');
  const change = kind === 'text' ? f.props(p => p['aria-label'] === 'Search SharePoint...').onChange : f.action('Pages');
  const enter = kind === 'text' ? f.props(p => p['aria-label'] === 'Search SharePoint...').onKeyDown : undefined;
  f.h.act(() => { if (kind === 'text') change({}, { value: 'B' }); else change(); void run(); if (enter) enter({ key: 'Enter' }); });
  await drain(f); assert.equal(f.calls.length, 0, 'retired same-tick commands never reach HTTP');
  f.click(kind === 'text' ? 'Search' : 'Search All'); const i = await queryIndex(f); assert.equal(kind === 'text' ? request(f, i).Querytext : request(f, i).SourceId, kind === 'text' ? 'B' : '5e34578e-4d68-4783-8c79-1f07d10bed4f'); await finish(f, i, json(page('Current atomic target'))); assert.match(f.container.textContent, /Current atomic target/);
});

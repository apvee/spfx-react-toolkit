const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const web = 'https://tenant.test/sites/root';
const sub = web + '/subweb';
const listId = '11111111-1111-4111-8111-111111111111';
const row = (StringId = '0x0101', Name = 'Document', extra = {}) => ({ StringId, Name, Group: 'Documents', Description: 'A description', Hidden: false, ReadOnly: true, Sealed: true, ...extra });
function setup(response, props = {}, origin = sub) {
  const calls = [];
  const h = createHarness({ '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } } });
  const context = { owner: {}, tenantId: listId, userId: 'u', webUrl: origin, siteCollectionUrl: web,
    spClient: { get: async uri => { calls.push(uri); const data = await response(uri); return { ok: !data?.status, status: data?.status ?? 200, json: async () => data }; } } };
  const cache = h.load('controls/pickers/picker-request-cache.internal.ts').createPickerRequestCache(context.owner);
  const create = h.load('controls/pickers/content-type-picker.source.internal.ts').createContentTypePickerSource;
  const source = create(props, context);
  const request = { query: '', pageSize: 30, generation: 0, context, cache, partition: 'ct-test:' };
  return { h, context, cache, request, source, create, calls };
}
const response = uri => uri.includes('/_api/site/rootWeb?') ? { Url: web } : uri.includes('/contenttypes') ? uri.includes('getById') ? row() : { value: [row()] } : { Id: listId, Title: 'Documents', BaseTemplate: 101 };
test('site content types resolve a subweb to its owning collection root without guessing remote paths', async () => {
  const remote = 'https://tenant.test/teams/remote/sub';
  const s = setup(response, { webUrl: remote });
  const page = await s.source.discover(s.request);
  assert.equal(page.options[0].value.scope, 'site'); assert.equal(page.options[0].value.webUrl, web);
  assert.ok(s.calls[0].startsWith(remote + '/_api/site/rootWeb?'));
  assert.ok(s.calls[1].startsWith(web + '/_api/web/contenttypes?'));
  assert.equal(s.source.completeCatalog, true);
  assert.deepEqual(page.options[0].readOnly, true); assert.deepEqual(page.options[0].sealed, true);
  const candidate = { ...page.options[0], key: 'saved-root' };
  assert.equal(s.source.accepts(candidate), true);
  assert.equal((await s.source.verifyCandidate(candidate, s.request)).key, 'saved-root');
});
test('list scope enumerates the exact list in its owning subweb and never requests the site catalog', async () => {
  const s = setup(response, { scope: 'list', list: 'Documents' });
  const option = (await s.source.discover(s.request)).options[0];
  assert.equal(option.value.webUrl, sub); assert.equal(option.value.listId, listId); assert.equal(option.value.scope, 'list');
  assert.ok(s.calls.every(uri => uri.startsWith(sub + '/_api/web/')));
  assert.ok(s.calls.some(uri => uri.includes(`/lists(guid'${listId}')/contenttypes`)));
});
test('complete catalog follows all pages, sorts Name then ID and performs case-insensitive local prefix search', async () => {
  const s = setup(uri => uri.includes('/rootWeb?') ? { Url: web } : new URL(uri).searchParams.has('$skiptoken') ?
    { d: { results: [row('0x0101AB', 'Alpha'), row('0x0101AC', 'Alpha')] } } :
    { value: [row('0x0101', 'Zebra')], '@odata.nextLink': web + '/_api/web/contenttypes?$skiptoken=next' });
  const [a, b] = await Promise.all([s.source.discover(s.request), s.create({}, s.context).discover({ ...s.request, query: 'AL' })]);
  assert.deepEqual(a.options.map(o => o.name), ['Alpha', 'Alpha', 'Zebra']); assert.equal(b.options.length, 2);
  assert.equal(a.completeness, 'complete'); assert.equal(a.continuation, undefined);
  const count = s.calls.length; await s.source.discover({ ...s.request, query: 'al' }); assert.equal(s.calls.length, count);
});
test('group, hidden and hierarchical ID filters apply to discovery, resolution and candidate verification', async () => {
  const data = [row(), row('0x0101AB', 'Child'), row('0x0101AC', 'Hidden', { Hidden: true }), row('0x0102', 'Other'), row('0x0101AD', 'Foreign', { Group: 'Other' })];
  const s = setup(uri => uri.includes('/rootWeb?') ? { Url: web } : uri.includes('getById') ? data.find(item => decodeURIComponent(uri).toLowerCase().includes("'" + item.StringId.toLowerCase() + "'")) : { value: data }, { group: 'dOcUmEnTs', includeHidden: false, contentTypeId: '0X0101', includeDescendants: true });
  const page = await s.source.discover(s.request); assert.deepEqual(page.options.map(o => o.id), ['0x0101AB', '0x0101']);
  const base = page.options[0].value;
  for (const id of ['0x0101AC', '0x0101AD', '0x0102']) {
    const result = await s.source.resolve([{ ...base, contentTypeId: id }], s.request); assert.equal(result.resolved.length, 0, id);
  }
  const exact = s.create({ contentTypeId: '0x0101' }, s.context);
  assert.deepEqual((await exact.discover(s.request)).options.map(o => o.id), ['0x0101']);
  assert.throws(() => s.create({ includeDescendants: true }, s.context), /contentTypeId/);
});
test('resolution rejects references from other sites, scope or lists and preserves valid aliases', async () => {
  const s = setup(response, { scope: 'list', list: { kind: 'id', id: listId } });
  const option = (await s.source.discover(s.request)).options[0]; const value = option.value;
  const invalid = [{ ...value, listId: '22222222-2222-4222-8222-222222222222' }, { ...value, webUrl: web }, { kind: 'contentType', scope: 'site', webUrl: web, contentTypeId: value.contentTypeId }];
  const result = await s.source.resolve(invalid, s.request); assert.equal(result.resolved.length, 0); assert.equal(result.unresolved.length, 3);
  const hydrated = await s.source.hydrate([{ kind: 'value', key: 'legacy-alias', value }], s.request); assert.equal(hydrated.resolved[0].option.key, 'legacy-alias');
  const site = setup(response, { webUrl: sub });
  assert.equal((await site.source.resolve([{ kind: 'contentType', scope: 'site', webUrl: web, contentTypeId: '0x0101' }], site.request)).resolved.length, 1);
  assert.equal((await site.source.resolve([{ kind: 'contentType', scope: 'site', webUrl: web + '/other', contentTypeId: '0x0101' }], site.request)).resolved.length, 0);
});
test('malformed metadata, REST failures and paging cycles reject the complete catalog', async () => {
  for (const data of [{ status: 403 }, { value: [row('0x0100AB')] }, { value: [{ ...row(), Sealed: undefined }] }]) {
    const s = setup(uri => uri.includes('/rootWeb?') ? { Url: web } : data);
    await assert.rejects(s.source.discover(s.request), /403|Invalid|boolean/);
  }
  const s = setup(uri => uri.includes('/rootWeb?') ? { Url: web } : { value: [row()], '@odata.nextLink': web + '/_api/web/contenttypes?$skiptoken=cycle' });
  await assert.rejects(s.source.discover(s.request), /Repeated/);
  const intermediate = setup(uri => uri.includes('/rootWeb?') ? { Url: web } : new URL(uri).searchParams.has('$skiptoken') ? { status: 403 } : { value: [row()], '__next': web + '/_api/web/contenttypes?$skiptoken=next' });
  await assert.rejects(intermediate.source.discover(intermediate.request), /403/);
});
test('catalog and entity reads coalesce, refresh and owner/client changes reload dependencies', async () => {
  const s = setup(response);
  const option = (await s.source.discover(s.request)).options[0]; const count = s.calls.length;
  await s.source.verifyCandidate(option, s.request); assert.equal(s.calls.length, count);
  const ids = [option.value, { ...option.value, contentTypeId: option.value.contentTypeId.toUpperCase() }];
  await Promise.all(ids.map(value => s.source.resolve([value], s.request))); assert.equal(s.calls.length, count + 1);
  s.cache.invalidate('ct-test:'); await s.source.discover(s.request); assert.equal(s.calls.length, count * 2 + 1);
  for (const change of [{ owner: {} }, { spClient: { ...s.context.spClient } }]) {
    const context = { ...s.context, ...change }; const before = s.calls.length;
    await s.create({}, context).discover({ ...s.request, context }); assert.equal(s.calls.length, before + 2);
  }
});

test('System root and empty legacy descriptions map without inventing metadata', async () => {
  const s = setup(uri => uri.includes('/rootWeb?') ? { Url: web } : { value: [row('0x', 'System', { Description: null })] });
  const option = (await s.source.discover(s.request)).options[0];
  assert.equal(option.value.contentTypeId, '0x'); assert.equal(option.description, undefined);
});
test('removed content types resolve terminally without a fabricated transport error', async () => {
  const s = setup(uri => uri.includes('/rootWeb?') ? { Url: web } : { status: 404 });
  const result = await s.source.resolve([{ kind: 'contentType', scope: 'site', webUrl: web, contentTypeId: '0x0101' }], s.request);
  assert.equal(result.resolved.length, 0); assert.equal(result.unresolved[0].error, undefined);
});

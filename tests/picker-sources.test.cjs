const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const web = 'https://tenant.test/sites/hr';
const guid = '11111111-1111-4111-8111-111111111111';
function setup(response) {
  const calls = [];
  const h = createHarness({ '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } } });
  const context = { owner: {}, tenantId: guid, userId: 'u', webUrl: web, siteCollectionUrl: web, spClient: { get: async (url, config, options) => { calls.push({ url, options }); return { ok: true, status: 200, json: async () => typeof response === 'function' ? response(url) : response }; } } };
  const cache = h.load('controls/pickers/picker-request-cache.internal.ts').createPickerRequestCache(context.owner);
  return { h, context, calls, request: { query: 'Pr', pageSize: 50, generation: 0, context, cache } };
}
test('metadata hydration and resolution share identical entity reads while retaining renderer aliases', async () => {
  const row = { Id: guid, InternalName: 'Title', Title: 'Projects', TypeAsString: 'Text', FieldTypeKind: 2, Hidden: false, ReadOnlyField: false, Group: 'Demo', BaseTemplate: 100, DefaultView: false, PersonalView: false };
  for (const domain of ['List', 'Field', 'View']) {
    const state = setup(row);
    const value = domain === 'List' ? { kind: 'list', webUrl: web, listId: guid } : domain === 'Field' ?
      { kind: 'field', webUrl: web, fieldId: guid, internalName: 'Title' } : { kind: 'view', webUrl: web, listId: guid, viewId: guid };
    const create = state.h.load(`controls/pickers/${domain.toLowerCase()}-picker.source.internal.ts`)[`create${domain}PickerSource`];
    const source = create(domain === 'View' ? { list: 'Projects' } : {}, state.context);
    const [hydrated, resolved] = await Promise.all([
      source.hydrate([{ kind: 'value', key: 'consumer-alias', value }], state.request), source.resolve([value], state.request),
    ]);
    assert.equal(hydrated.resolved[0].option.key, 'consumer-alias', domain);
    assert.equal(resolved.resolved[0].option.text, 'Projects', domain);
    assert.equal(state.calls.length, 1, `${domain} duplicate entity GET`);
  }
});
test('candidate verification shares exact filtered reads without reusing unconstrained entity hydration', async () => {
  const row = { Id: guid, InternalName: 'Title', Title: 'Projects', TypeAsString: 'Text', FieldTypeKind: 2, Hidden: false, ReadOnlyField: false, Group: 'Demo', BaseTemplate: 100, DefaultView: false, PersonalView: false };
  for (const domain of ['List', 'Field', 'View']) {
    const state = setup(uri => new URL(uri).searchParams.has('$top') ? { value: [row] } : row);
    const value = domain === 'List' ? { kind: 'list', webUrl: web, listId: guid } : domain === 'Field' ?
      { kind: 'field', webUrl: web, fieldId: guid, internalName: 'Title' } : { kind: 'view', webUrl: web, listId: guid, viewId: guid };
    const source = state.h.load(`controls/pickers/${domain.toLowerCase()}-picker.source.internal.ts`)[`create${domain}PickerSource`](domain === 'View' ? { list: 'Projects' } : {}, state.context);
    const candidate = (await source.resolve([value], state.request)).resolved[0].option;
    const [first, second] = await Promise.all([source.verifyCandidate(candidate, state.request), source.verifyCandidate(candidate, state.request)]);
    assert.equal(first.text, 'Projects'); assert.deepEqual(second, first);
    const verification = state.calls.filter(call => new URL(call.url).searchParams.has('$top'));
    assert.equal(verification.length, 1, `${domain} duplicate filtered GET`);
    assert.match(new URL(verification[0].url).searchParams.get('$filter'), /Id eq guid/);
  }
});
test('metadata entity reads isolate context identities and refresh invalidates every dependent source', async () => {
  const row = { Id: guid, InternalName: 'Title', Title: 'Old', TypeAsString: 'Text', FieldTypeKind: 2, Hidden: false, ReadOnlyField: false, Group: 'Demo' };
  let finish;
  const state = setup(() => new Promise(resolve => { finish = resolve; }));
  const create = state.h.load('controls/pickers/field-picker.source.internal.ts').createFieldPickerSource;
  const value = { kind: 'field', webUrl: web, fieldId: guid, internalName: 'Title' };
  const original = { ...state.request, partition: 'first:' };
  const first = create({}, state.context).resolve([value], original);
  const second = create({ includeHidden: false }, state.context).resolve([value], { ...original, partition: 'second:' });
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(state.calls.length, 1, 'same entity URI shares a request across source filters');
  const completeOld = finish;
  original.cache.invalidate('second:');
  const refreshed = create({}, state.context).resolve([value], original);
  await new Promise(resolve => setImmediate(resolve));
  finish({ ...row, Title: 'Fresh' });
  assert.equal((await refreshed).resolved[0].option.text, 'Fresh');
  completeOld(row);
  assert.equal((await first).resolved[0].option.text, 'Old'); await second;
  assert.equal((await create({}, state.context).resolve([value], original)).resolved[0].option.text, 'Fresh');
  for (const change of [{ owner: {} }, { spClient: { ...state.context.spClient } }, { tenantId: 'other' }, { userId: 'other' }]) {
    const context = { ...state.context, ...change };
    const pending = create({}, context).resolve([value], { ...original, context });
    await new Promise(resolve => setImmediate(resolve)); finish(row); await pending;
  }
  assert.equal(state.calls.length, 6, 'refresh plus each identity change has a fresh read');
});
test('current Site suggestions and selection hydration share the same web metadata read', async () => {
  const state = setup(uri => uri.includes('/_api/site?') ? { Id: guid, Url: web, HubSiteId: '00000000-0000-0000-0000-000000000000' } : { Title: 'Current', Id: guid });
  const source = state.h.load('controls/pickers/site-picker.source.internal.ts').createSitePickerSource({ includeHubSite: false }, state.context);
  const [suggestions, resolved] = await Promise.all([source.personalSuggestions(state.request), source.resolve([{ kind: 'site', url: web }], state.request)]);
  assert.equal(suggestions.currentSite[0].text, 'Current'); assert.equal(resolved.resolved[0].option.text, 'Current');
  assert.equal(state.calls.filter(call => call.url.includes('/_api/web?')).length, 1);
  assert.equal(state.calls.filter(call => call.url.includes('/_api/site?')).length, 2, 'distinct metadata projections remain separate');
});
test('failed entity reads are shared until explicit refresh permits recovery', async () => {
  let denied = true;
  const state = setup(() => {
    if (denied) throw new Error('Entity denied');
    return { Id: guid, Title: 'Projects', BaseTemplate: 100, Hidden: false };
  });
  const source = state.h.load('controls/pickers/list-picker.source.internal.ts').createListPickerSource({}, state.context);
  const value = { kind: 'list', webUrl: web, listId: guid };
  const request = { ...state.request, partition: 'entity:' };
  assert.match((await source.resolve([value], request)).unresolved[0].error.message, /Entity denied/);
  denied = false;
  assert.match((await source.resolve([value], request)).unresolved[0].error.message, /Entity denied/);
  assert.equal(state.calls.length, 1);
  request.cache.invalidate('entity:');
  assert.equal((await source.resolve([value], request)).resolved[0].option.text, 'Projects');
  assert.equal(state.calls.length, 2);
});
test('Field/View verification reuses positive rows from a warm complete catalog', async () => {
  const row = { Id: guid, InternalName: 'Title', Title: 'Projects', TypeAsString: 'Text', FieldTypeKind: 2, Hidden: false, ReadOnlyField: false, Group: 'Demo', BaseTemplate: 100, DefaultView: false, PersonalView: false };
  for (const domain of ['Field', 'View']) {
    const state = setup(uri => uri.includes('/fields?') || uri.includes('/views?') ? { value: [row] } : row);
    const create = state.h.load(`controls/pickers/${domain.toLowerCase()}-picker.source.internal.ts`)[`create${domain}PickerSource`];
    const props = domain === 'View' ? { list: 'Projects' } : {};
    const source = create(props, state.context);
    const page = await source.discover({ ...state.request, query: '' });
    const calls = state.calls.length;
    const candidate = { ...page.options[0], key: 'consumer', text: 'Consumer label' };
    const verified = await source.verifyCandidate(candidate, { ...state.request, partition: 'verification:' });
    assert.equal(verified.key, 'consumer'); assert.equal(verified.text, 'Consumer label');
    assert.equal(state.calls.length, calls, `${domain} warm catalog already verifies this row`);
    state.request.cache.invalidate('verification:');
    await source.verifyCandidate(candidate, state.request);
    assert.equal(state.calls.length, calls + 1, 'refresh invalidates the reused catalog dependency');
    assert.match(new URL(state.calls.at(-1).url).searchParams.get('$filter'), /Id eq guid/);
  }
});
test('Field/View verification joins an in-flight complete catalog instead of reading each candidate', async () => {
  const row = { Id: guid, InternalName: 'Title', Title: 'Projects', TypeAsString: 'Text', FieldTypeKind: 2, Hidden: false, ReadOnlyField: false, Group: 'Demo', BaseTemplate: 100, DefaultView: false, PersonalView: false };
  for (const domain of ['Field', 'View']) {
    let finish;
    const state = setup(uri => uri.includes('/fields?') || uri.includes('/views?') ? new Promise(resolve => { finish = resolve; }) : row);
    const source = state.h.load(`controls/pickers/${domain.toLowerCase()}-picker.source.internal.ts`)[`create${domain}PickerSource`](domain === 'View' ? { list: 'Projects' } : {}, state.context);
    const value = domain === 'Field' ? { kind: 'field', webUrl: web, fieldId: guid, internalName: 'Title' } : { kind: 'view', webUrl: web, listId: guid, viewId: guid };
    const candidate = domain === 'Field' ? { key: 'consumer', text: 'Consumer label', value, webUrl: web, id: guid, internalName: 'Title', title: 'Projects', typeAsString: 'Text', fieldTypeKind: 2, hidden: false, readOnly: false, group: 'Demo' } :
      { key: 'consumer', text: 'Consumer label', value, webUrl: web, listId: guid, id: guid, title: 'Projects', defaultView: false, personalView: false };
    const catalog = source.discover({ ...state.request, query: '' });
    await new Promise(resolve => setImmediate(resolve));
    const verification = source.verifyCandidate(candidate, state.request);
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(state.calls.filter(call => /\/(fields|views)\?/.test(call.url)).length, 1, `${domain} must join the catalog request`);
    finish({ value: [row] }); await catalog;
    assert.equal((await verification).key, 'consumer');
  }
});
test('catalog verification keeps filtered fallback for absent, incomplete, failed or missing positive evidence', async () => {
  const value = { kind: 'field', webUrl: web, fieldId: guid, internalName: 'Title' };
  const candidate = { key: 'consumer', text: 'Consumer label', value, title: 'Candidate' };
  for (const mode of ['absent', 'partial', 'error', 'failed', 'missing', 'differentIdentity', 'rejected', 'expired']) {
    const state = setup({});
    let now = 0;
    state.request.cache = state.h.load('controls/pickers/picker-request-cache.internal.ts').createPickerRequestCache(state.context.owner, { now: () => now });
    let reads = 0, catalogs = 0;
    const source = state.h.load('controls/pickers/picker-metadata-source.internal.ts').createMetadataPickerSource({
      identity: 'field-catalog', context: state.context, completeCatalog: true, kinds: ['field'], accepts: option => option.title !== 'Rejected',
      discover: async () => {
        catalogs++;
        if (mode === 'failed') throw new Error('Catalog denied');
        return { options: mode === 'missing' ? [] : [{ ...candidate, value: mode === 'differentIdentity' ? { ...value, fieldId: '22222222-2222-4222-8222-222222222222' } : value, title: mode === 'rejected' ? 'Rejected' : 'Cached' }], completeness: mode === 'partial' ? 'partial' : 'complete',
          errors: mode === 'error' ? [{ source: 'sharePoint', error: new Error('Catalog partial error') }] : undefined };
      },
      read: async () => { reads++; return { ...candidate, title: 'Filtered' }; },
    });
    if (mode !== 'absent') await source.discover({ ...state.request, query: '' }).catch(() => undefined);
    if (mode === 'expired') now = 300001;
    const result = await source.verifyCandidate(candidate, state.request);
    assert.equal(result.title, 'Filtered', mode); assert.equal(reads, 1, mode);
    assert.equal(catalogs, mode === 'absent' ? 0 : 1, `${mode} must not initiate another catalog`);
  }
});
test('catalog verification does not accept an invalidated or replaced in-flight catalog', async () => {
  const state = setup({});
  const candidate = { key: 'consumer', text: 'Consumer label', title: 'Candidate', value: { kind: 'field', webUrl: web, fieldId: guid, internalName: 'Title' } };
  let finish, reads = 0, catalogs = 0;
  const source = state.h.load('controls/pickers/picker-metadata-source.internal.ts').createMetadataPickerSource({
    identity: 'field-catalog', context: state.context, completeCatalog: true, kinds: ['field'], accepts: () => true,
    discover: async () => ++catalogs === 1 ? new Promise(resolve => { finish = resolve; }) : { options: [{ ...candidate, title: 'Replacement' }], completeness: 'complete' },
    read: async () => { reads++; return { ...candidate, title: 'Filtered' }; },
  });
  const request = { ...state.request, partition: 'catalog:' };
  const old = source.discover({ ...request, query: '' });
  await new Promise(resolve => setImmediate(resolve));
  const verification = source.verifyCandidate(candidate, request);
  request.cache.invalidate('catalog:');
  await source.discover({ ...request, query: '' });
  finish({ options: [{ ...candidate, title: 'Stale' }], completeness: 'complete' });
  await old;
  assert.equal((await verification).title, 'Filtered'); assert.equal(reads, 1);
  assert.equal((await source.verifyCandidate(candidate, request)).title, 'Replacement', 'later verification may use the new catalog');
  assert.equal(reads, 1);
});
test('catalog verification crossing the TTL boundary cannot renew the expired catalog', async () => {
  const state = setup({});
  const candidate = { key: 'consumer', text: 'Consumer label', title: 'Candidate', value: { kind: 'field', webUrl: web, fieldId: guid, internalName: 'Title' } };
  let crossingBoundary = false, clockReads = 0, reads = 0, catalogs = 0;
  state.request.cache = state.h.load('controls/pickers/picker-request-cache.internal.ts').createPickerRequestCache(state.context.owner, {
    now: () => !crossingBoundary ? 0 : ++clockReads === 1 ? 299999 : 300000,
  });
  const source = state.h.load('controls/pickers/picker-metadata-source.internal.ts').createMetadataPickerSource({
    identity: 'field-catalog', context: state.context, completeCatalog: true, kinds: ['field'], accepts: () => true,
    discover: async () => { catalogs++; return { options: [{ ...candidate, title: 'Cached' }], completeness: 'complete' }; },
    read: async () => { reads++; return { ...candidate, title: 'Filtered' }; },
  });
  await source.discover({ ...state.request, query: '' });
  crossingBoundary = true;
  assert.equal((await source.verifyCandidate(candidate, state.request)).title, 'Filtered');
  assert.equal((await source.verifyCandidate(candidate, state.request)).title, 'Filtered', 'the expired promise must not acquire a fresh TTL');
  assert.equal(reads, 2); assert.equal(catalogs, 1, 'verification must not initiate a catalog refresh');
});
test('catalog verification is isolated by source, client, owner, tenant, user and target context', async () => {
  const state = setup({});
  const candidate = { key: 'consumer', text: 'Consumer label', title: 'Candidate', value: { kind: 'field', webUrl: web, fieldId: guid, internalName: 'Title' } };
  let reads = 0;
  const create = (identity = 'field-catalog', context = state.context) => state.h.load('controls/pickers/picker-metadata-source.internal.ts').createMetadataPickerSource({
    identity, context, completeCatalog: true, kinds: ['field'], accepts: () => true,
    discover: async () => ({ options: [{ ...candidate, title: 'Cached' }], completeness: 'complete' }),
    read: async () => { reads++; return { ...candidate, title: 'Filtered' }; },
  });
  await create().discover({ ...state.request, query: '' });
  assert.equal((await create().verifyCandidate(candidate, state.request)).title, 'Cached');
  assert.equal(reads, 0);
  assert.equal((await create('other-source').verifyCandidate(candidate, state.request)).title, 'Filtered');
  for (const change of [{ owner: {} }, { spClient: {} }, { tenantId: 'other' }, { userId: 'other' }, { webUrl: web + '/sub' }, { siteCollectionUrl: 'https://tenant.test/sites/other' }]) {
    const context = { ...state.context, ...change };
    assert.equal((await create('field-catalog', context).verifyCandidate(candidate, { ...state.request, context })).title, 'Filtered');
  }
  assert.equal(reads, 7);
});
test('list REST filters preserve raw precedence, template OR grouping and typed search', async () => {
  const row = { Id: guid, Title: 'Projects', BaseTemplate: 100, Hidden: false, RootFolder: { ServerRelativeUrl: '/sites/hr/Lists/Projects' } };
  const { h, context, calls, request } = setup({ value: [row] });
  const { createListPickerSource } = h.load('controls/pickers/list-picker.source.internal.ts');
  const source = createListPickerSource({ baseTemplate: [100, 101], includeHidden: false }, context);
  const page = await source.discover(request);
  const filter = new URL(calls[0].url).searchParams.get('$filter');
  assert.match(filter, /\(BaseTemplate eq 100 or BaseTemplate eq 101\)/);
  assert.match(filter, /Hidden eq false/);
  assert.doesNotMatch(filter, /startswith|tolower|toupper/);
  assert.equal(page.options[0].value.listId, guid);
  const raw = createListPickerSource({ baseTemplate: 999, includeHidden: false, filter: 'Hidden eq true' }, context);
  await raw.discover(request);
  assert.equal(new URL(calls.at(-1).url).searchParams.get('$filter'), "(Hidden eq true)");
});
test('field source queries site columns when list is omitted and uses ReadOnlyField metadata', async () => {
  const row = { Id: guid, InternalName: 'Title', Title: 'Title', TypeAsString: 'Text', FieldTypeKind: 2, Hidden: false, ReadOnlyField: true, Group: 'Columns' };
  const { h, context, calls, request } = setup({ value: [row] });
  const { createFieldPickerSource } = h.load('controls/pickers/field-picker.source.internal.ts');
  const page = await createFieldPickerSource({}, context).discover({ ...request, query: '' });
  assert.match(calls[0].url, /_api\/web\/fields/);
  assert.equal(page.options[0].readOnly, true);
  assert.equal(page.options[0].value.listId, undefined);
});
test('view source requires resolved list identity and excludes titles without inventing blank references', async () => {
  const { h, context, request } = setup(url => url.includes('/views') ? { value: [{ Id: guid, Title: 'Hidden', DefaultView: false, PersonalView: true }, { Id: '22222222-2222-4222-8222-222222222222', Title: 'Public', DefaultView: true, PersonalView: false }] } : { Id: guid, Title: 'Projects', BaseTemplate: 100 });
  const { createViewPickerSource } = h.load('controls/pickers/view-picker.source.internal.ts');
  const source = createViewPickerSource({ list: 'Projects', excludedViewTitles: ['Hidden'], showBlankOption: true }, context);
  const page = await source.discover({ ...request, query: '' });
  assert.deepEqual(page.options.map(o => o.title), ['Public']);
  assert.equal(page.options[0].value.listId, guid);
  assert.equal(page.options[0].personalView, false);
});
test('SitePicker searches collections beyond its host without a current-collection constraint and retains verified identity', async () => {
  const { h, context, request } = setup({ value: [] });
  context.siteCollectionUrl = 'https://tenant.test/sites/root';
  context.webUrl = 'https://tenant.test/sites/root/sub';
  const calls = [];
  const search = async (query, options) => { calls.push({ query, options }); return { results: [{ id: 'random-fallback', data: { Title: 'HR', Path: 'https://tenant.test/sites/other', SiteId: guid } }], totalResults: 1 }; };
  const { createSitePickerSource } = h.load('controls/pickers/site-picker.source.internal.ts');
  const source = createSitePickerSource({ mode: 'site' }, context, search);
  const page = await source.discover({ ...request, context });
  assert.match(calls[0].query, /contentclass:STS_Site/);
  assert.doesNotMatch(calls[0].query, /Path=/);
  assert.equal(page.options[0].value.url, 'https://tenant.test/sites/other');
  assert.equal(source.accepts(page.options[0]), true);
  assert.doesNotMatch(calls[0].query, /root\/sub/);
  assert.notEqual(page.options[0].key, 'random-fallback');
  assert.equal(page.options[0].value.siteId, guid);
  assert.equal(calls[0].options.startRow, 0);
  assert.equal(page.continuation, undefined);
});
test('list content-type filtering accepts JSON-light expanded arrays case-insensitively', async () => {
  const row = { Id: guid, Title: 'Projects', BaseTemplate: 100, Hidden: false, RootFolder: { ServerRelativeUrl: '/sites/hr/Lists/Projects' }, ContentTypes: [{ StringId: '0x0100AB' }] };
  const { h, context, request } = setup({ value: [row] });
  const source = h.load('controls/pickers/list-picker.source.internal.ts').createListPickerSource({ contentTypeId: '0x0100ab' }, context);
  assert.equal((await source.discover(request)).options.length, 1);
});
test('decoded title/path selectors preserve hash, percent, Unicode and apostrophes in REST resources', () => {
  const { h } = setup({});
  const { pickerListResource } = h.load('controls/pickers/picker-list-target.internal.ts');
  const title = new URL(pickerListResource("R&D #1 50% O'Neil", web));
  assert.equal(title.hash, '');
  assert.match(decodeURIComponent(title.pathname), /R&D #1 50% O''Neil/);
  const root = new URL(pickerListResource({ kind: 'url', serverRelativeUrl: '/sites/hr/Lists/A#B%À' }, web));
  assert.equal(root.hash, '');
  assert.match(root.pathname, /GetListUsingPath/);
  assert.match(decodeURIComponent(root.pathname), /A#B%À/);
});
test('SitePicker associatedSites uses native collection metadata and returns empty without an association', async () => {
  const { h, context, request, calls } = setup({ HubSiteId: '00000000-0000-0000-0000-000000000000' });
  const { createSitePickerSource } = h.load('controls/pickers/site-picker.source.internal.ts');
  let searches = 0;
  const page = await createSitePickerSource({ mode: 'associatedSites' }, context, async () => { searches++; return { results: [], totalResults: 0 }; }).discover(request);
  assert.equal(page.options.length, 0);
  assert.equal(searches, 0);
  assert.match(calls[0].url, /_api\/site\?/);
});
test('metadata continuation binds query, page size, source and client while preserving REST filters', async () => {
  const { h, context, calls, request } = setup(url => new URL(url).searchParams.has('$skiptoken') ? { value: [] } : { value: [{ Id: guid, Title: 'Projects', BaseTemplate: 100, Hidden: false }], '@odata.nextLink': web + '/_api/web/lists?$skiptoken=Paged%3dTRUE' });
  const { createListPickerSource } = h.load('controls/pickers/list-picker.source.internal.ts');
  request.pageSize = 1;
  const source = createListPickerSource({ includeHidden: false }, context);
  const page = await source.discover(request);
  assert.ok(page.continuation);
  await source.discover({ ...request, query: 'PR', continuation: page.continuation });
  assert.match(new URL(calls.at(-1).url).searchParams.get('$filter'), /Hidden eq false/);
  await assert.rejects(source.discover({ ...request, query: 'Other', continuation: page.continuation }), /continuation/);
  await assert.rejects(source.discover({ ...request, pageSize: 20, continuation: page.continuation }), /continuation/);
  const other = createListPickerSource({ includeHidden: true }, context);
  await assert.rejects(other.discover({ ...request, continuation: page.continuation }), /continuation/);
});
test('hub picker exposes a scoped continuation for locally filtered hub results', async () => {
  const { h, context, request } = setup({ value: ['A', 'B', 'C'].map((Title, i) => ({ Title, SiteUrl: `${web}/${i}`, ID: guid })) });
  const source = h.load('controls/pickers/site-picker.source.internal.ts').createSitePickerSource({ mode: 'hub' }, context);
  const first = await source.discover({ ...request, query: '', pageSize: 2 });
  assert.equal(first.options.length, 2);
  assert.ok(first.continuation);
  const second = await source.discover({ ...request, query: '', pageSize: 2, continuation: first.continuation });
  assert.deepEqual(second.options.map(o => o.text), ['C']);
  await assert.rejects(source.discover({ ...request, pageSize: 2, continuation: first.continuation }), /continuation/);
});
test('SitePicker defaults to collections and rejects the former web mode and subweb references', async () => {
  const { h, context, request } = setup(url => url.includes('/_api/site?') ? { Id: guid, Url: web } : { Title: 'Subweb', Id: guid });
  const calls = [];
  const search = async query => { calls.push(query); return { results: [], totalResults: 0 }; };
  const create = h.load('controls/pickers/site-picker.source.internal.ts').createSitePickerSource;
  await create({}, context, search).discover(request);
  assert.match(calls[0], /contentclass:STS_Site/);
  assert.throws(() => create({ mode: 'web' }, context, search), /Unsupported/);
  const result = await create({}, context, search).resolve([{ kind: 'site', url: web + '/subweb' }], request);
  assert.equal(result.resolved.length, 0);
  assert.equal(result.unresolved.length, 1);
});
test('collection search sorts the complete retrieved set before paging without unsupported managed-property sorts', async () => {
  const { h, context, request } = setup({});
  const calls = [];
  const all = Array.from({ length: 201 }, (_, i) => ({ data: { Title: `Collection ${String(200 - i).padStart(3, '0')}`, Path: `https://tenant.test/sites/c${i}`, SiteId: guid } }));
  const search = async (query, options) => { calls.push({ query, options }); return { results: all.slice(options.startRow, options.startRow + options.pageSize), totalResults: all.length }; };
  const source = h.load('controls/pickers/site-picker.source.internal.ts').createSitePickerSource({}, context, search);
  const first = await source.discover({ ...request, pageSize: 2 });
  assert.deepEqual(first.options.map(o => o.text), ['Collection 000', 'Collection 001']);
  const second = await source.discover({ ...request, pageSize: 2, continuation: first.continuation });
  assert.deepEqual(second.options.map(o => o.text), ['Collection 002', 'Collection 003']);
  assert.equal(calls.length, 2);
  assert.equal(calls[0].options.orderBy, undefined);
  await assert.rejects(source.discover({ ...request, query: 'other', pageSize: 2, continuation: first.continuation }), /continuation/);
});
test('KQL literal quotes are doubled and backslashes are preserved according to string-token syntax', async () => {
  const { h, context, request } = setup({}); let query;
  const source = h.load('controls/pickers/site-picker.source.internal.ts').createSitePickerSource({}, context, async text => { query = text; return { results: [], totalResults: 0 }; });
  await source.discover({ ...request, query: 'A "quoted" \\ phrase' });
  assert.ok(query.includes('Title:"a ""quoted"" \\ phrase*"'));
  assert.equal(query.includes('\\"'), false);
});

test('metadata search is case insensitive beyond the first native page and preserves only consumer filters', async () => {
  const row = { Id: guid, Title: 'ÀREA Documenti', BaseTemplate: 100, Hidden: false };
  const state = setup(url => new URL(url).searchParams.has('$skiptoken') ? { value: [row] } : { value: [{ ...row, Title: 'Unrelated' }], 'odata.nextLink': `${web}/_api/web/lists?$skiptoken=next` });
  const source = state.h.load('controls/pickers/list-picker.source.internal.ts').createListPickerSource({ filter: "Hidden eq false" }, state.context);
  const result = await source.discover({ ...state.request, query: 'ÀREA', pageSize: 1 });
  assert.deepEqual(result.options.map(o => o.text), ['ÀREA Documenti']);
  assert.equal(state.calls.length, 2);
  assert.equal(new URL(state.calls[0].url).searchParams.get('$filter'), '(Hidden eq false)');
  const lower = await source.discover({ ...state.request, query: 'àrea', pageSize: 1 });
  assert.deepEqual(lower.options, result.options);
});

test('Field/View complete catalogs follow every native page and search locally without additional REST', async () => {
  for (const domain of ['Field', 'View']) {
    const rows = Array.from({ length: 245 }, (_, i) => ({ Id: `${String(i + 1).padStart(8, '0')}-1111-4111-8111-111111111111`, Title: `Àrea ${i}`, InternalName: `Field${i}`, TypeAsString: 'Text', FieldTypeKind: 2, Hidden: false, ReadOnlyField: false, Group: 'Demo', DefaultView: false, PersonalView: false }));
    const state = setup(uri => {
      if (!uri.includes('/fields?') && !uri.includes('/views?')) return { Id: guid, Title: 'Projects', BaseTemplate: 100 };
      const url = new URL(uri), offset = Number(url.searchParams.get('$skiptoken') ?? 0), next = offset + 35;
      return { value: rows.slice(offset, next), 'odata.nextLink': next < rows.length ? `${url.origin}${url.pathname}?$skiptoken=${next}` : undefined };
    });
    const create = state.h.load(`controls/pickers/${domain.toLowerCase()}-picker.source.internal.ts`)[`create${domain}PickerSource`];
    const src = create(domain === 'Field' ? {} : { list: 'Projects' }, state.context);
    const initial = await src.discover({ ...state.request, query: '', pageSize: 30 });
    assert.equal(initial.options.length, 245, domain);
    assert.equal(initial.continuation, undefined); assert.equal(initial.completeness, 'complete');
    const requests = state.calls.length;
    const typed = await src.discover({ ...state.request, query: 'ÀREA 244', pageSize: 1 });
    const lower = await src.discover({ ...state.request, query: 'àrea 244', pageSize: 1 });
    assert.deepEqual(typed.options.map(o => o.text), ['Àrea 244']); assert.deepEqual(lower.options, typed.options);
    assert.equal(state.calls.length, requests, 'typing must not request REST');
  }
});

test('complete metadata catalogs preserve errors, refresh and source/client isolation', async () => {
  let fail = true;
  const state = setup(() => { if (fail) throw new Error('Catalog denied'); return { value: [{ Id: guid, Title: 'Title', InternalName: 'Title', TypeAsString: 'Text', FieldTypeKind: 2, Hidden: false, ReadOnlyField: false, Group: 'Demo' }] }; });
  const create = state.h.load('controls/pickers/field-picker.source.internal.ts').createFieldPickerSource;
  const request = { ...state.request, partition: 'catalog-test:' };
  const src = create({}, state.context);
  await assert.rejects(src.discover({ ...request, query: '' }), /Catalog denied/);
  const failedCalls = state.calls.length; fail = false;
  await assert.rejects(src.discover({ ...request, query: 'title' }), /Catalog denied/);
  assert.equal(state.calls.length, failedCalls);
  request.cache.invalidate('catalog-test:');
  assert.equal((await src.discover({ ...request, query: 'TITLE' })).options.length, 1);
  const successfulCalls = state.calls.length;
  await create({}, state.context).discover({ ...request, query: '' });
  assert.equal(state.calls.length, successfulCalls);
  await create({ includeHidden: false }, state.context).discover({ ...request, query: '' });
  assert.equal(state.calls.length, successfulCalls + 1);
  const changed = { ...state.context, spClient: { ...state.context.spClient } };
  await create({}, changed).discover({ ...request, context: changed, query: '' });
  assert.equal(state.calls.length, successfulCalls + 2);
});

test('Site defaults load current collection and its associated hub without a text search, independently toggle and cache', async () => {
  const hubUrl = 'https://tenant.test/sites/hub', hubId = '22222222-2222-4222-8222-222222222222';
  const state = setup(uri => {
    if (uri.includes('GetById')) return { ID: hubId, SiteId: hubId, SiteUrl: hubUrl, Title: 'Hub' };
    if (uri.includes('/_api/site?')) return { Id: uri.startsWith(hubUrl) ? hubId : guid, Url: uri.startsWith(hubUrl) ? hubUrl : web, HubSiteId: hubId, IsHubSite: false };
    if (uri.includes('/_api/web?')) return { Id: guid, Title: uri.startsWith(hubUrl) ? 'Hub' : 'Current' };
    throw new Error(`Unexpected ${uri}`);
  });
  state.context.webUrl = web + '/subweb';
  let searches = 0;
  const create = state.h.load('controls/pickers/site-picker.source.internal.ts').createSitePickerSource;
  const src = create({}, state.context, async () => { searches++; return { results: [], totalResults: 0 }; });
  const data = await src.personalSuggestions({ ...state.request, query: '' });
  assert.equal(data.currentSite[0].url, web); assert.equal(data.hubSite[0].url, hubUrl); assert.equal(searches, 0);
  assert.equal(state.calls.some(call => call.url.includes('/subweb/_api')), false);
  const count = state.calls.length; await src.personalSuggestions({ ...state.request, query: '' }); assert.equal(state.calls.length, count);
  const off = create({ includeCurrentSite: false, includeHubSite: false }, state.context);
  assert.equal(off.personalSuggestions, undefined);
  const hubOnly = await create({ includeCurrentSite: false }, state.context, async () => ({ results: [], totalResults: 0 })).personalSuggestions(state.request);
  assert.equal(hubOnly.currentSite, undefined); assert.equal(hubOnly.hubSite.length, 1);
});
test('Site default source failure retains healthy current collection, while an absent hub is not an error', async () => {
  const hubId = '22222222-2222-4222-8222-222222222222'; let associated = true;
  const state = setup(uri => {
    if (uri.includes('GetById')) throw new Error('Hub denied');
    if (uri.includes('/_api/site?')) return { Id: guid, Url: web, HubSiteId: associated ? hubId : '00000000-0000-0000-0000-000000000000', IsHubSite: false };
    return { Id: guid, Title: 'Current' };
  });
  const create = state.h.load('controls/pickers/site-picker.source.internal.ts').createSitePickerSource;
  const data = await create({}, state.context).personalSuggestions(state.request);
  assert.equal(data.currentSite.length, 1); assert.equal(data.completeness, 'partial'); assert.equal(data.errors[0].source, 'hubSite');
  associated = false; const request = { ...state.request, cache: state.h.load('controls/pickers/picker-request-cache.internal.ts').createPickerRequestCache({}) };
  const emptyHub = await create({}, state.context).personalSuggestions(request);
  assert.equal(emptyHub.currentSite.length, 1); assert.equal(emptyHub.hubSite, undefined); assert.equal(emptyHub.completeness, 'complete');
});

test('Site hub default uses actual association rather than configured hub ID and honors source KQL', async () => {
  const actual = '22222222-2222-4222-8222-222222222222', configured = '33333333-3333-4333-8333-333333333333';
  const state = setup(uri => uri.includes('GetById') ? { ID: actual, SiteUrl: web } : uri.includes('/_api/site?') ? { Id: actual, Url: web, HubSiteId: actual, IsHubSite: true } : { Title: 'Current hub' });
  const create = state.h.load('controls/pickers/site-picker.source.internal.ts').createSitePickerSource;
  const data = await create({ hubSiteId: configured }, state.context).personalSuggestions(state.request);
  assert.equal(data.hubSite[0].siteId, actual); assert.equal(data.currentSite[0].siteId, actual);
  assert.equal(state.calls.some(call => call.url.includes(configured)), false);
  const queries = [];
  const filtered = await create({ additionalQuery: 'Title:Allowed' }, state.context, async query => { queries.push(query); return { results: [], totalResults: 0 }; }).personalSuggestions(state.request);
  assert.equal(filtered.currentSite, undefined); assert.equal(filtered.hubSite, undefined); assert.equal(filtered.errors.length, 0);
  assert.ok(queries.every(query => query.includes('(Title:Allowed)') && query.includes('Path=')));
});

test('associatedSites scopes discovery using actual collection HubSiteId metadata', async () => {
  const hub = '22222222-2222-4222-8222-222222222222';
  const state = setup(uri => uri.includes('/_api/site?') ? { Id: guid, Url: web, HubSiteId: hub, IsHubSite: false } : {});
  const queries = [];
  const source = state.h.load('controls/pickers/site-picker.source.internal.ts').createSitePickerSource({ mode: 'associatedSites' }, state.context, async query => { queries.push(query); return { results: [], totalResults: 0 }; });
  await source.discover(state.request);
  assert.equal(queries.length, 1); assert.ok(queries[0].includes(`DepartmentId:"${hub}"`));
});

test('a current hub with no association is not its own associated hub', async () => {
  const state = setup(uri => uri.includes('GetById') ? { ID: guid, SiteUrl: web } : uri.includes('/_api/site?') ? { Id: guid, Url: web, HubSiteId: '00000000-0000-0000-0000-000000000000', IsHubSite: true } : { Title: 'Current hub' });
  const src = state.h.load('controls/pickers/site-picker.source.internal.ts').createSitePickerSource({}, state.context);
  const data = await src.personalSuggestions(state.request);
  assert.equal(data.currentSite.length, 1); assert.equal(data.hubSite, undefined);
  const compose = state.h.load('controls/pickers/picker-suggestions.internal.ts').composePickerSuggestions;
  assert.deepEqual(compose(data).map(s => s.source), ['currentSite']);
  assert.equal(state.calls.filter(c => c.url.includes('/_api/site?')).length, 1);
});

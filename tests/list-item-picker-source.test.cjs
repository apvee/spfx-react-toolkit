const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const web = 'https://tenant.test/sites/hr';
const guid = '11111111-1111-4111-8111-111111111111';
function setup(type, responder, baseTemplate = 100, metadata = {}) {
  const calls = [];
  const h = createHarness({ '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } } });
  const send = async (url, config, options, post = false) => {
    calls.push({ url, options, post });
    const value = url.includes('/fields/') ? { InternalName: 'Title', EntityPropertyName: 'Title', TypeAsString: type, FieldTypeKind: type === 'Calculated' ? 17 : type === 'Number' ? 9 : 2, OutputType: 2, ...metadata } : !url.includes('/items') && !url.includes('/GetItems') ? { Id: guid, Title: 'Projects', BaseTemplate: baseTemplate } : await responder(url, options, post);
    return { ok: true, status: 200, json: async () => value };
  };
  const context = { owner: {}, tenantId: guid, userId: 'u', webUrl: web, siteCollectionUrl: web, spClient: { get: send, post: (url, config, options) => send(url, config, options, true) } };
  const cache = h.load('controls/pickers/picker-request-cache.internal.ts').createPickerRequestCache(context.owner);
  return { h, context, calls, request: { query: "O'Neil's", pageSize: 50, generation: 0, context, cache, partition: 'items:' } };
}
test('concurrent item hydration and resolution share entity reads and preserve distinct verification', async () => {
  const row = { Id: 7, Title: 'Selected', FSObjType: 0 };
  const state = setup('Text', url => new URL(url).pathname.endsWith('/items(7)') ? row : { value: [row] });
  const source = state.h.load('controls/pickers/list-item-picker.source.internal.ts').createListItemPickerSource({ list: { kind: 'id', id: guid }, textField: 'Title' }, state.context);
  const value = { kind: 'item', webUrl: web, listId: guid, itemId: 7 };
  const [hydrated, resolved] = await Promise.all([source.hydrate([{ kind: 'value', value, key: 'alias' }], state.request), source.resolve([value], state.request)]);
  assert.equal(hydrated.resolved[0].option.key, 'alias'); assert.equal(resolved.resolved[0].option.text, 'Selected');
  assert.equal(state.calls.filter(call => call.url.includes('/items(7)')).length, 1);
  await Promise.all([source.verifyCandidate(resolved.resolved[0].option, state.request), source.verifyCandidate(resolved.resolved[0].option, state.request)]);
  assert.equal(state.calls.filter(call => call.url.includes('/items?')).length, 1);
  assert.match(new URL(state.calls.find(call => call.url.includes('/items?')).url).searchParams.get('$filter'), /Id eq 7/);
});
test('text contains matches apostrophes case-insensitively and retains sparse native continuation rather than count', async () => {
  const next = `${web}/_api/web/lists(guid'${guid}')/items?$skiptoken=Paged%3dTRUE%26p_ID%3d900`;
  const { h, context, calls, request } = setup('Text', url => new URL(url).searchParams.has('$skiptoken') ? { value: [{ Id: 950, Title: "O'Neil's Last", FSObjType: 0 }] } : { value: [{ Id: 2, Title: "O'Neil's First", FSObjType: 0 }, { Id: 900, Title: "O'Neil's Middle", FSObjType: 0 }], 'odata.nextLink': next });
  const source = h.load('controls/pickers/list-item-picker.source.internal.ts').createListItemPickerSource({ list: 'Projects', textField: 'Title', searchMode: 'contains' }, context);
  request.pageSize = 2;
  const page = await source.discover(request);
  assert.equal(new URL(calls.find(c => c.url.includes('/items')).url).searchParams.get('$filter'), '(FSObjType eq 0)');
  assert.equal(page.options[1].value.itemId, 900);
  const second = await source.discover({ ...request, continuation: page.continuation });
  assert.equal(second.options[0].value.itemId, 950);
  await assert.rejects(source.discover({ ...request, query: 'changed', continuation: page.continuation }), /continuation/);
});
test('REST entity property names differ from CAML/internal field names and blank titles remain resolvable', async () => {
  const state = setup('Text', () => ({ value: [{ Id: 1, OData__Text: 'Text', FSObjType: 0 }, { Id: 2, OData__Text: null, FSObjType: 0 }] }), 100, { InternalName: '_Text', EntityPropertyName: 'OData__Text' });
  const source = state.h.load('controls/pickers/list-item-picker.source.internal.ts').createListItemPickerSource({ list: 'Projects', textField: '_Text' }, state.context);
  const page = await source.discover({ ...state.request, query: '' });
  assert.deepEqual(page.options.map(o => o.itemId), [1, 2]);
  assert.equal(page.options[1].text, 'Elemento 2');
  assert.match(new URL(state.calls.find(c => c.url.includes('/items')).url).searchParams.get('$select'), /OData__Text/);
});
test('calculated text follows native REST pages without relying on CAML casing', async () => {
  const state = setup('Calculated', url => new URL(url).searchParams.has('$skiptoken') ? { value: [{ Id: 2, Title: 'MATCH second', FSObjType: 0 }] } : { value: [{ Id: 1, Title: 'match first', FSObjType: 0 }], 'odata.nextLink': `${web}/_api/web/lists(guid'${guid}')/items?$skiptoken=second` });
  const source = state.h.load('controls/pickers/list-item-picker.source.internal.ts').createListItemPickerSource({ list: 'Projects', textField: 'Title' }, state.context);
  const page = await source.discover({ ...state.request, query: 'Match' });
  assert.deepEqual(page.options.map(o => o.itemId), [1, 2]);
  assert.equal(page.completeness, 'complete');
  assert.equal(state.calls.some(call => call.post), false);
});
test('custom-key uniqueness is not certified while a short endpoint page still has a continuation', async () => {
  const state = setup('Text', url => new URL(url).searchParams.get('$select') === 'Id' ? { value: [{ Id: 1 }], 'odata.nextLink': `${url.split('?')[0]}?$skiptoken=more` } : { value: [{ Id: 1, Title: 'Same', FSObjType: 0 }] });
  const source = state.h.load('controls/pickers/list-item-picker.source.internal.ts').createListItemPickerSource({ list: 'Projects', textField: 'Title', keyField: 'Title' }, state.context);
  const candidate = (await source.discover({ ...state.request, query: '' })).options[0];
  await assert.rejects(source.verifyCandidate(candidate, state.request), /unique|continuation/);
});
test('calculated text matches locally while preserving raw OData filters and order', async () => {
  const { h, context, calls, request } = setup('Calculated', (url, options, post) => post ? { value: [{ Id: 40 }] } : { value: [{ Id: 900, Title: 'Unmatched', FSObjType: 0 }, { Id: 40, Title: 'a<& matched', FSObjType: 0 }] });
  const source = h.load('controls/pickers/list-item-picker.source.internal.ts').createListItemPickerSource({ list: 'Projects', textField: 'Title', filter: 'Status eq 1', orderBy: [{ field: 'Id', direction: 'desc' }] }, context);
  const page = await source.discover({ ...request, query: 'A<&' });
  assert.deepEqual(page.options.map(o => o.itemId), [40]);
  assert.match(new URL(calls.find(c => c.url.includes('/items')).url).searchParams.get('$filter'), /Status eq 1/);
  assert.equal(new URL(calls.find(c => c.url.includes('/items')).url).searchParams.get('$orderby'), 'Id desc');
  assert.equal(calls.some(call => call.post), false);
});
test('nontext search uses source FieldValuesAsText and reports partial results after bounded scanning', async () => {
  let pages = 0;
  const { h, context, request } = setup('Number', url => ({ value: [{ Id: ++pages, Title: pages, FSObjType: 0, FieldValuesAsText: { Title: 'does not match' } }], 'odata.nextLink': `${url.split('?')[0]}?$skiptoken=${pages}` }));
  const source = h.load('controls/pickers/list-item-picker.source.internal.ts').createListItemPickerSource({ list: 'Projects', textField: 'Title' }, context);
  const page = await source.discover({ ...request, query: 'needle' });
  assert.equal(pages, 5);
  assert.equal(page.completeness, 'partial');
  assert.equal(page.options.length, 0);
  assert.ok(page.continuation);
});
test('library entries are files/folders only with verified metadata and selected values omit raw fields', async () => {
  const rows = [{ Id: 7, Title: 'Contract', FSObjType: 0, File: { UniqueId: guid, Name: 'Contract.pdf', ServerRelativeUrl: '/sites/hr/docs/Contract.pdf' } }, { Id: 9, Title: 'Folder', FSObjType: 1, Folder: { UniqueId: guid, ServerRelativeUrl: '/sites/hr/docs/Folder' } }];
  const { h, context, request } = setup('Text', () => ({ value: rows }), 101);
  const source = h.load('controls/pickers/list-item-picker.source.internal.ts').createListItemPickerSource({ list: 'Docs', textField: 'Title', includeFolders: true }, context);
  const page = await source.discover({ ...request, query: '' });
  assert.deepEqual(page.options.map(o => o.value.kind), ['file', 'folder']);
  assert.equal(page.options[0].value.extension, 'pdf');
  assert.equal(page.options[0].value.fields, undefined);
  const broken = setup('Text', () => ({ value: [{ Id: 7, Title: 'No metadata', FSObjType: 0 }] }), 101);
  await assert.rejects(broken.h.load('controls/pickers/list-item-picker.source.internal.ts').createListItemPickerSource({ list: 'Docs', textField: 'Title' }, broken.context).discover({ ...broken.request, query: '' }), /metadata/);
});

test('all item display-field types match mixed case locally rather than filtering out backend candidates', async () => {
  for (const type of ['Text', 'Note', 'Calculated', 'Number']) {
    const state = setup(type, () => ({ value: [{ Id: 2, Title: 'ÀREA Documenti', FieldValuesAsText: { Title: 'ÀREA Documenti' }, FSObjType: 0 }] }));
    const src = state.h.load('controls/pickers/list-item-picker.source.internal.ts').createListItemPickerSource({ list: 'Projects', textField: 'Title', searchMode: 'contains', filter: 'Id gt 0' }, state.context);
    const result = await src.discover({ ...state.request, query: 'àrea' });
    assert.equal(result.options.length, 1, type);
    const request = state.calls.find(call => call.url.includes('/items?'));
    assert.equal(new URL(request.url).searchParams.get('$filter'), '(Id gt 0) and (FSObjType eq 0)');
    assert.equal(state.calls.some(call => call.post), false);
  }
});

test('file names participate in case-insensitive search and label empty titles through hydration', async () => {
  for (const title of [null, '', '   ']) {
    const row = { Id: 7, Title: title, FSObjType: 0, File: { UniqueId: guid, Name: 'Quarterly-Report.docx', ServerRelativeUrl: '/sites/hr/docs/Quarterly-Report.docx' } };
    const state = setup('Text', url => new URL(url).pathname.endsWith('/items(7)') ? row : { value: [row] }, 101);
    const source = state.h.load('controls/pickers/list-item-picker.source.internal.ts').createListItemPickerSource({ list: 'Docs', textField: 'Title' }, state.context);
    const page = await source.discover({ ...state.request, query: 'QUARTERLY' });
    assert.equal(page.options.length, 1, 'an empty Title must not hide a matching file');
    assert.equal(page.options[0].text, 'Quarterly-Report.docx');
    const verified = await source.verifyCandidate(page.options[0], state.request);
    assert.equal(verified.text, 'Quarterly-Report.docx');
    state.request.cache.invalidate('items:');
    const hydrated = await source.resolve([verified.value], state.request);
    assert.equal(hydrated.resolved[0].option.text, 'Quarterly-Report.docx');
    assert.equal(hydrated.resolved[0].value.kind, 'file');
    assert.equal(state.calls.some(call => call.post), false);
  }
});
test('file-name matching honors contains versus startsWith without replacing a nonempty title', async () => {
  const row = { Id: 7, Title: 'Board materials', FSObjType: 0, File: { UniqueId: guid, Name: 'Quarterly-Report.docx', ServerRelativeUrl: '/sites/hr/docs/Quarterly-Report.docx' } };
  for (const [mode, query, count] of [['startsWith', 'REPORT', 0], ['contains', 'REPORT', 1], ['startsWith', 'quarterly', 1], ['startsWith', 'BOARD', 1]]) {
    const state = setup('Text', () => ({ value: [row] }), 101);
    const source = state.h.load('controls/pickers/list-item-picker.source.internal.ts').createListItemPickerSource({ list: 'Docs', textField: 'Title', searchMode: mode }, state.context);
    const page = await source.discover({ ...state.request, query });
    assert.equal(page.options.length, count, mode + ':' + query);
    if (count) assert.equal(page.options[0].text, 'Board materials');
  }
});

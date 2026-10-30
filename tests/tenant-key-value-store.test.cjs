const assert = require('node:assert/strict');
const test = require('node:test');
const { createHarness } = require('./services-test-harness.cjs');
const { createScriptedSPHttpClient } = require('./fixtures/key-value-store-transport.cjs');

const SPPermission = createHarness({ '@microsoft/sp-core-library': { _SPKillSwitch: { isActivated: () => false } } })
  .load(require.resolve('@microsoft/sp-page-context/lib-commonjs/SPPermission')).default;
const url = 'https://tenant.test/sites/catalog';
const list = `${url}/_api/web/lists/getByTitle('TenantKeyValueStore')`;
const fields = `${list}/fields?$filter=InternalName eq 'Title' or InternalName eq 'Value' or InternalName eq 'Description'&$select=InternalName,FieldTypeKind,Indexed,EnforceUniqueValues`;
const items = `${list}/items?$select=Id,Title,Value,Description&$orderby=Title`;
const find = key => `${list}/items?$filter=${encodeURIComponent(`Title eq '${key.replace(/'/g, "''")}'`).replace(/'/g, '%27')}&$select=Id%2CTitle%2CValue%2CDescription&$top=2`;
const schema = (unique = true) => [
  { InternalName: 'Title', FieldTypeKind: 2, Indexed: unique, EnforceUniqueValues: unique },
  { InternalName: 'Value', FieldTypeKind: 3 }, { InternalName: 'Description', FieldTypeKind: 3 }
];
const row = (Title = 'key', Value = '123', Id = 7) => ({ Id, Title, Value, Description: 'old', '@odata.etag': '"3"' });

function fixture(t) {
  const transport = createScriptedSPHttpClient();
  const h = createHarness({
    '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } },
    '@microsoft/sp-page-context': { SPPermission },
    './useSPFxSPHttpClient': { useSPFxSPHttpClient: () => ({ client: transport.client }) },
    './useSPFxPageContext': { useSPFxPageContext: () => ({ web: { absoluteUrl: url } }) }
  });
  t.after(h.close);
  const create = () => h.load('services/spfx-tenant-key-value-store.service.ts').createSPFxTenantKeyValueStoreService(transport.client);
  return { ...transport, h, create, service: create() };
}
function probe(f, value = schema()) {
  f.enqueue('GET', `${list}?$select=Id`);
  f.enqueue('GET', fields, { body: { value } });
}

test('tenant absent list makes reads and remove empty without provisioning', async t => {
  const f = fixture(t);
  for (const [run, expected] of [[() => f.service.get('key', url), undefined], [() => f.service.list(url), []], [() => f.service.remove('key', url), undefined]]) {
    f.enqueue('GET', `${list}?$select=Id`, { status: 404 });
    assert.deepEqual(await run(), expected);
  }
  assert.ok(f.calls.every(call => call.method === 'GET'));
  f.assertDrained();
});

test('tenant reads retain legacy values, nullable Note values, stored casing and escaped keys', async t => {
  const f = fixture(t), key = "O'Brien & #/?";
  probe(f);
  f.enqueue('GET', find(key), { body: { value: [{ ...row('Stored key'), Value: null, Description: null }] } });
  assert.deepEqual(await f.service.get(key, url), { id: 7, key: 'Stored key', value: null, description: undefined });
  probe(f);
  f.enqueue('GET', items, { body: { value: [row('number'), row('boolean', 'true', 8), row('text', 'hello', 9)] } });
  assert.deepEqual((await f.service.list(url)).map(item => item.value), [123, true, 'hello']);
  f.assertDrained();
});

for (const status of [401, 403, 500]) {
  test(`tenant read rejects ${status} instead of reporting an empty store`, async t => {
    const f = fixture(t);
    f.enqueue('GET', `${list}?$select=Id`, { status, statusText: 'Denied' });
    await assert.rejects(f.service.list(url), /Failed to check list existence: Denied/);
    f.assertDrained();
  });
}

test('tenant read rejects incompatible fields, malformed rows and duplicate keys', async t => {
  for (const value of [schema().slice(1), schema().map(field => field.InternalName === 'Value' ? { ...field, FieldTypeKind: 2 } : field)]) {
    const f = fixture(t); probe(f, value);
    await assert.rejects(f.service.get('key', url), /field|schema/i); f.assertDrained();
  }
  for (const value of [[{ ...row(), Id: '7' }], [{ ...row(), Value: {} }], [row(), row('key', '456', 8)]]) {
    const f = fixture(t); probe(f);
    f.enqueue('GET', find('key'), { body: { value } });
    await assert.rejects(f.service.get('key', url), /item|row|unique/i); f.assertDrained();
  }
});

test('tenant list follows both continuation forms and retains server ordering', async t => {
  const f = fixture(t); probe(f);
  const second = `${list}/items?$skiptoken=2`, third = `${list}/items?$skiptoken=3`;
  f.enqueue('GET', items, { body: { value: [row('A')], '@odata.nextLink': second.replace('https://tenant.test', '') } });
  f.enqueue('GET', second, { body: { value: [row('B', 'true', 8)], 'odata.nextLink': third } });
  f.enqueue('GET', third, { body: { value: [row('C', 'hello', 9)] } });
  assert.deepEqual((await f.service.list(url)).map(item => item.key), ['A', 'B', 'C']); f.assertDrained();
});

test('tenant list rejects foreign, repeated and malformed pages without returning partial data', async t => {
  for (const next of ['https://foreign.test/items', items, `${url}/_api/web/lists/getByTitle('Other')/items`, 12, '']) {
    const f = fixture(t); probe(f);
    f.enqueue('GET', items, { body: { value: [row()], '@odata.nextLink': next } });
    await assert.rejects(f.service.list(url)); f.assertDrained();
  }
  const f = fixture(t); probe(f);
  const next = `${list}/items?$skiptoken=2`;
  f.enqueue('GET', items, { body: { value: [row()], 'odata.nextLink': next } });
  f.enqueue('GET', next, { status: 500 });
  await assert.rejects(f.service.list(url)); f.assertDrained();
});

test('tenant setup requires unique Title on existing lists and retries failures', async t => {
  const f = fixture(t); probe(f, schema(false));
  f.enqueue('POST', `${list}/fields/getByInternalNameOrTitle('Title')`, { status: 403, statusText: 'Denied' });
  await assert.rejects(f.service.ensureListReady(url), /uniqueness.*Denied/);
  probe(f);
  await f.service.ensureListReady(url);
  await f.service.ensureListReady(` ${url}/ `);
  assert.equal(f.calls.length, 5); f.assertDrained();
});

test('tenant setup shares in-flight work and normalizes URLs while isolating factories', async t => {
  const f = fixture(t); let resolve;
  f.enqueue('GET', `${list}?$select=Id`, new Promise(done => { resolve = done; }));
  f.enqueue('GET', fields, { body: { value: schema() } });
  const first = f.service.ensureListReady(url);
  const duplicate = f.service.ensureListReady(` ${url}/ `);
  const settled = Promise.allSettled([first, duplicate]);
  resolve({}); await settled;
  assert.equal(duplicate, first);
  await f.service.ensureListReady(url);
  probe(f); await f.create().ensureListReady(url); f.assertDrained();
});

test('tenant setup recovers concurrent list/field creation only after validated rereads', async t => {
  const f = fixture(t);
  f.enqueue('GET', `${list}?$select=Id`, { status: 404 });
  f.enqueue('POST', `${url}/_api/web/lists`, { status: 409 });
  f.enqueue('GET', `${list}?$select=Id`);
  f.enqueue('GET', fields, { body: { value: schema().filter(field => field.InternalName !== 'Value') } });
  f.enqueue('POST', `${list}/fields`, { status: 409 });
  f.enqueue('GET', fields, { body: { value: schema() } });
  await f.service.ensureListReady(url); f.assertDrained();
});

test('tenant setup creates the original hidden list schema and verifies unique Title', async t => {
  const f = fixture(t);
  f.enqueue('GET', `${list}?$select=Id`, { status: 404 });
  f.enqueue('POST', `${url}/_api/web/lists`);
  f.enqueue('GET', fields, { body: { value: schema(false).slice(0, 1) } });
  f.enqueue('POST', `${list}/fields`); f.enqueue('POST', `${list}/fields`);
  f.enqueue('POST', `${list}/fields/getByInternalNameOrTitle('Title')`);
  f.enqueue('GET', fields, { body: { value: schema() } });
  await f.service.ensureListReady(url);
  assert.deepEqual(JSON.parse(f.calls[1].options.body), { BaseTemplate: 100, Title: 'TenantKeyValueStore', Hidden: true, NoCrawl: true });
  assert.deepEqual(JSON.parse(f.calls[3].options.body), { FieldTypeKind: 3, Title: 'Value' });
  assert.deepEqual(JSON.parse(f.calls[4].options.body), { FieldTypeKind: 3, Title: 'Description' });
  assert.deepEqual(JSON.parse(f.calls[5].options.body), { Indexed: true, EnforceUniqueValues: true });
  f.assertDrained();
});

test('tenant hook permission indicator uses list effective grants instead of administrator role', async t => {
  for (const exists of [true, false]) {
    const f = fixture(t);
    const useCatalog = f.h.load('hooks/useAppCatalogUrl.internal.ts').useAppCatalogUrl;
    const view = f.h.mount(() => useCatalog());
    f.enqueue('GET', `${list}?$select=Id`, { status: exists ? 200 : 404 });
    f.enqueue('GET', `${exists ? list : `${url}/_api/web`}/effectiveBasePermissions`, { body: { High: '0', Low: exists ? '14' : '2062' } });
    assert.equal(await view.current.checkWritePermission(url), true);
    f.assertDrained(); view.unmount();
  }
});

test('tenant permission indicator fails closed on missing bits, malformed grants and failed probes', async t => {
  for (const body of [{ High: '0', Low: '6' }, { High: '0', Low: '14junk' }, {}, { d: { EffectiveBasePermissions: { High: '0', Low: '6' } } }]) {
    const f = fixture(t), useCatalog = f.h.load('hooks/useAppCatalogUrl.internal.ts').useAppCatalogUrl;
    const view = f.h.mount(() => useCatalog());
    f.enqueue('GET', `${list}?$select=Id`);
    f.enqueue('GET', `${list}/effectiveBasePermissions`, { body });
    assert.equal(await view.current.checkWritePermission(url), false); f.assertDrained(); view.unmount();
  }
  for (const status of [401, 403, 500]) {
    const f = fixture(t), useCatalog = f.h.load('hooks/useAppCatalogUrl.internal.ts').useAppCatalogUrl;
    const view = f.h.mount(() => useCatalog());
    f.enqueue('GET', `${list}?$select=Id`, { status });
    assert.equal(await view.current.checkWritePermission(url), false); f.assertDrained(); view.unmount();
  }
});

test('tenant new string and bigint writes round-trip without changing legacy readers or stored values', async t => {
  const f = fixture(t); probe(f);
  const values = ['', '123', 'true', 'null', '{"enabled":true}', '"quoted"', 'hello', 9007199254740993n, new Date('2026-01-02T03:04:05Z')];
  const expected = ['', '123', 'true', 'null', '{"enabled":true}', '"quoted"', 'hello', '9007199254740993', '2026-01-02T03:04:05.000Z'];
  for (const [index, value] of values.entries()) {
    f.enqueue('GET', find('key'), { body: { value: [] } });
    f.enqueue('POST', `${list}/items`);
    await f.service.save('key', value, url);
    const stored = JSON.parse(f.calls.at(-1).options.body).Value;
    if (value === '') assert.equal(stored, '""', 'empty Note text must survive SharePoint null normalization');
    // The historical reader is JSON-first. New data needs no version marker or new column.
    let legacyRead; try { legacyRead = JSON.parse(stored); } catch { legacyRead = stored; }
    assert.equal(legacyRead, expected[index]);
    probe(f);
    f.enqueue('GET', find('key'), { body: { value: [row('key', stored === '' ? null : stored)] } });
    assert.equal((await f.service.get('key', url)).value, expected[index]);
  }
  f.assertDrained();
});

test('tenant invalid serializable values fail before schema or item writes', async t => {
  const f = fixture(t), circular = {}; circular.self = circular;
  for (const value of [circular, { number: 1n }, new Date('invalid')]) {
    await assert.rejects(f.service.save('key', value, url));
  }
  assert.equal(f.calls.length, 0); f.assertDrained();
});

test('tenant updates preserve or clear descriptions and send observed ETags', async t => {
  const f = fixture(t); probe(f);
  for (const [description, expected] of [[undefined, 'old'], ['', ''], ['new', 'new']]) {
    f.enqueue('GET', find('key'), { body: { value: [row()] } });
    f.enqueue('POST', `${list}/items(7)`);
    await f.service.save('key', { enabled: true }, url, description);
    assert.deepEqual(f.calls.at(-1).options, {
      headers: { 'X-HTTP-Method': 'MERGE', 'If-Match': '"3"' },
      body: JSON.stringify({ Value: '{"enabled":true}', Description: expected })
    });
  }
  f.assertDrained();
});

test('tenant missing collection ETag uses a fresh single-item snapshot and response header', async t => {
  const f = fixture(t); probe(f);
  f.enqueue('GET', find('key'), { body: { value: [{ ...row(), '@odata.etag': undefined }] } });
  f.enqueue('GET', `${list}/items(7)?$select=Id,Title,Value,Description`, {
    ok: true, status: 200, headers: new Headers({ ETag: '"4"' }),
    json: async () => ({ ...row(), '@odata.etag': undefined, Description: 'fresh' }), text: async () => ''
  });
  f.enqueue('POST', `${list}/items(7)`);
  await f.service.save('key', 5, url);
  assert.equal(f.calls.at(-1).options.headers['If-Match'], '"4"');
  assert.equal(JSON.parse(f.calls.at(-1).options.body).Description, 'fresh'); f.assertDrained();
});

for (const operation of ['save', 'remove']) {
  test(`tenant ${operation} rejects conflicts without overwriting or retrying`, async t => {
    const f = fixture(t); probe(f);
    f.enqueue('GET', find('key'), { body: { value: [row()] } });
    f.enqueue('POST', `${list}/items(7)`, { status: 412, statusText: 'Precondition Failed', text: 'changed concurrently' });
    await assert.rejects(operation === 'save' ? f.service.save('key', 5, url) : f.service.remove('key', url), /Precondition Failed/);
    assert.equal(f.calls.at(-1).options.headers['If-Match'], '"3"'); f.assertDrained();
  });
}

test('tenant missing ETag never falls back to wildcard mutation', async t => {
  const f = fixture(t); probe(f);
  const item = { ...row(), '@odata.etag': undefined };
  f.enqueue('GET', find('key'), { body: { value: [item] } });
  f.enqueue('GET', `${list}/items(7)?$select=Id,Title,Value,Description`, { body: item });
  await assert.rejects(f.service.save('key', 1, url), /ETag/);
  assert.ok(f.calls.every(call => call.method === 'GET')); f.assertDrained();
});

test('tenant concurrent key creation recovers with one conditional update and preserves description', async t => {
  const f = fixture(t); probe(f);
  f.enqueue('GET', find('key'), { body: { value: [] } });
  f.enqueue('POST', `${list}/items`, { status: 409 });
  f.enqueue('GET', fields, { body: { value: schema() } });
  f.enqueue('GET', find('key'), { body: { value: [row()] } });
  f.enqueue('POST', `${list}/items(7)`);
  await f.service.save('key', 5, url);
  assert.equal(f.calls.at(-1).options.headers['If-Match'], '"3"');
  assert.equal(JSON.parse(f.calls.at(-1).options.body).Description, 'old'); f.assertDrained();
});

test('tenant conditional recovery exposes the update conflict instead of the original create collision', async t => {
  const f = fixture(t); probe(f);
  f.enqueue('GET', find('key'), { body: { value: [] } });
  f.enqueue('POST', `${list}/items`, { status: 409, statusText: 'Create collision' });
  f.enqueue('GET', fields, { body: { value: schema() } });
  f.enqueue('GET', find('key'), { body: { value: [row()] } });
  f.enqueue('POST', `${list}/items(7)`, { status: 412, statusText: 'Precondition Failed' });
  await assert.rejects(f.service.save('key', 5, url), /Failed to update item: Precondition Failed/);
  f.assertDrained();
});

test('tenant remove accepts a vanished ETag snapshot only after confirming key absence', async t => {
  const f = fixture(t); probe(f);
  f.enqueue('GET', find('key'), { body: { value: [{ ...row(), '@odata.etag': undefined }] } });
  f.enqueue('GET', `${list}/items(7)?$select=Id,Title,Value,Description`, { status: 404, statusText: 'Not Found' });
  probe(f); f.enqueue('GET', find('key'), { body: { value: [] } });
  await f.service.remove('key', url);
  assert.ok(f.calls.every(call => call.method === 'GET')); f.assertDrained();
});

test('tenant remove keeps a snapshot 404 failure when the key was recreated or cannot be verified', async t => {
  for (const denied of [true, false]) {
    const f = fixture(t); probe(f);
    f.enqueue('GET', find('key'), { body: { value: [{ ...row(), '@odata.etag': undefined }] } });
    f.enqueue('GET', `${list}/items(7)?$select=Id,Title,Value,Description`, { status: 404, statusText: 'Not Found' });
    if (denied) f.enqueue('GET', `${list}?$select=Id`, { status: 403 });
    else { probe(f); f.enqueue('GET', find('key'), { body: { value: [row('key', 'new', 8)] } }); }
    await assert.rejects(f.service.remove('key', url), /Failed to read item for mutation: Not Found/);
    assert.ok(f.calls.every(call => call.method === 'GET')); f.assertDrained();
  }
});

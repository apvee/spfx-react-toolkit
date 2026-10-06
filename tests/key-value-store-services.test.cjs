const assert = require('node:assert/strict');
const test = require('node:test');
const { createHarness } = require('./services-test-harness.cjs');
const { createScriptedSPHttpClient } = require('./fixtures/key-value-store-transport.cjs');

const catalogUrl = 'https://tenant.test/sites/catalog';
const listUrl = url => `${url}/_api/web/lists/getByTitle('TenantKeyValueStore')`;
const fieldsUrl = url => `${listUrl(url)}/fields?$filter=InternalName eq 'Value' or InternalName eq 'Description'&$select=InternalName`;
const itemsUrl = url => `${listUrl(url)}/items?$select=Id,Title,Value,Description&$orderby=Title`;
const findUrl = (url, key) => {
  const encode = value => encodeURIComponent(value).replace(/'/g, '%27');
  return `${listUrl(url)}/items?$filter=${encode(`Title eq '${key.replace(/'/g, "''")}'`)}&$select=${encode('Id,Title,Value,Description')}&$top=1`;
};
const mergeHeaders = { 'X-HTTP-Method': 'MERGE', 'If-Match': '*' };

function setup(factoryName = 'tenant') {
  const transport = createScriptedSPHttpClient();
  const configuration = {};
  const harness = createHarness({ '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: configuration } } } });
  const createService = client => factoryName === 'tenant'
    ? harness.load('services/spfx-tenant-key-value-store.service.ts').createSPFxTenantKeyValueStoreService(client)
    : harness.load('services/spfx-list-key-value-store.internal.ts').createSPFxListKeyValueStoreService(client, 'tenant');
  return { ...transport, configuration, createService, service: createService(transport.client) };
}

function enqueueReady(transport, url = catalogUrl, fields = ['Value', 'Description']) {
  transport.enqueue('GET', `${listUrl(url)}?$select=Id`);
  transport.enqueue('GET', fieldsUrl(url), { body: { value: fields.map(InternalName => ({ InternalName })) } });
}

function assertOptions(call, body, headers) {
  assert.deepEqual(call.options, { ...(headers ? { headers } : {}), ...(body ? { body: JSON.stringify(body) } : {}) });
}

function characterize(factoryName) {
  const prefix = `${factoryName} tenant behavior: `;
  test(prefix + 'new list setup preserves exact requests, bodies and warning-only uniqueness failure', async () => {
    const t = setup(factoryName);
    t.enqueue('GET', `${listUrl(catalogUrl)}?$select=Id`, { status: 404 });
    t.enqueue('POST', `${catalogUrl}/_api/web/lists`);
    t.enqueue('POST', `${listUrl(catalogUrl)}/fields`);
    t.enqueue('POST', `${listUrl(catalogUrl)}/fields`);
    t.enqueue('POST', `${listUrl(catalogUrl)}/fields/getByInternalNameOrTitle('Title')`, { status: 500 });
    const warnings = [];
    const originalWarn = console.warn;
    console.warn = message => warnings.push(message);
    try { await t.service.ensureListReady(catalogUrl); } finally { console.warn = originalWarn; }
    assert.deepEqual(warnings, ['Failed to set Title uniqueness constraint. It may already be configured.']);
    assertOptions(t.calls[1], { BaseTemplate: 100, Title: 'TenantKeyValueStore', Hidden: true, NoCrawl: true });
    assertOptions(t.calls[2], { FieldTypeKind: 3, Title: 'Value' });
    assertOptions(t.calls[3], { FieldTypeKind: 3, Title: 'Description' });
    assertOptions(t.calls[4], { Indexed: true, EnforceUniqueValues: true }, mergeHeaders);
    assert.ok(t.calls.every(call => call.configuration === t.configuration));
    await t.service.ensureListReady(catalogUrl);
    assert.equal(t.calls.length, 5);
    t.assertDrained();
  });

  test(prefix + 'existing list repairs only missing fields in Value/Description order', async () => {
    for (const fields of [[], ['Value'], ['Description'], ['Value', 'Description']]) {
      const t = setup(factoryName);
      enqueueReady(t, catalogUrl, fields);
      const missing = ['Value', 'Description'].filter(field => !fields.includes(field));
      for (const field of missing) t.enqueue('POST', `${listUrl(catalogUrl)}/fields`);
      await t.service.ensureListReady(catalogUrl);
      missing.forEach((field, index) => assertOptions(t.calls[index + 2], { FieldTypeKind: 3, Title: field }));
      t.assertDrained();
    }
  });

  test(prefix + 'get/list never provision, preserve escaping/raw slash URLs, ignore nextLink and retain coercion', async () => {
    const t = setup(factoryName);
    const url = catalogUrl + '/';
    const key = "O'Brien & #/?";
    t.enqueue('GET', findUrl(url, key), { body: { value: [{ Id: 3, Title: key, Value: '123', Description: '' }] } });
    assert.deepEqual(await t.service.get(key, url), { id: 3, key, value: 123, description: undefined });
    t.enqueue('GET', findUrl(url, 'absent'), { body: { value: [] } });
    assert.equal(await t.service.get('absent', url), undefined);
    t.enqueue('GET', itemsUrl(url), { body: {
      value: [{ Id: 4, Title: 'bool', Value: 'true' }, { Id: 5, Title: 'text', Value: 'hello', Description: 'kept' }],
      '@odata.nextLink': 'https://tenant.test/unused', 'odata.nextLink': 'https://tenant.test/also-unused'
    } });
    assert.deepEqual(await t.service.list(url), [
      { id: 4, key: 'bool', value: true, description: undefined },
      { id: 5, key: 'text', value: 'hello', description: 'kept' }
    ]);
    assert.equal(t.calls.length, 3);
    assert.ok(t.calls.every(call => call.url.startsWith(url + '/_api/')));
    assert.deepEqual(Object.keys(t.service), ['ensureListReady', 'get', 'list', 'save', 'remove']);
    t.assertDrained();
  });

  test(prefix + 'save auto-provisions then creates with legacy serialization and description defaults', async () => {
    const t = setup(factoryName);
    enqueueReady(t);
    const values = ['123', 'true', new Date('2026-01-02T03:04:05Z'), { enabled: true }];
    const serialized = ['123', 'true', '2026-01-02T03:04:05.000Z', '{"enabled":true}'];
    for (const [index, value] of values.entries()) {
      t.enqueue('GET', findUrl(catalogUrl, 'key'), { body: { value: [] } });
      t.enqueue('POST', `${listUrl(catalogUrl)}/items`);
      await t.service.save('key', value, catalogUrl, index === 3 ? 'description' : undefined);
      assertOptions(t.calls.at(-1), { Title: 'key', Value: serialized[index], Description: index === 3 ? 'description' : '' });
    }
    t.assertDrained();
  });

  test(prefix + 'MERGE preserves absent description, permits explicit empty override and uses wildcard ETag', async () => {
    const t = setup(factoryName);
    enqueueReady(t);
    for (const [description, existingDescription, expected] of [[undefined, 'old', 'old'], ['', 'old', ''], ['new', 'old', 'new'], [undefined, undefined, '']]) {
      t.enqueue('GET', findUrl(catalogUrl, 'key'), { body: { value: [{ Id: 7, Title: 'key', Value: '0', Description: existingDescription }] } });
      t.enqueue('POST', `${listUrl(catalogUrl)}/items(7)`);
      await t.service.save('key', 12, catalogUrl, description);
      assertOptions(t.calls.at(-1), { Value: '12', Description: expected }, mergeHeaders);
    }
    t.assertDrained();
  });

  test(prefix + 'remove auto-provisions and deletes existing key; absent key is a no-op', async () => {
    const t = setup(factoryName);
    enqueueReady(t);
    t.enqueue('GET', findUrl(catalogUrl, 'key'), { body: { value: [{ Id: 8, Title: 'key', Value: '0' }] } });
    t.enqueue('POST', `${listUrl(catalogUrl)}/items(8)`);
    await t.service.remove('key', catalogUrl);
    assertOptions(t.calls.at(-1), undefined, { 'X-HTTP-Method': 'DELETE', 'If-Match': '*' });
    t.enqueue('GET', findUrl(catalogUrl, 'absent'), { body: { value: [] } });
    await t.service.remove('absent', catalogUrl);
    t.assertDrained();
  });

  test(prefix + 'same-instance provisioning shares promise and separate raw URLs have isolated readiness', async () => {
    const t = setup(factoryName);
    let resolve;
    t.enqueue('GET', `${listUrl(catalogUrl)}?$select=Id`, new Promise(done => { resolve = done; }));
    t.enqueue('GET', fieldsUrl(catalogUrl), { body: { value: [{ InternalName: 'Value' }, { InternalName: 'Description' }] } });
    const first = t.service.ensureListReady(catalogUrl);
    assert.equal(t.service.ensureListReady(catalogUrl), first);
    assert.equal(t.calls.length, 1);
    resolve({});
    await first;
    await t.service.ensureListReady(catalogUrl);
    for (const url of [catalogUrl + '/', 'https://tenant.test/sites/other']) {
      enqueueReady(t, url);
      await t.service.ensureListReady(url);
    }
    assert.equal(t.calls.length, 6);
    t.assertDrained();
  });

  test(prefix + 'failed setup clears the mutex, retries, and separate factories do not share readiness', async () => {
    const t = setup(factoryName);
    t.enqueue('GET', `${listUrl(catalogUrl)}?$select=Id`, { status: 500, statusText: 'Broken' });
    await assert.rejects(t.service.ensureListReady(catalogUrl), { message: 'Failed to check list existence: Broken' });
    enqueueReady(t);
    await t.service.ensureListReady(catalogUrl);
    enqueueReady(t);
    await t.createService(t.client).ensureListReady(catalogUrl);
    t.assertDrained();
  });

  test(prefix + 'read errors retain original messages and do not swallow 404', async () => {
    const t = setup(factoryName);
    t.enqueue('GET', findUrl(catalogUrl, 'key'), { status: 404, statusText: 'Not Found', text: 'ignored' });
    await assert.rejects(t.service.get('key', catalogUrl), { message: 'Failed to find item: Not Found' });
    t.enqueue('GET', itemsUrl(catalogUrl), { status: 403, statusText: 'Forbidden', text: 'ignored' });
    await assert.rejects(t.service.list(catalogUrl), { message: 'Failed to list items: Forbidden' });
    t.assertDrained();
  });

  test(prefix + 'field check, list creation and field creation retain error text behavior', async () => {
    for (const operation of ['fields', 'list', 'Value', 'Description']) {
      const t = setup(factoryName);
      if (operation === 'list') {
        t.enqueue('GET', `${listUrl(catalogUrl)}?$select=Id`, { status: 404 });
        t.enqueue('POST', `${catalogUrl}/_api/web/lists`, { status: 500, statusText: 'Broken', text: 'details' });
      } else {
        t.enqueue('GET', `${listUrl(catalogUrl)}?$select=Id`);
        t.enqueue('GET', fieldsUrl(catalogUrl), operation === 'fields'
          ? { status: 500, statusText: 'Broken', text: 'ignored' }
          : { body: { value: [] } });
        if (operation === 'Description') t.enqueue('POST', `${listUrl(catalogUrl)}/fields`);
        if (operation !== 'fields') t.enqueue('POST', `${listUrl(catalogUrl)}/fields`, { status: 500, statusText: 'Broken', text: 'details' });
      }
      const message = operation === 'fields' ? 'Failed to check list fields: Broken'
        : operation === 'list' ? 'Failed to create list: Broken. details'
          : `Failed to create ${operation} field: Broken. details`;
      await assert.rejects(t.service.ensureListReady(catalogUrl), { message });
      t.assertDrained();
    }
  });

  test(prefix + 'create, update and delete retain exact errors without recovery requests', async () => {
    for (const operation of ['create', 'update', 'remove']) {
      const t = setup(factoryName);
      enqueueReady(t);
      t.enqueue('GET', findUrl(catalogUrl, 'key'), { body: { value: operation === 'create' ? [] : [{ Id: 9, Title: 'key', Value: 'old' }] } });
      t.enqueue('POST', `${listUrl(catalogUrl)}/items${operation === 'create' ? '' : '(9)'}`, { status: 500, statusText: 'Broken', text: 'details' });
      await assert.rejects(operation === 'remove' ? t.service.remove('key', catalogUrl) : t.service.save('key', 1, catalogUrl), {
        message: `Failed to ${operation} item: Broken. details`
      });
      t.assertDrained();
    }
  });
}

characterize('tenant');
characterize('core');

const SPPermission = createHarness({ '@microsoft/sp-core-library': { _SPKillSwitch: { isActivated: () => false } } }).load(require.resolve('@microsoft/sp-page-context/lib-commonjs/SPPermission')).default;
const rootUrl = 'https://tenant.test/sites/root';
const siteList = url => `${url}/_api/web/lists/getByTitle('SiteKeyValueStore')`;
const siteFields = url => `${siteList(url)}/fields?$filter=InternalName eq 'Title' or InternalName eq 'Value' or InternalName eq 'Description'&$select=InternalName,FieldTypeKind,Indexed,EnforceUniqueValues`;
const siteItems = url => `${siteList(url)}/items?$select=Id,Title,Value,Description&$orderby=Title`;
const siteFind = (url, key) => findUrl(url, key).replace('TenantKeyValueStore', 'SiteKeyValueStore').replace('$top=1', '$top=2');
const schema = (unique = true) => [
  { InternalName: 'Title', FieldTypeKind: 2, Indexed: unique, EnforceUniqueValues: unique },
  { InternalName: 'Value', FieldTypeKind: 3 }, { InternalName: 'Description', FieldTypeKind: 3 }
];
const row = (key = 'key', id = 7) => ({ Id: id, Title: key, Value: '123', Description: 'old' });
function siteSetup() {
  const t = createScriptedSPHttpClient();
  const harness = createHarness({
    '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } },
    '@microsoft/sp-page-context': { SPPermission }
  });
  const createService = client => harness.load('services/spfx-site-key-value-store.service.ts').createSPFxSiteKeyValueStoreService(client);
  return { ...t, createService, service: createService(t.client) };
}
function probe(t, url = rootUrl, fields = schema()) {
  t.enqueue('GET', `${siteList(url)}?$select=Id`);
  t.enqueue('GET', siteFields(url), { body: { value: fields } });
}
function mask(...permissions) {
  return permissions.reduce((acc, permission) => ({ High: acc.High | permission.value.High, Low: acc.Low | permission.value.Low }), { High: 0, Low: 0 });
}
const writeBits = [SPPermission.addListItems, SPPermission.editListItems, SPPermission.deleteListItems];

test('site public factory isolates roots, normalizes URLs and reads without provisioning', async () => {
  const t = siteSetup();
  for (const [url, id] of [[rootUrl, 7], ['https://tenant.test/sites/other', 8]]) {
    probe(t, url);
    t.enqueue('GET', siteFind(url, "O'Brien & #/?"), { body: { value: [row("O'Brien & #/?", id)] } });
    assert.deepEqual(await t.service.get("O'Brien & #/?", ` ${url}/// `), { id, key: "O'Brien & #/?", value: 123, description: 'old' });
  }
  assert.ok(t.calls.every(call => call.method === 'GET'));
  assert.ok(t.calls.every(call => !/appcatalog|TenantKeyValueStore|CorporateCatalog/.test(call.url)));
  t.assertDrained();
});

test('site confirmed list absence returns empty reads and remove without provisioning', async () => {
  const t = siteSetup();
  for (const [operation, expected] of [[() => t.service.get('key', rootUrl), undefined], [() => t.service.list(rootUrl), []], [() => t.service.remove('key', rootUrl), undefined]]) {
    t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`, { status: 404 });
    assert.deepEqual(await operation(), expected);
  }
});

test('site denies failed list/field/item probes and incompatible schema or rows', async () => {
  for (const status of [403, 500]) {
    const t = siteSetup();
    t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`, { status });
    await assert.rejects(t.service.list(rootUrl));
    t.assertDrained();
  }
  for (const stage of ['fields', 'items']) for (const status of [404, 403, 500]) {
    const t = siteSetup();
    if (stage === 'fields') {
      t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`);
      t.enqueue('GET', siteFields(rootUrl), { status });
    } else {
      probe(t); t.enqueue('GET', siteItems(rootUrl), { status });
    }
    await assert.rejects(t.service.list(rootUrl)); t.assertDrained();
  }
  for (const fields of [schema().slice(0, 2), schema().map(f => f.InternalName === 'Value' ? { ...f, FieldTypeKind: 2 } : f)]) {
    const t = siteSetup(); probe(t, rootUrl, fields);
    await assert.rejects(t.service.get('key', rootUrl), /schema|field/i); t.assertDrained();
  }
  for (const invalid of [{ ...row(), Id: '7' }, { ...row(), Value: {} }, { ...row(), Title: null }]) {
    const t = siteSetup(); probe(t); t.enqueue('GET', siteItems(rootUrl), { body: { value: [invalid] } });
    await assert.rejects(t.service.list(rootUrl), /item|row/i); t.assertDrained();
  }
});

test('site list completes both continuation forms and preserves server order', async () => {
  const t = siteSetup(); probe(t);
  const page2 = `${siteList(rootUrl)}/items?$skiptoken=2`;
  const page3 = `${siteList(rootUrl)}/items?$skiptoken=3`;
  t.enqueue('GET', siteItems(rootUrl), { body: { value: [row('A', 1)], '@odata.nextLink': page2.replace('https://tenant.test', '') } });
  t.enqueue('GET', page2, { body: { value: [row('B', 2)], 'odata.nextLink': page3 } });
  t.enqueue('GET', page3, { body: { value: [row('C', 3)] } });
  assert.deepEqual((await t.service.list(rootUrl)).map(item => item.key), ['A', 'B', 'C']); t.assertDrained();
});

test('site pagination rejects foreign, repeated, malformed links and failed later page', async () => {
  for (const next of ['https://foreign.test/items', siteItems(rootUrl), `${rootUrl}/_api/web/lists/getByTitle('Other')/items`, 12]) {
    const t = siteSetup(); probe(t);
    t.enqueue('GET', siteItems(rootUrl), { body: { value: [row()], '@odata.nextLink': next } });
    await assert.rejects(t.service.list(rootUrl)); t.assertDrained();
  }
  const t = siteSetup(); probe(t);
  const next = `${siteList(rootUrl)}/items?$skiptoken=2`;
  t.enqueue('GET', siteItems(rootUrl), { body: { value: [row()], 'odata.nextLink': next } });
  t.enqueue('GET', next, { status: 500 });
  await assert.rejects(t.service.list(rootUrl)); t.assertDrained();
});

test('site setup normalizes mutex/readiness, isolates factories and retries failure', async () => {
  const t = siteSetup(); let resolve;
  t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`, new Promise(done => { resolve = done; }));
  t.enqueue('GET', siteFields(rootUrl), { body: { value: schema() } });
  const first = t.service.ensureListReady(rootUrl);
  assert.equal(t.service.ensureListReady(` ${rootUrl}/ `), first);
  resolve({}); await first;
  await t.service.ensureListReady(rootUrl);
  probe(t); await t.createService(t.client).ensureListReady(rootUrl);
  const other = 'https://tenant.test/sites/other';
  t.enqueue('GET', `${siteList(other)}?$select=Id`, { status: 500 });
  await assert.rejects(t.service.ensureListReady(other));
  probe(t, other); await t.service.ensureListReady(other); t.assertDrained();
});

test('site setup provisions missing list/Note fields and strictly configures Title uniqueness', async () => {
  const t = siteSetup();
  t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`, { status: 404 });
  t.enqueue('POST', `${rootUrl}/_api/web/lists`);
  t.enqueue('GET', siteFields(rootUrl), { body: { value: schema(false).slice(0, 1) } });
  t.enqueue('POST', `${siteList(rootUrl)}/fields`);
  t.enqueue('POST', `${siteList(rootUrl)}/fields`);
  t.enqueue('POST', `${siteList(rootUrl)}/fields/getByInternalNameOrTitle('Title')`);
  t.enqueue('GET', siteFields(rootUrl), { body: { value: schema() } });
  await t.service.ensureListReady(rootUrl);
  assertOptions(t.calls[1], { BaseTemplate: 100, Title: 'SiteKeyValueStore', Hidden: true, NoCrawl: true });
  assertOptions(t.calls[3], { FieldTypeKind: 3, Title: 'Value' });
  assertOptions(t.calls[4], { FieldTypeKind: 3, Title: 'Description' });
  assertOptions(t.calls[5], { Indexed: true, EnforceUniqueValues: true }, mergeHeaders); t.assertDrained();
});

test('site setup rejects wrong field types and failed unique constraint without deleting data', async () => {
  const t = siteSetup(); probe(t, rootUrl, schema().map(f => f.InternalName === 'Title' ? { ...f, FieldTypeKind: 3 } : f));
  await assert.rejects(t.service.ensureListReady(rootUrl)); t.assertDrained();
  const u = siteSetup(); probe(u, rootUrl, schema(false));
  u.enqueue('POST', `${siteList(rootUrl)}/fields/getByInternalNameOrTitle('Title')`, { status: 500 });
  await assert.rejects(u.service.ensureListReady(rootUrl)); u.assertDrained();
});

test('site creation races recover list and field only after bounded validated rereads', async () => {
  const t = siteSetup();
  t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`, { status: 404 });
  t.enqueue('POST', `${rootUrl}/_api/web/lists`, { status: 409 });
  t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`);
  t.enqueue('GET', siteFields(rootUrl), { body: { value: schema().filter(f => f.InternalName !== 'Value') } });
  t.enqueue('POST', `${siteList(rootUrl)}/fields`, { status: 500 });
  t.enqueue('GET', siteFields(rootUrl), { body: { value: schema() } });
  await t.service.ensureListReady(rootUrl); t.assertDrained();
  for (const status of [401, 403]) {
    const denied = siteSetup(); denied.enqueue('GET', `${siteList(rootUrl)}?$select=Id`, { status: 404 });
    denied.enqueue('POST', `${rootUrl}/_api/web/lists`, { status });
    await assert.rejects(denied.service.ensureListReady(rootUrl)); denied.assertDrained();
  }
});

test('site save shares CRUD serialization, description preservation, clear and escaping', async () => {
  const t = siteSetup(); probe(t);
  const key = "O'Brien";
  t.enqueue('GET', siteFind(rootUrl, key), { body: { value: [] } });
  t.enqueue('POST', `${siteList(rootUrl)}/items`);
  await t.service.save(key, { ok: true }, rootUrl, 'note');
  assertOptions(t.calls.at(-1), { Title: key, Value: '{"ok":true}', Description: 'note' });
  for (const [description, expected] of [[undefined, 'old'], ['', '']]) {
    t.enqueue('GET', siteFind(rootUrl, key), { body: { value: [row(key)] } });
    t.enqueue('POST', `${siteList(rootUrl)}/items(7)`);
    await t.service.save(key, 'true', rootUrl, description);
    assertOptions(t.calls.at(-1), { Value: 'true', Description: expected }, mergeHeaders);
  }
  t.assertDrained();
});

test('site item collision performs at most one exact-key MERGE and preserves original failure otherwise', async () => {
  for (const outcome of ['success', 'missing', 'merge-fails', 'lookup-fails', 'duplicate', 'forbidden']) {
    const t = siteSetup(); probe(t);
    t.enqueue('GET', siteFind(rootUrl, 'key'), { body: { value: [] } });
    t.enqueue('POST', `${siteList(rootUrl)}/items`, { status: outcome === 'forbidden' ? 403 : 500, statusText: 'Collision', text: 'original' });
    if (outcome !== 'forbidden') {
      t.enqueue('GET', siteFields(rootUrl), { body: { value: schema() } });
      t.enqueue('GET', siteFind(rootUrl, 'key'), outcome === 'lookup-fails' ? { status: 500 } : { body: { value: outcome === 'missing' ? [] : outcome === 'duplicate' ? [row(), row('key', 8)] : [row()] } });
      if (outcome === 'success' || outcome === 'merge-fails') t.enqueue('POST', `${siteList(rootUrl)}/items(7)`, { status: outcome === 'success' ? 200 : 500 });
    }
    if (outcome === 'success') await t.service.save('key', 5, rootUrl);
    else await assert.rejects(t.service.save('key', 5, rootUrl), { message: 'Failed to create item: Collision. original' });
    t.assertDrained();
  }
});

test('site remove never repairs schema and accepts only verified concurrent absence', async () => {
  for (const outcome of ['absent', 'gone', 'present', 'denied']) {
    const t = siteSetup(); probe(t, rootUrl, schema(false));
    t.enqueue('GET', siteFind(rootUrl, 'key'), { body: { value: outcome === 'absent' ? [] : [row()] } });
    if (outcome !== 'absent') {
      t.enqueue('POST', `${siteList(rootUrl)}/items(7)`, { status: outcome === 'denied' ? 403 : 404, statusText: 'Delete failed' });
      if (outcome !== 'denied') {
        t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`);
        t.enqueue('GET', siteFields(rootUrl), { body: { value: schema(false) } });
        t.enqueue('GET', siteFind(rootUrl, 'key'), { body: { value: outcome === 'gone' ? [] : [row()] } });
      }
    }
    if (outcome === 'present' || outcome === 'denied') await assert.rejects(t.service.remove('key', rootUrl));
    else await t.service.remove('key', rootUrl);
    t.assertDrained();
  }
});

test('site permission checks list effective grants or absent root creation grants and fails closed', async () => {
  for (const absent of [false, true]) for (const permissions of [writeBits, [...writeBits, SPPermission.manageLists], writeBits.slice(0, 2)]) {
    const t = siteSetup();
    t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`, { status: absent ? 404 : 200 });
    const bits = mask(...permissions);
    t.enqueue('GET', absent ? `${rootUrl}/_api/web/effectiveBasePermissions` : `${siteList(rootUrl)}/effectiveBasePermissions`, { body: { High: String(bits.High), Low: String(bits.Low) } });
    assert.equal(await t.service.canCurrentUserWrite(rootUrl), permissions.length >= (absent ? 4 : 3));
    t.assertDrained();
  }
  for (const body of [{}, { High: 'x', Low: '14' }, { High: 0, Low: -1 }, { High: 0, Low: '14junk' }]) {
    const t = siteSetup(); t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`);
    t.enqueue('GET', `${siteList(rootUrl)}/effectiveBasePermissions`, { body });
    assert.equal(await t.service.canCurrentUserWrite(rootUrl), false); t.assertDrained();
  }
  const t = siteSetup(); t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`, { status: 403 });
  assert.equal(await t.service.canCurrentUserWrite(rootUrl), false); t.assertDrained();
});

test('site normalizes nullable empty SharePoint Note values and distinguishes serialized null', async () => {
  const t = siteSetup(); probe(t);
  t.enqueue('GET', siteFind(rootUrl, 'key'), { body: { value: [{ ...row(), Value: null, Description: null }] } });
  assert.deepEqual(await t.service.get('key', rootUrl), { key: 'key', id: 7, value: '', description: undefined });
  probe(t);
  t.enqueue('GET', siteItems(rootUrl), { body: { value: [{ ...row('empty'), Value: null }, { ...row('null', 8), Value: 'null' }] } });
  assert.deepEqual((await t.service.list(rootUrl)).map(item => item.value), ['', null]);
  t.assertDrained();
});

test('site creation recovery rejects missing or incompatible postconditions and retains original errors', async () => {
  for (const outcome of ['absent', 'reread-denied']) {
    const t = siteSetup(); t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`, { status: 404 });
    t.enqueue('POST', `${rootUrl}/_api/web/lists`, { status: 500, statusText: 'Create failed', text: 'original' });
    t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`, { status: outcome === 'absent' ? 404 : 403 });
    await assert.rejects(t.service.ensureListReady(rootUrl), { message: 'Failed to create list: Create failed. original' }); t.assertDrained();
  }
  for (const outcome of ['missing', 'wrong-type', 'denied']) {
    const t = siteSetup(); probe(t, rootUrl, schema().filter(f => f.InternalName !== 'Value'));
    t.enqueue('POST', `${siteList(rootUrl)}/fields`, { status: outcome === 'denied' ? 403 : 409, statusText: 'Create failed', text: 'original' });
    if (outcome !== 'denied') t.enqueue('GET', siteFields(rootUrl), { body: { value: outcome === 'missing' ? schema().filter(f => f.InternalName !== 'Value') : schema().map(f => f.InternalName === 'Value' ? { ...f, FieldTypeKind: 2 } : f) } });
    await assert.rejects(t.service.ensureListReady(rootUrl), { message: 'Failed to create Value field: Create failed. original' }); t.assertDrained();
  }
});

test('site permissions require each item bit and reject denied or malformed grant response', async () => {
  for (let omitted = 0; omitted < writeBits.length; omitted++) {
    const t = siteSetup(); t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`);
    t.enqueue('GET', `${siteList(rootUrl)}/effectiveBasePermissions`, { body: mask(...writeBits.filter((_, index) => index !== omitted), SPPermission.manageLists) });
    assert.equal(await t.service.canCurrentUserWrite(rootUrl), false); t.assertDrained();
  }
  for (const status of [403, 500]) {
    const t = siteSetup(); t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`);
    t.enqueue('GET', `${siteList(rootUrl)}/effectiveBasePermissions`, { status });
    assert.equal(await t.service.canCurrentUserWrite(rootUrl), false); t.assertDrained();
  }
});

test('site permissions accept minimal and verbose EffectiveBasePermissions envelopes but reject malformed words', async () => {
  for (const wrap of [bits => ({ EffectiveBasePermissions: bits }), bits => ({ d: { EffectiveBasePermissions: bits } })]) {
    for (const valid of [true, false]) {
      const t = siteSetup(); t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`);
      t.enqueue('GET', `${siteList(rootUrl)}/effectiveBasePermissions`, { body: wrap({ High: '0', Low: valid ? '14' : '14junk' }) });
      assert.equal(await t.service.canCurrentUserWrite(rootUrl), valid); t.assertDrained();
    }
  }
});

test('site respects SharePoint key collation and preserves server Title casing when retrieving and updating', async () => {
  const t = siteSetup(); probe(t);
  t.enqueue('GET', siteFind(rootUrl, 'feature'), { body: { value: [row('Feature')] } });
  assert.deepEqual(await t.service.get('feature', rootUrl), { id: 7, key: 'Feature', value: 123, description: 'old' });
  probe(t);
  t.enqueue('GET', siteFind(rootUrl, 'FEATURE'), { body: { value: [row('Feature')] } });
  t.enqueue('POST', `${siteList(rootUrl)}/items(7)`);
  await t.service.save('FEATURE', true, rootUrl);
  assertOptions(t.calls.at(-1), { Value: 'true', Description: 'old' }, mergeHeaders);
  t.assertDrained();
});

test('site rejects empty collection URLs before dispatching relative current-web HTTP', async () => {
  for (const url of ['', '   ', ' /// ']) {
    const t = siteSetup();
    for (const operation of [() => t.service.ensureListReady(url), () => t.service.get('key', url), () => t.service.list(url), () => t.service.save('key', 1, url), () => t.service.remove('key', url)]) {
      await assert.rejects(operation(), /site collection URL/i);
    }
    assert.equal(await t.service.canCurrentUserWrite(url), false);
    assert.equal(t.calls.length, 0); t.assertDrained();
  }
});

for (const stage of ['fields-denied', 'wrong-type', 'missing-title', 'field-create-denied', 'title-merge-denied', 'title-verify-denied']) {
  test(`site list recovery preserves original create error through ${stage} while successful create exposes schema failure`, async () => {
    for (const recovered of [true, false]) {
      const t = siteSetup();
      t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`, { status: 404 });
      t.enqueue('POST', `${rootUrl}/_api/web/lists`, recovered ? { status: 500, statusText: 'Original list failure', text: 'creation details' } : {});
      if (recovered) t.enqueue('GET', `${siteList(rootUrl)}?$select=Id`);
      const fields = stage === 'wrong-type' ? schema().map(field => field.InternalName === 'Value' ? { ...field, FieldTypeKind: 2 } : field)
        : stage === 'missing-title' ? schema().slice(1)
          : stage === 'field-create-denied' ? schema().filter(field => field.InternalName !== 'Value')
            : schema(false);
      t.enqueue('GET', siteFields(rootUrl), stage === 'fields-denied' ? { status: 403, statusText: 'Fields denied' } : { body: { value: fields } });
      if (stage === 'field-create-denied') {
        t.enqueue('POST', `${siteList(rootUrl)}/fields`, { status: 403, statusText: 'Field creation denied' });
      }
      if (stage === 'title-merge-denied' || stage === 'title-verify-denied') {
        t.enqueue('POST', `${siteList(rootUrl)}/fields/getByInternalNameOrTitle('Title')`, stage === 'title-merge-denied' ? { status: 403, statusText: 'Title denied' } : {});
      }
      if (stage === 'title-verify-denied') t.enqueue('GET', siteFields(rootUrl), { status: 403, statusText: 'Verification denied' });
      const expected = {
        'fields-denied': 'Failed to check list fields: Fields denied',
        'wrong-type': 'Incompatible Value field type in site store schema.',
        'missing-title': 'Missing Title field in site store schema.',
        'field-create-denied': 'Failed to create Value field: Field creation denied. ',
        'title-merge-denied': 'Failed to set Title uniqueness constraint: Title denied. ',
        'title-verify-denied': 'Failed to check list fields: Verification denied'
      };
      await assert.rejects(t.service.ensureListReady(rootUrl), { message: recovered ? 'Failed to create list: Original list failure. creation details' : expected[stage] });
      t.assertDrained();
      // A failed recovery must not leave readiness or a pending setup cached.
      probe(t); await t.service.ensureListReady(rootUrl); t.assertDrained();
    }
  });
}

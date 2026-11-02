const assert = require('node:assert/strict');
const { test } = require('node:test');
const { createStoreFixture, response, schema } = require('./fixtures/executable-examples/stores-harness.cjs');
const catalog = 'https://tenant.test/sites/catalog-a';
const row = (Value = '50', Id = 7, Title = 'StoredKey') => ({ Id, Title, Value, Description: 'Local metadata', '@odata.etag': '"3"' });
function protocol(siteRows = [], tenantRows = [], override = () => undefined) { return r => override(r) ?? (
  r.url.endsWith('/_api/SP_TenantSettings_Current') ? response({ CorporateCatalogUrl: catalog }) :
  r.url.includes('/fields?') ? response({ value: schema }) :
  r.url.includes('/items?') ? response({ value: r.url.includes('SiteKeyValueStore') ? siteRows : tenantRows }) : response({ Id: 'list' })); }
for (const kind of ['configurationHook', 'configurationService']) test(`${kind} import, drafts and Reset are request-free`, async t => {
  const f = createStoreFixture(t, kind, { revision: 7 }); await f.flush(); assert.equal(f.calls.length, 0); assert.equal(f.initializations.length + f.tokens.length + f.downloads.length, 0);
  f.change('Key', 'draft'); f.reset(); await f.flush(); assert.equal(f.calls.length, 0);
});
for (const [raw, expected] of [['false', false], ['0', 0], ['""', ''], ['', ''], ['null', null], [null, ''], ['{"nested":{"y":2}}', { nested: { y: 2 } }]]) test(`actual emitted layered Read preserves present site ${JSON.stringify(raw)} without tenant traffic`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.click('Run'); await f.drain(protocol([row(raw)], [row('{"a":1,"nested":{"x":1}}')]));
  assert.equal(f.dt('Source'), 'site'); assert.equal(f.dt('Selected key'), 'StoredKey'); assert.equal(f.dt('Description'), 'Local metadata'); assert.equal(f.dt('Physical item ID'), '7');
  assert.equal(f.dt('Selected value'), JSON.stringify(expected) + (expected === '' ? ' (empty string)' : '')); assert.equal(f.dt('Value type'), typeof expected);
  assert.ok(f.calls.every(r => r.url.startsWith('https://tenant.test/sites/clients-one/_api/web/lists/getByTitle(\'SiteKeyValueStore\')')));
});
for (const absentList of [false, true]) test(`only confirmed site ${absentList ? 'list' : 'key'} absence permits discovery and tenant fallback`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.click('Run');
  await f.drain(protocol([], [row('20', 9, 'TENANT')], r => absentList && r.url.includes('SiteKeyValueStore') ? response({}, 404) : undefined));
  assert.equal(f.dt('Source'), 'tenant'); assert.equal(f.dt('Selected value'), '20'); assert.equal(f.dt('Selected key'), 'TENANT'); assert.equal(f.dt('Physical item ID'), '9');
  assert.equal(f.calls.filter(r => r.url.endsWith('SP_TenantSettings_Current')).length, 1);
  assert.ok(f.calls.filter(r => r.url.includes('TenantKeyValueStore')).every(r => r.url.startsWith(catalog)));
});
for (const failure of ['403', 'schema', 'duplicate', 'item404']) test(`site ${failure} rejects without tenant fallback or absence success`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.click('Run'); await f.drain(protocol([], [], r => {
    if (failure === '403') return response({}, 403);
    if (failure === 'schema' && r.url.includes('/fields?')) return response({ value: schema.slice(1) });
    if (failure === 'duplicate' && r.url.includes('/items?')) return response({ value: [row(), row('1', 8)] });
    if (failure === 'item404' && r.url.includes('/items?')) return response({}, 404);
  }));
  assert.ok(f.container.querySelector('[role="alert"]')); assert.ok(!f.container.textContent.includes('not found.'));
  assert.ok(f.calls.every(r => !r.url.includes('TenantKeyValueStore') && !r.url.endsWith('SP_TenantSettings_Current')));
});
for (const failedPage of [false, true]) test(`physical inventory preserves duplicate order and ${failedPage ? 'rejects page failure without partial success' : 'complete pages'}`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.change('Operation', 'list'); f.click('Run');
  const next = 'https://tenant.test/sites/clients-one/_api/web/lists/getByTitle(\'SiteKeyValueStore\')/items?$skiptoken=2';
  await f.drain(protocol([], [], r => r.url.includes('/items?') ? r.url === next ? failedPage ? response({}, 403) : response({ value: [row('2', 8, 'feature')] }) : response({ value: [row('1', 7, 'Feature')], '@odata.nextLink': next }) : undefined));
  if (failedPage) { assert.ok(f.container.querySelector('[role="alert"]')); assert.ok(!f.container.textContent.includes('Loaded ')); }
  else { assert.match(f.container.textContent, /Loaded 2 site physical/); assert.deepEqual([...f.container.querySelector('[aria-label="First ten configuration records"]').querySelectorAll('dt')].map(n => n.textContent), ['Feature', 'feature']); }
  assert.ok(f.calls.every(r => r.url.includes('SiteKeyValueStore')));
});

const recipeCommand = kind => kind === 'configurationHook' ? 'Run default and override recipe' : 'Run service recipe';
function authored(value, predicate) {
  if (!value || typeof value !== 'object') return undefined;
  if (predicate(value.props ?? {})) return value.props;
  for (const child of Array.isArray(value) ? value : [value.props?.children]) { const found = authored(child, predicate); if (found) return found; }
  return undefined;
}
function control(f, label) { const found = authored(f.regions.controls, p => typeof p.onChange === 'function' && (p['aria-label'] === label || p.label === label)); assert.ok(found, label); return found.onChange; }
function consent(f) { f.check('I consent to these exact disposable mutations'); }
for (const kind of ['configurationHook', 'configurationService']) test(`${kind} recipe requires consent and preserves numeric defaults, provenance and site-only captured removal`, async t => {
  const f = createStoreFixture(t, kind); const command = recipeCommand(kind);
  f.click(command); await f.flush(); assert.equal(f.calls.length, 0);
  consent(f); f.click(command); let site = false, tenant = false;
  const route = protocol([], [], r => {
    if (r.method === 'POST') { const body = r.init.body ? JSON.parse(r.init.body) : {}; if (r.init.headers?.['X-HTTP-Method'] === 'DELETE') site = false; else if (r.url.includes('SiteKeyValueStore')) { assert.equal(body.Value, '50'); assert.equal(body.Description, 'Local override'); site = true; } else { assert.equal(body.Value, '20'); tenant = true; } return response({}); }
    if (r.url.includes('/items?')) return response({ value: r.url.includes('SiteKeyValueStore') ? site ? [row('50', 7)] : [] : tenant ? [row('20', 7)] : [] });
  });
  await f.drain(route); assert.match(f.container.textContent, /Saved tenant default 20/); assert.match(f.container.textContent, /Saved site override 50/); assert.equal(f.dt('Source'), 'site');
  assert.equal(f.calls.filter(r => r.method === 'POST').length, 2); assert.ok(f.find('Confirm removal'));
  f.click('Confirm removal'); await f.drain(route);
  assert.equal(f.dt('Source'), 'tenant'); assert.equal(f.dt('Selected value'), '20'); assert.equal(f.dt('Value type'), 'number');
  const writes = f.calls.filter(r => r.method === 'POST'); assert.equal(writes.length, 3); assert.ok(writes.at(-1).url.includes('SiteKeyValueStore')); assert.match(f.container.textContent, /Removed only the site override/);
  assert.ok(f.container.textContent.includes('Physical catalog URL is unavailable through this facade'));
});
for (const kind of ['configurationHook', 'configurationService']) for (const phase of ['early', 'terminal']) for (const rejected of [false, true]) test(`${kind} whole original public write ${phase} ${rejected ? 'reject' : 'fulfill'} keeps Reset locked and never chains`, async t => {
  const f = createStoreFixture(t, kind); const command = recipeCommand(kind), button = f.find(command), old = f.handler(command);
  consent(f); const approved = f.handler(command); f.click(command); await f.flush();
  if (phase === 'terminal') await f.drain(r => r.method === 'POST' ? undefined : protocol()(r));
  const held = f.calls.find(r => !r.settled); assert.ok(held); assert.equal(phase === 'terminal' ? held.method : held.url.endsWith('SP_TenantSettings_Current'), phase === 'terminal' ? 'POST' : true);
  f.reset(); assert.equal(f.find(command), button); assert.match(f.container.textContent, /Previous write still pending/); assert.equal(f.find(command).getAttribute('aria-disabled'), 'true'); const before = f.calls.length;
  await f.flush(() => { old(); approved(); }); f.click(command); assert.equal(f.calls.length, before);
  held.settled = true; await f.flush(() => held.resolve(rejected ? response({}, 403) : held.url.endsWith('SP_TenantSettings_Current') ? response({ CorporateCatalogUrl: catalog }) : response({})));
  await f.drain(protocol()); assert.ok(!f.container.textContent.includes('Previous write still pending')); assert.ok(!f.container.textContent.includes('Saved tenant default'));
  assert.ok(f.calls.every(r => !r.url.includes('SiteKeyValueStore')), 'no source-owned site step after obsolete tenant public Save');
  consent(f); f.click(command); await f.drain(r => r.method === 'POST' ? undefined : protocol()(r));
  assert.ok(f.calls.some(r => r.method === 'POST' && !r.settled), 'fresh consent still admits a real write');
  f.reset(); await f.drain(protocol());
});
for (const [label, next] of [['Mode', 'site'], ['Physical scope', 'tenant'], ['Operation', 'remove'], ['Key', 'spfx-toolkit-config-demo-B'], ['Value', '42'], ['Value format', 'json'], ['Replace/clear destination description', true], ['Description', 'New metadata']]) test(`rich ${label} edit synchronously retires retained consent and Run before React commit`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.change('Operation', 'save'); f.change('Value', '123');
  if (label === 'Description') f.check('Replace/clear destination description');
  consent(f); const edit = control(f, label), approve = control(f, 'I consent to these exact disposable mutations'), old = f.handler('Run');
  f.h.act(() => { edit(undefined, typeof next === 'boolean' ? { checked: next } : { value: next }); approve(undefined, { checked: true }); old(); assert.equal(f.calls.length, 0); });
  assert.equal(f.calls.length, 0); assert.equal(f.container.querySelector('[aria-label="Key"]').disabled, false);
  consent(f); f.click('Run'); if (label === 'Operation') f.click('Confirm removal'); await f.drain(protocol()); assert.ok(f.calls.length > 0);
});
for (const kind of ['configurationHook', 'configurationService']) test(`${kind} accepted ABA retires old controls, no-op remains usable and fresh A consent works`, async t => {
  const f = createStoreFixture(t, kind); if (kind === 'configurationHook') f.change('Operation', 'save');
  const command = kind === 'configurationHook' ? 'Run' : recipeCommand(kind), a = f.container.querySelector('[aria-label="Key"]').value;
  consent(f); const oldA = f.handler(command), consentA = control(f, 'I consent to these exact disposable mutations'); f.change('Key', 'spfx-toolkit-config-demo-B');
  consent(f); const oldB = f.handler(command), consentB = control(f, 'I consent to these exact disposable mutations'), editB = control(f, 'Key');
  f.h.act(() => { editB(undefined, { value: a }); consentA(undefined, { checked: true }); oldA(); consentB(undefined, { checked: true }); oldB(); }); assert.equal(f.calls.length, 0);
  const currentEdit = control(f, 'Key'); consent(f); const currentConsent = control(f, 'I consent to these exact disposable mutations'), run = f.handler(command);
  f.h.act(() => { currentEdit(undefined, { value: a }); currentConsent(undefined, { checked: true }); run(); }); await f.drain(r => r.method === 'POST' ? undefined : protocol()(r)); assert.ok(f.calls.some(r => r.method === 'POST'));
  f.reset(); await f.drain(protocol());
});

for (const kind of ['configurationHook', 'configurationService']) test(`${kind} synchronous duplicate recipe events admit only one original public write`, async t => {
  const f = createStoreFixture(t, kind); consent(f); const run = f.handler(recipeCommand(kind)); f.h.act(() => { run(); run(); }); await f.flush();
  assert.equal(f.calls.length, 1); assert.ok(f.calls[0].url.endsWith('SP_TenantSettings_Current')); f.reset(); await f.drain(protocol()); assert.equal(f.calls.filter(r => r.method === 'POST').length, 1);
});
for (const kind of ['configurationHook', 'configurationService']) for (const point of ['site-save', 'site-read', ...(kind === 'configurationService' ? ['inventory'] : []), 'confirmation', 'site-remove']) test(`${kind} Reset at ${point} ends the sequence without stale next source step`, async t => {
  const f = createStoreFixture(t, kind); let posts = 0;
  const route = r => {
    if (r.method === 'POST') { posts++; if (point === 'site-save' && posts === 2 || point === 'site-remove' && r.init.headers?.['X-HTTP-Method'] === 'DELETE') { posts--; return undefined; } return response({}); }
    if (posts >= 2 && r.url.includes('/items?') && (point === 'site-read' && !r.url.includes('$orderby') || point === 'inventory' && r.url.includes('$orderby'))) return undefined;
    return protocol(posts >= 2 ? [row()] : [])(r);
  };
  consent(f); f.click(recipeCommand(kind)); await f.drain(route);
  if (point === 'site-remove') { f.click('Confirm removal'); await f.drain(route); }
  const old = point === 'confirmation' ? f.handler('Confirm removal') : f.handler(recipeCommand(kind)), held = f.calls.find(r => !r.settled), count = f.calls.length;
  f.reset(); await f.flush(() => old()); assert.equal(f.calls.length, count);
  if (held) { held.settled = true; await f.flush(() => held.resolve(response(held.url.includes('/items?') ? { value: [row()] } : {}))); }
  await f.drain(protocol()); assert.ok(!f.container.textContent.includes('Saved tenant default')); assert.ok(!f.container.textContent.includes('Removed only'));
  if (point === 'site-save') assert.equal(f.calls.length, count, 'site Save settlement never starts an effective Read');
  if (point === 'site-read' || point === 'inventory' || point === 'confirmation') assert.equal(f.calls.filter(r => r.method === 'POST').length, 2, 'old read/dialog cannot initiate site removal');
  if (point === 'site-remove') assert.equal(f.calls.length, count, 'old removal cannot initiate final effective Read');
});
for (const kind of ['configurationHook', 'configurationService']) test(`${kind} host loss and actual unmount preserve original write settlement without publication`, async t => {
  const f = createStoreFixture(t, kind); consent(f); const old = f.handler(recipeCommand(kind)); f.click(recipeCommand(kind)); await f.flush(); const held = f.calls[0];
  f.changeScope(f.host('configuration-new')); const before = f.calls.length; await f.flush(() => old()); assert.equal(f.calls.length, before); assert.match(f.container.textContent, /Previous write still pending/);
  const errors = t.mock.method(console, 'error', () => {}); f.unmount(); held.settled = true; await f.flush(() => held.resolve(response({ CorporateCatalogUrl: catalog }))); await f.drain(protocol()); assert.equal(errors.mock.callCount(), 0); assert.ok(f.calls.every(r => !r.url.includes('SiteKeyValueStore')));
});
for (const mode of ['site', 'tenant']) test(`${mode} manual Save whole promise stays locked after Reset changes mode to layered`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.change('Mode', mode); f.change('Operation', 'save'); consent(f); f.click('Run'); await f.flush(); const held = f.calls[0];
  f.reset(); assert.equal(f.container.querySelector('[aria-label="Mode"]').value, 'layered'); assert.match(f.container.textContent, /Previous write still pending/);
  const count = f.calls.length; f.click('Run'); assert.equal(f.calls.length, count); held.settled = true; await f.flush(() => held.resolve(held.url.endsWith('SP_TenantSettings_Current') ? response({ CorporateCatalogUrl: catalog }) : response({})));
  await f.drain(protocol()); assert.ok(!f.container.textContent.includes('Previous write still pending')); assert.ok(!f.container.textContent.includes('Saved ')); assert.equal(f.calls.filter(r => r.method === 'POST').length, 1);
});
test('actual Configuration facade reuses shared property discovery and Site/Tenant native list queues', async t => {
  const path = require('node:path'); const { createHarness } = require('./services-test-harness.cjs');
  const h = createHarness({ '@microsoft/sp-core-library': { _SPKillSwitch: { isActivated: () => false } } }); t.after(h.close);
  const native = h.load(path.resolve(__dirname, 'fixtures/demo-presentation/spfx-services.ts'));
  const SPPermission = h.load(require.resolve('@microsoft/sp-page-context/lib-commonjs/SPPermission')).default;
  const f = createHarness({ '@microsoft/sp-http': native, '@microsoft/sp-page-context': { SPPermission }, './spfx-services': native }); t.after(f.close);
  const names = ['propertiesFixture', 'trimFixture', 'clientsFixture']; const previous = names.map(n => Object.getOwnPropertyDescriptor(window, n));
  t.after(() => names.forEach((n, i) => { if (previous[i]) Object.defineProperty(window, n, previous[i]); else delete window[n]; }));
  const transport = f.load(path.resolve(__dirname, 'fixtures/demo-presentation/trim-transports.ts')); const client = transport.nativeFixtureService(native.SPHttpClient.serviceKey, {});
  const store = f.load('services/spfx-configuration-store.service.ts').createSPFxConfigurationStoreService(client, { mode: 'layered', siteCollectionUrl: 'https://fixture.invalid/sites/root', discoveryWebUrl: 'https://fixture.invalid/sites/root/child' });
  assert.equal(transport.propertiesNative.requests.length + transport.storesNative.requests.length + transport.contextRequests.length + transport.clientsNative.requests.length, 0);
  transport.propertiesNative.configure({ catalogUrl: catalog }); assert.equal(await store.get('absent'), undefined);
  assert.deepEqual(transport.propertiesNative.requests.map(r => r.target), ['https://fixture.invalid/sites/root/child/_api/SP_TenantSettings_Current']);
  assert.equal(transport.storesNative.requests.filter(r => r.target.includes('SiteKeyValueStore')).length, 3); assert.equal(transport.storesNative.requests.filter(r => r.target.includes('TenantKeyValueStore')).length, 3);
  await store.save('spfx-toolkit-config-demo-native', 20, { scope: 'tenant' }); const post = transport.storesNative.requests.find(r => r.settings.method === 'POST'); assert.equal(JSON.parse(post.settings.options.body).Value, '20'); assert.ok(post.target.startsWith(catalog));
  assert.equal(transport.propertiesNative.requests.length, 1); assert.equal(transport.contextRequests.length, 0); assert.equal(transport.clientsNative.requests.length, 0);
});
for (const kind of ['configurationHook', 'configurationService']) test(`${kind} Reset retains actual facade/cache and explicit Read shows pending state`, async t => {
  const f = createStoreFixture(t, kind); const command = kind === 'configurationHook' ? 'Run' : 'Read effective value', button = f.find(command);
  f.click(command); await f.flush(); assert.ok([...f.container.querySelectorAll('[role="status"]')].some(node => node.textContent === 'Reading...'));
  await f.drain(protocol()); assert.equal(f.calls.filter(r => r.url.endsWith('SP_TenantSettings_Current')).length, 1);
  f.reset(); assert.equal(f.find(command), button); f.click(command); await f.drain(protocol()); assert.equal(f.calls.filter(r => r.url.endsWith('SP_TenantSettings_Current')).length, 1, 'unchanged dependency Reset retained the actual facade discovery cache');
});
test('malformed native site row rejects rather than enabling tenant fallback', async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.click('Run'); await f.drain(protocol([row(42)])); assert.ok(f.container.querySelector('[role="alert"]')); assert.ok(f.calls.every(r => !r.url.endsWith('SP_TenantSettings_Current')));
});
test('confirmed absence in both physical stores reports absence without provisioning', async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.click('Run'); await f.drain(protocol()); assert.match(f.container.textContent, /not found\./); assert.ok(f.calls.every(r => r.method === 'GET'));
});
for (const kind of ['configurationHook', 'configurationService']) test(`${kind} revision ABA and empty disposable suffix cannot revive old mutation admission`, async t => {
  const f = createStoreFixture(t, kind, { revision: 7 }); consent(f); const old = f.handler(recipeCommand(kind)); f.setRevision(8); f.setRevision(7); await f.flush(() => old()); assert.equal(f.calls.length, 0);
  f.change('Key', 'spfx-toolkit-config-demo-'); consent(f); await f.flush(() => f.handler(recipeCommand(kind))()); assert.equal(f.calls.length, 0);
  f.change('Key', 'production-setting'); consent(f); await f.flush(() => f.handler(recipeCommand(kind))()); assert.equal(f.calls.length, 0);
});

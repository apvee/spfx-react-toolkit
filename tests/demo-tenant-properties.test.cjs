const assert = require('node:assert/strict');
const { test } = require('node:test');
const { createStoreFixture, response } = require('./fixtures/executable-examples/stores-harness.cjs');
const version = 'spfx-toolkit-test-version', counter = 'spfx-toolkit-test-counter';
const catalogA = 'https://tenant.test/sites/catalog-a', catalogB = 'https://tenant.test/sites/catalog-b';
const propertyUrl = (catalog, key) => `${catalog}/_api/web/GetStorageEntity('${key}')`;
function protocol(catalog = catalogA, text = 'Current A') {
  return request => {
    if (request.url.endsWith('/_api/SP_TenantSettings_Current')) return response({ CorporateCatalogUrl: catalog });
    if (request.url === propertyUrl(catalog, version)) return response({ Value: JSON.stringify(text), Description: 'Version description' });
    if (request.url === propertyUrl(catalog, counter)) return response({ Value: '0', Description: 'Counter description' });
    throw new Error('Unexpected property protocol: ' + request.url);
  };
}
test('manual properties: zero mount/activation HTTP, two real discovery caches and fresh scalar-navigation session', async t => {
  const f = createStoreFixture(t, 'propertyHook', { revision: 7 });
  await f.flush(); assert.deepEqual(f.calls, []);
  const activate = f.find('Activate tenant properties'); f.click('Activate tenant properties'); await f.flush();
  assert.deepEqual(f.calls, []); assert.ok(f.find('Run'));
  f.click('Run'); await f.flush();
  assert.equal(f.calls.length, 2);
  assert.ok(f.calls.every(r => r.url === 'https://tenant.test/sites/clients-one/_api/SP_TenantSettings_Current'));
  await f.drain(protocol());
  assert.equal(f.dt(version), 'Current A'); assert.equal(f.dt(counter), '0');
  assert.equal(JSON.parse(f.dt('Version description')), 'Version description'); assert.equal(JSON.parse(f.dt('Counter description')), 'Counter description');
  const old = f.handler('Run'); f.click('Run'); await f.drain(protocol());
  assert.equal(f.calls.filter(r => r.url.endsWith('SP_TenantSettings_Current')).length, 2, 'warm independent caches');
  const sameClient = f.instance.native.services.get(f.keys.sp), samePage = f.instance.context.pageContext;
  samePage.web.absoluteUrl = 'https://tenant.test/sites/b'; f.render();
  assert.equal(f.find('Activate tenant properties'), activate); assert.ok(!f.container.textContent.includes('Current A'));
  const before = f.calls.length; await f.flush(() => old()); assert.equal(f.calls.length, before);
  assert.equal(f.instance.native.services.get(f.keys.sp), sameClient); assert.equal(f.instance.context.pageContext, samePage);
  f.click('Activate tenant properties'); await f.flush(); assert.equal(f.calls.length, before);
  f.click('Run'); await f.flush();
  assert.deepEqual(f.calls.slice(before).map(r => r.url), Array(2).fill('https://tenant.test/sites/b/_api/SP_TenantSettings_Current'));
  await f.drain(protocol(catalogB, 'Current B')); assert.equal(f.dt(version), 'Current B');
});

test('direct property factory validates unknown boolean after explicit known-catalog Read', async t => {
  const f = createStoreFixture(t, 'propertyService', { revision: 9 });
  await f.flush(); assert.equal(f.calls.length, 0); assert.equal(f.find('Read').disabled, true);
  f.change('Catalog URL', catalogA); f.change('Property key', "owner'&flag");
  f.click('Read'); await f.flush(); assert.equal(f.calls.length, 1);
  assert.equal(f.calls[0].url, propertyUrl(catalogA, 'owner%27%27%26flag'));
  assert.equal(f.calls[0].configuration, f.native.SPHttpClient.configurations.v1);
  await f.settle(0, { Value: 'false', Description: '' });
  assert.match(f.container.textContent, /Validated boolean: false/); assert.equal(f.dt('Description'), '""');
});

for (const identity of ['client', 'page', 'root', 'instance']) test(`Tenant Properties completed execution is invalidated by actual ${identity} identity`, async t => {
  const f = createStoreFixture(t, 'propertyHook'); f.click('Activate tenant properties'); f.click('Run'); await f.drain(protocol());
  const run = f.handler('Run'), before = f.calls.length;
  if (identity === 'root') { f.instance.context.pageContext.site.absoluteUrl += '-other'; f.render(); }
  else if (identity === 'instance') f.replace(f.host('property-B'));
  else {
    const next = f.host('clients-one');
    if (identity === 'client') next.native.services.set(f.keys.page, f.instance.context.pageContext);
    else next.native.services.set(f.keys.sp, f.instance.native.services.get(f.keys.sp));
    f.changeScope(next);
  }
  assert.ok(!f.container.textContent.includes('Current A')); await f.flush(() => run()); assert.equal(f.calls.length, before);
  f.click('Activate tenant properties'); assert.equal(f.calls.length, before); f.click('Run'); await f.drain(protocol(catalogB, 'Current B'));
  assert.equal(f.dt(version), 'Current B');
});
for (const stage of ['discovery', 'property']) for (const ending of ['resolve', 'reject']) test(`Tenant Properties obsolete pending ${stage} ${ending} cannot replace newer host result`, async t => {
  const f = createStoreFixture(t, 'propertyHook'); f.click('Activate tenant properties'); const old = f.handler('Run'); f.click('Run'); await f.flush();
  if (stage === 'property') await f.drain(r => r.url.endsWith('SP_TenantSettings_Current') ? protocol()(r) : undefined);
  const pending = f.calls.filter(r => !r.settled); assert.equal(pending.length, 2);
  f.replace(f.host('property-B')); f.click('Activate tenant properties'); f.click('Run');
  await f.drain(r => pending.includes(r) ? undefined : protocol(catalogB, 'Current B')(r));
  // The oldest held requests lead the queue; resolve only the new host explicitly.
  for (let step = 0; step < 6; step++) {
    const request = f.calls.find(r => !r.settled && !pending.includes(r)); if (!request) break;
    request.settled = true; await f.flush(() => request.resolve(protocol(catalogB, 'Current B')(request)));
  }
  assert.equal(f.dt(version), 'Current B');
  for (const request of pending) { request.settled = true; await f.flush(() => request[ending](ending === 'resolve' ? protocol()(request) : new Error('Obsolete A'))); }
  // An already-started load may continue from old discovery to old property HTTP.
  await f.drain(protocol()); const count = f.calls.length; await f.flush(() => old());
  assert.equal(f.calls.length, count); assert.equal(f.dt(version), 'Current B'); assert.ok(!f.container.textContent.includes('Obsolete A'));
});
for (const stage of ['discovery', 'property']) for (const ending of ['resolve', 'reject']) test(`Reset retains outer DOM and blocks old Run across held ${stage} ${ending}`, async t => {
  const f = createStoreFixture(t, 'propertyHook', { revision: 5 }); const activate = f.find('Activate tenant properties');
  f.click('Activate tenant properties'); const run = f.handler('Run'); f.click('Run'); await f.flush();
  if (stage === 'property') await f.drain(r => r.url.endsWith('SP_TenantSettings_Current') ? protocol()(r) : undefined);
  const held = f.calls.filter(r => !r.settled); assert.equal(held.length, 2);
  f.reset(); assert.equal(f.find('Activate tenant properties'), activate); assert.equal(activate.disabled, false);
  const count = f.calls.length; await f.flush(() => run()); assert.equal(f.calls.length, count);
  for (const request of held) { request.settled = true; await f.flush(() => request[ending](ending === 'resolve' ? protocol()(request) : new Error('Old failed'))); }
  await f.drain(protocol()); assert.ok(!f.container.textContent.includes('Current A')); assert.ok(!f.container.textContent.includes('Old failed'));
  f.click('Activate tenant properties'); f.click('Run'); await f.drain(protocol(catalogB, 'Reactivated')); assert.equal(f.dt(version), 'Reactivated');
});
for (const failedKey of [version, counter]) test(`current ${failedKey} failure beats retained data and successful dashboard on later render`, async t => {
  const f = createStoreFixture(t, 'propertyHook'); f.click('Activate tenant properties'); f.click('Run'); await f.drain(protocol());
  assert.equal(f.dt(counter), '0'); f.click('Run'); await f.drain(r => r.url === propertyUrl(catalogA, failedKey) ? response({}, 403) : protocol()(r));
  assert.equal(f.container.querySelector('[role="alert"]').textContent, 'Failed to read property: Forbidden');
  assert.equal(f.container.querySelector('[aria-label="Tenant properties"]'), null); f.render();
  assert.ok(!f.container.textContent.includes('Tenant property reads completed.')); assert.ok(!f.container.textContent.includes('Current A'));
  f.click('Run'); await f.drain(protocol()); assert.equal(f.dt(version), 'Current A');
});
test('latest Run owns reverse property completions and rejection; departure blocks retained Run', async t => {
  const f = createStoreFixture(t, 'propertyHook'); f.click('Activate tenant properties'); f.click('Run'); await f.drain(protocol());
  const run = f.handler('Run'); f.click('Run'); await f.flush(); const first = f.calls.filter(r => !r.settled); assert.equal(first.length, 2);
  await f.flush(() => run()); const next = f.calls.filter(r => !r.settled && !first.includes(r)); assert.equal(next.length, 2);
  for (const request of next) { request.settled = true; await f.flush(() => request.resolve(protocol(catalogA, 'Newest')(request))); }
  for (const request of first) { request.settled = true; await f.flush(() => request.reject(new Error('Stale read'))); }
  assert.equal(f.dt(version), 'Newest'); assert.ok(!f.container.textContent.includes('Stale read'));
  const current = f.handler('Run'), count = f.calls.length; f.unmount(); await f.flush(() => current()); assert.equal(f.calls.length, count);
  assert.equal(f.themeHandlers.size, 0);
});
test('actual property absence clears descriptions, while zero and empty metadata stay observable', async t => {
  const f = createStoreFixture(t, 'propertyHook'); f.click('Activate tenant properties'); f.click('Run');
  await f.drain(r => r.url.endsWith('SP_TenantSettings_Current') ? protocol()(r) : response({ Value: null, Description: 'Must disappear' }));
  assert.equal(f.dt(version), '(not set)'); assert.equal(f.dt('Version description'), '(not set)'); assert.equal(f.dt('Version ready'), 'false');
  f.click('Run'); await f.drain(r => response({ Value: '0', Description: '' }));
  assert.equal(f.dt(counter), '0'); assert.equal(f.dt('Counter ready'), 'true'); assert.equal(f.dt('Counter description'), '""');
});

const serviceCases = [
  ['true', { Value: 'true', Description: 'Enabled' }, 'Validated boolean: true', '"Enabled"'],
  ['false', { Value: 'false', Description: '' }, 'Validated boolean: false', '""'],
  ['string', { Value: '"false"' }, 'Expected a boolean', '(not set)'],
  ['raw string', { Value: 'hello' }, 'Expected a boolean', '(not set)'],
  ['object', { Value: '{"enabled":true}' }, 'Expected a boolean', '(not set)'],
  ['stored malformed JSON', { Value: '{broken' }, 'Expected a boolean', '(not set)'],
  ['decoded JSON null', { Value: 'null', Description: 'Retained' }, 'Expected a boolean', '"Retained"'],
  ['absent', { Description: 'Discarded' }, 'Property not set', '(not set)'],
  ['REST-null Value', { Value: null, Description: 'Discarded' }, 'Property not set', '(not set)'],
  ['legacy null Description', { Value: 'true', Description: null }, 'Validated boolean: true', '(legacy description unavailable)'],
];
for (const [name, body, message, description] of serviceCases) test(`actual direct property service: ${name}`, async t => {
  const f = createStoreFixture(t, 'propertyService'); f.change('Catalog URL', catalogA); f.click('Read'); await f.settle(0, body);
  assert.equal(f.container.querySelector('[role="status"]').textContent, message); assert.equal(f.dt('Description'), description);
  assert.equal(f.calls.length, 1); assert.equal(f.calls[0].init, undefined);
});
for (const status of [403, 404]) test(`actual direct property HTTP ${status} is failure after prior success, never absence`, async t => {
  const f = createStoreFixture(t, 'propertyService'); f.change('Catalog URL', catalogA); f.click('Read'); await f.settle(0, { Value: 'true' });
  f.click('Read'); await f.settle(1, {}, status); assert.ok(f.container.querySelector('[role="alert"]').textContent.startsWith('Failed to read property:'));
  assert.equal(f.container.querySelector('[aria-label="StorageEntity result"]'), null);
});
test('actual HTTP response JSON rejection is failure distinct from malformed stored text', async t => {
  const f = createStoreFixture(t, 'propertyService'); f.change('Catalog URL', catalogA); f.click('Read');
  await f.flush(() => f.calls[0].resolve({ ...response(), json: async () => { throw new SyntaxError('Malformed HTTP JSON'); } }));
  assert.equal(f.container.querySelector('[role="alert"]').textContent, 'Malformed HTTP JSON');
});
for (const change of ['catalog', 'key', 'reset', 'host', 'departure']) for (const ending of ['resolve', 'reject']) test(`direct Read ${change} invalidates retained callback and pending ${ending}`, async t => {
  const f = createStoreFixture(t, 'propertyService'); f.change('Catalog URL', catalogA); const read = f.handler('Read'); f.click('Read'); await f.flush(); const held = f.calls[0];
  if (change === 'catalog') f.change('Catalog URL', catalogB);
  else if (change === 'key') f.change('Property key', 'Other');
  else if (change === 'reset') { const input = f.container.querySelector('[aria-label="Catalog URL"]'); f.reset(); assert.equal(f.container.querySelector('[aria-label="Catalog URL"]'), input); assert.equal(input.value, ''); }
  else if (change === 'host') f.changeScope(f.host('other'));
  else f.unmount();
  await f.flush(() => read()); assert.equal(f.calls.length, 1);
  await f.flush(() => held[ending](ending === 'resolve' ? response({ Value: 'true' }) : new Error('Obsolete failed')));
  assert.ok(!f.container.textContent.includes('Validated boolean: true')); assert.ok(!f.container.textContent.includes('Obsolete failed'));
  if (change !== 'departure') { if (change === 'reset') f.change('Catalog URL', catalogA); f.click('Read'); await f.settle(1, { Value: 'false' }); assert.match(f.container.textContent, /Validated boolean: false/); }
});
test('direct Read latest operation and ABA target/reset generations remain current-only', async t => {
  const f = createStoreFixture(t, 'propertyService', { revision: 9 }); f.change('Catalog URL', catalogA); const original = f.handler('Read'); f.click('Read');
  await f.flush(() => original()); assert.equal(f.calls.length, 2); await f.settle(1, { Value: 'false' }); await f.settle(0, { Value: 'true' });
  assert.equal(f.container.querySelector('[role="status"]').textContent, 'Validated boolean: false');
  f.change('Property key', 'B'); f.change('Property key', 'my-app:featureEnabled'); await f.flush(() => original()); assert.equal(f.calls.length, 2);
  const beforeReset = f.handler('Read'); f.setRevision(10); f.setRevision(9); f.change('Catalog URL', catalogA); await f.flush(() => beforeReset()); assert.equal(f.calls.length, 2);
});
for (const kind of ['propertyHook', 'propertyService']) test(`${kind} portable fallback uses the actual provider and no initial I/O`, async t => {
  const f = createStoreFixture(t, kind, { fallback: true, directProvider: true, revision: 11 }); await f.flush(); assert.equal(f.calls.length, 0);
  assert.ok(f.find(kind === 'propertyHook' ? 'Activate tenant properties' : 'Read'));
});

test('property native fixture separates actual discovery and StorageEntity from Site/trim/Graph protocols', async t => {
  const path = require('node:path'); const { createHarness } = require('./services-test-harness.cjs');
  const h = createHarness({ '@microsoft/sp-core-library': { _SPKillSwitch: { isActivated: () => false } } }); t.after(h.close);
  const native = h.load(path.resolve(__dirname, 'fixtures/demo-presentation/spfx-services.ts'));
  const transport = createHarness({ '@microsoft/sp-http': native, './spfx-services': native }); t.after(transport.close);
  const globalNames = ['propertiesFixture', 'trimFixture', 'clientsFixture'];
  const previousGlobals = globalNames.map(name => Object.getOwnPropertyDescriptor(window, name));
  t.after(() => globalNames.forEach((name, index) => { if (previousGlobals[index]) Object.defineProperty(window, name, previousGlobals[index]); else delete window[name]; }));
  const seam = transport.load(path.resolve(__dirname, 'fixtures/demo-presentation/trim-transports.ts'));
  const client = seam.nativeFixtureService(native.SPHttpClient.serviceKey, {});
  const catalog = transport.load('services/spfx-app-catalog.service.ts').createSPFxAppCatalogService(client, { web: { absoluteUrl: 'https://fixture.invalid/sites/current' } });
  const factory = transport.load('services/spfx-tenant-property.service.ts').createSPFxTenantPropertyService;
  seam.propertiesNative.configure({ catalogUrl: catalogA, entities: { "owner'&flag": { Value: 'false', Description: '' } } });
  const service = factory(client); assert.equal(seam.propertiesNative.requests.length, 0);
  assert.equal(await catalog.discoverUrl(), catalogA); assert.equal(await catalog.discoverUrl(), catalogA);
  assert.deepEqual(await service.get("owner'&flag", catalogA), { value: false, description: '' });
  assert.deepEqual(seam.propertiesNative.requests.map(r => r.target), ['https://fixture.invalid/sites/current/_api/SP_TenantSettings_Current', propertyUrl(catalogA, 'owner%27%27%26flag')]);
  assert.ok(seam.propertiesNative.requests.every(r => r.settings.method === 'GET' && r.settings.configuration === native.SPHttpClient.configurations.v1));
  assert.equal(seam.storesNative.requests.length, 0); assert.equal(seam.contextRequests.length, 0); assert.equal(seam.clientsNative.requests.length, 0);
  seam.propertiesNative.configure({ hold: true }); const pending = service.get('held', catalogA);
  const request = seam.propertiesNative.requests.at(-1); request.resolve(new Response('{', { status: 200 })); await assert.rejects(pending, SyntaxError);
});
test('two actual provider instances keep independent discovery caches and property results', async t => {
  const f = createStoreFixture(t, 'propertyHook'); const peer = f.mountPeer();
  f.click('Activate tenant properties'); f.click('Run'); await f.drain(protocol());
  assert.ok(!peer.container.textContent.includes('Current A')); const firstClient = f.calls[0].client;
  peer.click('Activate tenant properties'); peer.click('Run'); await f.flush();
  const discoveries = f.calls.filter(r => !r.settled); assert.equal(discoveries.length, 2); assert.ok(discoveries.every(r => r.client !== firstClient));
  await f.drain(protocol(catalogB, 'Peer B')); assert.equal(f.dt(version), 'Current A'); assert.match(peer.container.textContent, /Peer B/);
  peer.unmount(); f.unmount(); assert.equal(f.themeHandlers.size, 0);
});
test('hook retained Run checks current URL scalars even before the next host rerender', async t => {
  const f = createStoreFixture(t, 'propertyHook'); f.click('Activate tenant properties'); const run = f.handler('Run');
  f.instance.context.pageContext.web.absoluteUrl += '-changed'; await f.flush(() => run()); assert.equal(f.calls.length, 0);
});

test('named target-guard negative control exposes a retained direct Read dispatch after target change', async t => {
  const f = createStoreFixture(t, 'propertyService', { propertyMutation: 'target' }); f.change('Catalog URL', catalogA); const old = f.handler('Read');
  f.change('Property key', 'B'); await f.flush(() => old()); assert.equal(f.calls.length, 1, 'the deliberately removed target guard admits stale work');
  await f.settle(0, { Value: 'false' });
});
test('named hook error negative control exposes retained data as a false successful dashboard', async t => {
  const f = createStoreFixture(t, 'propertyHook', { propertyMutation: 'errors' }); f.click('Activate tenant properties'); f.click('Run'); await f.drain(protocol());
  f.click('Run'); await f.drain(() => response({}, 403));
  assert.ok(f.container.textContent.includes('Tenant property reads completed.'), 'removing current error precedence exposes the regression');
  assert.equal(f.dt(version), 'Current A');
});

test('known-catalog inputs reject invalid URL shapes and empty keys without dispatch', async t => {
  const f = createStoreFixture(t, 'propertyService');
  for (const value of ['invalid', 'file:///catalog', 'https://tenant.test/sites/catalog?query=1', 'https://user:secret@tenant.test/sites/catalog']) {
    f.change('Catalog URL', value); f.click('Read'); await f.flush();
    assert.ok(f.container.querySelector('[role="alert"]')); assert.equal(f.calls.length, 0);
  }
  f.change('Catalog URL', catalogA); const read = f.handler('Read'); f.change('Property key', '');
  assert.equal(f.find('Read').disabled, true); await f.flush(() => read()); assert.equal(f.calls.length, 0);
});
test('hook reset revision ABA cannot resurrect the retained activation or Run session', async t => {
  const f = createStoreFixture(t, 'propertyHook', { revision: 9 }); const activation = f.handler('Activate tenant properties');
  f.click('Activate tenant properties'); const run = f.handler('Run'); f.click('Run'); await f.drain(protocol());
  const count = f.calls.length; f.setRevision(10); f.setRevision(9); await f.flush(() => { activation(); run(); }); assert.equal(f.calls.length, count);
  assert.equal(f.find('Activate tenant properties').disabled, false); f.click('Activate tenant properties'); f.click('Run'); await f.drain(protocol(catalogB, 'Fresh ABA'));
  assert.equal(f.dt(version), 'Fresh ABA');
});

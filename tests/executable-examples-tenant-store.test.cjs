const assert = require('node:assert/strict');
const { test } = require('node:test');
const { createStoreFixture, response, schema } = require('./fixtures/executable-examples/stores-harness.cjs');
const catalogA = 'https://tenant.test/sites/catalog-a', catalogB = 'https://tenant.test/sites/catalog-b';
const command = 'Save and read numeric-looking text';
function authored(value, predicate) {
  if (!value || typeof value !== 'object') return undefined;
  if (predicate(value.props ?? {})) return value.props;
  for (const child of Array.isArray(value) ? value : [value.props?.children]) { const found = authored(child, predicate); if (found) return found; }
  return undefined;
}
function protocol(catalog = catalogA, mask = 14) { return r => {
  if (r.url.endsWith('/_api/SP_TenantSettings_Current')) return response({ CorporateCatalogUrl: catalog });
  if (r.url.endsWith('/effectiveBasePermissions')) return response({ High: '0', Low: String(mask) });
  if (r.url.includes('/fields?')) return response({ value: schema });
  if (r.url.includes('/items?')) return response({ value: [] });
  return response({ Id: 'list' });
}; }
async function writable(t, kind, sequence = true) {
  const f = createStoreFixture(t, kind, { revision: 7 });
  if (kind === 'tenantHook') { f.click('Activate tenant store'); await f.drain(protocol()); f.change('Operation', sequence ? 'roundtrip' : 'save'); }
  else f.change('Catalog URL', catalogA);
  f.check('I consent to this exact disposable mutation'); return f;
}
test('actual TenantKV selection/reset/drafts are silent; activation alone discovers then checks effective grants', async t => {
  const f = createStoreFixture(t, 'tenantHook', { revision: 7 }); await f.flush(); assert.equal(f.calls.length, 0);
  f.reset(); await f.flush(); assert.equal(f.calls.length, 0);
  f.click('Activate tenant store'); await f.drain(protocol());
  assert.deepEqual(f.calls.map(r => r.url), ['https://tenant.test/sites/clients-one/_api/SP_TenantSettings_Current', `${catalogA}/_api/web/lists/getByTitle('TenantKeyValueStore')?$select=Id`, `${catalogA}/_api/web/lists/getByTitle('TenantKeyValueStore')/effectiveBasePermissions`]);
  assert.equal(f.dt('Can Write'), 'true'); assert.equal(f.dt('Client ready'), 'true');
  f.change('Key', 'draft'); assert.equal(f.calls.length, 3); assert.match(f.container.textContent, /Physical catalog URL is unavailable through this hook/);
});
test('hook draft edit immediately invalidates retained consent and Run before a batched React commit', async t => {
  const f = await writable(t, 'tenantHook', false); const run = f.handler('Run');
  const input = authored(f.regions.controls, p => p['aria-label'] === 'Key').onChange;
  const consent = authored(f.regions.controls, p => p.label === 'I consent to this exact disposable mutation').onChange;
  const count = f.calls.length;
  await f.flush(() => require('react-dom').unstable_batchedUpdates(() => { input(undefined, { value: 'spfx-toolkit-tenant-demo-changed' }); consent(undefined, { checked: true }); run(); }));
  assert.equal(f.calls.length, count, 'a draft edit retires handlers synchronously, before rendering commits');
});
test('actual direct hook warm cache stays A on same PageContext scalar mutation; wrapper explicitly remounts fresh B', async t => {
  const f = createStoreFixture(t, 'tenantHook');
  const Wrapper = ({ children }) => f.h.React.createElement(f.toolkit.SPFxWebPartProvider, { instance: f.instance }, children);
  const direct = f.h.mount(() => f.toolkit.useSPFxTenantKeyValueStore(), {}, Wrapper); t.after(direct.unmount);
  await f.drain(protocol()); const page = f.instance.context.pageContext, client = f.instance.native.services.get(f.keys.sp);
  page.web.absoluteUrl += '/B'; direct.render();
  let pending; await f.flush(() => { pending = direct.current.get('warm'); }); await f.drain(protocol()); await pending;
  assert.equal(f.calls.filter(r => r.url.endsWith('SP_TenantSettings_Current')).length, 1);
  assert.ok(f.calls.filter(r => r.url.includes('TenantKeyValueStore')).every(r => r.url.startsWith(catalogA)));
  direct.unmount();
  f.render(); f.click('Activate tenant store'); await f.drain(protocol()); const node = f.find('Activate tenant store'), old = f.handler('Run');
  page.web.absoluteUrl += '/C'; f.render(); const count = f.calls.length;
  await f.flush(() => old()); assert.equal(f.calls.length, count); assert.equal(f.find('Activate tenant store'), node);
  f.click('Activate tenant store'); await f.drain(protocol(catalogB));
  assert.equal(f.instance.context.pageContext, page); assert.equal(f.instance.native.services.get(f.keys.sp), client);
  assert.ok(f.calls.slice(count).some(r => r.url === page.web.absoluteUrl + '/_api/SP_TenantSettings_Current'));
  assert.ok(f.calls.slice(count).filter(r => r.url.includes('TenantKeyValueStore')).every(r => r.url.startsWith(catalogB)));
});
for (const kind of ['tenantHook', 'tenantService']) for (const stage of ['early', 'mutation']) for (const ending of ['resolve', 'reject']) test(`${kind} whole original write ticket survives Reset at ${stage} through ${ending} without next public read`, async t => {
  const f = await writable(t, kind); const old = f.handler(command), activation = kind === 'tenantHook' ? f.handler('Activate tenant store') : undefined;
  const node = f.find(kind === 'tenantHook' ? 'Activate tenant store' : command);
  f.click(command); await f.flush();
  if (stage === 'mutation') await f.drain(r => r.method === 'POST' ? undefined : protocol()(r));
  const held = f.calls.find(r => !r.settled); assert.ok(held); assert.equal(held.method, stage === 'mutation' ? 'POST' : 'GET');
  f.reset(); assert.equal(f.find(kind === 'tenantHook' ? 'Activate tenant store' : command), node); assert.match(f.container.textContent, /Previous write still pending/);
  const count = f.calls.length;
  await f.flush(() => { old(); activation?.(); f.handler(kind === 'tenantHook' ? 'Activate tenant store' : command)(); }); assert.equal(f.calls.length, count);
  held.settled = true; await f.flush(() => held[ending](ending === 'resolve' ? protocol()(held) : new Error('Retired failure')));
  if (stage === 'early' && ending === 'resolve') {
    assert.match(f.container.textContent, /Previous write still pending/);
    await f.drain(r => r.method === 'POST' ? undefined : protocol()(r)); assert.match(f.container.textContent, /Previous write still pending/);
  }
  await f.drain(protocol());
  assert.ok(!f.container.textContent.includes('Previous write still pending')); assert.ok(!f.container.textContent.includes('Retired failure')); assert.ok(!f.container.textContent.includes('Saved '));
  const snapshots = f.calls.filter(r => r.url.includes('/items?')).length;
  assert.equal(snapshots, ending === 'reject' && stage === 'early' ? 0 : 1, 'only internal Save snapshot; no next source public get');
  if (kind === 'tenantHook') { f.click('Activate tenant store'); await f.drain(protocol(catalogB)); f.change('Operation', 'save'); }
  else f.change('Catalog URL', catalogB);
  f.check('I consent to this exact disposable mutation'); f.click(kind === 'tenantHook' ? 'Run' : command);
  await f.drain(r => r.url.includes('/items?') && kind === 'tenantService' && f.calls.filter(x => x.method === 'POST').length > (ending === 'resolve' || stage === 'mutation' ? 1 : 0) ? response({ value: [{ Id: 1, Title: 'spfx-toolkit-tenant-demo-service', Value: '"123"' }] }) : protocol(catalogB)(r));
  assert.match(f.container.textContent, /Saved /);
});
test('actual tenant native transport separates catalog discovery from tenant-list GET/POST without generic trim payloads', async t => {
  const path = require('node:path'); const { createHarness } = require('./services-test-harness.cjs');
  const h = createHarness({ '@microsoft/sp-core-library': { _SPKillSwitch: { isActivated: () => false } } }); t.after(h.close);
  const native = h.load(path.resolve(__dirname, 'fixtures/demo-presentation/spfx-services.ts'));
  const SPPermission = h.load(require.resolve('@microsoft/sp-page-context/lib-commonjs/SPPermission')).default;
  const factoryHarness = createHarness({ '@microsoft/sp-http': native, '@microsoft/sp-page-context': { SPPermission }, './spfx-services': native }); t.after(factoryHarness.close);
  const names = ['propertiesFixture', 'trimFixture', 'clientsFixture']; const previous = names.map(n => Object.getOwnPropertyDescriptor(window, n));
  t.after(() => names.forEach((n, i) => { if (previous[i]) Object.defineProperty(window, n, previous[i]); else delete window[n]; }));
  const transport = factoryHarness.load(path.resolve(__dirname, 'fixtures/demo-presentation/trim-transports.ts'));
  const client = transport.nativeFixtureService(native.SPHttpClient.serviceKey, {});
  const service = factoryHarness.load('services/spfx-tenant-key-value-store.service.ts').createSPFxTenantKeyValueStoreService(client);
  const catalog = factoryHarness.load('services/spfx-app-catalog.service.ts').createSPFxAppCatalogService(client, { web: { absoluteUrl: 'https://fixture.invalid/sites/discovery' } });
  assert.equal(transport.propertiesNative.requests.length, 0); assert.equal(transport.storesNative.requests.length, 0);
  transport.propertiesNative.configure({ catalogUrl: catalogA }); assert.equal(await catalog.discoverUrl(), catalogA);
  assert.equal(await service.get('key', catalogA), undefined);
  await service.save('spfx-toolkit-tenant-demo-native', '123', catalogA);
  assert.equal(transport.propertiesNative.requests.length, 1);
  assert.ok(transport.storesNative.requests.every(r => r.target.startsWith(catalogA)));
  assert.equal(JSON.parse(transport.storesNative.requests.find(r => r.settings.method === 'POST').settings.options.body).Value, '"123"');
  assert.equal(transport.contextRequests.length, 0); assert.equal(transport.clientsNative.requests.length, 0);
});
for (const [mask, expected] of [[14, true], [12, false], [10, false], [6, false]]) test(`actual TenantKV SDK effective grants ${mask} yield ${expected}`, async t => {
  const f = createStoreFixture(t, 'tenantHook'); f.click('Activate tenant store'); await f.drain(protocol(catalogA, mask)); assert.equal(f.dt('Can Write'), String(expected)); assert.equal(f.calls.length, 3);
});
for (const [mask, expected] of [[14, false], [2062, true]]) test(`actual absent tenant list requires ManageLists with mask ${mask}`, async t => {
  const f = createStoreFixture(t, 'tenantHook'); f.click('Activate tenant store');
  await f.settle(0, { CorporateCatalogUrl: catalogA }); await f.settle(1, {}, 404); await f.settle(2, { High: '0', Low: String(mask) });
  assert.equal(f.dt('Can Write'), String(expected)); assert.equal(f.calls[2].url, catalogA + '/_api/web/effectiveBasePermissions');
});
test('Can Write false is advisory; a consented actual Save still dispatches and no readback is automatic', async t => {
  const f = createStoreFixture(t, 'tenantHook'); f.click('Activate tenant store'); await f.drain(protocol(catalogA, 6));
  f.change('Operation', 'save'); f.check('I consent to this exact disposable mutation'); f.click('Run'); await f.drain(protocol(catalogA, 6));
  assert.match(f.container.textContent, /Saved /); assert.equal(f.calls.filter(r => r.url.includes('/items?')).length, 1);
});
for (const kind of ['tenantHook', 'tenantService']) {
  const label = kind === 'tenantHook' ? 'Run' : command;
  test(`${kind} rejects same-turn duplicates and revoked/ABA target consent`, async t => {
    const f = await writable(t, kind, false); const run = f.handler(label);
    f.check('I consent to this exact disposable mutation', false); const count = f.calls.length; await f.flush(() => run()); assert.equal(f.calls.length, count);
    f.check('I consent to this exact disposable mutation'); const old = f.handler(label);
    const key = kind === 'tenantHook' ? 'spfx-toolkit-tenant-demo-disposable' : 'spfx-toolkit-tenant-demo-service';
    f.change('Key', key + '-B'); f.change('Key', key); await f.flush(() => old()); assert.equal(f.calls.length, count);
    f.check('I consent to this exact disposable mutation'); const fresh = f.handler(label); await f.flush(() => { fresh(); fresh(); }); await f.drain(protocol());
    assert.equal(f.calls.filter(r => r.method === 'POST').length, 1);
  });
  test(`${kind} original write remains locked across client ownership loss and outer unmount`, async t => {
    const f = await writable(t, kind); const old = f.handler(command); f.click(command);
    await f.drain(r => r.method === 'POST' ? undefined : protocol()(r)); const post = f.calls.at(-1);
    f.changeScope(f.host('tenant-new')); const count = f.calls.length; assert.match(f.container.textContent, /Previous write still pending/);
    await f.flush(() => { old(); f.handler(kind === 'tenantHook' ? 'Activate tenant store' : command)(); }); assert.equal(f.calls.length, count);
    const consoleError = t.mock.method(console, 'error', () => {}); f.unmount();
    post.settled = true; await f.flush(() => { old(); post.resolve(response({})); });
    assert.equal(f.calls.length, count); assert.equal(consoleError.mock.callCount(), 0); assert.equal(f.themeHandlers.size, 0);
  });
  test(`${kind} baseline nonzero reset, ABA revision, nonempty disposable suffix and fallback`, async t => {
    const f = createStoreFixture(t, kind, { revision: 9, fallback: true, directProvider: true }); await f.flush(); assert.equal(f.calls.length, 0);
    if (kind === 'tenantHook') { f.click('Activate tenant store'); await f.drain(protocol()); f.change('Operation', 'save'); } else f.change('Catalog URL', catalogA);
    f.check('I consent to this exact disposable mutation'); const old = f.handler(label); f.setRevision(10); f.setRevision(9); const count = f.calls.length; await f.flush(() => old()); assert.equal(f.calls.length, count);
    if (kind === 'tenantHook') { f.click('Activate tenant store'); await f.drain(protocol()); f.change('Operation', 'save'); } else f.change('Catalog URL', catalogA);
    for (const key of ['production-setting', 'spfx-toolkit-tenant-demo-']) { f.change('Key', key); assert.equal(f.find(label).disabled, true); }
  });
}
test('hook public Save can remain in discovery through Reset then finish internal setup without source Get', async t => {
  const f = createStoreFixture(t, 'tenantHook'); f.click('Activate tenant store'); f.change('Operation', 'roundtrip'); f.check('I consent to this exact disposable mutation'); f.click(command); await f.flush();
  const discoveries = f.calls.filter(r => r.url.endsWith('SP_TenantSettings_Current')); assert.equal(discoveries.length, 2); f.reset();
  const activate = f.handler('Activate tenant store'); await f.flush(() => activate()); assert.equal(f.calls.length, 2);
  for (const r of discoveries) { r.settled = true; await f.flush(() => r.resolve(response({ CorporateCatalogUrl: catalogA }))); }
  assert.match(f.container.textContent, /Previous write still pending/); await f.drain(r => r.method === 'POST' ? undefined : protocol()(r));
  assert.match(f.container.textContent, /Previous write still pending/); await f.drain(protocol());
  assert.equal(f.calls.filter(r => r.url.includes('/items?')).length, 1); assert.ok(!f.container.textContent.includes('Saved '));
});
test('tenant removal captures exact key/callback/ETag, closes before restoring focus and retires tokens', async t => {
  const { Simulate } = require('react-dom/test-utils'); const f = createStoreFixture(t, 'tenantHook'); f.click('Activate tenant store'); await f.drain(protocol());
  f.change('Operation', 'remove'); f.check('I consent to this exact disposable mutation'); const run = f.find('Run'); run.focus(); f.click('Run'); const old = f.handler('Confirm removal');
  assert.equal(f.container.querySelector('[aria-label="Key"]').disabled, true); f.click('Cancel'); assert.equal(document.activeElement, run);
  const count = f.calls.length; await f.flush(() => old()); assert.equal(f.calls.length, count);
  f.change('Key', 'spfx-toolkit-tenant-demo-exact'); f.check('I consent to this exact disposable mutation'); f.click('Run');
  f.h.act(() => Simulate.keyDown(document.querySelector('[role="dialog"]'), { key: 'Escape', code: 'Escape', keyCode: 27 })); await f.flush();
  assert.equal(document.querySelector('[role="dialog"]'), null); assert.equal(document.activeElement, run);
  f.click('Run'); const confirm = f.handler('Confirm removal'); await f.flush(() => old()); assert.equal(f.calls.length, count); f.click('Confirm removal');
  await f.drain(r => r.url.includes('/items?') ? response({ value: [{ Id: 8, Title: 'spfx-toolkit-tenant-demo-exact', Value: '"x"', '@odata.etag': '"4"' }] }) : r.method === 'POST' ? undefined : protocol()(r));
  const post = f.calls.at(-1); assert.equal(post.init.headers['If-Match'], '"4"'); assert.equal(post.init.headers['X-HTTP-Method'], 'DELETE'); assert.match(f.calls.find(r => r.url.includes('/items?')).url, /exact/);
  f.reset(); await f.flush(() => confirm()); assert.match(f.container.textContent, /Previous write still pending/); await f.drain(protocol()); assert.ok(!f.container.textContent.includes('Removed '));
});
test('retained Run checks live PageContext scalars before the next host render', async t => {
  const f = await writable(t, 'tenantHook', false); const run = f.handler('Run'); const count = f.calls.length;
  f.instance.context.pageContext.web.absoluteUrl += '-changed'; await f.flush(() => run()); assert.equal(f.calls.length, count);
});
test('actual TenantKV older failed read and permission cannot erase current record/indicator', async t => {
  const f = createStoreFixture(t, 'tenantHook'); f.click('Activate tenant store'); await f.drain(protocol()); const run = f.handler('Run');
  f.click('Run'); await f.drain(r => r.url.includes('/items?') ? undefined : protocol()(r)); await f.flush(() => run());
  for (let i = 0; i < 6; i++) { const r = f.calls.find(r => !r.settled && !r.url.includes('/items?')); if (!r) break; r.settled = true; await f.flush(() => r.resolve(protocol()(r))); }
  const items = f.calls.filter(r => !r.settled && r.url.includes('/items?')); assert.equal(items.length, 2);
  items[1].settled = true; await f.flush(() => items[1].resolve(response({ value: [{ Id: 2, Title: 'latest', Value: '"current"' }] })));
  await f.drain(r => !r.url.includes('/items?') ? protocol(catalogA, 6)(r) : undefined);
  items[0].settled = true; await f.flush(() => items[0].resolve(response({}, 403))); await f.drain(protocol(catalogA, 6));
  assert.equal(f.dt('Selected key'), 'latest'); assert.equal(f.dt('Can Write'), 'false'); assert.ok(!f.container.textContent.includes('Forbidden'));
});
test('latest actual tenant effective-permission response wins reverse completion', async t => {
  const f = createStoreFixture(t, 'tenantHook'); f.click('Activate tenant store'); await f.settle(0, { CorporateCatalogUrl: catalogA }); await f.settle(1, {});
  const oldPermission = f.calls[2]; assert.match(oldPermission.url, /effectiveBasePermissions$/); f.click('Run');
  for (let i = 0; i < 12; i++) {
    await f.flush(); const r = f.calls.find(r => r !== oldPermission && !r.settled); if (!r) break;
    r.settled = true; await f.flush(() => r.resolve(protocol(catalogA, 6)(r)));
  }
  assert.equal(f.dt('Can Write'), 'false'); oldPermission.settled = true; await f.flush(() => oldPermission.resolve(response({ High: '0', Low: '14' })));
  assert.equal(f.dt('Can Write'), 'false');
});
for (const identity of ['page', 'root', 'instance']) test(`TenantKV ${identity} change invalidates old commands and requires fresh discovery`, async t => {
  const f = await writable(t, 'tenantHook', false); const old = f.handler('Run');
  if (identity === 'root') { f.instance.context.pageContext.site.absoluteUrl += '-other'; f.render(); }
  else if (identity === 'instance') f.replace(f.host('tenant-other'));
  else { const next = f.host('clients-one'); next.native.services.set(f.keys.sp, f.instance.native.services.get(f.keys.sp)); f.changeScope(next); }
  const count = f.calls.length; await f.flush(() => old()); assert.equal(f.calls.length, count);
  f.click('Activate tenant store'); await f.drain(protocol(catalogB)); assert.ok(f.calls.slice(count).some(r => r.url.endsWith('SP_TenantSettings_Current')));
});
test('direct service rejects missing/unsafe known catalog input without discovery or writes', async t => {
  const f = createStoreFixture(t, 'tenantService'); const old = f.handler(command); await f.flush(() => old()); assert.equal(f.calls.length, 0);
  for (const url of ['invalid', 'file:///catalog', 'https://user:secret@tenant.test/catalog', 'https://tenant.test/catalog?query=1']) {
    f.change('Catalog URL', url); f.check('I consent to this exact disposable mutation'); f.click(command); await f.flush(); assert.equal(f.calls.length, 0); assert.ok(f.container.querySelector('[role="alert"]'));
  }
});
for (const ending of ['host', 'unmount']) test(`pending exact tenant confirmation cannot dispatch after ${ending}`, async t => {
  const f = await writable(t, 'tenantHook', false); f.change('Operation', 'remove'); f.check('I consent to this exact disposable mutation'); f.click('Run'); const confirm = f.handler('Confirm removal');
  if (ending === 'host') f.changeScope(f.host('tenant-confirm-B')); else f.unmount();
  const count = f.calls.length; await f.flush(() => confirm()); assert.equal(f.calls.length, count);
});
test('held public Save discovery rejection retires only its original ticket after Reset', async t => {
  const f = createStoreFixture(t, 'tenantHook'); f.click('Activate tenant store'); f.change('Operation', 'roundtrip'); f.check('I consent to this exact disposable mutation'); f.click(command); await f.flush();
  assert.equal(f.calls.length, 2); f.reset(); assert.match(f.container.textContent, /Previous write still pending/);
  for (const r of f.calls) { r.settled = true; await f.flush(() => r.reject(new Error('Obsolete discovery failed'))); }
  assert.ok(!f.container.textContent.includes('Previous write still pending')); assert.ok(!f.container.textContent.includes('Obsolete discovery failed')); assert.equal(f.calls.length, 2);
});
test('failed automatic discovery does not claim catalog readiness; explicit current Read retries', async t => {
  const f = createStoreFixture(t, 'tenantHook'); f.click('Activate tenant store'); await f.settle(0, {}, 403);
  assert.equal(f.dt('Client ready'), 'true'); assert.equal(f.dt('Can Write'), 'false'); assert.ok(!f.container.textContent.includes('catalog ready'));
  f.click('Run'); await f.drain(protocol()); assert.equal(f.calls.filter(r => r.url.endsWith('SP_TenantSettings_Current')).length, 2); assert.match(f.container.textContent, /not found\./);
});
for (const field of ['Catalog URL', 'Key']) test(`direct service batched ${field} edit retires captured consent/command before commit and fresh target works`, async t => {
  const f = await writable(t, 'tenantService'); const run = f.handler(command);
  const input = authored(f.regions.controls, p => p['aria-label'] === field).onChange;
  const consent = authored(f.regions.controls, p => p.label === 'I consent to this exact disposable mutation').onChange;
  const next = field === 'Catalog URL' ? catalogB : 'spfx-toolkit-tenant-demo-batched';
  await f.flush(() => require('react-dom').unstable_batchedUpdates(() => { input(undefined, { value: next }); consent(undefined, { checked: true }); run(); }));
  assert.equal(f.calls.length, 0, 'retained approval cannot resurrect the superseded exact target');
  assert.equal(f.container.querySelector('input[type="checkbox"]').checked, false);
  f.check('I consent to this exact disposable mutation'); f.click(command); await f.drain(protocol(field === 'Catalog URL' ? catalogB : catalogA));
  const post = f.calls.find(r => r.method === 'POST'); assert.ok(post, 'fresh target remains writable after fresh consent');
  assert.ok(f.calls.every(r => r.url.startsWith(field === 'Catalog URL' ? catalogB : catalogA)));
  assert.equal(JSON.parse(post.init.body).Title, field === 'Key' ? next : 'spfx-toolkit-tenant-demo-service');
});

const assert = require('node:assert/strict');
const { test } = require('node:test');
const path = require('node:path');
const { createStoreFixture, response, schema } = require('./fixtures/executable-examples/stores-harness.cjs');
function protocol(mask = 14) { return r => {
  if (r.url.endsWith('/effectiveBasePermissions')) return response({ High: '0', Low: String(mask) });
  if (r.url.includes('/fields?')) return response({ value: schema });
  if (r.url.includes('/items?')) return response({ value: [] });
  if (r.method === 'POST') return response({});
  return response({ Id: 'list' });
}; }
for (const [mask, expected] of [[14, true], [12, false], [10, false], [6, false]]) {
  test(`actual SDK hook permission mask ${mask} yields ${expected} only after explicit activation`, async t => {
    const f = createStoreFixture(t); await f.flush(); assert.equal(f.calls.length, 0);
    f.click('Activate site store'); await f.drain(protocol(mask));
    assert.equal(f.dt('Can Write'), String(expected));
    assert.equal(f.calls.length, 2); assert.ok(f.calls.every(r => r.method === 'GET'));
    assert.ok(!f.calls.some(r => /fields|\/items/.test(r.url)));
  });
}
for (const [mask, expected] of [[14, false], [2062, true]]) test(`actual SDK absent-list root mask ${mask} yields ${expected}`, async t => {
  const f = createStoreFixture(t); f.click('Activate site store');
  await f.settle(0, {}, 404); await f.settle(1, { High: '0', Low: String(mask) });
  assert.equal(f.dt('Can Write'), String(expected)); assert.match(f.calls[1].url, /\/_api\/web\/effectiveBasePermissions$/);
});
test('native browser SDK permission shim implements the actual required members and bit semantics', t => {
  const { createHarness } = require('./services-test-harness.cjs'); const h = createHarness(); t.after(h.close);
  const { SPPermission } = h.load(path.resolve(__dirname, 'fixtures/demo-presentation/spfx-services.ts'));
  assert.equal(typeof SPPermission.prototype.hasAllPermissions, 'function');
  for (const [mask, expected] of [[14, true], [12, false], [10, false], [6, false]]) assert.equal(new SPPermission({ High: 0, Low: mask }).hasAllPermissions(SPPermission.addListItems, SPPermission.editListItems, SPPermission.deleteListItems), expected);
});
for (const kind of ['hook', 'service']) for (const ending of ['resolve', 'reject']) test(`${kind} original public write ticket survives Reset until ${ending}`, async t => {
  const f = createStoreFixture(t, kind); const oldActivate = kind === 'hook' ? f.handler('Activate site store') : undefined; if (kind === 'hook') { f.click('Activate site store'); await f.drain(protocol()); f.change('Operation', 'save'); }
  else { f.click('Read'); await f.drain(protocol()); }
  f.check('I consent to this exact disposable mutation');
  const old = f.handler(kind === 'hook' ? 'Run' : 'Save');
  f.click(kind === 'hook' ? 'Run' : 'Save');
  await f.drain(r => r.method === 'POST' ? undefined : protocol()(r));
  const write = f.calls.find(r => r.method === 'POST' && !r.settled); assert.ok(write);
  const control = f.find(kind === 'hook' ? 'Activate site store' : 'Save');
  f.reset(); assert.equal(f.find(kind === 'hook' ? 'Activate site store' : 'Save'), control);
  assert.match(f.container.textContent, /Previous write still pending/);
  const count = f.calls.length;
  await f.flush(() => { old(); oldActivate?.(); f.handler(kind === 'hook' ? 'Activate site store' : 'Save')(); });
  assert.equal(f.calls.length, count);
  write.settled = true; await f.flush(() => write[ending](ending === 'resolve' ? response({}) : new Error('Old rejected')));
  assert.ok(!f.container.textContent.includes('Old rejected')); assert.ok(!f.container.textContent.includes('Saved ')); assert.ok(!f.container.textContent.includes('Previous write still pending'));
  assert.equal(f.calls.length, count);
  if (kind === 'hook') { f.click('Activate site store'); await f.drain(protocol()); f.change('Operation', 'save'); }
  else { f.click('Read'); await f.drain(protocol()); }
  f.check('I consent to this exact disposable mutation'); f.click(kind === 'hook' ? 'Run' : 'Save'); await f.drain(protocol());
  assert.equal(f.calls.filter(r => r.method === 'POST').length, 2); assert.match(f.container.textContent, /Saved /);
});
async function writable(t, kind = 'hook') {
  const f = createStoreFixture(t, kind);
  if (kind === 'hook') { f.click('Activate site store'); await f.drain(protocol()); f.change('Operation', 'save'); }
  else { f.click('Read'); await f.drain(protocol()); }
  f.check('I consent to this exact disposable mutation'); return f;
}
for (const kind of ['hook', 'service']) {
  test(`${kind} retained Save after consent revoked cannot dispatch`, async t => {
    const f = await writable(t, kind); const run = f.handler(kind === 'hook' ? 'Run' : 'Save');
    f.check('I consent to this exact disposable mutation', false); const count = f.calls.length;
    await f.flush(() => run()); assert.equal(f.calls.length, count);
  });
  test(`${kind} same-turn repeated authored handler writes once and publishes the current success`, async t => {
    const f = await writable(t, kind); const run = f.handler(kind === 'hook' ? 'Run' : 'Save');
    await f.flush(() => { run(); run(); }); await f.drain(protocol());
    assert.equal(f.calls.filter(r => r.method === 'POST').length, 1); assert.match(f.container.textContent, /Saved /);
  });
}
test('hook Reset makes an actual fresh setup cache while outer activation node survives', async t => {
  const f = await writable(t); const activate = f.find('Activate site store');
  f.click('Run'); await f.drain(protocol());
  const first = f.calls.filter(r => r.url.includes('/fields?')).length; assert.equal(first, 1);
  f.check('I consent to this exact disposable mutation'); f.click('Run'); await f.drain(protocol());
  assert.equal(f.calls.filter(r => r.url.includes('/fields?')).length, first, 'warm service cache skips setup');
  f.reset(); assert.equal(f.find('Activate site store'), activate); f.click('Activate site store'); await f.drain(protocol());
  f.change('Operation', 'save'); f.check('I consent to this exact disposable mutation'); f.click('Run'); await f.drain(protocol());
  assert.equal(f.calls.filter(r => r.url.includes('/fields?')).length, first + 1, 'new actual service repeats setup at same client/root');
});
for (const kind of ['hook', 'service']) test(`${kind} early native setup stage held across Reset retains ticket through later internal public RPCs`, async t => {
  const f = await writable(t, kind); f.click(kind === 'hook' ? 'Run' : 'Save');
  const first = f.calls.at(-1); assert.equal(first.method, 'GET'); f.reset(); const count = f.calls.length;
  assert.match(f.container.textContent, /Previous write still pending/);
  await f.drain(r => r.method === 'POST' ? undefined : protocol()(r));
  assert.ok(f.calls.length > count); assert.match(f.container.textContent, /Previous write still pending/);
  const post = f.calls.at(-1); assert.equal(post.method, 'POST'); post.settled = true; await f.flush(() => post.resolve(response({})));
  assert.ok(!f.container.textContent.includes('Previous write still pending')); assert.ok(!f.container.textContent.includes('Saved '));
});
for (const kind of ['hook', 'service']) test(`${kind} pending-ledger removal negative control exposes Reset lock failure`, async t => {
  const f = createStoreFixture(t, kind, { siteMutation: true });
  if (kind === 'hook') { f.click('Activate site store'); await f.drain(protocol()); f.change('Operation', 'save'); } else { f.click('Read'); await f.drain(protocol()); }
  f.check('I consent to this exact disposable mutation'); f.click(kind === 'hook' ? 'Run' : 'Save');
  await f.drain(r => r.method === 'POST' ? undefined : protocol()(r)); assert.ok(f.calls.some(r => r.method === 'POST'));
  f.reset(); assert.throws(() => assert.match(f.container.textContent, /Previous write still pending/), /Previous write still pending/);
  await f.drain(protocol());
});
for (const body of [{ High: '0', Low: '14junk' }, {}]) test('actual permission malformed mask fails closed without CRUD', async t => {
  const f = createStoreFixture(t); f.click('Activate site store'); await f.settle(0, {}); await f.settle(1, body); assert.equal(f.dt('Can Write'), 'false'); assert.equal(f.calls.length, 2);
});
test('actual permission HTTP 403 fails closed without provisioning', async t => {
  const f = createStoreFixture(t); f.click('Activate site store'); await f.settle(0, {}, 403); assert.equal(f.dt('Can Write'), 'false'); assert.equal(f.calls.length, 1);
});
for (const kind of ['hook', 'service']) test(`${kind} host scope loss keeps original write locked and stale closures cannot dispatch`, async t => {
  const f = await writable(t, kind); const old = f.handler(kind === 'hook' ? 'Run' : 'Save'); f.click(kind === 'hook' ? 'Run' : 'Save');
  await f.drain(r => r.method === 'POST' ? undefined : protocol()(r)); const post = f.calls.at(-1);
  f.changeScope(f.host('site-B')); const count = f.calls.length; assert.match(f.container.textContent, /Previous write still pending/);
  await f.flush(() => { old(); f.handler(kind === 'hook' ? 'Activate site store' : 'Save')(); }); assert.equal(f.calls.length, count);
  post.settled = true; await f.flush(() => post.resolve(response({}))); assert.ok(!f.container.textContent.includes('Saved '));
  if (kind === 'hook') { f.click('Activate site store'); await f.drain(protocol()); assert.ok(f.calls.slice(count).every(r => r.url.includes('/site-B/_api/'))); }
});
test('direct service requires a current true observed indicator, target-specific consent and catches rejecting Read', async t => {
  const f = createStoreFixture(t, 'service'); await f.flush(); assert.equal(f.calls.length, 0);
  const old = f.handler('Save'); await f.flush(() => old()); assert.equal(f.calls.length, 0);
  f.click('Read'); await f.drain(protocol(6)); assert.equal(f.dt('Can Write'), 'false');
  await f.flush(() => f.handler('Save')()); assert.equal(f.calls.filter(r => r.method === 'POST').length, 0);
  f.click('Read'); await f.drain(protocol()); f.check('I consent to this exact disposable mutation'); const allowed = f.handler('Save');
  f.change('Key', 'spfx-toolkit-site-demo-other'); await f.flush(() => allowed()); assert.equal(f.calls.filter(r => r.method === 'POST').length, 0);
  f.click('Read'); await f.drain(r => response({}, 403)); assert.match(f.container.textContent, /Forbidden/); assert.ok(!f.container.textContent.includes('not found.'));
});
test('hook consented mutation can dispatch despite advisory false indicator and does not auto-read', async t => {
  const f = createStoreFixture(t); f.click('Activate site store'); await f.drain(protocol(6)); assert.equal(f.dt('Can Write'), 'false');
  f.change('Operation', 'save'); f.check('I consent to this exact disposable mutation'); f.click('Run'); await f.drain(protocol(6));
  assert.match(f.container.textContent, /Saved /); assert.equal(f.calls.filter(r => r.url.includes('/items?')).length, 1, 'only Save snapshot, no automatic read');
});
test('captured exact remove, cancellation, newer dialog, Reset and departure guard tokens and focus', async t => {
  const f = createStoreFixture(t); f.click('Activate site store'); await f.drain(protocol());
  f.change('Operation', 'remove'); f.check('I consent to this exact disposable mutation'); const run = f.find('Run'); run.focus(); f.click('Run');
  assert.equal(f.container.querySelector('select').disabled, true); assert.equal(f.container.querySelector('[aria-label="Key"]').disabled, true);
  const old = f.handler('Confirm removal'); f.click('Cancel'); assert.ok(document.activeElement === run, 'Cancel restores the connected current trigger');
  await f.flush(() => old()); assert.equal(f.calls.length, 2);
  f.change('Key', 'spfx-toolkit-site-demo-exact'); f.check('I consent to this exact disposable mutation'); f.click('Run');
  await f.flush(() => old()); assert.equal(f.calls.length, 2);
  const current = f.handler('Confirm removal'); f.click('Confirm removal');
  await f.drain(r => r.url.includes('/items?') ? response({ value: [{ Id: 5, Title: 'spfx-toolkit-site-demo-exact', Value: '"x"', '@odata.etag': '"7"' }] }) : protocol()(r));
  const post = f.calls.find(r => r.method === 'POST'); assert.equal(post.init.headers['X-HTTP-Method'], 'DELETE'); assert.equal(post.init.headers['If-Match'], '"7"');
  assert.match(f.calls.find(r => r.url.includes('/items?')).url, /exact/); assert.match(f.container.textContent, /Removed spfx-toolkit-site-demo-exact/);
  const count = f.calls.length; await f.flush(() => current()); assert.equal(f.calls.length, count);
  f.check('I consent to this exact disposable mutation'); f.click('Run'); const after = f.handler('Confirm removal');
  f.reset(); await f.flush(() => after()); assert.equal(f.calls.length, count);
  f.click('Activate site store'); await f.drain(protocol()); f.change('Operation', 'remove'); f.check('I consent to this exact disposable mutation'); f.click('Run'); const departed = f.handler('Confirm removal'); f.unmount(); const final = f.calls.length; await f.flush(() => departed()); assert.equal(f.calls.length, final);
});
test('actual initial nonzero revision is not Reset; revision/target ABA cannot resurrect retained callback', async t => {
  const f = createStoreFixture(t, 'service', { revision: 9 }); f.click('Read'); await f.drain(protocol()); f.check('I consent to this exact disposable mutation'); const old = f.handler('Save');
  f.change('Key', 'spfx-toolkit-site-demo-B'); f.change('Key', 'spfx-toolkit-site-demo-service'); await f.flush(() => old()); assert.equal(f.calls.filter(r => r.method === 'POST').length, 0);
  f.setRevision(10); f.setRevision(9); await f.flush(() => old()); assert.equal(f.calls.filter(r => r.method === 'POST').length, 0);
});
test('native storage terminal has a separate full HTTP GET/POST queue and preserves original Graph/fetch behavior', async t => {
  const { createHarness } = require('./services-test-harness.cjs');
  const h = createHarness({ '@microsoft/sp-core-library': { _SPKillSwitch: { isActivated: () => false } } }); t.after(h.close);
  const native = h.load(path.resolve(__dirname, 'fixtures/demo-presentation/spfx-services.ts'));
  const SPPermission = h.load(require.resolve('@microsoft/sp-page-context/lib-commonjs/SPPermission')).default;
  // Use the actual public factory with the native fixture transport and actual SDK.
  const factoryHarness = createHarness({ '@microsoft/sp-http': native, '@microsoft/sp-page-context': { SPPermission }, './spfx-services': native }); t.after(factoryHarness.close);
  const transport = factoryHarness.load(path.resolve(__dirname, 'fixtures/demo-presentation/trim-transports.ts'));
  const sp = transport.nativeFixtureService(native.SPHttpClient.serviceKey, {});
  const service = factoryHarness.load('services/spfx-site-key-value-store.service.ts').createSPFxSiteKeyValueStoreService(sp);
  const root = 'https://fixture.invalid/sites/store';
  assert.equal(await service.canCurrentUserWrite(root), true); transport.storesNative.configure({ mask: { High: '0', Low: '6' } }); assert.equal(await service.canCurrentUserWrite(root), false);
  assert.equal(await service.get('key', root), undefined); assert.equal(transport.clientsNative.requests.length, 0); assert.equal(transport.clientsNative.initializations.length, 0);
  await service.save('spfx-toolkit-site-demo-native', '', root, '');
  const write = transport.storesNative.requests.find(r => r.settings.method === 'POST'); assert.ok(write); assert.equal(JSON.parse(write.settings.options.body).Value, '""');
  assert.equal(transport.clientsNative.downloads.length, 0);
});
for (const kind of ['hook', 'service']) test(`${kind} standalone fallback mounts with actual provider and no initial I/O`, async t => {
  const f = createStoreFixture(t, kind, { fallback: true, directProvider: true, revision: 11 }); await f.flush(); assert.equal(f.calls.length, 0); assert.ok(f.find(kind === 'hook' ? 'Activate site store' : 'Read'));
});
for (const ending of ['resolve', 'reject']) test(`captured actual Remove stays pending through Reset and ${ending} suppresses old publication`, async t => {
  const f = createStoreFixture(t); const activate = f.handler('Activate site store'); f.click('Activate site store'); await f.drain(protocol());
  f.change('Operation', 'remove'); f.check('I consent to this exact disposable mutation'); f.click('Run'); const confirm = f.handler('Confirm removal'); f.click('Confirm removal');
  await f.drain(r => r.method === 'POST' ? undefined : r.url.includes('/items?') ? response({ value: [{ Id: 8, Title: 'spfx-toolkit-site-demo-disposable', Value: '"x"', '@odata.etag': '"4"' }] }) : protocol()(r));
  const post = f.calls.at(-1); assert.equal(post.init.headers['X-HTTP-Method'], 'DELETE'); f.reset(); const count = f.calls.length;
  await f.flush(() => { confirm(); activate(); f.handler('Activate site store')(); }); assert.equal(f.calls.length, count); assert.match(f.container.textContent, /Previous write still pending/);
  post.settled = true; await f.flush(() => post[ending](ending === 'resolve' ? response({}) : new Error('Retired delete failed')));
  assert.ok(!f.container.textContent.includes('Removed ')); assert.ok(!f.container.textContent.includes('Retired delete failed')); assert.ok(!f.container.textContent.includes('Previous write still pending'));
});
for (const kind of ['hook', 'service']) test(`${kind} only a nonempty disposable suffix can be consented or written`, async t => {
  const f = await writable(t, kind); const old = f.handler(kind === 'hook' ? 'Run' : 'Save'); const before = f.calls.length;
  for (const key of ['production-setting', 'spfx-toolkit-site-demo-']) {
    f.change('Key', key); await f.flush(() => old()); assert.equal(f.calls.length, before);
    assert.equal(f.find(kind === 'hook' ? 'Run' : 'Save').disabled, true);
  }
});
test('confirmation Escape cancels exact token and restores focus after closed commit', async t => {
  const { Simulate } = require('react-dom/test-utils'); const f = createStoreFixture(t); f.click('Activate site store'); await f.drain(protocol());
  f.change('Operation', 'remove'); f.check('I consent to this exact disposable mutation'); const trigger = f.find('Run'); trigger.focus(); f.click('Run'); const confirm = f.handler('Confirm removal');
  const surface = document.querySelector('[role="dialog"]'); assert.ok(surface);
  f.h.act(() => Simulate.keyDown(surface, { key: 'Escape', code: 'Escape', keyCode: 27 })); await f.flush();
  assert.equal(document.querySelector('[role="dialog"]'), null); assert.ok(document.activeElement === trigger); await f.flush(() => confirm()); assert.equal(f.calls.length, 2);
});
test('actual native reverse permission completion cannot replace the later indicator or current record', async t => {
  const f = createStoreFixture(t); f.click('Activate site store'); f.click('Run');
  await f.settle(1, {}); await f.settle(2, { value: schema }); await f.settle(3, { value: [{ Id: 1, Title: 'current-record', Value: '"current"' }] });
  await f.settle(4, {}); await f.settle(5, { High: '0', Low: '6' }); assert.equal(f.dt('Can Write'), 'false');
  await f.settle(0, {}); await f.settle(6, { High: '0', Low: '14' }); assert.equal(f.dt('Can Write'), 'false'); assert.equal(f.dt('Selected key'), 'current-record');
});
test('actual native older failed read cannot erase a newer successful record or publish fallback success', async t => {
  const f = createStoreFixture(t); f.click('Activate site store'); await f.drain(protocol()); const run = f.handler('Run'); f.click('Run');
  await f.drain(r => r.url.includes('/items?') ? undefined : protocol()(r));
  await f.flush(() => run());
  for (let step = 0; step < 8; step++) {
    const request = f.calls.find(r => !r.settled && !r.url.includes('/items?')); if (!request) break;
    request.settled = true; await f.flush(() => request.resolve(protocol()(request)));
  }
  const reads = f.calls.filter(r => !r.settled && r.url.includes('/items?')); assert.equal(reads.length, 2);
  reads[1].settled = true; await f.flush(() => reads[1].resolve(response({ value: [{ Id: 2, Title: 'latest', Value: '"current result"' }] })));
  reads[0].settled = true; await f.flush(() => reads[0].resolve(response({}, 403))); await f.drain(protocol());
  assert.equal(f.dt('Selected key'), 'latest'); assert.equal(f.dt('Selected value'), '"current result"'); assert.ok(!f.container.textContent.includes('Forbidden')); assert.ok(!f.container.textContent.includes('not found.'));
});

// Retain the source-owned Checkbox callback, bypassing both disabled DOM and
// Fluent's live event trampoline, exactly as a consumer may retain an old prop.
function capturedServiceConsent(f) {
  function visit(value) {
    if (!value || typeof value !== 'object') return undefined;
    if (value.props?.label === 'I consent to this exact disposable mutation') return value.props.onChange;
    for (const child of Array.isArray(value) ? value : [value.props?.children]) { const found = visit(child); if (found) return found; }
    return undefined;
  }
  const change = visit(f.regions.controls); assert.equal(typeof change, 'function');
  return checked => change(undefined, { checked });
}
test('direct service retained Save cannot reuse a successful Read observation after the Save rejects', async t => {
  const f = await writable(t, 'service'); const savedHandler = f.handler('Save');
  f.click('Save'); await f.drain(r => r.method === 'POST' ? response({}, 403) : protocol()(r));
  assert.match(f.container.textContent, /Forbidden/); assert.equal(f.dt('Can Write'), '(not observed)'); assert.equal(f.find('Save').disabled, true);
  const count = f.calls.length; await f.flush(() => savedHandler());
  assert.equal(f.calls.length, count, 'a retained pre-failure Save cannot dispatch without a current observed permission');
});
for (const nextRead of ['denied', 'failed', 'allowed']) test(`direct service retained consent and Save cannot reuse an observation superseded by a newer ${nextRead} Read`, async t => {
  const f = await writable(t, 'service'); const consent = capturedServiceConsent(f), save = f.handler('Save');
  f.click('Read');
  await f.drain(nextRead === 'failed' ? () => response({}, 403) : protocol(nextRead === 'denied' ? 6 : 14));
  assert.equal(f.dt('Can Write'), nextRead === 'failed' ? '(not observed)' : String(nextRead === 'allowed'));
  assert.equal(f.container.querySelector('input[type="checkbox"]').checked, false, 'new Read retires consent');
  const count = f.calls.length;
  await f.flush(() => { consent(true); save(); });
  assert.equal(f.calls.length, count, 'the superseded observation cannot admit a native request');
  assert.equal(f.container.querySelector('input[type="checkbox"]').checked, false, 'captured consent cannot revive the superseded observation');
  if (nextRead === 'allowed') {
    f.check('I consent to this exact disposable mutation'); await f.flush(() => save());
    assert.equal(f.calls.length, count, 'fresh consent does not make a handler from an older observation current');
    f.click('Save'); await f.drain(protocol()); assert.equal(f.calls.filter(r => r.method === 'POST').length, 1);
    assert.match(f.container.textContent, /Saved /);
  }
});

const assert = require('node:assert/strict');
const { test } = require('node:test');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const titles = {
  useSPFxUserPhoto: 'Read the current user photo with managed URL cleanup',
  useSPFxOneDriveAppData: 'Read and explicitly replace OneDrive app data',
};
test('GraphData has exactly the two adopted owning source scenarios', () => {
  for (const [owner, title] of Object.entries(titles)) {
    const entry = scenario(title, `packages/spfx-react-toolkit/src/hooks/${owner}.ts#${owner}`).entry;
    assert.equal(entry.metadata.host, 'webPart');
    assert.equal(entry.metadata.kind, owner === 'useSPFxUserPhoto' ? 'read' : 'write');
    assert.equal(entry.metadata.contentPlacement, 'configuration');
  }
});
const { createGraphDataFixture } = require('./fixtures/executable-examples/graph-data-scenarios.cjs');
const drivePath = '/me/drive/special/approot:/spfx-react-toolkit-demo/test-data.json';
const downloadUrl = 'https://download.fixture.invalid/app-data';
const button = (f, name = 'Run') => [...f.container.querySelectorAll('button')].find(node => node.textContent === name);
async function activated(t, owner, options) {
  const f = createGraphDataFixture(t, owner, options);
  f.click('Start'); await f.ready();
  assert.ok(f.initializations.length > 0);
  assert.ok(f.initializations.every(request => request.target === '3'));
  assert.equal(f.calls.length, 0);
  return f;
}
async function readData(f, data, index = f.calls.length) {
  f.click('Run');
  await f.flush(() => f.calls[index].resolve({ '@microsoft.graph.downloadUrl': downloadUrl }));
  const download = f.downloads.at(-1);
  assert.deepEqual(download.init, { credentials: 'omit' });
  await f.flush(() => download.resolve(new Response(JSON.stringify(data), { status: 200 })));
}
for (const owner of Object.keys(titles)) {
  test(`${owner} selection rerender Reset and first revision are dormant before Start`, async t => {
    const f = createGraphDataFixture(t, owner, { revision: 7 });
    f.render(); f.reset(); await f.flush();
    assert.equal(f.initializations.length, 0); assert.equal(f.calls.length, 0); assert.equal(f.downloads.length, 0);
    assert.ok(button(f, 'Start'));
    f.click('Start'); await f.ready();
    assert.equal(f.calls.length, 0); assert.equal(f.downloads.length, 0);
    const init = f.initializations.length; f.reset(); await f.flush();
    assert.equal(f.initializations.length, init, 'Reset preserves activated child and native callbacks');
    assert.ok(!button(f, 'Start'));
  });
  test(`${owner} delayed and failed native factory cannot execute business I/O`, async t => {
    const f = createGraphDataFixture(t, owner);
    f.click('Start'); await f.flush();
    assert.equal(button(f).disabled, true); f.click('Run');
    await f.flush(() => f.initializations.forEach(r => r.reject(new Error('Factory denied'))));
    assert.equal(button(f).disabled, true); f.click('Run');
    assert.equal(f.calls.length, 0); assert.equal(f.downloads.length, 0);
    assert.match(f.container.textContent, /Factory denied/);
  });
  test(`${owner} retained actions do not resurrect after revision ABA and scope ABA`, async t => {
    const f = await activated(t, owner);
    const original = f.action('Run');
    const scopeA = f.instance.context.serviceScope;
    f.changeScope(f.host('graph-B')); await f.ready();
    f.changeScope({ context: { serviceScope: scopeA } }); await f.ready();
    await f.flush(() => { void original(); });
    assert.equal(f.calls.length, 0);
    const beforeReset = f.action('Run');
    f.setRevision(1); f.setRevision(0);
    await f.flush(() => { void beforeReset(); });
    assert.equal(f.calls.length, 0);
  });
  test(`${owner} portable standalone fallback uses actual provider and explicit Start`, async t => {
    const f = await activated(t, owner, { fallback: true, directProvider: true, revision: 9 });
    assert.ok(button(f)); assert.equal(f.calls.length, 0);
  });
}
test('photo uses default 240 request, displays 96 image, hides owned URL on Reset and releases on replacement/unmount', async t => {
  const creates = t.mock.method(URL, 'createObjectURL', (() => { let i = 0; return () => 'blob:photo-' + ++i; })());
  const revokes = t.mock.method(URL, 'revokeObjectURL', () => {});
  const f = await activated(t, 'useSPFxUserPhoto');
  assert.match(f.container.textContent, /Fixture User/);
  f.click('Run');
  assert.equal(f.calls[0].url, '/me/photos/240x240/$value');
  await f.flush(() => f.calls[0].resolve(new Blob(['photo-A'])));
  const image = f.container.querySelector('img');
  assert.equal(image.getAttribute('width'), '96'); assert.equal(image.getAttribute('height'), '96');
  assert.equal(image.alt, 'Fixture User'); assert.match(f.container.textContent, /7 bytes/);
  const oldAction = f.action('Run');
  f.reset(); await f.flush();
  assert.equal(f.container.querySelector('img'), null); assert.ok(!f.container.textContent.includes('7 bytes'));
  assert.equal(revokes.mock.callCount(), 0);
  await f.flush(() => { void oldAction(); }); assert.equal(f.calls.length, 1);
  f.click('Run'); f.reset();
  assert.equal(button(f).disabled, true);
  await f.flush(() => f.calls[1].resolve(new Blob(['hidden'])));
  assert.equal(creates.mock.callCount(), 2, 'current hook can create hidden owned URL after local Reset');
  assert.equal(f.container.querySelector('img'), null);
  assert.deepEqual(revokes.mock.calls.map(call => call.arguments[0]), ['blob:photo-1']);
  f.unmount(); assert.deepEqual(revokes.mock.calls.map(call => call.arguments[0]), ['blob:photo-1', 'blob:photo-2']);
});
for (const [status, text] of [[401, 'Authentication failed'], [403, 'Insufficient permissions'], [404, 'Photo not found']])
  test(`photo current ${status} is visible, Reset hides it and obsolete unmount cannot create a URL`, async t => {
    t.mock.method(console, 'error', () => {});
    const creates = t.mock.method(URL, 'createObjectURL', () => 'blob:unexpected');
    const f = await activated(t, 'useSPFxUserPhoto'); f.click('Run');
    await f.flush(() => f.calls[0].reject(Object.assign(new Error('Graph failed'), { statusCode: status })));
    assert.match(f.container.textContent, new RegExp(text)); f.reset();
    assert.ok(!f.container.textContent.includes(text)); f.click('Run'); f.unmount();
    await f.flush(() => f.calls[1].resolve(new Blob(['late'])));
    assert.equal(creates.mock.callCount(), 0);
  });
test('OneDrive missing read never writes and operation/draft edits perform no native work', async t => {
  const f = await activated(t, 'useSPFxOneDriveAppData');
  assert.equal(f.container.querySelector('select').value, 'read'); f.click('Run');
  assert.equal(f.calls[0].url, drivePath); assert.equal(f.calls[0].method, 'GET');
  assert.equal(f.calls[0].fields, undefined);
  await f.flush(() => f.calls[0].reject(Object.assign(new Error('Not found'), { statusCode: 404 })));
  assert.match(f.container.textContent, /File not found\./); assert.equal(f.downloads.length, 0);
  f.change('OneDrive operation', 'write');
  assert.equal(f.container.querySelector('input').value, 'Project update');
  f.change('Message', 'Changed'); assert.equal(f.calls.length, 1);
  assert.ok(!f.container.textContent.includes('File not found.'));
});
test('OneDrive native write has exact payload/header, locks through Reset and never clears the restored draft', async t => {
  const f = await activated(t, 'useSPFxOneDriveAppData');
  await readData(f, { message: 'Previous', counter: 4, timestamp: 1 });
  f.change('OneDrive operation', 'write'); f.change('Message', 'New message');
  t.mock.timers.enable({ apis: ['Date'], now: 123456 });
  f.click('Run');
  const request = f.calls[1];
  assert.equal(request.url, drivePath + ':/content'); assert.equal(request.method, 'PUT');
  assert.deepEqual(request.headers, { 'Content-Type': 'application/json' });
  assert.equal(request.body, '{"message":"New message","counter":5,"timestamp":123456}');
  assert.equal(f.container.querySelector('input').disabled, true); assert.equal(f.container.querySelector('select').disabled, true);
  f.reset();
  assert.equal(f.container.querySelector('select').disabled, true); assert.equal(button(f).disabled, true);
  await f.flush(() => request.resolve({}));
  assert.match(f.container.textContent, /Run the selected operation\./);
  f.change('OneDrive operation', 'write'); assert.equal(f.container.querySelector('input').value, 'Project update');
});
test('OneDrive current successful write alone clears draft, empty disables, whitespace remains valid', async t => {
  const f = await activated(t, 'useSPFxOneDriveAppData'); f.change('OneDrive operation', 'write'); f.click('Run');
  assert.equal(JSON.parse(f.calls[0].body).counter, 1);
  await f.flush(() => f.calls[0].resolve({}));
  assert.equal(f.container.querySelector('input').value, ''); assert.equal(button(f).disabled, true);
  assert.match(f.container.textContent, /Project update/);
  f.change('Message', '   '); assert.equal(button(f).disabled, false); f.click('Run');
  assert.equal(JSON.parse(f.calls[1].body).message, '   ');
  await f.flush(() => f.calls[1].resolve({}));
});
for (const outcome of ['reject', 'resolve'])
  test(`OneDrive old identity write ${outcome} cannot replace current read or erase edited draft`, async t => {
    t.mock.method(console, 'error', () => {});
    const f = await activated(t, 'useSPFxOneDriveAppData'); f.change('OneDrive operation', 'write');
    const stale = f.action('Run'); f.click('Run'); const old = f.calls[0];
    f.changeScope(f.host('drive-B')); await f.ready();
    f.change('OneDrive operation', 'read'); await readData(f, { message: 'Current service read', counter: 4, timestamp: 1 });
    assert.match(f.container.textContent, /Current service read/);
    await f.flush(() => old[outcome](outcome === 'reject' ? new Error('Obsolete failed') : {}));
    assert.match(f.container.textContent, /Current service read/); assert.ok(!f.container.textContent.includes('Obsolete failed'));
    f.change('OneDrive operation', 'write'); f.change('Message', 'Draft for the new service');
    await f.flush(() => { void stale(); });
    assert.equal(f.calls.length, 2); assert.equal(f.container.querySelector('input').value, 'Draft for the new service');
  });
test('OneDrive empty text/zero are present, failed downloads and reload errors never publish cached success', async t => {
  t.mock.method(console, 'error', () => {});
  const f = await activated(t, 'useSPFxOneDriveAppData');
  await readData(f, { message: '', counter: 0, timestamp: 1 });
  assert.match(f.container.textContent, /\(empty\)/); assert.match(f.container.textContent, /Counter\s*0/);
  for (const [body, status] of [['', 200], ['missing download', 404]]) {
    f.click('Run'); await f.flush(() => f.calls.at(-1).resolve({ '@microsoft.graph.downloadUrl': downloadUrl }));
    await f.flush(() => f.downloads.at(-1).resolve(new Response(body, { status })));
    assert.match(f.container.textContent, /Operation not completed\./);
    assert.ok(!f.container.textContent.includes('Saved message')); assert.ok(!f.container.textContent.includes('File not found.'));
  }
  assert.ok(f.calls.every(call => call.method === 'GET'));
});
test('OneDrive Reset invalidates a pending download without cancelling or enabling another read', async t => {
  const f = await activated(t, 'useSPFxOneDriveAppData'); f.click('Run');
  await f.flush(() => f.calls[0].resolve({ '@microsoft.graph.downloadUrl': downloadUrl }));
  const old = f.action('Run'); f.reset(); assert.equal(button(f).disabled, true);
  await f.flush(() => { void old(); }); assert.equal(f.calls.length, 1);
  await f.flush(() => f.downloads[0].resolve(new Response('{"message":"Late read","counter":2,"timestamp":1}')));
  assert.ok(!f.container.textContent.includes('Late read')); assert.match(f.container.textContent, /Run the selected operation\./);
});
test('OneDrive read does not expose stale writeError and departure suppresses late write publication', async t => {
  t.mock.method(console, 'error', () => {});
  const f = await activated(t, 'useSPFxOneDriveAppData'); f.change('OneDrive operation', 'write'); f.click('Run');
  await f.flush(() => f.calls[0].reject(new Error('Write denied'))); assert.match(f.container.textContent, /Write denied/);
  f.change('OneDrive operation', 'read'); await readData(f, { message: 'Read works', counter: 0, timestamp: 1 });
  assert.match(f.container.textContent, /Read works/); assert.ok(!f.container.textContent.includes('Write denied'));
  f.change('OneDrive operation', 'write'); f.click('Run'); const pending = f.calls.at(-1);
  f.select(titles.useSPFxUserPhoto, 'packages/spfx-react-toolkit/src/hooks/useSPFxUserPhoto.ts#useSPFxUserPhoto');
  assert.ok(button(f, 'Start')); await f.flush(() => pending.resolve({}));
  assert.ok(button(f, 'Start')); assert.equal(f.calls.length, 3);
});
test('OneDrive undefined successful read is distinct in an explicitly service-doubled fixture', async t => {
  const f = await activated(t, 'useSPFxOneDriveAppData', { oneDriveService: { read: async () => ({ data: undefined, isNotFound: false }), write: async () => {} } });
  f.click('Run'); await f.flush();
  assert.match(f.container.textContent, /No data returned\./); assert.equal(f.calls.length, 0);
});
test('photo actual client replacement owns latest result and obsolete success creates no URL', async t => {
  const creates = t.mock.method(URL, 'createObjectURL', () => 'blob:latest');
  const revokes = t.mock.method(URL, 'revokeObjectURL', () => {});
  const f = await activated(t, 'useSPFxUserPhoto'); f.click('Run'); const old = f.calls[0];
  f.changeScope(f.host('photo-B')); await f.ready(); f.click('Run');
  await f.flush(() => f.calls[1].resolve(new Blob(['B'])));
  await f.flush(() => old.resolve(new Blob(['Old A'])));
  assert.equal(creates.mock.callCount(), 1); assert.match(f.container.textContent, /1 bytes/);
  f.unmount(); assert.equal(revokes.mock.callCount(), 1);
});
test('OneDrive old successful write cannot erase a draft edited before completion on the new client', async t => {
  const f = await activated(t, 'useSPFxOneDriveAppData'); f.change('OneDrive operation', 'write'); f.click('Run');
  f.changeScope(f.host('draft-B')); await f.ready();
  f.change('Message', 'Draft for the new service');
  await f.flush(() => f.calls[0].resolve({}));
  assert.equal(f.container.querySelector('input').value, 'Draft for the new service');
  assert.ok(!f.container.textContent.includes('Saved message'));
});
test('existing integrated native terminal preserves profile projection and adds Blob metadata PUT and separate fake downloads', async t => {
  const { createHarness } = require('./services-test-harness.cjs');
  const path = require('node:path');
  const native = Object.fromEntries(['PageContext', 'HttpClient', 'SPHttpClient', 'MSGraphClientFactory', 'AadHttpClientFactory', 'AadTokenProviderFactory', 'ThemeProvider'].map(name => [name, { serviceKey: name }]));
  const fetchDescriptor = Object.getOwnPropertyDescriptor(window, 'fetch');
  Object.defineProperty(window, 'fetch', { configurable: true, writable: true, value: async () => { throw new Error('Unexpected external fetch'); } });
  t.after(() => { if (fetchDescriptor) Object.defineProperty(window, 'fetch', fetchDescriptor); else delete window.fetch; });
  const h = createHarness({ './spfx-services': native }); t.after(h.close);
  const transport = h.load(path.resolve(__dirname, 'fixtures/demo-presentation/trim-transports.ts'));
  assert.equal(transport.clientsNative.initializations.length, 0);
  const factory = transport.nativeFixtureService(native.MSGraphClientFactory.serviceKey, {});
  const graph = await factory.getClient('3');
  assert.equal(transport.clientsNative.initializations[0].target, '3');
  assert.equal(transport.clientsNative.requests.length, 0);
  const profile = await graph.api('/me').select('displayName,mail,userPrincipalName').get();
  assert.equal(profile.displayName, 'Fixture User'); assert.equal(profile.select, undefined);
  assert.deepEqual(transport.clientsNative.requests[0].settings, { select: 'displayName,mail,userPrincipalName' });
  const blob = await graph.api('/me/photos/240x240/$value').get(); assert.ok(blob instanceof Blob); assert.ok(blob.size > 0);
  const metadata = await graph.api(drivePath).get(); assert.equal(metadata['@microsoft.graph.downloadUrl'], downloadUrl);
  const body = '{"message":"Native saved","counter":1,"timestamp":1}';
  await graph.api(drivePath + ':/content').header('Content-Type', 'application/json').put(body);
  const write = transport.clientsNative.requests.at(-1);
  assert.deepEqual(write.settings, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body });
  transport.clientsNative.configure({ holdDownloads: true });
  const fetchPromise = window.fetch(downloadUrl, { credentials: 'omit' });
  const download = transport.clientsNative.downloads[0];
  assert.deepEqual(download.settings, { credentials: 'omit' });
  download.resolve(new Response(body, { status: 200 }));
  assert.equal(await (await fetchPromise).text(), body);
});
for (const owner of Object.keys(titles)) test(`${owner} direct same-turn commands dispatch once and Reset keeps pending lock while hiding late rejection`, async t => {
  t.mock.method(console, 'error', () => {});
  const f = await activated(t, owner);
  const run = f.action('Run');
  await f.flush(() => { run(); run(); });
  assert.equal(f.calls.length, 1);
  const init = f.initializations.length;
  f.reset(); assert.equal(button(f).disabled, true);
  await f.flush(() => { f.action('Run')(); }); assert.equal(f.calls.length, 1);
  await f.flush(() => f.calls[0].reject(new Error('Retired request failure')));
  assert.ok(!f.container.textContent.includes('Retired request failure'));
  assert.equal(button(f).disabled, false); assert.equal(f.initializations.length, init);
});
test('OneDrive Reset hides same-client write rejection without clearing the restored draft or cancelling work', async t => {
  t.mock.method(console, 'error', () => {});
  const f = await activated(t, 'useSPFxOneDriveAppData'); f.change('OneDrive operation', 'write'); f.click('Run');
  const pending = f.calls[0]; f.reset(); assert.equal(button(f).disabled, true);
  assert.equal(f.container.querySelector('select').disabled, true);
  await f.flush(() => pending.reject(new Error('Retired same-client write')));
  assert.ok(!f.container.textContent.includes('Retired same-client write'));
  f.change('OneDrive operation', 'write'); assert.equal(f.container.querySelector('input').value, 'Project update');
});
test('OneDrive legacy JSON null remains present without crashing or claiming a missing file', async t => {
  t.mock.method(console, 'error', () => {});
  const f = await activated(t, 'useSPFxOneDriveAppData');
  await readData(f, null); // JSON interoperability, not new application state or a validated shape.
  const result = f.container.querySelector('[data-demo-result]').textContent;
  assert.match(result, /Saved message/);
  assert.ok(!result.includes('File not found.')); assert.ok(!result.includes('No data returned.'));
  assert.equal(f.calls.length, 1);
});

// Legacy GraphData concerns transferred to actual owning source/lib scenarios.
// Native Graph factory/terminal/fetch alone are controlled; no subject hook doubles.
const assert = require('node:assert/strict');
const { test } = require('node:test');
const { createGraphDataFixture } = require('./fixtures/executable-examples/graph-data-scenarios.cjs');
const drivePath = '/me/drive/special/approot:/spfx-react-toolkit-demo/test-data.json';
const runButton = f => [...f.container.querySelectorAll('button')].find(node => node.textContent === 'Run');
async function fixture(t) {
  const f = createGraphDataFixture(t, 'useSPFxOneDriveAppData');
  assert.equal(f.initializations.length, 0);
  f.click('Start'); await f.ready();
  return f;
}
test('OneDrive missing-file read does not create; configuration is passive and explicit write locks its fixed target', async t => {
  const f = await fixture(t);
  assert.deepEqual(f.calls, []); f.click('Run');
  assert.equal(f.calls[0].url, drivePath);
  await f.flush(() => f.calls[0].reject(Object.assign(new Error('Missing'), { statusCode: 404 })));
  assert.match(f.container.querySelector('[data-demo-result]').textContent, /File not found/);
  assert.equal(f.downloads.length, 0);
  f.change('OneDrive operation', 'write');
  assert.equal(f.calls.length, 1);
  assert.ok(!f.container.querySelector('[data-demo-result]').textContent.includes('File not found'));
  f.click('Run'); assert.equal(f.calls.length, 2); assert.equal(f.calls[1].method, 'PUT');
  assert.equal(f.calls[1].url, drivePath + ':/content');
  assert.equal(f.container.querySelector('select').disabled, true);
  assert.equal(f.container.querySelector('input').disabled, true);
  assert.equal(runButton(f).disabled, true);
  await f.flush(() => f.calls[1].resolve({}));
  assert.equal(f.container.querySelector('select').disabled, false);
});
for (const state of ['initializing', 'missing']) test(`unavailable OneDrive client (${state}) cannot execute or claim a persisted default`, async t => {
  const f = createGraphDataFixture(t, 'useSPFxOneDriveAppData');
  f.click('Start'); await f.flush();
  if (state === 'missing') await f.flush(() => f.initializations.forEach(request => request.reject(new Error('Graph unavailable'))));
  assert.equal(runButton(f).disabled, true);
  await f.flush(() => { f.action('Run')(); });
  assert.equal(f.calls.length, 0);
  const result = f.container.querySelector('[data-demo-result]').textContent;
  assert.match(result, /Graph unavailable/); assert.ok(!result.includes('Saved message')); assert.ok(!result.includes('Counter'));
});
test('obsolete OneDrive write rejection cannot replace a newer read on another actual native client identity', async t => {
  const f = await fixture(t); f.change('OneDrive operation', 'write'); f.click('Run'); const old = f.calls[0];
  f.changeScope(f.host('legacy-B')); await f.ready(); f.change('OneDrive operation', 'read'); f.click('Run');
  await f.flush(() => f.calls[1].resolve({ '@microsoft.graph.downloadUrl': 'https://download.fixture.invalid/app-data' }));
  await f.flush(() => f.downloads[0].resolve(new Response('{"message":"Current service read","counter":4,"timestamp":1}')));
  await f.flush(() => old.reject(new Error('Obsolete write failed')));
  assert.ok(!f.container.textContent.includes('Obsolete write failed'));
  assert.match(f.container.querySelector('[data-demo-result]').textContent, /Current service read/);
  assert.notEqual(old.client, f.calls[1].client);
});
test('obsolete successful OneDrive write cannot erase a draft edited after Graph identity changes', async t => {
  const f = await fixture(t); f.change('OneDrive operation', 'write'); f.click('Run');
  f.changeScope(f.host('draft-B')); await f.ready(); f.change('Message', 'Draft for the new service');
  await f.flush(() => f.calls[0].resolve({}));
  assert.equal(f.container.querySelector('input').value, 'Draft for the new service');
  assert.ok(!f.container.querySelector('[data-demo-result]').textContent.includes('Saved message'));
});

const test = require('node:test');
const assert = require('node:assert/strict');
const { createTrimsFixture, single, nested, spGroup, entra, groupId } = require('./fixtures/executable-examples/trims-scenarios.cjs');
test('canonical permission workflow stages real native reads and protects lifetimes', async t => {
  const f = createTrimsFixture(t, single); await f.flush(); assert.equal(f.calls.length, 0);
  f.click(); await f.flush(); assert.ok(f.container.querySelector('[data-trim-state="loading"]')); assert.equal(f.diagnostics.mounts, 0);
  await f.settle(f.calls[0], { High: 0, Low: 1 }); f.click('Protected counter: 0'); assert.equal(f.diagnostics.counter, 1);
  const target = f.diagnostics.target; f.change('Web URL', 'https://tenant.test/sites/remote'); assert.deepEqual(f.diagnostics.target, target); assert.equal(f.calls.length, 1);
  f.click(); await f.flush(); await f.settle(f.calls[1], { High: 0, Low: 0 }); assert.ok(f.container.querySelector('[data-trim-state="denied"]')); assert.equal(f.diagnostics.active, false);
  f.change('Web URL', 'invalid URL'); f.click(); await f.flush(); assert.ok(f.container.querySelector('[data-trim-state="error"]')); assert.equal(f.calls.length, 2);
});
test('canonical nested AND starts Graph only after actual outer permission grant', async t => {
  const f = createTrimsFixture(t, nested); f.click(); await f.flush(); await f.settle(f.calls[0], { High: 0, Low: 0 });
  assert.equal(f.initializations.length, 0); assert.equal(f.diagnostics.mounts, 0);
  f.click(); await f.flush(); await f.settle(f.calls[1], { High: 0, Low: 1 }); assert.equal(f.initializations.length, 1);
  await f.initialize(); await f.settle(f.calls[2], { value: [{ id: groupId, displayName: 'Project Readers' }] });
  await f.settle(f.calls[3], { value: [groupId] }); assert.equal(f.diagnostics.mounts, 1); assert.equal(f.diagnostics.active, true);
});

test('canonical SharePoint groups stage native checks and retain direct-membership lifetime parity', async t => {
  const f = createTrimsFixture(t, spGroup); assert.equal(f.calls.length, 0);
  f.click(); await f.flush(); assert.ok(f.container.querySelector('[data-trim-state="loading"]'));
  await f.settle(f.calls[0], { value: [{ Id: 2, Title: 'Project Readers' }] });
  await f.settle(f.calls[1], { Id: 17 }); await f.settle(f.calls[2], { value: [{ Id: 17 }] });
  f.click('Protected counter: 0'); assert.equal(f.diagnostics.counter, 1); assert.equal(f.diagnostics.mounts, 1);
  const groups = f.diagnostics.groups; f.change('Group values', 'Other Readers'); assert.deepEqual(f.diagnostics.groups, groups); assert.equal(f.calls.length, 3);
  f.click(); await f.flush(); await f.settle(f.calls[3], { value: [{ Id: 3, Title: 'Other Readers' }] });
  await f.settle(f.calls[4], { Id: 17 }); await f.settle(f.calls[5], { value: [] }); assert.ok(f.container.querySelector('[data-trim-state="denied"]')); assert.equal(f.diagnostics.unmounts, 1);
  f.click(); await f.flush(); await f.reject(f.calls[6]); assert.ok(f.container.querySelector('[data-trim-state="error"]')); assert.equal(f.initializations.length, 0);
});

test('canonical Entra workflow stages native reads, preserves drafts and protected lifetimes, and separates denial/error', async t => {
  const f = createTrimsFixture(t, entra); assert.equal(f.calls.length, 0); assert.equal(f.initializations.length, 0);
  f.click(); await f.flush(); assert.ok(f.container.querySelector('[data-trim-state="loading"]')); assert.equal(f.diagnostics.mounts, 0);
  await f.initialize(); await f.settle(f.calls[0], { value: [{ id: groupId, displayName: 'Project Readers' }] }); await f.settle(f.calls[1], { value: [groupId] });
  f.click('Protected counter: 0'); assert.equal(f.diagnostics.counter, 1); assert.equal(f.diagnostics.mounts, 1);
  const applied = f.diagnostics.groups; f.change('Group values', 'Other Readers'); assert.deepEqual(f.diagnostics.groups, applied); assert.equal(f.diagnostics.counter, 1); assert.equal(f.calls.length, 2);
  f.click(); await f.flush(); await f.initialize(); await f.settle(f.calls[2], { value: [{ id: groupId, displayName: 'Other Readers' }] }); await f.settle(f.calls[3], { value: [] }); assert.ok(f.container.querySelector('[data-trim-state="denied"]')); assert.equal(f.diagnostics.unmounts, 1);
  f.click(); await f.flush(); await f.initialize(); await f.reject(f.calls[4]); assert.ok(f.container.querySelector('[data-trim-state="error"]')); assert.equal(f.diagnostics.active, false);
});

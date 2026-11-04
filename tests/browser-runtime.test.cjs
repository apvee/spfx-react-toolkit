const test = require('node:test');
const assert = require('node:assert/strict');
const { withBrowserCleanup } = require('./browser/helpers/browser-runtime.cjs');

test('server acquired before browser launch failure is closed and primary error survives', async () => {
  const primary = new Error('launch failed'); const events = [];
  await assert.rejects(withBrowserCleanup(async defer => { defer(() => events.push('server closed')); throw primary; }), error => error === primary);
  assert.deepEqual(events, ['server closed']);
});
test('browser close failure still closes server, with observable cleanup errors', async () => {
  const failure = new Error('browser close failed'); const events = [];
  await assert.rejects(withBrowserCleanup(async defer => { defer(() => events.push('server')); defer(() => { events.push('browser'); throw failure; }); }), error => error instanceof AggregateError && error.errors[0] === failure);
  assert.deepEqual(events, ['browser', 'server']);
});
test('report failure preserves primary and completes all teardown with secondary failures', async () => {
  const primary = new Error('assertion'); const report = new Error('report'); const close = new Error('close'); const events = [];
  await assert.rejects(withBrowserCleanup(async defer => {
    defer(() => events.push('server')); defer(() => { events.push('browser'); throw close; }); defer(() => { events.push('report'); throw report; }); throw primary;
  }), error => error instanceof AggregateError && error.cause === primary && error.errors.length === 3 && error.errors[0] === primary && error.errors[1] === report && error.errors[2] === close);
  assert.deepEqual(events, ['report', 'browser', 'server']);
});
test('awaits asynchronous LIFO disposers, registers immediately and disposes each exactly once', async () => {
  const events = []; let release; const gate = new Promise(resolve => { release = resolve; });
  const pending = withBrowserCleanup(async defer => {
    const dispose = defer(async () => { events.push('first start'); await gate; events.push('first end'); });
    defer(async () => { events.push('last'); });
    const immediate = dispose(); assert.equal(dispose(), immediate); release(); await immediate;
    return 42;
  });
  assert.equal(await pending, 42); assert.deepEqual(events, ['first start', 'first end', 'last']);
  const ordered = [];
  await withBrowserCleanup(async defer => { defer(() => ordered.push('server')); defer(async () => { ordered.push('browser start'); await new Promise(resolve => setImmediate(resolve)); ordered.push('browser end'); }); });
  assert.deepEqual(ordered, ['browser start', 'browser end', 'server']);
});
test('report-only failure rejects after resource cleanup; rejected undefined is preserved', async () => {
  const failure = new Error('report write'); const events = [];
  await assert.rejects(withBrowserCleanup(async defer => { defer(() => events.push('closed')); defer(() => { throw failure; }); }), error => error.errors[0] === failure);
  assert.deepEqual(events, ['closed']);
  let rejected = false; await withBrowserCleanup(async () => { throw undefined; }).catch(value => { rejected = true; assert.equal(value, undefined); }); assert.equal(rejected, true);
});

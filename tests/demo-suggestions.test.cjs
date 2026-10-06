const assert = require('node:assert/strict');
const path = require('node:path');
const { test } = require('node:test');
const { createHarness } = require('./services-test-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
require('@fluentui/react').registerIcons({ icons: { SearchBookmark: 'search', InfoSolid: 'info', Info: 'info' } });
const demo = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/demos/PnPSearchSuggestionsDemo.tsx');
function deferred() { let resolve, reject; const promise = new Promise((a, b) => { resolve = a; reject = b; }); return { promise, resolve, reject }; }
function fixture(t) {
  const requests = [];
  const suggest = query => { const request = { query, ...deferred() }; requests.push(request); return request.promise; };
  const h = createHarness({ '@apvee/spfx-react-toolkit': { useSPFxPnPSearch: () => ({ suggest }) } });
  t.after(h.close);
  t.mock.timers.enable({ apis: ['setTimeout'] });
  // In a browser window.setTimeout and the global clearTimeout share timer IDs.
  t.mock.method(window, 'setTimeout', (callback, delay) => setTimeout(callback, delay));
  const rendered = h.mountComponent(h.load(demo).PnPSearchSuggestionsDemo);
  function change(value) { h.act(() => { Simulate.change(rendered.container.querySelector('input'), { target: { value } }); }); }
  function debounce() { h.act(() => { t.mock.timers.tick(300); }); }
  async function resolve(index, values) { await h.act(async () => { requests[index].resolve(values); await requests[index].promise; }); }
  return { ...rendered, ...h, requests, change, debounce, resolve };
}
test('search suggestions keep the newest response when requests finish in reverse order', async t => {
  const f = fixture(t);
  f.change('old'); f.debounce();
  f.change('new'); f.debounce();
  assert.deepEqual(f.requests.map(r => r.query), ['old', 'new']);
  await f.resolve(1, ['new suggestion']);
  await f.resolve(0, ['old suggestion']);
  assert.ok(f.container.textContent.includes('new suggestion'));
  assert.ok(!f.container.textContent.includes('old suggestion'));
});
test('clearing search input invalidates pending suggestions immediately', async t => {
  const f = fixture(t);
  f.change('old'); f.debounce();
  f.change('');
  await f.resolve(0, ['old suggestion']);
  assert.ok(!f.container.textContent.includes('old suggestion'));
  f.debounce();
  assert.equal(f.requests.length, 1);
});
test('editing during debounce invalidates an already running suggestion request', async t => {
  const f = fixture(t);
  f.change('old'); f.debounce();
  f.change('new');
  await f.resolve(0, ['old suggestion']);
  assert.ok(!f.container.textContent.includes('old suggestion'));
  f.debounce();
  assert.equal(f.requests[1].query, 'new');
  await f.resolve(1, ['new suggestion']);
});
test('unmount cancels debounce and ignores an already running suggestion request', async t => {
  const f = fixture(t);
  f.change('old'); f.debounce();
  f.change('new'); f.unmount();
  const errors = t.mock.method(console, 'error', () => {});
  f.debounce();
  await f.resolve(0, ['old suggestion']);
  assert.equal(f.requests.length, 1);
  assert.equal(errors.mock.callCount(), 0);
});
test('selecting a suggestion cancels pending updates for the previous query', async t => {
  const f = fixture(t);
  f.change('old'); f.debounce();
  await f.resolve(0, ['selected suggestion']);
  f.change('new');
  const item = Array.from(f.container.querySelectorAll('div')).find(e => e.textContent === 'selected suggestion');
  f.act(() => { Simulate.click(item); });
  f.debounce();
  assert.equal(f.requests.length, 1);
  assert.equal(f.container.querySelector('input').value, 'selected suggestion');
});

const assert = require('node:assert/strict');
const path = require('node:path');
const { test } = require('node:test');
const { createHarness } = require('./services-test-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const demo = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/demos/PnPSearchSuggestionsDemo.tsx');
function deferred() { let resolve, reject; const promise = new Promise((a, b) => { resolve = a; reject = b; }); return { promise, resolve, reject }; }
function fixture(t) {
  const requests = [];
  const suggest = query => { const request = { query, ...deferred() }; requests.push(request); return request.promise; };
  const h = createHarness({ '@apvee/spfx-react-toolkit': { useSPFxPnPSearch: () => ({ suggest }) } });
  t.after(h.close);
  // JSDOM has no layout observer; messages still use the actual Fluent DOM.
  window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  // Tabster's Card root creates a UID while this test advances fake timers.
  Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  t.mock.timers.enable({ apis: ['setTimeout'] });
  // In a browser window.setTimeout and the global clearTimeout share timer IDs.
  t.mock.method(window, 'setTimeout', (callback, delay) => setTimeout(callback, delay));
  const Component = h.load(demo).PnPSearchSuggestionsDemo;
  const rendered = h.mountComponent(() => h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Component)));
  function change(value) { h.act(() => { Simulate.change(rendered.container.querySelector('input'), { target: { value } }); }); }
  function debounce() { h.act(() => { t.mock.timers.tick(300); }); }
  async function resolve(index, values) { await h.act(async () => { requests[index].resolve(values); await requests[index].promise; }); }
  async function reject(index, error) {
    await h.act(async () => {
      requests[index].reject(error);
      await requests[index].promise.catch(() => undefined);
    });
  }
  return { ...rendered, ...h, requests, change, debounce, resolve, reject };
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
  f.change('new'); f.debounce();
  await f.resolve(1, ['selected suggestion']);
  const item = Array.from(f.container.querySelectorAll('button')).find(e => e.textContent === 'selected suggestion');
  f.act(() => { Simulate.click(item); });
  await f.resolve(0, ['obsolete suggestion']);
  f.debounce();
  assert.equal(f.requests.length, 2);
  assert.equal(f.container.querySelector('input').value, 'selected suggestion');
  assert.ok(!f.container.textContent.includes('obsolete suggestion'));
});

test('autocomplete shows loading throughout debounce and the pending request, never an empty response', async t => {
  const f = fixture(t);
  f.change('ricevuta');
  assert.ok(f.container.textContent.includes('Loading suggestions...'));
  assert.ok(!f.container.textContent.includes('No suggestions available'));
  assert.equal(f.requests.length, 0);
  f.act(() => { t.mock.timers.tick(299); });
  assert.equal(f.requests.length, 0);
  f.act(() => { t.mock.timers.tick(1); });
  assert.equal(f.requests[0].query, 'ricevuta');
  assert.ok(f.container.textContent.includes('Loading suggestions...'));
  await f.resolve(0, ['ricevuta pagamento']);
  assert.ok(f.container.textContent.includes('ricevuta pagamento'));
  assert.ok(!f.container.textContent.includes('Loading suggestions...'));
  assert.ok(!f.container.textContent.includes('No suggestions available'));
});

test('an empty successful response explains SharePoint query history instead of indexed document volume', async t => {
  const f = fixture(t);
  f.change('ricevuta'); f.debounce();
  await f.resolve(0, []);
  assert.ok(f.container.textContent.includes('No suggestions available for this text.'));
  assert.ok(f.container.textContent.includes('search history and tenant configuration'));
  assert.ok(f.container.textContent.includes('Indexed documents alone do not guarantee suggestions.'));
  assert.ok(!f.container.textContent.includes('Loading suggestions...'));
  assert.ok(!f.container.textContent.includes('production environment'));
  assert.ok(!f.container.textContent.includes('Could not load suggestions'));
});

for (const failure of [new Error('Access denied'), 'unstructured failure']) {
  test(`autocomplete distinguishes request failure from an empty response: ${String(failure)}`, async t => {
    t.mock.method(console, 'error', () => {});
    const f = fixture(t);
    f.change('ricevuta'); f.debounce();
    await f.reject(0, failure);
    assert.ok(f.container.textContent.includes('Could not load suggestions.'));
    assert.ok(f.container.textContent.includes(failure instanceof Error ? 'Access denied' : 'The suggestion request failed.'));
    assert.ok(!f.container.textContent.includes('No suggestions available'));
    assert.ok(!f.container.textContent.includes('Loading suggestions...'));
    f.change('ri');
    assert.ok(!f.container.textContent.includes('Could not load suggestions'));
    assert.ok(f.container.textContent.includes('Loading suggestions...'));
    f.debounce();
    await f.resolve(1, ['ricevuta pagamento']);
    assert.ok(f.container.textContent.includes('ricevuta pagamento'));
    assert.ok(!f.container.textContent.includes('Loading suggestions...'));
  });
}

test('editing hides suggestions for the previous input until the current response arrives', async t => {
  const f = fixture(t);
  f.change('old'); f.debounce();
  await f.resolve(0, ['old suggestion']);
  f.change('new');
  assert.ok(!f.container.textContent.includes('old suggestion'));
  assert.ok(f.container.textContent.includes('Loading suggestions...'));
  f.debounce();
  await f.resolve(1, []);
  assert.ok(f.container.textContent.includes('No suggestions available'));
});

test('obsolete suggestion errors cannot replace current loading or successful suggestions', async t => {
  t.mock.method(console, 'error', () => {});
  const f = fixture(t);
  f.change('old'); f.debounce();
  f.change('older'); f.debounce();
  f.change('new');
  await f.reject(0, new Error('Obsolete failure during debounce'));
  assert.ok(f.container.textContent.includes('Loading suggestions...'));
  assert.ok(!f.container.textContent.includes('Could not load suggestions'));
  f.debounce();
  await f.resolve(2, ['new suggestion']);
  await f.reject(1, new Error('Obsolete failure after success'));
  assert.ok(f.container.textContent.includes('new suggestion'));
  assert.ok(!f.container.textContent.includes('Could not load suggestions'));
});

test('short input resets every feedback state and ignores a late error', async t => {
  t.mock.method(console, 'error', () => {});
  const f = fixture(t);
  f.change('old'); f.debounce();
  f.change('r');
  await f.reject(0, new Error('Late failure'));
  assert.ok(!f.container.textContent.includes('Loading suggestions...'));
  assert.ok(!f.container.textContent.includes('No suggestions available'));
  assert.ok(!f.container.textContent.includes('Could not load suggestions'));
  f.debounce();
  assert.equal(f.requests.length, 1);
  f.change('retry'); f.debounce();
  await f.reject(1, new Error('Current failure'));
  assert.ok(f.container.textContent.includes('Could not load suggestions'));
  f.change('');
  assert.ok(!f.container.textContent.includes('Could not load suggestions'));
  f.change('empty'); f.debounce();
  await f.resolve(2, []);
  assert.ok(f.container.textContent.includes('No suggestions available'));
  f.change('r');
  assert.ok(!f.container.textContent.includes('No suggestions available'));
});

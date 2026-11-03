const assert = require('node:assert/strict');
const path = require('node:path');
const { test } = require('node:test');
const { createHarness } = require('./services-test-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const demo = () => scenario('Request debounced suggestions after explicit start', 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPSearch.ts#useSPFxPnPSearch').file;
const ReactDOM = require('react-dom');
function deferred() { let resolve, reject; const promise = new Promise((a, b) => { resolve = a; reject = b; }); return { promise, resolve, reject }; }
function fixture(t, { autoStart = true, revision = 37, transform } = {}) {
  const requests = [];
  let suggest = query => { const request = { query, ...deferred() }; requests.push(request); return request.promise; };
  const context = { sp: {}, isInitialized: true };
  let hookError, regions;
  const h = createHarness({ '@apvee/spfx-react-toolkit': { useSPFxPnPSearch: () => ({ suggest, error: hookError }), useSPFxPnPContext: () => context } });
  t.after(h.close);
  // JSDOM has no layout observer; messages still use the actual Fluent DOM.
  window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  // Tabster's Card root creates a UID while this test advances fake timers.
  Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const timerFires = [];
  const setTimer = setTimeout;
  t.mock.method(global, 'setTimeout', (callback, delay, ...args) => setTimer(() => { if (delay === 300) timerFires.push(delay); callback(...args); }, delay));
  // In a browser window.setTimeout and the global clearTimeout share timer IDs.
  t.mock.method(window, 'setTimeout', (callback, delay) => setTimeout(callback, delay));
  const fs = require('node:fs'), selected = demo(); let Component;
  if (transform) { const read = fs.readFileSync; const mock = t.mock.method(fs, 'readFileSync', (file, ...args) => { const value = read(file, ...args); return String(file) === selected ? transform(value) : value; }); try { Component = h.load(selected).Demo; } finally { mock.mock.restore(); } } else Component = h.load(selected).Demo;
  function Root() { return h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Component, { resetRevision: revision, render: value => { regions = value; return h.React.createElement('div', {}, value.controls, value.commands, value.status, value.children, value.diagnostics); } })); }
  const rendered = h.mountComponent(Root);
  const rerender = () => h.act(() => { ReactDOM.render(h.React.createElement(Root), rendered.container); });
  const click = name => h.act(() => { const button = [...rendered.container.querySelectorAll('button')].find(b => b.textContent === name); assert.ok(button, name); Simulate.click(button); });
  if (autoStart) click('Start suggestions');
  function change(value) { h.act(() => { Simulate.change(rendered.container.querySelector('input'), { target: { value } }); }); }
  function debounce() { h.act(() => { t.mock.timers.tick(300); }); }
  async function resolve(index, values) { await h.act(async () => { requests[index].resolve(values); await requests[index].promise; }); }
  async function reject(index, error) {
    await h.act(async () => {
      requests[index].reject(error);
      await requests[index].promise.catch(() => undefined);
    });
  }
  return { ...rendered, ...h, requests, timerFires, change, debounce, resolve, reject, click, rerender, regions: () => regions, reset() { revision++; rerender(); }, replace() { const previous = suggest; suggest = query => previous(query); rerender(); }, hookError(error) { hookError = error; rerender(); } };
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

// Whole-hook double tier: real Fluent DOM, source controller and fake timers, not native transport.
function findProps(value, match) {
  if (Array.isArray(value)) { for (const child of value) { const found = findProps(child, match); if (found) return found; } return; }
  if (!value || typeof value !== 'object') return;
  return match(value.props ?? {}) ? value.props : findProps(value.props?.children, match);
}
const command = (f, name) => findProps(f.regions().commands, p => p.children === name).onClick;
const inputEdit = f => findProps(f.regions().controls, p => p['aria-label'] === 'Search with Autocomplete').onChange;
const choose = (f, name) => findProps(f.regions().children, p => p.children === name).onClick;
test('unarmed raw drafts are inert; Start schedules pretyped raw whitespace and repeated Start preserves 299+1 deadline', async t => {
  const f = fixture(t, { autoStart: false }); const input = f.container.querySelector('input'); input.focus();
  f.change('  '); f.debounce(); assert.equal(f.requests.length, 0); const start = command(f, 'Start suggestions'); f.click('Start suggestions');
  f.act(() => t.mock.timers.tick(299)); f.click('Start suggestions'); f.act(() => start()); f.change('  ');
  assert.equal(f.requests.length, 0); f.act(() => t.mock.timers.tick(1)); assert.deepEqual(f.requests.map(r => r.query), ['  ']);
  await f.resolve(0, []); assert.match(f.container.textContent, /No suggestions available/); assert.equal(f.container.querySelector('input') === input, true); assert.equal(document.activeElement === input, true);
});
test('short Start is idle and only the newest genuine edit owns a fresh 300ms deadline', async t => {
  const f = fixture(t, { autoStart: false }); f.change('r'); f.click('Start suggestions'); f.debounce(); assert.equal(f.requests.length, 0);
  f.change('raw'); f.act(() => t.mock.timers.tick(299)); f.change('raw2'); f.act(() => t.mock.timers.tick(299)); assert.equal(f.requests.length, 0); f.act(() => t.mock.timers.tick(1)); assert.equal(f.requests[0].query, 'raw2'); await f.resolve(0, []);
});
for (const boundary of ['stop', 'reset', 'method', 'unmount']) test(boundary + ' cancels local debounce and retires retained Start Stop edit callbacks', async t => {
  const f = fixture(t, { autoStart: false }); f.change('retained'); const start = command(f, 'Start suggestions'); f.click('Start suggestions'); const stop = command(f, 'Stop suggestions'), edit = inputEdit(f);
  if (boundary === 'stop') f.click('Stop suggestions'); else if (boundary === 'reset') f.reset(); else if (boundary === 'method') f.replace(); else f.unmount();
  f.act(() => { start(); stop(); edit({}, { value: 'obsolete' }); }); f.debounce(); assert.equal(f.requests.length, 0);
  if (boundary !== 'unmount') { assert.equal(f.container.querySelector('input').value, boundary === 'reset' ? '' : 'retained'); assert.doesNotMatch(f.container.textContent, /Loading suggestions|No suggestions available|Could not load/); f.click('Start suggestions'); if (boundary === 'reset') f.change('fresh'); f.debounce(); assert.equal(f.requests.length, 1); await f.resolve(0, []); }
});
for (const boundary of ['stop', 'reset', 'method']) for (const failure of [false, true]) test(boundary + ' retires held ' + (failure ? 'failure' : 'success') + ' and permits new explicit work', async t => {
  const f = fixture(t); const input = f.container.querySelector('input'); input.focus(); f.change('old'); f.debounce();
  if (boundary === 'stop') f.click('Stop suggestions'); else if (boundary === 'reset') f.reset(); else f.replace();
  assert.equal(f.container.querySelector('input') === input, true); assert.equal(document.activeElement === input, true);
  f.click('Start suggestions'); f.change('fresh'); f.debounce(); if (failure) await f.reject(0, new Error('Retired error')); else await f.resolve(0, ['Retired result']);
  assert.match(f.container.textContent, /Loading suggestions/); assert.doesNotMatch(f.container.textContent, /Retired/); await f.resolve(1, ['Fresh result']); assert.match(f.container.textContent, /Fresh result/);
});
test('selection leaves armed for the next genuine edit and old Select Stop Start cannot revive ABA or reset work', async t => {
  const f = fixture(t); f.change('same'); f.debounce(); await f.resolve(0, ['exact selection']); const select = choose(f, 'exact selection'), oldStop = command(f, 'Stop suggestions'), oldStart = command(f, 'Start suggestions'), edit = inputEdit(f);
  f.change('other'); f.change('same'); f.act(() => { select(); oldStop(); oldStart(); edit({}, { value: 'retired' }); }); assert.equal(f.container.querySelector('input').value, 'same');
  f.debounce(); await f.resolve(1, ['exact selection']); f.click('exact selection'); f.debounce(); assert.equal(f.requests.length, 2); assert.equal(f.container.querySelector('input').value, 'exact selection'); f.change('next edit'); f.debounce(); assert.equal(f.requests[2].query, 'next edit'); await f.resolve(2, []);
  const retained = command(f, 'Start suggestions'); f.reset(); f.act(() => retained()); f.debounce(); assert.equal(f.requests.length, 3);
});
test('equivalent method rerender preserves timer; obsolete hook error during debounce never becomes current failure', async t => {
  const f = fixture(t); f.change('before'); f.debounce(); f.change('after'); f.act(() => t.mock.timers.tick(299)); f.hookError(new Error('Old hook error')); f.rerender(); await f.reject(0, new Error('Obsolete rejection')); assert.match(f.container.textContent, /Loading suggestions/); assert.doesNotMatch(f.container.textContent, /Old hook error|Could not load/); f.act(() => t.mock.timers.tick(1)); assert.equal(f.requests[1].query, 'after'); await f.resolve(1, ['Current']);
});

for (const boundary of ['stop','reset','unmount']) test('local debounce is canceled rather than merely suppressing the callback after ' + boundary, t => {
  const f = fixture(t); f.change('scheduled'); if (boundary === 'stop') f.click('Stop suggestions'); else if (boundary === 'reset') f.reset(); else f.unmount(); f.debounce(); assert.equal(f.timerFires.length, 0); assert.equal(f.requests.length, 0);
});
test('fault control detects removal of local timer cancellation independently of publication guards', t => {
  const f = fixture(t, { transform: s => s.replace('clearTimeout(timer.current);', '') }); f.change('scheduled'); f.click('Stop suggestions'); f.debounce(); assert.throws(() => assert.equal(f.timerFires.length,0,'Stop cancels scheduled timer'), /Stop cancels/); assert.equal(f.requests.length,0,'remaining ownership defense still suppresses dispatch');
});
test('fault control detects Start restarting an already armed deadline', async t => {
  const f = fixture(t, { autoStart:false, transform:s=>s.replace(' || session.armed || !context.isInitialized',' || !context.isInitialized') }); f.change('scheduled'); f.click('Start suggestions'); f.act(()=>t.mock.timers.tick(299)); f.click('Start suggestions'); f.act(()=>t.mock.timers.tick(1)); assert.throws(()=>assert.equal(f.requests.length,1,'Start must preserve original deadline'),/Start must preserve/); f.debounce(); await f.resolve(0,[]);
});
test('fault control detects selection scheduling an unwanted follow-up request', async t => {
  const f=fixture(t,{transform:s=>s.replace('setDraft({ raw: suggestion, armed: session.armed });','setDraft({ raw: suggestion, armed: session.armed }); schedule(session.version);')}); f.change('typed');f.debounce();await f.resolve(0,['selected']);f.click('selected');f.debounce();assert.throws(()=>assert.equal(f.requests.length,1,'Selection sends no follow-up'),/Selection sends no follow-up/);await f.resolve(1,[]);
});

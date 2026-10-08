const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const harness = require('./fixtures/formatted-text-harness.cjs');
const entry = path.resolve(__dirname, '../packages/spfx-react-toolkit/src/controls/relative-time/relative-time.tsx');
const instant = Date.parse('2026-01-31T12:00:00Z');
let React, ReactDOM, act, RelativeTime, dom;
test.before(async () => {
  assert.ok(fs.existsSync(entry), 'RelativeTime implementation missing');
  dom = harness.createSxDom();
  for (const name of ['Node', 'Element', 'HTMLElement', 'NodeFilter', 'MutationObserver', 'Event', 'CustomEvent', 'FocusEvent']) global[name] = window[name];
  global.getComputedStyle = window.getComputedStyle.bind(window);
  Object.defineProperty(document, 'hidden', { configurable: true, value: false });
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  ({ RelativeTime } = await harness.loadModule(entry));
});
test.after(async () => { await harness.cleanup(); dom?.window.close(); });
function mount(props, targetDocument = document, withoutVisibilityDocument = false) {
  const { FluentProvider } = require('@fluentui/react-provider');
  const { webLightTheme } = require('@fluentui/react-theme');
  const host = targetDocument.createElement('div'); targetDocument.body.appendChild(host);
  function render(next = props, nextDocument = targetDocument) {
    props = next; targetDocument = nextDocument;
    let child = React.createElement(RelativeTime, { 'data-time': '', ...props });
    if (withoutVisibilityDocument) child = React.createElement(require('@fluentui/react-shared-contexts').Provider_unstable, { value: { dir: 'ltr', targetDocument: undefined } }, child);
    act(() => { ReactDOM.render(React.createElement(FluentProvider, { theme: webLightTheme, targetDocument }, child), host); });
  }
  try { render(); } catch (error) { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); throw error; }
  return { host, render, get time() { return host.querySelector('time'); }, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
const fixed = { dateTime: instant, referenceTime: instant + 300000, locale: 'en', showTooltip: false };
test('semantic root preserves instant, native ref/attributes and external descriptions', () => {
  const ref = React.createRef(); let focuses = 0;
  const f = mount({ ...fixed, ref, id: 'created', className: 'consumer', lang: 'en-GB', 'aria-describedby': 'external', onFocus: () => focuses++ });
  try {
    assert.equal(f.time.textContent, '5 minutes ago'); assert.equal(ref.current, f.time);
    assert.equal(f.time.dateTime, '2026-01-31T12:00:00.000Z'); assert.equal(f.time.id, 'created');
    assert.equal(f.time.lang, 'en-GB'); assert.equal(f.time.className, 'consumer');
    assert.equal(f.time.hasAttribute('style'), false); assert.equal(f.time.hasAttribute('tabindex'), false);
    assert.equal(f.time.hasAttribute('aria-live'), false);
    const ids = f.time.getAttribute('aria-describedby').split(' '); assert.equal(ids[0], 'external');
    assert.ok(document.getElementById(ids[1]).hidden); assert.match(document.getElementById(ids[1]).textContent, /2026/);
    act(() => { f.time.dispatchEvent(new window.FocusEvent('focusin', { bubbles: true })); }); assert.equal(focuses, 1);
  } finally { f.close(); }
});
test('Intl styles/locales and time-zone presentation are independent of relative calculation', () => {
  for (const locale of ['it-IT', 'ar', ['zz-ZZ', 'fr-FR']]) for (const format of ['long', 'short', 'narrow']) {
    const f = mount({ ...fixed, locale, format, timeZone: 'UTC' });
    try {
      assert.equal(f.time.textContent, new Intl.RelativeTimeFormat(locale, { style: format, numeric: 'always' }).format(-5, 'minute'));
      const id = f.time.getAttribute('aria-describedby'); const full = document.getElementById(id).textContent;
      f.render({ ...fixed, locale, format, timeZone: 'Europe/Rome' });
      assert.equal(f.time.textContent, new Intl.RelativeTimeFormat(locale, { style: format, numeric: 'always' }).format(-5, 'minute'));
      assert.notEqual(document.getElementById(id).textContent, full);
    } finally { f.close(); }
  }
});
test('invalid instants render fallback before Intl configuration and no authored content can replace time', () => {
  const f = mount({ dateTime: undefined, locale: 'bad_locale', timeZone: 'Mars/OlympusMons', fallback: 'Unknown' });
  try {
    assert.equal(Boolean(f.time), false); assert.equal(f.host.textContent, 'Unknown');
    f.render({ ...fixed, referenceTime: 'invalid', fallback: 'Invalid reference' }); assert.equal(Boolean(f.time), false);
    assert.equal(f.host.textContent, 'Invalid reference');
    f.render({ ...fixed, children: 'Injected', dangerouslySetInnerHTML: { __html: '<b>Injected</b>' }, style: { color: 'red' } });
    assert.equal(f.time.textContent, '5 minutes ago'); assert.equal(f.time.querySelectorAll('b').length, 0);
    assert.equal(f.time.style.color, 'red'); assert.equal(f.time.dateTime, '2026-01-31T12:00:00.000Z');
  } finally { f.close(); }
});
test('valid instants keep Intl errors and tooltip keyboard defaults', () => {
  for (const option of [{ locale: 'bad_locale' }, { timeZone: 'Mars/OlympusMons' }]) {
    const errors = []; const original = console.error; console.error = (...args) => errors.push(args);
    try { assert.throws(() => mount({ ...fixed, ...option }), RangeError); }
    finally { console.error = original; }
  }
  const f = mount({ ...fixed, showTooltip: true });
  try {
    assert.equal(f.time.tabIndex, 0); const id = f.time.getAttribute('aria-describedby');
    assert.ok(document.getElementById(id));
    f.render({ ...fixed, showTooltip: true, tabIndex: -1 }); assert.equal(f.time.tabIndex, -1);
  } finally { f.close(); }
});
test('live clock follows minute boundaries and cancels work on invalidation/unmount', t => {
  t.mock.timers.enable({ apis: ['Date', 'setTimeout'], now: instant + 59000 });
  const f = mount({ dateTime: instant, locale: 'en', showTooltip: false });
  try {
    assert.equal(f.time.textContent, 'now'); act(() => { t.mock.timers.tick(1000); });
    assert.equal(f.time.textContent, '1 minute ago'); act(() => { t.mock.timers.tick(60000); });
    assert.equal(f.time.textContent, '2 minutes ago');
    f.render({ dateTime: undefined, fallback: 'Unknown', showTooltip: false });
    const spy = t.mock.method(Date, 'now'); act(() => { t.mock.timers.tick(3600000); }); assert.equal(spy.mock.callCount(), 0);
  } finally { f.close(); }
});
test('fixed/disabled/live references resample only on the declared clock triggers', t => {
  t.mock.timers.enable({ apis: ['Date', 'setTimeout'], now: instant + 59000 });
  const f = mount({ dateTime: instant, autoUpdate: false, locale: 'en', showTooltip: false });
  try {
    act(() => { t.mock.timers.tick(61000); }); assert.equal(f.time.textContent, 'now');
    f.render({ dateTime: new Date(instant), autoUpdate: false, locale: 'it', format: 'short', showTooltip: false });
    assert.equal(f.time.textContent, new Intl.RelativeTimeFormat('it', { numeric: 'auto', style: 'short' }).format(0, 'second'));
    f.render({ dateTime: instant, locale: 'en', showTooltip: false }); assert.equal(f.time.textContent, '2 minutes ago');
    f.render(fixed); act(() => { t.mock.timers.tick(600000); }); assert.equal(f.time.textContent, '5 minutes ago');
    f.render({ dateTime: instant + 60000, autoUpdate: false, locale: 'en', showTooltip: false });
    assert.equal(f.time.textContent, '11 minutes ago');
  } finally { f.close(); }
});
test('hidden host document pauses timers and resume/document swap refresh correctly', t => {
  t.mock.timers.enable({ apis: ['Date', 'setTimeout'], now: instant + 59000 });
  const other = document.implementation.createHTMLDocument('Other');
  Object.defineProperty(other, 'hidden', { configurable: true, value: true, writable: true });
  const f = mount({ dateTime: instant, locale: 'en', showTooltip: false }, other);
  try {
    act(() => { t.mock.timers.tick(61000); }); assert.equal(f.time.textContent, 'now');
    other.hidden = false; act(() => { other.dispatchEvent(new window.Event('visibilitychange')); });
    assert.equal(f.time.textContent, '2 minutes ago');
    other.hidden = true; act(() => { other.dispatchEvent(new window.Event('visibilitychange')); });
    act(() => { t.mock.timers.tick(60000); }); assert.equal(f.time.textContent, '2 minutes ago');
    f.render({ dateTime: instant, locale: 'en', showTooltip: false }, document); assert.equal(f.time.textContent, '3 minutes ago');
    act(() => { t.mock.timers.tick(60000); }); assert.equal(f.time.textContent, '4 minutes ago');
  } finally { f.close(); }
});

test('unrepresentable next boundaries use the hourly cap and valid unmount cancels the timer', t => {
  t.mock.timers.enable({ apis: ['Date', 'setTimeout'], now: 8640000000000000 });
  const timeout = t.mock.method(global, 'setTimeout');
  const f = mount({ dateTime: -8640000000000000, locale: 'en', timeZone: 'UTC', showTooltip: false });
  assert.ok(timeout.mock.calls.some(call => call.arguments[1] === 3600000));
  f.close();
  const clock = t.mock.method(Date, 'now');
  act(() => { t.mock.timers.tick(3600000); });
  assert.equal(clock.mock.callCount(), 0);
});
test('mounted control without a visibility document still refreshes even if the global document is hidden', t => {
  t.mock.timers.enable({ apis: ['Date', 'setTimeout'], now: instant + 59000 });
  Object.defineProperty(document, 'hidden', { configurable: true, value: true });
  const listener = t.mock.method(document, 'addEventListener');
  const f = mount({ dateTime: instant, locale: 'en', showTooltip: false }, document, true);
  try {
    assert.equal(listener.mock.calls.filter(call => call.arguments[0] === 'visibilitychange').length, 0);
    act(() => { t.mock.timers.tick(1000); });
    assert.equal(f.time.textContent, '1 minute ago');
  } finally { f.close(); Object.defineProperty(document, 'hidden', { configurable: true, value: false }); }
});

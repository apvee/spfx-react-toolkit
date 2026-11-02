process.env.NODE_ENV = 'test';
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const { Simulate } = require('react-dom/test-utils');
const ReactDOM = require('react-dom');
const titles = {
  skeleton: 'Explore mounted profile loading and native focus',
  time: 'Explore relative clocks and fallback slots',
  text: 'Explore semantic text and inline Markdown',
  edit: 'Edit a controlled block with policy and dictionaries',
};
const globals = new Map();
for (const name of ['Node', 'NodeFilter', 'Element', 'HTMLElement', 'Event', 'CustomEvent', 'MutationObserver']) {
  globals.set(name, Object.getOwnPropertyDescriptor(global, name));
  Object.defineProperty(global, name, { configurable: true, writable: true, value: window[name] });
}
// JSDOM has no text/layout geometry; native geometry/inert remains browser evidence.
const fixtureDocument = document;
const range = fixtureDocument.createRange;
document.createRange = () => ({ selectNodeContents() {}, getClientRects() { return []; } });
Object.defineProperty(document, 'hidden', { configurable: true, value: false });
test.after(() => { fixtureDocument.createRange = range; for (const [name, descriptor] of globals) {
  if (descriptor) Object.defineProperty(global, name, descriptor); else delete global[name];
} });
const { FluentProvider, webLightTheme, webDarkTheme } = require('@fluentui/react-components');
const { RendererProvider } = require('@griffel/react');
const { createDOMRenderer } = require('@griffel/core');
function fixture(t, kind, standalone = false) {
  const actual = scenario(titles[kind]), toolkit = {};
  const h = createHarness({ '@apvee/spfx-react-toolkit': toolkit });
  t.after(h.close);
  Object.assign(toolkit, h.load('helpers/styles/index.ts'), h.load('hooks/useSx.ts'), h.load('hooks/useStableCallback.ts'));
  const files = { skeleton: 'auto-skeleton/auto-skeleton', time: 'relative-time/relative-time', text: 'formatted-text/formatted-text' };
  Object.assign(toolkit, h.load('controls/' + files[kind] + '.tsx'));
  const { Demo } = h.load(actual.file), renderer = createDOMRenderer(null);
  const { DemoCard } = h.load(path.resolve(actual.file, '../../../shared/DemoCard.tsx'));
  let revision = 7, dir = 'ltr', theme = webLightTheme;
  const render = regions => h.React.createElement(DemoCard, { title: titles[kind], iconName: 'Info', ...regions });
  function Root() { return h.React.createElement(RendererProvider, { renderer }, h.React.createElement(FluentProvider, { dir, theme }, h.React.createElement(Demo, { resetRevision: revision, render: standalone ? undefined : render }))); }
  const view = h.mountComponent(Root), prefix = { skeleton: 'auto-skeleton', time: 'relative-time', text: 'formatted-text' }[kind];
  const query = marker => view.container.querySelector(`[data-${prefix}-demo="${marker}"]`);
  const find = marker => { const node = query(marker); assert.ok(node, marker); return node; };
  const click = marker => h.act(() => Simulate.click(find(marker)));
  const change = (marker, value) => h.act(() => { const node = find(marker); node.value = value; Simulate.change(node, { target: { value } }); });
  const check = (marker, checked) => h.act(() => { const node = find(marker); node.checked = checked; Simulate.change(node); });
  const update = () => h.act(() => ReactDOM.render(h.React.createElement(Root), view.container));
  return { h, view, find, query, click, change, check, renderer, reset() { revision++; update(); }, host(direction, dark = false) { dir = direction; theme = dark ? webDarkTheme : webLightTheme; update(); } };
}
for (const [kind, owner] of Object.entries({ skeleton: 'AutoSkeleton', time: 'RelativeTime', text: 'FormattedText', edit: 'EditableText' })) {
  test(`${owner} exposes its real source-owned scenario`, () => {
    const { entry, source } = scenario(titles[kind]);
    assert.equal(entry.featureTitle, owner); assert.ok(entry.ownerKey.endsWith('#' + owner));
    assert.deepEqual(entry.metadata, { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: entry.metadata.observe });
    assert.ok(entry.metadata.observe.length > 0);
    assert.doesNotMatch(source, /components\/|fixtures\/|style=|\.style\./);
  });
}
test('profile source retains actual note and DOM through loading policies, host updates and Reset', t => {
  const f = fixture(t, 'skeleton'), note = f.find('focus-target'), root = f.find('preview');
  f.change('focus-target', 'Retained note'); f.check('loading', true);
  for (const animation of ['none', 'pulse', 'wave']) f.change('animation', animation);
  for (const size of ['16', '32', '8']) f.change('minSize', size);
  f.check('resize', true); f.check('ignoreSelectors', false); f.host('rtl', true);
  assert.equal(f.find('preview'), root); assert.equal(f.find('focus-target'), note); assert.equal(note.value, 'Retained note');
  assert.equal(root.getAttribute('aria-busy'), 'true'); f.click('inspect-ref');
  assert.match(f.find('observations').textContent, /Content mounts: 1/);
  f.reset(); assert.equal(f.find('preview'), root); assert.equal(f.find('focus-target'), note);
  assert.equal(note.value, 'Documentation update'); assert.equal(root.getAttribute('aria-busy'), 'false');
  assert.equal(f.find('animation').value, 'wave'); assert.equal(f.find('minSize').value, '8');
  f.check('empty', true); assert.equal(f.query('focus-target'), null);
  f.check('empty', false); assert.notEqual(f.find('focus-target'), note);
  assert.match(f.find('observations').textContent, /Content mounts: 2/);
});
test('relative source keeps fixed/frozen/live clock authority and cleans up timers on Reset/departure', t => {
  const instant = Date.parse('2026-01-31T12:00:00Z');
  t.mock.timers.enable({ apis: ['Date', 'setTimeout'], now: instant + 59000 });
  const f = fixture(t, 'time'), root = f.find('preview');
  assert.equal(root.tagName, 'TIME'); assert.equal(root.textContent, '1 month ago');
  f.change('mode', 'frozen'); assert.equal(root.textContent, 'now');
  f.h.act(() => t.mock.timers.tick(61000)); f.change('locale', 'it-IT'); f.change('format', 'short');
  assert.equal(root.textContent, new Intl.RelativeTimeFormat('it-IT', { numeric: 'auto', style: 'short' }).format(0, 'second'));
  f.change('locale', 'en-US'); f.change('format', 'long'); f.change('mode', 'live');
  assert.equal(root.textContent, '2 minutes ago'); f.h.act(() => t.mock.timers.tick(60000)); assert.equal(root.textContent, '3 minutes ago');
  f.reset(); assert.equal(f.find('preview'), root); assert.equal(root.textContent, '1 month ago');
  f.h.act(() => t.mock.timers.tick(60000)); assert.equal(root.textContent, '1 month ago');
  f.change('mode', 'live'); f.view.unmount(); const now = t.mock.method(Date, 'now');
  f.h.act(() => t.mock.timers.tick(3600000)); assert.equal(now.mock.callCount(), 0);
});
test('relative source transfers invalid and callback fallback refs/events and returns to actual TIME', t => {
  const f = fixture(t, 'time', true); f.change('scenario', 'invalid'); assert.equal(f.query('preview'), null);
  assert.match(f.view.container.textContent, /Date unavailable/); f.change('scenario', 'slot');
  const fallback = f.find('fallback-slot'); assert.equal(fallback.tagName, 'SPAN'); f.click('fallback-slot'); f.click('inspect-slot');
  assert.match(f.view.container.textContent, /Fallback element: span. Fallback clicks: 1/);
  f.change('fallback', 'Updated fallback'); assert.equal(f.find('fallback-slot'), fallback); assert.equal(fallback.textContent, 'Updated fallback');
  f.change('scenario', 'event'); f.click('preview'); f.click('inspect-ref');
  assert.match(f.find('observations').textContent, /Ref element: time. Preview clicks: 1/);
});
test('formatted source retains integer width drafts, semantic independence and one real JSX action', t => {
  const f = fixture(t, 'text'), width = f.find('width');
  f.change('variant', 'title2'); assert.equal(f.find('preview').tagName, 'SPAN');
  f.change('width', '640'); const rules = Object.keys(f.renderer.insertionCache).join(' '); assert.match(rules, /640px/);
  for (const draft of ['', '6', '64', '900', '160.5']) { f.change('width', draft); assert.equal(width.value, draft); assert.match(f.view.container.textContent, /640px/); }
  f.h.act(() => Simulate.blur(width)); assert.equal(width.value, '640');
  f.change('as', 'h4'); f.click('inspect-ref'); assert.match(f.find('observations').textContent, /Ref element: h4/);
  f.change('text', 'Read **complete** ~~changes~~ and `code` [docs](#docs).');
  assert.equal(f.find('preview').querySelector('strong').textContent, 'complete'); assert.ok(f.find('preview').querySelector('del')); assert.ok(f.find('preview').querySelector('code'));
  f.change('content', 'jsx'); assert.equal(f.find('mixed-jsx').tagName, 'BUTTON'); f.click('mixed-jsx');
  assert.match(f.find('observations').textContent, /Preview clicks: 1/);
  f.check('markdown', false); assert.equal(f.find('preview').querySelector('strong'), null); assert.ok(f.query('mixed-jsx'));
  f.change('content', 'text'); f.check('markdown', true); assert.equal(f.find('preview').querySelector('strong').textContent, 'complete');
  f.change('maxLines', '2'); f.check('truncate', true); f.check('wrap', false);
  const root = f.find('preview'); f.host('rtl', true); assert.equal(f.find('preview'), root);
  f.reset(); assert.equal(f.find('width'), width); assert.equal(width.value, '320'); assert.equal(f.find('preview').tagName, 'SPAN');
  assert.equal(f.find('variant').value, 'body1'); assert.equal(f.find('maxLines').value, '0'); assert.equal(f.find('markdown').checked, true);
});

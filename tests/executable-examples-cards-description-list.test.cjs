const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const ReactDOM = require('react-dom');
const { FluentProvider, webLightTheme, webDarkTheme } = require('@fluentui/react-components');
const { RendererProvider } = require('@griffel/react');
const { createDOMRenderer } = require('@griffel/core');
const previousGlobals = new Map();
for (const name of ['Node', 'NodeFilter', 'Element', 'HTMLElement', 'MutationObserver']) {
  previousGlobals.set(name, Object.getOwnPropertyDescriptor(global, name));
  Object.defineProperty(global, name, { configurable: true, writable: true, value: window[name] });
}
test.after(() => { for (const [name, descriptor] of previousGlobals) { if (descriptor) Object.defineProperty(global, name, descriptor); else delete global[name]; } });
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples');
const titles = {
  content: 'Configure content slots and native card interaction',
  preview: 'Inspect preview media and image recovery',
  retention: 'Retain caller data when switching card types',
  details: 'Inspect responsive metadata with data and compound APIs',
};
function fixture(t, kind, standalone = false) {
  const toolkit = {};
  const h = createHarness({ '@apvee/spfx-react-toolkit': toolkit });
  t.after(h.close);
  Object.assign(toolkit, h.load('helpers/styles/index.ts'), h.load('hooks/useSx.ts'));
  for (const file of ['content-card/content-card', 'preview-card/preview-card', 'description-list/description-list', 'auto-skeleton/auto-skeleton', 'formatted-text/formatted-text']) Object.assign(toolkit, h.load(`controls/${file}.tsx`));
  const catalogue = h.load(path.join(base, 'catalogue.ts'));
  const entry = catalogue.executableExamples.find(item => item.title === titles[kind]);
  assert.ok(entry, `Missing actual source-owned scenario: ${titles[kind]}`);
  const owner = { content: 'ContentCard', preview: 'PreviewCard', retention: 'PreviewCard', details: 'DescriptionList' }[kind];
  assert.equal(entry.featureTitle, owner); assert.ok(entry.ownerKey.endsWith('#' + owner));
  assert.ok(entry.publicNames.includes(owner)); assert.equal(entry.metadata.host, 'none');
  assert.equal(entry.metadata.kind, 'snapshot'); assert.equal(entry.metadata.contentPlacement, 'subject');
  const { Demo } = h.load(path.join(base, entry.id + '.tsx'));
  const renderer = createDOMRenderer(null);
  // JSDOM 15 lacks WebCrypto required by actual Fluent/Tabster focus handling.
  Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  const oldRange = document.createRange;
  document.createRange = () => ({ selectNodeContents() {}, getClientRects() { return []; } });
  t.after(() => { document.createRange = oldRange; });
  let revision = 7, dir = 'ltr', theme = webLightTheme;
  const { DemoCard } = h.load(path.resolve(base, '../../shared/DemoCard.tsx'));
  const render = regions => h.React.createElement(DemoCard, { title: titles[kind], iconName: 'Info', ...regions });
  function Root() { return h.React.createElement(RendererProvider, { renderer }, h.React.createElement(FluentProvider, { theme, dir }, h.React.createElement(Demo, { resetRevision: revision, render: standalone ? undefined : render }))); }
  const view = h.mountComponent(Root);
  if (!standalone) {
    assert.equal(view.container.querySelectorAll('[data-demo-configuration]').length, 1);
    assert.equal(view.container.querySelectorAll('[data-demo-subject]').length, 1);
  }
  if (kind !== 'details') {
    assert.equal(view.container.querySelectorAll('[data-cards-demo=preview-card],[data-cards-demo=content-card]').length, 1);
    assert.equal(view.container.querySelector('[data-cards-demo=native-card]'), null);
    if (kind !== 'content') assert.equal(view.container.querySelector('[data-cards-demo=preview-card]').getAttribute('aria-checked'), null);
  }
  const update = () => h.act(() => { ReactDOM.render(h.React.createElement(Root), view.container); });
  const find = marker => {
    const node = view.container.querySelector(`[data-${kind === 'details' ? 'description-list' : 'cards'}-demo="${marker}"]`);
    assert.ok(node, marker); return node;
  };
  const click = marker => h.act(() => Simulate.click(find(marker)));
  const select = (marker, value) => h.act(() => Simulate.change(find(marker), { target: { value } }));
  const check = (marker, checked) => h.act(() => { const target = find(marker); const input = target.tagName === 'INPUT' ? target : target.querySelector('input'); assert.ok(input, marker + ' native checkbox'); input.checked = checked; Simulate.change(input); });
  return { h, view, find, click, select, check, entry, reset() { revision++; update(); }, host(direction, dark = false) { dir = direction; theme = dark ? webDarkTheme : webLightTheme; update(); } };
}
// These tests catch omitted handlers/state ownership, media branch precedence and
// accidental remount/reset keys in the actual emitted source examples.
test('ContentCard source case owns real interaction, formatting and reset without replacing the subject', t => {
  const f = fixture(t, 'content');
  const root = f.find('content-card'), note = f.find('retained-input');
  assert.equal(root.querySelector('h2').textContent, 'Quarterly report');
  assert.equal(f.find('description-input').value, 'Team results, goals and next steps.');
  assert.equal(f.view.container.querySelectorAll('[data-demo-configuration]').length, 1);
  assert.ok(root.querySelector('.fui-CardHeader__image svg path[d]'));
  f.select('title-input', '**Literal title**'); f.select('description-input', '**Literal description**');
  assert.equal(root.querySelector('strong'), null);
  f.select('retained-input', 'Caller edit'); f.click('header-action'); f.click('footer-action');
  assert.match(f.find('counters').textContent, /Header:1 Footer:1/);
  assert.match(f.find('counters').textContent, /Root bubbled:2/);
  f.check('selectable', true); f.h.act(() => Simulate.click(root));
  assert.match(f.find('counters').textContent, /Selection requests:3 Selected:true/);
  assert.ok(root.querySelector('input[type=checkbox]'));
  f.check('selectable', false); f.h.act(() => Simulate.click(root));
  assert.match(f.find('counters').textContent, /Selection requests:4 Selected:true/);
  f.select('format', 'formatted'); assert.equal(root.querySelector('strong').textContent, 'Literal description');
  f.select('format', 'callback'); f.click('inspect-refs');
  assert.match(f.find('observations').textContent, /root:DIV title:SPAN button:BUTTON/);
  assert.match(f.find('observations').textContent, /description callbacks:[1-9]/);
  f.select('format', 'absent'); assert.equal(root.querySelector('.fui-CardHeader__description'), null);
  f.select('width', 'compact'); f.check('loading', true); assert.equal(f.find('skeleton').getAttribute('aria-busy'), 'true');
  f.check('loading', false); f.host('rtl', true); assert.equal(f.find('retained-input'), note); assert.equal(note.value, 'Caller edit');
  f.check('direction', true); f.host('ltr'); assert.equal(root.closest('[dir]').getAttribute('dir'), 'rtl');
  f.check('direction', false); assert.equal(root.closest('[dir]').getAttribute('dir'), 'ltr');
  f.check('disabled', true); assert.equal(f.find('footer-action').disabled, false);
  f.reset(); assert.equal(f.find('content-card'), root); assert.equal(f.find('retained-input'), note); assert.equal(note.value, 'To share at the team meeting.');
  assert.match(f.find('counters').textContent, /Header:0 Footer:0 Root bubbled:0 Selection requests:0 Selected:false/);
});
test('PreviewCard source case preserves media branches, same-src failure and caller nodes through presentation/reset', t => {
  const f = fixture(t, 'preview'); const root = f.find('preview-card'), note = f.find('retained-input');
  assert.equal(f.find('orientation').value, 'horizontal');
  assert.equal(f.find('width').value, 'standard');
  assert.equal(f.find('title-input').value, 'Quarterly report');
  assert.equal(f.find('description-input').value, 'Team results, goals and next steps.');
  assert.ok(f.find('logo-slot').querySelector('svg path[d]'));
  assert.match(root.querySelector('[data-card-preview] > img').src, /^data:image\/svg\+xml/);
  f.select('retained-input', 'Kept preview note');
  f.click('header-action'); f.click('footer-action'); f.click('inspect-refs');
  assert.match(f.find('counters').textContent, /Header:1 Footer:1/);
  assert.match(f.find('observations').textContent, /root:DIV title:SPAN preview:absent logo:DIV button:BUTTON/);
  for (const mode of ['custom', 'empty-object', 'empty-callback']) {
    f.select('preview-mode', mode); assert.ok(root.querySelector('[data-card-preview]'));
    assert.equal(root.querySelectorAll('img[alt="Illustrated quarterly report cover"]').length, 0);
  }
  for (const mode of ['false', 'empty-fragment']) { f.select('preview-mode', mode); assert.equal(root.querySelector('[data-card-preview]'), null); }
  f.select('preview-mode', 'image'); f.click('fail-image');
  f.h.act(() => Simulate.error(root.querySelector('[data-card-preview] > img')));
  let fallback = root.querySelector('[data-card-image-fallback]');
  assert.equal(fallback.getAttribute('aria-label'), 'Illustrated quarterly report cover');
  assert.equal(fallback.textContent, 'Cover unavailable');
  f.check('custom-fallback', true); assert.equal(root.querySelector('[data-card-image-fallback]'), fallback);
  assert.equal(fallback.textContent, 'Custom unavailable cover');
  f.check('decorative', true); assert.equal(fallback.getAttribute('aria-hidden'), 'true');
  f.check('logo-callback', true); f.click('inspect-refs'); assert.match(f.find('observations').textContent, /fallback callbacks:[1-9].*logo callbacks:[1-9]/);
  f.click('new-src'); assert.equal(root.querySelector('[data-card-image-fallback]'), null); assert.ok(root.querySelector('[data-card-preview] > img'));
  f.select('orientation', 'vertical'); f.select('width', 'compact'); f.check('loading', true); f.check('loading', false); f.host('rtl', true);
  assert.equal(f.find('retained-input'), note); assert.equal(note.value, 'Kept preview note');
  f.reset(); assert.equal(f.find('preview-card'), root); assert.equal(f.find('retained-input'), note); assert.equal(f.find('orientation').value, 'horizontal');
  assert.equal(note.value, 'To share at the team meeting.'); assert.equal(root.querySelector('img').alt, 'Illustrated quarterly report cover');
});
test('card switch source keeps caller title description and note in both directions and resets the same Demo', t => {
  const f = fixture(t, 'retention', true); const config = f.find('kind');
  f.select('title-input', 'Shared title'); f.select('description-input', 'Shared description'); f.select('retained-input', 'Shared note');
  let expectedTitle = 'Shared title', expectedDescription = 'Shared description', expectedNote = 'Shared note';
  for (const kind of ['content', 'preview', 'content']) {
    f.select('kind', kind); const root = f.find(kind + '-card');
    assert.equal(f.view.container.querySelectorAll('[data-cards-demo=preview-card],[data-cards-demo=content-card]').length, 1);
    assert.equal(root.getAttribute('aria-checked'), null, 'both replacement card types remain nonselectable');
    assert.equal(root.querySelector('h2').textContent, expectedTitle); assert.ok(root.textContent.includes(expectedDescription));
    assert.equal(f.find('retained-input').value, expectedNote);
    expectedTitle = 'Latest ' + kind + ' title'; expectedDescription = 'Latest ' + kind + ' description'; expectedNote = 'Latest ' + kind + ' note';
    f.select('title-input', expectedTitle); f.select('description-input', expectedDescription); f.select('retained-input', expectedNote);
  }
  f.reset(); assert.equal(f.find('kind'), config); assert.equal(config.value, 'preview'); assert.equal(f.find('title-input').value, 'Quarterly report');
  assert.equal(f.find('description-input').value, 'Team results, goals and next steps.'); assert.equal(f.find('retained-input').value, 'To share at the team meeting.');
});
test('DescriptionList source retains keyed child state across each API layout width order hidden and reset', t => {
  const f = fixture(t, 'details');
  for (const api of ['data', 'compound']) {
    f.reset(); f.select('composition', api); const list = f.find('list');
    assert.equal(list.querySelectorAll('dt').length, 5); assert.ok(list.querySelector('a')); assert.ok(list.querySelector('.fui-Badge'));
    const counter = f.find('counter'); f.click('counter');
    for (const pixels of ['320', '479', '480', '639', '640', '900']) { f.select('width', pixels); assert.equal(f.find('counter'), counter); assert.equal(counter.textContent, 'Revisions: 1'); }
    for (const layout of ['auto', 'horizontal', 'vertical']) { f.select('layout', layout); assert.equal(f.find('list').getAttribute('data-description-list-layout'), layout); }
    f.check('reverse', true); assert.equal(list.querySelector('dt').textContent, 'Revisions'); assert.equal(f.find('counter'), counter);
    f.check('hidden', true); assert.equal(list.hidden, true); assert.equal(f.find('counter'), counter); f.check('hidden', false);
    if (api === 'compound') {
      f.check('item-hidden', true); assert.equal(list.firstElementChild.hidden, true); f.click('inspect-ref');
      assert.match(f.find('ref').textContent, /root=DL; pairs=5; root hidden=false; item=DIV; item hidden=true/);
      f.check('item-hidden', false); f.h.act(() => Simulate.click(list.firstElementChild));
      assert.match(f.find('clicks').textContent, /Bubbled list clicks: 2 · Item clicks: 2/);
    }
    f.check('reverse', false);
  }
  f.check('empty', true); assert.equal(f.find('list').querySelectorAll('dt').length, 0);
  f.select('composition', 'data'); assert.equal(f.view.container.querySelector('dl'), null); f.check('empty', false);
  f.host('rtl', true); const list = f.find('list'); assert.equal(list.closest('[dir]').getAttribute('dir'), 'rtl');
  f.check('rtl', true); f.host('ltr'); assert.equal(list.closest('[dir]').getAttribute('dir'), 'rtl');
  f.check('rtl', false); assert.equal(list.closest('[dir]').getAttribute('dir'), 'ltr');
  const counter = f.find('counter'); f.click('counter'); f.reset(); assert.equal(f.find('counter'), counter); assert.equal(counter.textContent, 'Revisions: 0'); assert.equal(f.find('list'), list);
  f.click('inspect-ref'); assert.match(f.find('ref').textContent, /root=DL; pairs=5/);
});

for (const kind of ['content', 'preview', 'retention', 'details']) {
  test(`${kind} standalone source scenario resets independently and unmounts its native subject`, t => {
    const first = fixture(t, kind, true), second = fixture(t, kind, true);
    const marker = kind === 'details' ? 'counter' : 'retained-input';
    if (kind === 'details') first.click('counter'); else first.select(marker, 'First instance edit');
    assert.equal(second.find(marker)[kind === 'details' ? 'textContent' : 'value'], kind === 'details' ? 'Revisions: 0' : 'To share at the team meeting.');
    const retained = first.find(marker); first.reset(); assert.equal(first.find(marker), retained);
    assert.equal(first.find(marker)[kind === 'details' ? 'textContent' : 'value'], kind === 'details' ? 'Revisions: 0' : 'To share at the team meeting.');
    first.view.unmount(); assert.equal(retained.isConnected, false);
  });
}
test('DescriptionList portable anchors belong to each mounted Demo and survive Reset', t => {
  const first = fixture(t, 'details');
  const second = fixture(t, 'details');
  const link = f => f.view.container.querySelector('a[href^="#"]');
  const a = link(first), b = link(second);
  assert.ok(a && b);
  const firstId = a.getAttribute('href').slice(1), secondId = b.getAttribute('href').slice(1);
  assert.notEqual(firstId, secondId, 'simultaneous instances own distinct anchors');
  const firstTarget = first.view.container.querySelector(`[id="${firstId}"]`);
  const secondTarget = second.view.container.querySelector(`[id="${secondId}"]`);
  assert.ok(firstTarget && secondTarget, 'no fixture wrapper supplies the target');
  first.reset();
  assert.equal(link(first).getAttribute('href'), '#' + firstId);
  assert.equal(first.view.container.querySelector(`[id="${firstId}"]`), firstTarget);
  assert.equal(link(second), b);
  assert.equal(second.view.container.querySelector(`[id="${secondId}"]`), secondTarget);
});

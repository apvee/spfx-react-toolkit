process.env.NODE_ENV = 'test';
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const { Simulate } = require('react-dom/test-utils');
const ReactDOM = require('react-dom');
const titles = {
  stage: 'Edit a report with guarded Stage dismissal',
  settings: 'Preserve notification settings across group presentations',
  accordion: 'Control disclosure selection and retained content',
  state: 'Configure state artwork semantics and actions',
};
const previousGlobals = new Map();
for (const name of ['Node', 'NodeFilter', 'Element', 'HTMLElement', 'Event', 'CustomEvent', 'MutationObserver']) {
  previousGlobals.set(name, Object.getOwnPropertyDescriptor(global, name));
  Object.defineProperty(global, name, { configurable: true, writable: true, value: window[name] });
}
Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
window.matchMedia ??= () => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} });
test.after(() => { for (const [name, descriptor] of previousGlobals) { if (descriptor) Object.defineProperty(global, name, descriptor); else delete global[name]; } });
const { FluentProvider, webLightTheme, webDarkTheme } = require('@fluentui/react-components');
const { RendererProvider } = require('@griffel/react');
const { createDOMRenderer } = require('@griffel/core');
function fixture(t, kind, standalone = false) {
  const actual = scenario(titles[kind]);
  const toolkit = {};
  const h = createHarness({ '@apvee/spfx-react-toolkit': toolkit });
  t.after(h.close);
  Object.assign(toolkit, h.load('helpers/styles/index.ts'), h.load('hooks/useSx.ts'));
  for (const file of ['stage/stage', 'form/create-form', 'settings-group/settings-group', 'accordion-group/accordion-group', 'state-message/state-message']) Object.assign(toolkit, h.load(`controls/${file}.tsx`.replace('create-form.tsx', 'create-form.ts')));
  const { Demo } = h.load(actual.file);
  const renderer = createDOMRenderer(null);
  const { DemoCard } = h.load(path.resolve(actual.file, '../../../shared/DemoCard.tsx'));
  let revision = 7, dir = 'ltr', theme = webLightTheme;
  const render = regions => h.React.createElement(DemoCard, { title: titles[kind], iconName: 'Info', contentPlacement: actual.entry.metadata.contentPlacement, ...regions });
  function Root() { return h.React.createElement(RendererProvider, { renderer }, h.React.createElement(FluentProvider, { dir, theme }, h.React.createElement(Demo, { resetRevision: revision, render: standalone ? undefined : render }))); }
  const view = h.mountComponent(Root);
  const prefix = { stage: 'stage', settings: 'groups', accordion: 'accordion', state: 'state-message' }[kind];
  const query = marker => view.container.querySelector(`[data-${prefix}-demo="${marker}"]`) ?? document.querySelector(`[data-${prefix}-demo="${marker}"]`);
  const find = marker => { const node = query(marker); assert.ok(node, marker); return node; };
  const act = async run => { await h.act(async () => { await run(); }); };
  const click = async marker => act(() => Simulate.click(find(marker)));
  const select = async (marker, value) => act(() => { const input = find(marker); input.value = value; Simulate.change(input, { target: { value } }); });
  const check = async (marker, checked) => act(() => { const node = find(marker); const input = node.tagName === 'INPUT' ? node : node.querySelector('input'); assert.ok(input, marker); input.checked = checked; Simulate.change(input); });
  const update = () => act(() => ReactDOM.render(h.React.createElement(Root), view.container));
  return { h, view, find, query, click, select, check, act, entry: actual.entry, reset() { revision++; return update(); }, host(direction, dark = false) { dir = direction; theme = dark ? webDarkTheme : webLightTheme; return update(); } };
}
for (const [kind, owner] of Object.entries({ stage: 'Stage', settings: 'SettingsGroup', accordion: 'AccordionGroup', state: 'StateMessage' })) {
  test(`${owner} has its own actual source-owned executable scenario`, () => {
    const { entry, source } = scenario(titles[kind]);
    assert.equal(entry.featureTitle, owner); assert.ok(entry.ownerKey.endsWith('#' + owner));
    assert.ok(entry.publicNames.includes(owner)); assert.equal(entry.metadata.host, 'none'); assert.equal(entry.metadata.kind, 'snapshot');
    assert.equal(entry.metadata.contentPlacement, kind === 'stage' ? 'configuration' : 'subject');
    assert.doesNotMatch(source, /components\/|fixtures\/|style=|\.style\./);
  });
}
test('Stage source keeps real portaled Form ownership, dirty refusal and held-save cancellation through Reset', async t => {
  const f = fixture(t, 'stage');
  const trigger = f.find('trigger'), config = f.find('mode');
  assert.equal(f.view.container.querySelectorAll('[data-demo-result]').length, 1);
  assert.equal(f.view.container.querySelector('[data-demo-subject]'), null);
  await f.click('trigger');
  const surface = f.find('surface'), field = surface.querySelector('[name=name]');
  assert.equal(surface.querySelectorAll('form').length, 1);
  assert.equal(surface.querySelector('[data-stage-demo=save]').form, surface.querySelector('form'));
  await f.act(() => Simulate.change(field, { target: { value: 'Unsaved report' } }));
  await f.click('footer-close'); assert.equal(f.find('surface'), surface); assert.ok(f.query('discard-prompt'));
  await f.click('keep-editing'); assert.equal(f.query('discard-prompt'), null);
  await f.check('hold-save', true);
  await f.act(() => Simulate.submit(surface.querySelector('form')));
  assert.match(f.find('pending').textContent, /true/);
  await f.click('footer-close'); assert.match(f.find('status').textContent, /save pending/);
  const oldCompletion = f.find('complete-save');
  await f.reset();
  await f.act(() => Simulate.click(oldCompletion));
  assert.equal(f.find('mode'), config); assert.equal(f.find('trigger'), trigger);
  assert.equal(config.value, 'dialog'); assert.equal(f.find('status').textContent, 'Not saved');
  assert.match(f.find('save-count').textContent, /0/); assert.match(f.find('open').textContent, /false/);
  await f.click('trigger');
  assert.equal(f.find('surface').querySelector('[name=name]').value, 'Quarterly report');
  await f.act(() => Simulate.change(f.find('surface').querySelector('[name=name]'), { target: { value: 'Fresh report' } }));
  await f.act(() => Simulate.submit(f.find('surface').querySelector('form')));
  assert.equal(f.find('status').textContent, 'Saved Fresh report'); assert.match(f.find('pending').textContent, /false/);
  assert.match(f.find('save-count').textContent, /1/);
});
test('Stage source invalidates held work on departure and retains surface children only when selected', async t => {
  const f = fixture(t, 'stage', true);
  await f.check('retention', true); await f.click('trigger');
  const note = f.find('surface-note'); await f.select('surface-note', 'Retained edit');
  await f.click('footer-close'); await f.click('trigger');
  assert.equal(f.find('surface-note'), note); assert.equal(note.value, 'Retained edit');
  await f.check('hold-save', true); await f.act(() => Simulate.submit(f.find('surface').querySelector('form')));
  const completion = f.find('complete-save');
  f.view.unmount(); await f.act(() => Simulate.click(completion));
  assert.equal(completion.isConnected, false); assert.equal(document.querySelector('[data-stage-demo=surface]'), null);
});
test('SettingsGroup source retains caller settings through actual type transitions and reset', async t => {
  const f = fixture(t, 'settings'); const config = f.find('kind');
  assert.equal(f.find('settings').querySelectorAll('[data-settings-group-item]').length, 2);
  await f.select('email', 'draft@example.com');
  await f.act(() => { const input = f.find('settings').querySelector('[role=switch]'); input.checked = false; Simulate.change(input); });
  await f.click('action'); await f.select('kind', 'accordion');
  assert.equal(f.query('settings'), null); assert.equal(f.find('accordion').querySelector('[role=switch]').checked, false);
  const note = f.find('note'); await f.select('note', 'Kept draft');
  const profile = f.find('accordion').querySelector('button[aria-controls]');
  await f.act(() => profile.click()); await f.act(() => profile.click()); assert.equal(f.find('note'), note); assert.equal(note.value, 'Kept draft');
  await f.select('kind', 'settings'); assert.equal(f.find('email').value, 'draft@example.com'); assert.match(f.find('status').textContent, /Actions: 1/);
  const email = f.find('email'); await f.host('rtl', true); assert.equal(f.find('email'), email);
  await f.reset(); assert.equal(f.find('kind'), config); assert.equal(f.find('email'), email); assert.equal(email.value, 'operations@example.com');
  assert.match(f.find('status').textContent, /Actions: 0/);
});
test('AccordionGroup independent source accepts native toggle data and keeps controlled selection synchronized', async t => {
  const f = fixture(t, 'accordion');
  const profile = f.find('group').querySelector('button[aria-controls]');
  await f.check('multiple', true); await f.select('open-items', 'profile,advanced');
  assert.equal(f.find('open-items').value, 'profile,advanced'); assert.match(f.find('status').textContent, /Toggles: 0/);
  await f.act(() => profile.click());
  assert.equal(f.find('open-items').value, 'advanced'); assert.match(f.find('status').textContent, /Toggles: 1/);
  await f.select('open-items', 'profile'); await f.select('note', 'Uncontrolled edit'); const note = f.find('note');
  await f.select('appearance', 'outline'); await f.host('rtl', true); assert.equal(f.find('note'), note); assert.equal(note.value, 'Uncontrolled edit');
  await f.click('inspect'); assert.match(f.find('observations').textContent, /root:DIV title:SPAN/);
  await f.click('close-all'); assert.equal(f.find('open-items').value, '');
  await f.reset(); assert.equal(f.find('open-items').value, 'profile'); assert.equal(f.find('note').value, 'Retained note');
});
test('StateMessage source preserves real presets, formatted shorthand, literal callbacks, actions and native refs', async t => {
  const f = fixture(t, 'state'); const root = f.find('preview');
  assert.equal(root.querySelector('h3').textContent, 'No documents found'); assert.equal(root.querySelector('p strong').textContent, 'search filters');
  const artwork = root.querySelector('svg').innerHTML;
  await f.select('preset', 'error'); assert.notEqual(root.querySelector('svg').innerHTML, artwork);
  await f.check('slots', true); await f.select('heading', '4'); await f.click('inspect');
  assert.match(f.find('observations').textContent, /div, heading=h4, slot=span/);
  assert.equal(root.querySelector('h4').textContent, '**Literal slot title**'); assert.match(root.textContent, /Callback keeps \*\*literal markers\*\*\./);
  const primary = root.querySelector('button'); await f.act(() => primary.click()); assert.match(f.find('observations').textContent, /Actions: 1/);
  assert.equal(root.querySelectorAll('button')[1].disabled, true);
  await f.act(() => root.querySelectorAll('button')[1].click()); assert.match(f.find('observations').textContent, /Actions: 1/);
  assert.match(root.querySelector('a').href, /sharepoint-framework-overview/);
  await f.check('role', true); assert.equal(root.querySelector('[data-state-message-text]').getAttribute('role'), 'alert');
  assert.equal(root.querySelector('[role=alert]').contains(primary), false);
  await f.select('visual', 'hidden'); assert.equal(root.querySelector('[data-state-message-visual]'), null);
  await f.select('visual', 'custom'); assert.equal(root.querySelector('[role=img]').getAttribute('aria-label'), 'Custom visual');
  await f.host('rtl', true); assert.equal(f.find('preview'), root); assert.equal(root.querySelector('button'), primary);
  await f.reset(); assert.equal(f.find('preview'), root); assert.equal(root.querySelector('h3').textContent, 'No documents found'); assert.match(f.find('observations').textContent, /Actions: 0/);
});

for (const kind of ['settings', 'accordion', 'state']) {
  test(`${kind} standalone scenario keeps independent baseline state and native roots through reset`, async t => {
    const first = fixture(t, kind, true);
    const marker = kind === 'settings' ? 'email' : kind === 'accordion' ? 'title' : 'title';
    const value = first.find(marker).value;
    await first.select(marker, 'Independent local edit'); const input = first.find(marker);
    await first.reset(); assert.equal(first.find(marker), input); assert.equal(input.value, value);
    first.view.unmount(); assert.equal(input.isConnected, false);
    const second = fixture(t, kind, true); assert.equal(second.find(marker).value, value);
  });
}
for (const mode of ['dialog', 'drawer', 'popover']) {
  test(`Stage ${mode} in-surface Reset editor cancels the held save and permits a new real submission`, async t => {
    const f = fixture(t, 'stage');
    await f.select('mode', mode); await f.check('hold-save', true); await f.click('trigger');
    await f.act(() => Simulate.change(f.find('surface').querySelector('[name=name]'), { target: { value: 'Cancelled draft' } }));
    await f.act(() => Simulate.submit(f.find('surface').querySelector('form')));
    assert.match(f.find('pending').textContent, /true/);
    await f.click('reset-editor');
    assert.equal(f.find('mode').value, 'dialog'); assert.equal(f.find('status').textContent, 'Not saved');
    assert.match(f.find('open').textContent, /false/); assert.equal(f.find('trigger').disabled, false);
    await f.click('trigger');
    assert.equal(f.find('surface').querySelector('[name=name]').value, 'Quarterly report');
    await f.act(() => Simulate.submit(f.find('surface').querySelector('form')));
    assert.equal(f.find('status').textContent, 'Saved Quarterly report');
    assert.match(f.find('save-count').textContent, /1/); assert.match(f.find('pending').textContent, /false/);
  });
}
test('Stage Reset before the asynchronous Form callback starts cancels that submission without stranding the gate', async t => {
  const f = fixture(t, 'stage');
  await f.check('hold-save', true); await f.click('trigger');
  await f.act(() => {
    Simulate.submit(f.find('surface').querySelector('form'));
    Simulate.click(f.find('reset-editor'));
  });
  assert.equal(f.find('trigger').disabled, false, 'cancelled pre-callback submission must settle');
  assert.equal(f.find('status').textContent, 'Not saved');
  await f.click('trigger'); await f.act(() => Simulate.submit(f.find('surface').querySelector('form')));
  assert.equal(f.find('status').textContent, 'Saved Quarterly report');
  await f.click('footer-close'); await f.check('hold-save', true); await f.click('trigger');
  await f.act(() => Simulate.submit(f.find('surface').querySelector('form')));
  assert.match(f.find('pending').textContent, /true/);
  await f.click('complete-save'); assert.match(f.find('pending').textContent, /false/);
  assert.equal(f.find('status').textContent, 'Saved Quarterly report'); assert.match(f.find('save-count').textContent, /2/);
});

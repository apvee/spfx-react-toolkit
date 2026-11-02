const assert = require('node:assert/strict');
const path = require('node:path');
const { test } = require('node:test');
const { Simulate } = require('react-dom/test-utils');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const { createHarness } = require('./services-test-harness.cjs');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const modes = [
  ['Id', 'id', 'id', 'List GUID', 'Query by GUID', '11111111-2222-3333-4444-555555555555', 'Read items by list GUID'],
  ['Url', 'url', 'serverRelativeUrl', 'Server-relative list URL', 'Query by URL', '/sites/projects/Lists/Tasks', 'Read items by decoded server-relative URL'],
  ['Path', 'path', 'webRelativePath', 'Web-relative list path', 'Query by path', 'Lists/Tasks', 'Read items by decoded web-relative path'],
];
// Real emitted examples and selector hooks; context and service are deliberately doubled.
// Native HTTP serialization/provider evidence lives in executable-examples-pnp-selectors.
function fixture(t, mode, { crud = false } = {}) {
  const factories = [], calls = [];
  let resolve, reject;
  const context = { sp: {}, isInitialized: true, siteUrl: 'https://tenant.test/sites/projects' };
  const boundaries = {
    './useSPFxPnPContext': { useSPFxPnPContext: () => context },
    '../services/spfx-pnp-list.service': { createSPFxPnPListService: (sp, target, size) => {
      factories.push({ sp, target, size });
      return { query: (builder, options) => {
        const operations = [];
        const items = { select: (...fields) => { operations.push(['select', ...fields]); return items; }, orderBy: (...args) => { operations.push(['orderBy', ...args]); return items; } };
        assert.equal(builder(items), items);
        calls.push({ target, options, operations });
        return new Promise((a, b) => { resolve = a; reject = b; });
      } };
    } },
  };
  const h = createHarness(boundaries);
  const toolkit = { ...boundaries['./useSPFxPnPContext'], ...h.load('hooks/useSPFxPnPList.ts'), ...Object.assign({}, ...modes.map(([suffix]) => h.load('hooks/useSPFxPnPListBy' + suffix + '.ts'))) };
  Object.assign(toolkit, h.load('hooks/useSx.ts'), h.load('helpers/styles/index.ts'));
  boundaries['@apvee/spfx-react-toolkit'] = toolkit;
  t.after(h.close);
  window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  t.mock.method(window, 'setTimeout', (callback, delay) => setTimeout(callback, delay));
  const [suffix, , , , , , title] = mode;
  const selected = crud ? scenario('Manage list items with paging and captured delete consent', 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPList.ts#useSPFxPnPList') : scenario(title, 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPListBy' + suffix + '.ts#useSPFxPnPListBy' + suffix);
  const Component = h.load(selected.file).Demo;
  const { DemoCard } = h.load(path.resolve('apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard.tsx'));
  const view = h.mountComponent(() => h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Component, { render: regions => h.React.createElement(DemoCard, { title: 'useSPFxPnPListBy' + suffix, contentPlacement: 'configuration', ...regions }) })));
  const flush = async () => h.act(async () => { for (let i = 0; i < 8; i++) await Promise.resolve(); await new Promise(r => setTimeout(r, 0)); });
  function button(name) { const element = Array.from(view.container.querySelectorAll('button')).find(b => b.textContent === name); assert.ok(element, name); return element; }
  function input(label, value) {
    const labelNode = Array.from(view.container.querySelectorAll('label')).find(l => l.textContent === label);
    assert.ok(labelNode, label);
    const field = view.container.querySelector('#' + labelNode.htmlFor);
    h.act(() => Simulate.change(field, { target: { value } }));
  }
  return { ...h, ...view, factories, calls, context, flush, button, input,
    click: name => h.act(() => Simulate.click(button(name))),
    succeed: async items => { await h.act(async () => { resolve({ items, hasMore: false, nextSkip: items.length, effectivePageSize: 10 }); }); await flush(); },
    fail: async () => { await h.act(async () => { reject(new Error('Read denied')); }); await flush(); },
  };
}
for (const mode of modes) {
  const [suffix, kind, field, label, action, value] = mode;
  test(suffix + ' emitted selector mounts without queries; blank/whitespace disable commands', async t => {
    const f = fixture(t, mode); await f.flush();
    assert.deepEqual(f.calls, []); assert.deepEqual(f.factories.map(x => x.target.kind), [kind]);
    assert.ok(f.factories.every(x => x.sp === f.context.sp && x.size === 10));
    assert.equal(f.button(action).disabled, true); f.input(label, '   '); assert.equal(f.button(action).disabled, true); assert.deepEqual(f.calls, []);
  });
  test(suffix + ' forwards raw input and first-page query, displays pending/results/error and recovers', async t => {
    const f = fixture(t, mode); f.input(label, value); assert.equal(f.button(action).disabled, false); assert.deepEqual(f.calls, []);
    f.click(action); await f.flush(); assert.equal(f.button(action).disabled, true); assert.ok(f.container.textContent.includes('Loading'));
    assert.deepEqual(f.calls[0], { target: { kind, [field]: value }, options: undefined, operations: [['select', 'Id', 'Title'], ['orderBy', 'Id', false]] });
    await f.succeed([{ Id: 42, Title: 'Result title' }, { Id: 43, Title: null }]);
    assert.ok(f.container.textContent.includes('#42 - Result title')); assert.ok(f.container.textContent.includes('#43 - (untitled)')); assert.equal(f.button(action).disabled, false);
    f.click(action); await f.fail(); assert.ok(f.container.textContent.includes('Read denied')); assert.equal(f.button(action).disabled, false);
    f.click(action); await f.succeed([]); assert.ok(!f.container.textContent.includes('Read denied')); assert.ok(f.container.textContent.includes('No items loaded.'));
  });
  test(suffix + ' query region keeps inputs separate from commands without blank service previews', t => {
    const f = fixture(t, mode); assert.equal(f.container.querySelectorAll('[aria-label$=" preview"]').length, 0); assert.equal(f.container.querySelectorAll('[data-demo-section]').length, 1);
    const section = f.container.querySelector(`[data-demo-section][aria-label="useSPFxPnPListBy${suffix}"]`); assert.ok(section);
    assert.ok(section.querySelector(`[aria-label="useSPFxPnPListBy${suffix} configuration"] input`));
    const commands = section.querySelector(`[aria-label="useSPFxPnPListBy${suffix} commands"]`); assert.equal(commands.querySelectorAll('input').length, 0); assert.equal(commands.textContent, action);
  });
}
// Retargeted at retirement: actual title-list runtime is still deliberately service-doubled here.
test('canonical CRUD defaults to blank explicit title list without requests', async t => {
  const f = fixture(t, modes[0], { crud: true }); await f.flush();
  assert.equal(f.button('Load').disabled, true); assert.deepEqual(f.calls, []);
});

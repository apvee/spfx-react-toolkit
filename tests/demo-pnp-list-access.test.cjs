const assert = require('node:assert/strict');
const path = require('node:path');
const { test } = require('node:test');
const { Simulate } = require('react-dom/test-utils');
const { createHarness } = require('./services-test-harness.cjs');
require('@fluentui/react').registerIcons({ icons: { BulletedList: 'list', Info: 'info', InfoSolid: 'info', ErrorBadge: 'error', Cancel: 'cancel' } });
const demo = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/demos/PnPListAccessDemo.tsx');
const modes = [
  ['Id', 'id', 'id', 'List GUID', 'Query by GUID', '11111111-2222-3333-4444-555555555555'],
  ['Url', 'url', 'serverRelativeUrl', 'Server-relative list URL', 'Query by URL', '/sites/projects/Lists/Tasks'],
  ['Path', 'path', 'webRelativePath', 'Web-relative list path', 'Query by path', 'Lists/Tasks'],
];
function fixture(t) {
  const factories = [], calls = [];
  let resolve, reject;
  const context = { sp: {}, isInitialized: true };
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
  boundaries['@apvee/spfx-react-toolkit'] = Object.assign({}, ...modes.map(([suffix]) => h.load('hooks/useSPFxPnPListBy' + suffix + '.ts')));
  t.after(h.close);
  t.mock.method(window, 'setTimeout', (callback, delay) => setTimeout(callback, delay));
  const view = h.mountComponent(h.load(demo).PnPListAccessDemo);
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
test('all three concrete modes mount without queries and blank inputs disable actions', async t => {
  const f = fixture(t); await f.flush();
  assert.deepEqual(f.calls, []);
  assert.deepEqual(f.factories.map(x => x.target.kind).sort(), ['id', 'path', 'url']);
  assert.ok(f.factories.every(x => x.sp === f.context.sp && x.size === 10));
  for (const [, , , label, action] of modes) {
    assert.equal(f.button(action).disabled, true);
    f.input(label, '   ');
    assert.equal(f.button(action).disabled, true);
  }
  assert.deepEqual(f.calls, []);
});
for (const [suffix, kind, field, label, action, value] of modes) {
  test(suffix + ' forwards input and first-page query, displays pending/results/error and recovers', async t => {
    const f = fixture(t);
    f.input(label, value);
    assert.equal(f.button(action).disabled, false);
    assert.deepEqual(f.calls, []);
    f.click(action);
    await f.flush();
    assert.equal(f.button(action).disabled, true);
    assert.ok(f.container.textContent.includes('Loading'));
    assert.deepEqual(f.calls[0], { target: { kind, [field]: value }, options: undefined, operations: [['select', 'Id', 'Title'], ['orderBy', 'Id', false]] });
    await f.succeed([{ Id: 42, Title: 'Result title' }, { Id: 43, Title: null }]);
    assert.ok(f.container.textContent.includes('#42 - Result title'));
    assert.ok(f.container.textContent.includes('#43 - (untitled)'));
    assert.equal(f.button(action).disabled, false);
    f.click(action); await f.fail();
    assert.ok(f.container.textContent.includes('Read denied'));
    assert.equal(f.button(action).disabled, false);
    f.click(action); await f.succeed([]);
    assert.ok(!f.container.textContent.includes('Read denied'));
    assert.ok(f.container.textContent.includes('No items loaded.'));
  });
}

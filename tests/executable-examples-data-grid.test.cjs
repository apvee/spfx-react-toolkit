process.env.NODE_ENV = 'test';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const ts = require('typescript');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const host = require('./fixtures/form-selection-harness.cjs');
const source = require('./fixtures/formatted-text-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const titles = {
  typed: 'Explore typed document columns and controlled interaction',
  descriptors: 'Supply descriptor columns with explicit sort and row order',
  dynamic: 'Classify dynamic row values and preserve controlled selection',
};
let library;
test.before(async () => {
  await host.setup();
  document.createRange = () => ({ selectNodeContents() {}, getClientRects() { return []; } });
  library = await source.loadModule('tests/fixtures/sx-browser/data-grid-toolkit-boundary.ts');
});
test.after(async () => host.cleanup());
function fixture(t, kind, standalone = false, transform = text => text) {
  const actual = scenario(titles[kind]);
  assert.equal(actual.entry.featureTitle, kind === 'dynamic' ? 'DataGrid' : 'createDataGrid');
  const compiled = ts.transpileModule(transform(actual.source), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React } });
  const module = { exports: {} };
  new Function('require', 'module', 'exports', compiled.outputText)(name => name === '@apvee/spfx-react-toolkit' ? library : require(name), module, module.exports);
  let revision = 7, renders = 0, retry;
  const render = regions => {
    renders++;
    function inspect(node) {
      if (!node || typeof node !== 'object') return;
      if (Array.isArray(node)) { node.forEach(inspect); return; }
      if (node.props?.id === 'data-grid-main') retry = node.props.onRetry;
      inspect(node.props?.children);
    }
    inspect(regions.children);
    return host.React.createElement('section', {}, regions.controls, regions.commands, regions.status, regions.children, regions.diagnostics);
  };
  const element = () => host.React.createElement(module.exports.Demo, { resetRevision: revision, render: standalone ? undefined : render });
  const view = host.mount(element()); t.after(view.close);
  const find = id => { const node = view.host.querySelector(`#${id}`); assert.ok(node, id); return node; };
  const change = (id, value) => host.change(find(id), value);
  const click = id => host.click(find(id));
  const check = (id, checked) => { const node = find(id); const input = node.tagName === 'INPUT' ? node : node.querySelector('input'); input.checked = checked; host.act(() => Simulate.change(input)); };
  const json = id => JSON.parse(find(id).textContent);
  const grid = () => view.host.querySelector('[data-grid-root]');
  const rows = () => [...view.host.querySelectorAll('#data-grid-main [data-grid-row-id]')];
  const ids = () => rows().map(row => row.getAttribute('data-grid-row-id'));
  const cell = (id, column) => find('data-grid-main').querySelector(`[data-grid-row-id="${id}"] [data-grid-column-id="${column}"]`);
  return { view, find, change, click, check, json, grid, rows, ids, cell, renders: () => renders, retry: () => retry,
    reset() { revision++; view.render(element()); },
    rerender() { view.render(element()); },
    setRevision(value) { revision = value; view.render(element()); } };
}
for (const kind of Object.keys(titles)) {
  test(`${kind} native selection preserves off-page IDs, single proposal and none activation`, t => {
    const f = fixture(t, kind), first = kind === 'dynamic' ? 'valid' : 'a';
    assert.equal(f.view.host.querySelectorAll('[role=grid]').length, 1);
    assert.deepEqual(f.json('dg-selection'), ['off-page']);
    host.click(f.rows()[0]); assert.deepEqual(f.json('dg-selection').sort(), [first, 'off-page'].sort());
    const second = kind === 'dynamic' ? 'missing' : 'b';
    host.click(f.rows()[1]); f.check('dg-page', true); host.click(f.rows()[0]); assert.deepEqual(f.json('dg-selection').sort(), [second, 'off-page'].sort());
    f.change('dg-selection-mode', 'single'); host.click(f.rows()[0]); assert.deepEqual(f.json('dg-selection'), [first]);
    assert.equal(f.find('dg-activation-output').textContent, 'none');
    f.change('dg-selection-mode', 'none');
    assert.equal(f.grid().querySelectorAll('input[type=checkbox],input[type=radio]').length, 0);
    const cell = f.rows()[0].querySelector('[role=gridcell]'); cell.focus(); host.key(cell, 'Enter');
    assert.deepEqual(f.json('dg-activation-output'), { rowId: first });
    f.check('dg-loading', true); f.reset(); assert.deepEqual(f.json('dg-selection'), ['off-page']);
  });
  test(`${kind} native status priority, error retry, ref detach/restore, strings and reset`, t => {
    const f = fixture(t, kind), outer = f.find('data-grid-demo'), original = f.grid();
    f.click('dg-inspect-ref'); assert.equal(f.find('dg-ref').textContent, 'div');
    f.check('dg-strings', true); assert.equal(f.grid().querySelector('.fui-DataGridHeader input').getAttribute('aria-label'), 'Select every document');
    f.check('dg-loading', true); assert.equal(f.grid(), original); assert.equal(f.grid().getAttribute('aria-busy'), 'true');
    assert.equal(f.grid().querySelectorAll('[data-grid-loading-row]').length, kind === 'dynamic' ? 1 : 2);
    assert.equal(f.grid().querySelectorAll('input[type=checkbox]').length, 0);
    assert.equal(f.grid().querySelectorAll('[data-grid-loading-kind]').length, 4 * (kind === 'typed' ? 5 : 4));
    f.check('dg-loading', false); f.check('dg-error', true); f.click('dg-inspect-slot-refs');
    assert.match(f.find('dg-slot-refs').textContent, /error=DIV/); assert.equal(f.grid(), original);
    host.click([...f.view.host.querySelectorAll('button')].find(node => node.textContent === 'Retry')); assert.equal(f.find('dg-retries').textContent, '1');
    f.check('dg-error', true); f.check('dg-empty', true); assert.equal(f.grid(), null); assert.match(f.find('dg-frame').textContent, /Consumer-owned loading error/);
    f.check('dg-error', false); assert.match(f.find('dg-frame').textContent, /No team documents/);
    f.click('dg-inspect-ref'); assert.equal(f.find('dg-ref').textContent, 'absent');
    f.change('dg-empty-state','noResults'); f.check('dg-custom-status-slots',true); assert.match(f.find('dg-frame').textContent,/Consumer-owned no-results slot/);
    f.change('dg-empty-state','empty'); assert.match(f.find('dg-frame').textContent,/Consumer-owned empty slot/);
    f.check('dg-slot-suppress',true); assert.equal(f.find('dg-frame').textContent, ''); assert.equal(f.grid(),null);
    f.check('dg-hidden',true); assert.match(f.find('dg-frame').textContent,/No visible columns/);
    f.reset(); assert.equal(f.find('data-grid-demo'),outer); assert.ok(f.grid()); assert.equal(f.find('dg-retries').textContent,'0');
    assert.equal(f.find('dg-strings').checked,false); f.click('dg-inspect-slot-refs'); assert.match(f.find('dg-slot-refs').textContent,/error=absent/);
  });
}
for (const kind of ['typed','descriptors']) test(`${kind} sort intent is controlled and consumer apply/reverse changes supplied row order`, t => {
  const f = fixture(t,kind), initial = f.ids();
  const header = f.grid().querySelector('[role=columnheader][data-grid-column-id=title]').querySelector('button,[role=button]'); host.click(header);
  assert.deepEqual(f.json('dg-sort'),{ columnId:'title',direction:'ascending' }); assert.deepEqual(f.ids(),initial);
  host.click(header); assert.deepEqual(f.json('dg-sort'),{ columnId:'title',direction:'descending' }); assert.deepEqual(f.ids(),initial);
  host.click(header); f.click('dg-apply-sort'); assert.deepEqual(f.ids(),['b','a']);
  f.click('dg-reverse'); assert.deepEqual(f.ids(),['a','b']);
  f.check('dg-page',true); assert.deepEqual(f.ids(),['a']);
});
test('typed exact five presets retain controlled models, real chart values, slots and isolated actions', t => {
  const f=fixture(t,'typed'); host.click(f.rows()[0]);
  const expected={ overview:['title','total','status','owner','actions'], dates:['title','enabled','day','instant'], team:['title','link','owner','people','tags'], metrics:['title','progress','rating','trend'], custom:['title','enabled','icon','custom'] };
  for(const [preset,columns] of Object.entries(expected)) {
    f.change('dg-preset',preset);
    assert.deepEqual([...f.grid().querySelectorAll('.fui-DataGridHeader [data-grid-column-id]')].map(node=>node.getAttribute('data-grid-column-id')),columns);
    assert.deepEqual(f.json('dg-selection').sort(),['a','off-page']);
  }
  f.change('dg-preset','overview'); f.click('dg-inspect-slot-refs'); assert.match(f.find('dg-slot-refs').textContent,/header=SPAN/);
  host.click(f.grid().querySelector('[data-grid-slot-callback=header]')); assert.equal(f.find('dg-slot-clicks').textContent,'1');
  assert.match(f.cell('a','total').textContent,/12\.50/);
  const selectBefore=f.json('dg-selection'); host.click([...f.grid().querySelectorAll('button')].find(node=>node.textContent==='Open report'));
  assert.deepEqual(f.json('dg-action'),{rowId:'a',actionId:'review'}); assert.deepEqual(f.json('dg-selection'),selectBefore);
  const menu = f.cell('a','actions').querySelector('button[aria-label="More actions"]'); assert.ok(menu); host.click(menu);
  const archive = [...document.querySelectorAll('[role=menuitem]')].find(node => node.textContent==='Archive record'); assert.ok(archive); host.click(archive);
  assert.deepEqual(f.json('dg-action'),{rowId:'a',actionId:'archive'});
  f.change('dg-preset','team'); const link=f.cell('a','link').querySelector('a'); host.click(link);
  assert.deepEqual(f.json('dg-navigation'),{rowId:'a',value:'Quarterly report',href:'#details-a'});
  f.change('dg-selection-mode','none'); host.click(link); assert.equal(f.find('dg-activation-output').textContent,'none');
  f.change('dg-preset','metrics'); assert.ok(f.cell('a','trend').querySelector('[role=img]'));
  const summary=f.cell('a','trend').textContent; f.click('dg-update-sparkline'); assert.notEqual(f.cell('a','trend').textContent,summary);
  f.change('dg-preset','dates'); assert.match(f.cell('b','day').textContent,/7/); assert.equal(f.cell('b','enabled').querySelector('[role=img]').getAttribute('aria-label'),'Enabled no');
});
test('typed mounted custom button survives width/loading and retained reset callback restores baseline without Demo remount', t=>{
  const f=fixture(t,'typed'),outer=f.find('data-grid-demo'); f.change('dg-preset','custom'); const custom=f.cell('a','custom').querySelector('button'); host.click(custom); assert.equal(custom.textContent,'Custom 1');
  f.change('dg-width','320'); assert.equal(f.cell('a','custom').querySelector('button'),custom);
  f.check('dg-loading',true); f.check('dg-loading',false); assert.equal(f.cell('a','custom').querySelector('button'),custom); assert.equal(custom.textContent,'Custom 1');
  f.rerender(); assert.equal(custom.textContent,'Custom 1');
  f.reset(); assert.equal(f.find('data-grid-demo'),outer); assert.equal(f.find('dg-preset').value,'overview');
  f.change('dg-preset','custom'); assert.equal(f.cell('a','custom').querySelector('button').textContent,'Custom 0');
  f.change('dg-selection-mode','none'); host.click(f.cell('a','custom').querySelector('button')); assert.equal(f.find('dg-activation-output').textContent,'none');
});
test('dynamic empty text, zero and false remain valid; missing/invalid stay distinct and placeholder ref suppression detaches', t=>{
  const f=fixture(t,'dynamic'); assert.deepEqual(f.ids(),['valid','missing','invalid']);
  assert.equal(f.cell('valid','text').textContent,''); assert.equal(f.cell('valid','number').textContent,'0'); assert.equal(f.cell('valid','enabled').textContent,'No');
  assert.equal(f.cell('missing','text').textContent,'Missing text'); assert.equal(f.cell('invalid','text').textContent,'Invalid text');
  assert.equal(f.cell('missing','number').textContent,'Missing number'); assert.equal(f.cell('invalid','number').textContent,'Invalid number');
  assert.notEqual(f.cell('invalid','enabled').textContent,'Yes');
  f.click('dg-inspect-slot-refs'); assert.match(f.find('dg-slot-refs').textContent,/missing=SPAN/);
  f.check('dg-slot-suppress',true); f.click('dg-inspect-slot-refs'); assert.match(f.find('dg-slot-refs').textContent,/missing=absent/);
  assert.equal(f.view.host.querySelector('#dg-apply-sort,#dg-update-sparkline,#dg-color,#dg-preset'),null);
  f.click('dg-reverse'); assert.deepEqual(f.ids(),['invalid','missing','valid']); f.reset(); assert.deepEqual(f.ids(),['valid','missing','invalid']);
});
test('all three actual emitted modules are portable, metadata-correct and keep independent conformance reachable',()=>{
  for(const [kind,title] of Object.entries(titles)) {
    const {entry,source}=scenario(title); assert.equal(entry.metadata.host,'none'); assert.equal(entry.metadata.kind,'snapshot'); assert.equal(entry.metadata.contentPlacement,'subject'); assert.ok(entry.metadata.observe.trim());
    assert.doesNotMatch(source,/components\/|panels\/|style=|\.style\.|\bany\b|key=\{resetRevision\}/);
    assert.match(source,/resetRevision/); assert.equal(entry.featureTitle,kind==='dynamic'?'DataGrid':'createDataGrid');
  }
  const fixture=fs.readFileSync('tests/fixtures/demo-control-isolation/data-grid-regressions.tsx','utf8');
  assert.match(fixture,/Review this localized record in detail/); assert.match(fixture,/data-grid-activation/);
});
test('dynamic standalone fallback mounts the same real grid',t=>{ const f=fixture(t,'dynamic',true); assert.ok(f.grid()); assert.equal(f.ids().length,3); });

test('first nonzero reset observation does nothing and the original native retry callback belongs to the retained Demo',t=>{
  const f=fixture(t,'descriptors'); assert.equal(f.renders(),1);
  const retry=f.retry(),outer=f.find('data-grid-demo'); assert.equal(typeof retry,'function');
  f.click('dg-reverse'); f.reset(); assert.equal(f.find('data-grid-demo'),outer); assert.deepEqual(f.ids(),['a','b']);
  host.act(()=>retry()); assert.equal(f.find('dg-retries').textContent,'1');
});
test('sort regression oracle rejects a consumer apply command that reports intent without reordering rows',t=>{
  const f=fixture(t,'descriptors',false,text=>{
    const command='setRows(previous => [...previous].sort((left, right) => direction * left.title.localeCompare(right.title)));';
    assert.ok(text.includes(command),'mutates the actual emitted consumer sort command'); return text.replace(command,'void direction;');
  });
  host.click(f.grid().querySelector('[role=columnheader][data-grid-column-id=title]').querySelector('button,[role=button]')); f.click('dg-apply-sort');
  assert.throws(()=>assert.deepEqual(f.ids(),['b','a']),assert.AssertionError,'same independent order oracle rejects the actual no-op mutation');
});

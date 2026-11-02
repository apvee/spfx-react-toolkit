process.env.NODE_ENV = 'test';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const { Simulate } = require('react-dom/test-utils');
const ReactDOM = require('react-dom');
const titles = {
  typed: 'Edit typed reports across native fields and composition',
  schema: 'Validate and transform quantity with recoverable failures',
  dynamic: 'Edit a dynamic report with explicit save and reset',
};
const savedGlobals = new Map();
for (const name of ['Node', 'NodeFilter', 'Element', 'HTMLElement', 'Event', 'CustomEvent', 'MutationObserver', 'FileList']) {
  savedGlobals.set(name, Object.getOwnPropertyDescriptor(global, name));
  Object.defineProperty(global, name, { configurable: true, writable: true, value: window[name] });
}
Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
window.matchMedia ??= () => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} });
test.after(() => { for (const [name, descriptor] of savedGlobals) { if (descriptor) Object.defineProperty(global, name, descriptor); else delete global[name]; } });
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const { RendererProvider } = require('@griffel/react');
const { createDOMRenderer } = require('@griffel/core');
function fixture(t, kind, standalone = false) {
  const actual = scenario(titles[kind]), toolkit = {};
  const h = createHarness({ '@apvee/spfx-react-toolkit': toolkit }); t.after(h.close);
  Object.assign(toolkit, h.load('helpers/styles/index.ts'), h.load('hooks/useSx.ts'), h.load('controls/form/create-form.ts'), h.load('controls/form/form.tsx'), h.load('controls/settings-group/settings-group.tsx'));
  const { Demo } = h.load(actual.file); const renderer = createDOMRenderer(null);
  let revision = 7, completion;
  function inspect(node) {
    if (!node || typeof node !== 'object') return;
    if (Array.isArray(node)) { node.forEach(inspect); return; }
    if (node.props?.['data-form-demo'] === 'complete-save') completion = node.props.onClick;
    inspect(node.props?.children);
  }
  const render = regions => { Object.values(regions).forEach(inspect); return h.React.createElement('section', {}, regions.controls, regions.commands, regions.status, regions.children, regions.diagnostics); };
  function Root() { return h.React.createElement(RendererProvider, { renderer }, h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Demo, { resetRevision: revision, render: standalone ? undefined : render }))); }
  const view = h.mountComponent(Root);
  const query = marker => view.container.querySelector(`[data-form-demo="${marker}"]`);
  const find = marker => { const node = query(marker); assert.ok(node, marker); return node; };
  const act = async run => h.act(async () => { await run(); });
  const click = marker => act(() => Simulate.click(find(marker)));
  const check = (marker, checked) => act(() => { const node = find(marker); const input = node.tagName === 'INPUT' ? node : node.querySelector('input'); input.checked = checked; Simulate.change(input); });
  const change = (node, value) => act(() => { node.value = value; Simulate.change(node, { target: { value } }); });
  const select = (marker, value) => change(find(marker), value);
  const submit = () => act(() => Simulate.submit(find('root')));
  const live = () => JSON.parse(find('values').textContent);
  const output = () => find('saved').textContent ? JSON.parse(find('saved').textContent) : undefined;
  return { h, view, act, find, query, click, check, change, select, submit, live, output, entry: actual.entry, completion: () => completion,
    resetNow() { revision++; ReactDOM.render(h.React.createElement(Root), view.container); },
    reset() { return act(() => { revision++; ReactDOM.render(h.React.createElement(Root), view.container); }); } };
}
for (const kind of Object.keys(titles)) {
  test(`${kind} real Form cancels a pre-callback Reset, drains its gate, then accepts normal and held submissions`, async t => {
    const f = fixture(t, kind); const root = f.find('root');
    assert.equal(f.entry.featureTitle, kind === 'dynamic' ? 'Form' : 'createForm');
    await f.check('save-delay', true);
    await f.act(() => { Simulate.submit(root); f.resetNow(); });
    assert.equal(f.find('root'), root, 'scenario reset preserves the real mounted Form');
    assert.match(f.find('pending').textContent, /false/); assert.match(f.find('status').textContent, /Saves: 0/);
    assert.equal(f.output(), undefined);
    await f.submit(); assert.ok(f.output()); assert.match(f.find('status').textContent, /Saves: 1/);
    await f.check('save-delay', true); await f.submit();
    assert.match(f.find('pending').textContent, /true/); await f.click('complete-save');
    assert.match(f.find('pending').textContent, /false/); assert.match(f.find('status').textContent, /Saves: 2/);
  });
  test(`${kind} held Reset and departure silence the original completion and preserve failure/retry snapshots`, async t => {
    const f = fixture(t, kind), input = () => f.find('root').querySelector(`input[name="${kind === 'typed' ? 'reference' : kind === 'schema' ? 'quantity' : 'dynamic'}"]`);
    await f.check('save-delay', true); await f.change(input(), kind === 'schema' ? '23' : 'Before'); await f.submit();
    assert.equal(typeof f.completion(), 'function');
    await f.change(input(), kind === 'schema' ? '24' : 'During');
    await f.act(() => { Simulate.submit(f.find('root')); Simulate.submit(f.find('root')); });
    assert.match(f.find('status').textContent, /Saves: 1/);
    await f.click('complete-save'); assert.equal(f.output()[kind === 'typed' ? 'reference' : kind === 'schema' ? 'quantity' : 'dynamic'], kind === 'schema' ? 23 : 'Before');
    await f.submit(); const retained = f.completion(); await f.reset(); await f.act(() => retained());
    assert.match(f.find('status').textContent, /Saves: 0/); assert.equal(f.output(), undefined); assert.match(f.find('pending').textContent, /false/);
    await f.check('save-fail', true); await f.submit(); assert.match(f.find('status').textContent, /Save failed/); assert.deepEqual(JSON.parse(f.find('errors').textContent), ['root.submit: Unable to submit the form. Please try again.']);
    await f.check('save-fail', false); await f.submit(); assert.ok(f.output()); assert.match(f.find('pending').textContent, /false/);
    await f.check('save-delay', true); await f.submit(); await f.act(() => retained()); assert.match(f.find('pending').textContent, /true/, 'retained old completion must not release the new gate'); const departed = f.completion();
    f.view.unmount(); await f.act(() => departed()); assert.equal(f.view.container.querySelector('form'), null);
  });
}
test('typed report retains native domains and the draft across all adapters and three composition modes', async t => {
  const f = fixture(t, 'typed');
  const adapters = ['Input','Textarea','SearchBox','Checkbox','Switch','SpinButton','Slider','Select','RadioGroup','Dropdown','DropdownMulti','CheckboxGroup','Combobox','ComboboxMulti','TagPicker','Rating','ColorPicker','SwatchPicker','DatePicker','TimePicker','DateTime','FileInput'];
  for (const adapter of adapters) { await f.select('adapter', adapter); assert.ok(f.view.container.querySelector(`[data-form-control="${adapter}"]`), adapter); assert.equal(f.view.container.querySelectorAll('form').length, 1); }
  assert.equal(f.live().files, null); assert.equal(f.view.container.querySelector('input[type=file]').files instanceof window.FileList, true, 'real browser-owned empty FileList');
  await f.select('adapter', 'Input'); await f.change(f.find('root').querySelector('[name=custom]'), 'Text draft');
  await f.select('adapter', 'Textarea'); assert.equal(f.find('root').querySelector('[name=notes]').value, 'Text draft');
  await f.change(f.find('root').querySelector('[name=reference]'), 'Report draft');
  await f.select('adapter', 'SpinButton'); await f.change(f.find('root').querySelector('[name=quantity]'), '0');
  await f.act(() => Simulate.blur(f.find('root').querySelector('[name=quantity]'))); assert.equal(f.live().quantity, 0);
  await f.submit(); assert.equal(f.output().quantity, 0); assert.equal(f.output().approved, false);
  await f.select('adapter', 'Select'); await f.change(f.find('root').querySelector('[name=country]'), ''); await f.submit(); assert.equal(f.output().country, '');
  await f.select('adapter', 'SpinButton'); assert.equal(f.live().quantity, 2); assert.equal(f.live().notes, 'Text draft');
  for (const mode of ['object','array','rules']) { await f.select('mode', mode); assert.equal(f.live().reference, 'Report draft'); }
  await f.select('mode','object'); assert.ok(f.find('root').querySelector('fieldset')); assert.ok(f.find('root').querySelector('[name="customer.name"]'));
  await f.select('mode','array'); for (const action of ['append','prepend','insert','remove','update','replace','move','swap']) assert.ok(f.find('root').querySelector(`[data-form-array-action="${action}"]`));
});
test('schema transforms without replacing live input, contains invalid/throw, and dynamic active dirty/disabled/load are real', async t => {
  const f = fixture(t,'schema'), input = f.find('root').querySelector('[name=quantity]');
  await f.change(input,'23'); await f.submit(); assert.deepEqual(f.output(), { quantity: 23 }); assert.equal(f.live().quantity,'23');
  for (const invalid of ['', 'Infinity', 'throw']) { await f.change(input,invalid); await f.submit(); assert.equal(f.output().quantity,23); assert.notEqual(f.find('errors').textContent,'[]'); }
  await f.change(input,'25'); await f.submit(); assert.deepEqual(f.output(),{ quantity:25 });
  const d = fixture(t,'dynamic'), dynamic = d.find('root').querySelector('[name=dynamic]');
  await d.change(dynamic,'Edited'); assert.match(d.find('status').textContent,/Dirty: true/); await d.check('disabled',true); assert.equal(dynamic.disabled,true); await d.check('disabled',false);
  await d.click('new-record'); assert.equal(dynamic.value,'Annual report'); assert.match(d.find('status').textContent,/Dirty: false/);
});
test('typed source commits empty and Date domains, rejects invalid dates, and exposes real file clearing', async t => {
  const f = fixture(t,'typed');
  await f.select('adapter','SpinButton'); let input = f.find('root').querySelector('[name=quantity]');
  await f.change(input,''); await f.act(() => Simulate.blur(input)); assert.equal(f.live().quantity,null);
  await f.submit(); assert.equal(f.output().quantity,null);
  await f.select('adapter','DatePicker'); input=f.find('root').querySelector('[name=date]');
  assert.deepEqual(f.live().date,{ year:2026,month:10,day:8,hour:0,minute:0 });
  await f.change(input,'2026-02-31'); await f.act(() => Simulate.blur(input)); const count=f.find('status').textContent.match(/Saves: \d+/)[0];
  await f.submit(); assert.equal(f.find('status').textContent.match(/Saves: \d+/)[0],count); assert.equal(input.value,'2026-02-31');
  await f.change(input,'2026-10-12'); await f.act(() => Simulate.blur(input)); await f.submit(); assert.deepEqual(f.output().date,{ year:2026,month:10,day:12,hour:0,minute:0 });
  await f.select('adapter','FileInput'); const file=f.find('root').querySelector('input[type=file]');
  assert.equal(file.files instanceof window.FileList,true); assert.equal(file.files.length,0);
  await f.click('clear-file'); assert.equal(f.live().files,null); await f.submit(); assert.equal(f.output().files,null);
  await f.select('adapter','SearchBox'); input=f.find('root').querySelector('[name=query]'); await f.change(input,''); await f.submit(); assert.equal(f.output().query,'');
});
test('typed source executes complete row operations, section UI bounds, Settings naming and scoped disable', async t => {
  const f=fixture(t,'typed'); await f.select('mode','object');
  const settings=f.find('settings-reference'), reference=settings.querySelector('input'); assert.equal(settings.querySelectorAll('label').length,1);
  await f.check('scope-disabled',true); assert.equal(f.find('root').querySelector('[name="address.city"]').disabled,true); assert.equal(f.find('root').querySelector('[name="customer.name"]').disabled,false);
  await f.submit(); assert.equal(f.output().address,undefined); assert.equal(f.output().reference,'Quarterly report');
  await f.select('mode','array');
  const action=name=>f.act(()=>Simulate.click(f.find('root').querySelector(`[data-form-array-action="${name}"]`)));
  await action('append'); assert.equal(f.live().rows.length,3); assert.deepEqual(f.live().rows[2],{ code:'New row',note:'' });
  await action('prepend'); assert.equal(f.live().rows[0].code,'Prepended'); await action('insert'); assert.equal(f.live().rows[1].code,'Inserted');
  await action('update'); assert.deepEqual(f.live().rows[0],{ code:'Updated',note:'Replacement' });
  const keys=()=>[...f.find('root').querySelectorAll('[data-form-headless-key]')].map(node=>node.getAttribute('data-form-headless-key'));
  const first=keys()[0]; await action('move'); assert.equal(f.live().rows.at(-1).code,'Updated'); assert.equal(keys().at(-1),first);
  const before=keys(); await action('swap'); assert.deepEqual(keys().slice(0,2),[before[1],before[0]]);
  const length=f.live().rows.length; await action('remove'); assert.equal(f.live().rows.length,length-1);
  await action('replace'); assert.deepEqual(f.live().rows,[{ code:'Replaced',note:'Only row' }]);
  const section=f.find('array-section'); const add=[...section.querySelectorAll('button')].find(node=>node.textContent.includes('Add activity')); assert.ok(add);
  await f.act(()=>Simulate.click(add)); assert.equal(f.live().lines.length,3); assert.equal(add.disabled,true);
  const remove=[...section.querySelectorAll('button')].find(node=>node.getAttribute('aria-label')==='Remove Activities 1'); assert.ok(remove); await f.act(()=>Simulate.click(remove)); assert.equal(f.live().lines.length,2);
  await f.submit(); assert.deepEqual(f.output().rows,[{ code:'Replaced',note:'Only row' }]);
});
test('typed source submits false and empty arrays without coercion and keeps DateTime siblings independent', async t => {
  const f=fixture(t,'typed');
  await f.select('adapter','Switch'); let input=f.find('root').querySelector('[name=enabled]');
  await f.act(()=>{ input.checked=false; Simulate.change(input); }); await f.submit(); assert.equal(f.output().enabled,false);
  await f.select('adapter','CheckboxGroup'); input=f.find('root').querySelector('[data-form-control=CheckboxGroup] input');
  await f.act(()=>{ input.checked=false; Simulate.change(input); }); await f.submit(); assert.deepEqual(f.output().checks,[]);
  await f.select('adapter','DateTime');
  const appointment=f.live().appointment; assert.deepEqual(appointment,{ date:{ year:2026,month:10,day:8,hour:0,minute:0 },time:{ year:1970,month:1,day:1,hour:10,minute:0 } });
  input=f.find('root').querySelector('[data-form-control=DateTime] input'); await f.change(input,'2026-10-15'); await f.act(()=>Simulate.blur(input));
  await f.submit(); assert.deepEqual(f.output().appointment.time,appointment.time); assert.deepEqual(f.output().appointment.date,{ year:2026,month:10,day:15,hour:0,minute:0 });
});
for (const kind of Object.keys(titles)) test(`${kind} pending failure configuration is snapshotted and cancelled failure never reaches the new baseline`, async t => {
  const f=fixture(t,kind); await f.check('save-delay',true); await f.check('save-fail',true); await f.submit(); await f.check('save-fail',false); await f.click('complete-save');
  assert.match(f.find('status').textContent,/Save failed/); assert.equal(f.output(),undefined);
  await f.check('save-fail',true); await f.submit(); const retained=f.completion(); await f.reset(); await f.act(()=>retained());
  assert.equal(f.find('errors').textContent,'[]'); assert.match(f.find('status').textContent,/Saves: 0/);
  await f.submit(); assert.ok(f.output());
});
test('schema Reset drains a deferred invalid validation before clearing errors and restores a usable baseline', async t => {
  const f=fixture(t,'schema'); await f.change(f.find('root').querySelector('[name=quantity]'),'');
  await f.act(()=>{ Simulate.submit(f.find('root')); f.resetNow(); });
  assert.equal(f.live().quantity,'12'); assert.equal(f.find('errors').textContent,'[]'); assert.match(f.find('pending').textContent,/false/);
  await f.submit(); assert.deepEqual(f.output(),{ quantity:12 });
});
for (const kind of Object.keys(titles)) test(`${kind} native per-form dictionary reaches current save failures`, async t => {
  const f=fixture(t,kind); await f.check('strings',true); await f.check('save-fail',true); await f.submit();
  assert.deepEqual(JSON.parse(f.find('errors').textContent),['root.submit: The report could not be saved. Please retry.']);
  if(kind==='typed') {
    await f.select('adapter','FileInput'); assert.ok([...f.find('root').querySelectorAll('button')].find(node=>node.textContent==='Attach report'));
    await f.select('mode','array'); assert.ok([...f.find('array-section').querySelectorAll('button')].find(node=>node.textContent==='Add report line'));
  }
});

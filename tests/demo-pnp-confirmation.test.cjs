const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const { createDOMRenderer } = require('@griffel/core');
const { RendererProvider } = require('@griffel/react');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const selected = scenario('Manage list items with paging and captured delete consent', 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPList.ts#useSPFxPnPList');
function fixture(t, remove = async () => {}) {
  const calls = [], operations = [];
  let replaceOperation;
  const context = { sp: {}, isInitialized: true, siteUrl: 'https://confirmation.fixture.invalid/sites/demo' };
  let confirmCalls = 0;
  const previousConfirm = Object.getOwnPropertyDescriptor(global, 'confirm');
  Object.defineProperty(global, 'confirm', { configurable: true, value: () => { confirmCalls++; return false; } });
  t.after(() => { if (previousConfirm) Object.defineProperty(global, 'confirm', previousConfirm); else delete global.confirm; assert.equal(confirmCalls, 0, 'native confirm is never called'); });
  let publishRows;
  const toolkit = { useSPFxPnPContext: () => context, useSPFxPnPList: () => {
    const [items, setItems] = h.React.useState([{ Id: 42, Title: 'Intended item' }, { Id: 73, Title: 'Other item' }]);
    const [operation, setOperation] = h.React.useState(0);
    replaceOperation = () => setOperation(value => value + 1);
    publishRows = id => setItems(current => current.filter(item => item.Id !== id));
    return { items, loading: false, loadingMore: false, error: undefined, isEmpty: items.length === 0, hasMore: false, query: async () => items, clearError: () => {}, create: async () => { throw new Error('Other writes forbidden'); }, update: async () => { throw new Error('Other writes forbidden'); }, loadMore: async () => [], remove: id => { calls.push(id); operations.push({id, operation}); return remove(id); } };
  } };
  const boundaries = { '@apvee/spfx-react-toolkit': toolkit };
  const h = createHarness(boundaries);
  Object.assign(toolkit, h.load('hooks/useSx.ts'), h.load('helpers/styles/index.ts'));
  for (const name of ['@apvee/spfx-react-toolkit/hooks/useSPFxPnPContext', '@apvee/spfx-react-toolkit/hooks/useSPFxPnPList', '@apvee/spfx-react-toolkit/styles']) boundaries[name] = toolkit;
  for (const name of ['Node', 'NodeFilter', 'HTMLElement', 'Element']) {
    const previous = Object.getOwnPropertyDescriptor(global, name);
    Object.defineProperty(global, name, { configurable: true, value: window[name] });
    t.after(() => { if (previous) Object.defineProperty(global, name, previous); else delete global[name]; });
  }
  Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  const Component = h.load(selected.file).Demo;
  const renderer = createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  const view = h.mountComponent(() => h.React.createElement(RendererProvider, { renderer }, h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Component))));
  t.after(h.close);
  const buttons = name => [...document.querySelectorAll('button')].filter(node => node.textContent === name);
  const click = button => h.act(() => Simulate.click(button));
  const listInput = view.container.querySelector('[aria-label="List Title"]');
  h.act(() => { listInput.value = 'Mock list'; Simulate.change(listInput, { target: { value: 'Mock list' } }); });
  h.act(() => { Simulate.click(buttons('Load')[0]); });
  const open = () => { const trigger = buttons('Delete')[0]; trigger.focus(); click(trigger); return trigger; };
  return { ...h, ...view, calls, operations, replaceOperation: () => h.act(() => replaceOperation()), buttons, click, open, publishRows: id => h.act(() => publishRows(id)) };
}
test('PnP delete opens Fluent consent and cancellation restores the intended trigger without removal', async t => {
  const f = fixture(t); await f.act(async () => {}); const trigger = f.open();
  assert.ok(document.querySelector('[role="dialog"],[role="alertdialog"]'), 'Fluent confirmation surface must open');
  assert.deepEqual(f.calls, []);
  f.click(f.buttons('Cancel')[0]);
  assert.deepEqual(f.calls, []);
  assert.equal(document.activeElement === trigger, true, 'Cancel focus: ' + document.activeElement?.outerHTML?.slice(0, 180));
});
test('PnP consent captures intended ID and approves once even before removal settles', async t => {
  let finish; const pending = new Promise(resolve => { finish = resolve; });
  const f = fixture(t, () => pending); await f.act(async () => {}); f.open();
  assert.deepEqual(f.calls, []);
  const approve = f.buttons('Delete item')[0]; assert.ok(approve, 'explicit approval action');
  f.act(() => { Simulate.click(approve); Simulate.click(approve); });
  assert.deepEqual(f.calls, [42]);
  await f.act(async () => { finish(); await pending; });
  assert.deepEqual(f.calls, [42]);
});
test('PnP deletion retains original error reporting', async t => {
  const failure = new Error('Mock removal failed');
  const f = fixture(t, async () => { throw failure; }); await f.act(async () => {}); f.open();
  assert.ok(f.buttons('Delete item')[0], 'explicit approval action');
  await f.act(async () => { f.click(f.buttons('Delete item')[0]); });
  assert.deepEqual(f.calls, [42]);
  assert.match(f.container.textContent, /Mock removal failed/);
});


test('PnP approval retains stable focus when removal resolves before later refetch row publication', async t => {
  const f = fixture(t); await f.act(async () => {});
  const trigger = f.open();
  await f.act(async () => { f.click(f.buttons('Delete item')[0]); });
  assert.deepEqual(f.calls, [42]);
  assert.equal(trigger.isConnected, true, 'real remove resolves while row is still rendered');
  f.publishRows(42);
  assert.equal(trigger.isConnected, false, 'later debounced query publication removes the row');
  assert.equal(document.activeElement === f.container.querySelector('input[aria-label="List Title"]'), true, 'approval focus remains on a stable list control after deferred row removal');
});

test('PnP late removal and refetch preserve the focus chosen by subsequent user navigation', async t => {
  let finish; const pending = new Promise(resolve => { finish = resolve; });
  const f = fixture(t, () => pending); await f.act(async () => {}); f.open();
  f.click(f.buttons('Delete item')[0]);
  const destination = f.container.querySelector('[aria-label="New item title"]'); destination.focus();
  await f.act(async () => { finish(); await pending; });
  assert.equal(document.activeElement === destination, true, 'removal completion must not steal deliberate focus');
  f.publishRows(42);
  assert.equal(document.activeElement === destination, true, 'later refetch must not steal deliberate focus');
  assert.deepEqual(f.calls, [42]);
});

test('PnP captured consent keeps operation0 for42 across same-target callback replacement then operation1 for73',async t=>{
 const f=fixture(t);await f.act(async()=>{});f.open();f.replaceOperation();
 await f.act(async()=>{f.click(f.buttons('Delete item')[0]);});assert.deepEqual(f.operations,[{id:42,operation:0}]);f.publishRows(42);
 f.open();await f.act(async()=>{f.click(f.buttons('Delete item')[0]);});assert.deepEqual(f.operations,[{id:42,operation:0},{id:73,operation:1}]);
});
test('PnP Escape restores the connected trigger without dispatch',async t=>{
 const f=fixture(t);await f.act(async()=>{});const trigger=f.open();const surface=document.querySelector('[role="dialog"]');
 f.act(()=>Simulate.keyDown(surface,{key:'Escape'}));assert.deepEqual(f.calls,[]);assert.equal(document.activeElement===trigger,true);
});
test('PnP cancellation falls back to stable input when captured trigger has disappeared',async t=>{
 const f=fixture(t);await f.act(async()=>{});const trigger=f.open();f.publishRows(42);assert.equal(trigger.isConnected,false);
 f.click(f.buttons('Cancel')[0]);assert.deepEqual(f.calls,[]);assert.equal(document.activeElement===f.container.querySelector('[aria-label="List Title"]'),true);
});
function negativeFixture(t, transform, remove) {
 const fs=require('node:fs'),os=require('node:os'); const original=fs.readFileSync;
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'pnp-consent-negative-')),copy=path.join(dir,'scenario.tsx');const altered=transform(selected.source);assert.notEqual(altered,selected.source);fs.writeFileSync(copy,altered);
 t.after(()=>fs.rmSync(dir,{recursive:true,force:true}));const mock=t.mock.method(fs,'readFileSync',(file,...args)=>original(String(file)===selected.file?copy:file,...args));
 try{return fixture(t,remove);}finally{mock.mock.restore();}
}
test('PnP negative control detects current callback substitution for captured remove',async t=>{
 const f=negativeFixture(t,source=>source.replace('request.remove(request.id)','list.remove(request.id)'));await f.act(async()=>{});f.open();f.replaceOperation();
 await f.act(async()=>{f.click(f.buttons('Delete item')[0]);});assert.throws(()=>assert.deepEqual(f.operations,[{id:42,operation:0}]),/Expected values/);
});
test('PnP negative control detects focus moved after asynchronous mutation settlement',async t=>{
 let done;const pending=new Promise(resolve=>{done=resolve;});
 const f=negativeFixture(t,source=>source.replace('flushSync(() => setApproval(undefined)); input.current?.focus();','flushSync(() => setApproval(undefined));').replace('const result = await ticket.promise;','const result = await ticket.promise; input.current?.focus();'),()=>pending);
 await f.act(async()=>{});f.open();f.click(f.buttons('Delete item')[0]);const destination=f.container.querySelector('[aria-label="New item title"]');destination.focus();
 await f.act(async()=>{done();await pending;});assert.throws(()=>assert.equal(document.activeElement===destination,true,'later focus steal'),/later focus steal/);
});

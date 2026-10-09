const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, act, api, dom;
test.before(async () => {
  dom = harness.createSxDom();
  for (const key of ['HTMLElement', 'HTMLInputElement', 'File', 'FileList', 'Blob']) global[key] = dom.window[key];
  React = require('react'); ({ act } = require('react-dom/test-utils'));
  api = await harness.loadModule('packages/spfx-react-toolkit/src/controls/form/index.ts', { reactHookForm: true });
});
test.after(async () => { dom?.window.close(); await harness.cleanup(); });
function mount(element) {
  const ReactDOM = require('react-dom'); const { RendererProvider } = require('@griffel/react'); const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
  const renderer = require('@griffel/core').createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  const host = document.createElement('div'); document.body.appendChild(host);
  function render(next = element) { element = next; act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer }, React.createElement(FluentProvider, { theme: webLightTheme }, element)), host); }); }
  render(); return { host, render, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
async function submit(host) { await act(async () => { host.querySelector('form').dispatchEvent(new window.Event('submit', { bubbles: true, cancelable: true })); }); }
function input(host, value) { const node = host.querySelector('input'); act(() => { require('react-dom/test-utils').Simulate.change(node, { target: { value } }); }); }


test('SpinButton commits finite display strings and stepped values before callback; empty/nonfinite becomes null',()=>{const F=api.createForm();let methods;const ref=React.createRef();const seen=[];function Owner(){methods=F.useForm({defaultValues:{value:2}});return React.createElement(F.Root,{methods,onSubmit:()=>{}},React.createElement(F.SpinButton,{name:'value',ref,id:'number-id',label:'Amount',onChange:()=>seen.push(methods.getValues('value')),onBlur:()=>seen.push(methods.getFieldState('value').isTouched)}));}const f=mount(React.createElement(Owner));try{assert.equal(ref.current.tagName,'INPUT');assert.equal(f.host.querySelector('label').htmlFor,'number-id');act(()=>require('react-dom/test-utils').Simulate.keyDown(ref.current,{key:'ArrowUp'}));act(()=>require('react-dom/test-utils').Simulate.keyUp(ref.current,{key:'ArrowUp'}));assert.equal(seen[0],3);for(const [display,want] of [['4.25',4.25],['',null],['Infinity',null],['junk',null],['5tail',null]]){act(()=>require('react-dom/test-utils').Simulate.change(ref.current,{target:{value:display}}));act(()=>require('react-dom/test-utils').Simulate.keyDown(ref.current,{key:'Enter'}));assert.equal(methods.getValues('value'),want);}act(()=>require('react-dom/test-utils').Simulate.blur(ref.current));assert.equal(seen.at(-1),true);assert.equal(methods.getFieldState('value').isDirty,true);}finally{f.close();}});
test('Slider preserves caller default/bounds and binds native callbacks',()=>{const F=api.createForm();let methods;const seen=[];const ref=React.createRef();function Owner(){methods=F.useForm({defaultValues:{value:30}});return React.createElement(F.Root,{methods,onSubmit:()=>{}},React.createElement(F.Slider,{name:'value',ref,min:20,max:40,step:2,label:'Level',onChange:()=>seen.push(methods.getValues('value')),onBlur:()=>seen.push(methods.getFieldState('value').isTouched)}));}const f=mount(React.createElement(Owner));try{assert.equal(ref.current.tagName,'INPUT');assert.equal(ref.current.value,'30');assert.equal(ref.current.min,'20');assert.equal(ref.current.max,'40');act(()=>require('react-dom/test-utils').Simulate.change(ref.current,{target:{value:'34'}}));assert.deepEqual(seen,[34]);act(()=>require('react-dom/test-utils').Simulate.blur(ref.current));assert.equal(seen[1],true);}finally{f.close();}});
for(const member of ['SpinButton','Slider'])test(`${member} disabled input and buttons cannot be bypassed`,()=>{const F=api.createForm();const f=mount(React.createElement(F.Root,{disabled:true,defaultValues:{value:25},onSubmit:()=>{}},React.createElement(F[member],{name:'value',disabled:false,input:{disabled:false},...(member==='SpinButton'?{incrementButton:{disabled:false},decrementButton:{disabled:false}}:{min:20,max:40})})));try{assert.ok([...f.host.querySelectorAll('input,button')].every(n=>n.disabled));}finally{f.close();}});

test('SpinButton required metadata does not invent validation; explicit rules own nullable validity', async () => {
 const F = api.createForm(); let methods;
 function Owner({ rules }) {
  methods = F.useForm({ defaultValues: { value: null } });
  return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.SpinButton, { name: 'value', required: true, rules }));
 }
 const f = mount(React.createElement(Owner));
 try {
  let valid; await act(async () => { valid = await methods.trigger('value'); }); assert.equal(valid, true);
  assert.equal(methods.getValues('value'), null);
  f.render(React.createElement(Owner, { rules: { validate: value => value !== null || 'Enter a number' } }));
  await act(async () => { valid = await methods.trigger('value'); }); assert.equal(valid, false);
  assert.equal(methods.getFieldState('value').error.message, 'Enter a number');
  assert.match(f.host.textContent, /Enter a number/);
 } finally { f.close(); }
});

test('SpinButton slot false cannot bypass Fluent rounded range-bound defaults', () => {
 const F = api.createForm();
 const f = mount(React.createElement(F.Root, { defaultValues: { value: 4.999 }, onSubmit: () => {} }, React.createElement(F.SpinButton, { name: 'value', min: 0, max: 5, precision: 0, incrementButton: { disabled: false }, decrementButton: { disabled: false } })));
 try {
  assert.equal(f.host.querySelector('[aria-label="Increment value"]').disabled, true);
  assert.equal(f.host.querySelector('[aria-label="Decrement value"]').disabled, false);
 } finally { f.close(); }
});

for (const reset of ['reset', 'resetField']) test(`SpinButton ${reset} clears same-valued uncommitted text and keeps its ref/focus truthful`, () => {
 const F = api.createForm(); let methods; const ref = React.createRef();
 function Owner() {
  methods = F.useForm({ defaultValues: { value: 2 } });
  return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.SpinButton, { name: 'value', ref }));
 }
 const f = mount(React.createElement(Owner));
 try {
  const original = ref.current;
  act(() => { original.focus(); require('react-dom/test-utils').Simulate.change(original, { target: { value: 'Uncommitted' } }); });
  assert.equal(original.value, 'Uncommitted'); assert.equal(methods.getValues('value'), 2);
  act(() => { if (reset === 'reset') methods.reset(); else methods.resetField('value'); });
  assert.equal(ref.current.value, '2'); assert.equal(ref.current, f.host.querySelector('input'));
  assert.equal(document.activeElement, ref.current); assert.equal(methods.getValues('value'), 2);
  act(() => require('react-dom/test-utils').Simulate.blur(ref.current));
  assert.equal(methods.getValues('value'), 2);
 } finally { f.close(); }
});

test('SpinButton targeted reset leaves unrelated drafts intact; keepDirtyValues preserves only the RHF dirty owner', () => {
 const F = api.createForm(); let methods; const left = React.createRef(); const right = React.createRef();
 function Owner() {
  methods = F.useForm({ defaultValues: { left: 2, right: 4 } });
  const dirty = methods.formState.dirtyFields;
  return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.SpinButton, { name: 'left', ref: left }), React.createElement(F.SpinButton, { name: 'right', ref: right }), React.createElement('output', {}, Object.keys(dirty).join(',')));
 }
 const f = mount(React.createElement(Owner));
 const edit = (node, value) => act(() => require('react-dom/test-utils').Simulate.change(node, { target: { value } }));
 try {
  edit(left.current, 'Left draft'); edit(right.current, 'Right draft'); const untouched = right.current;
  act(() => methods.resetField('left'));
  assert.equal(left.current.value, '2'); assert.equal(right.current.value, 'Right draft'); assert.equal(right.current, untouched);
  edit(left.current, '3'); act(() => require('react-dom/test-utils').Simulate.keyDown(left.current, { key: 'Enter' }));
  assert.equal(methods.getFieldState('left').isDirty, true); assert.equal(methods.getFieldState('right').isDirty, false);
  edit(left.current, 'Dirty owner draft'); const preserved = left.current;
  act(() => methods.reset({ left: 10, right: 20 }, { keepDirtyValues: true }));
  assert.deepEqual(methods.getValues(), { left: 3, right: 20 });
  assert.equal(left.current.value, 'Dirty owner draft'); assert.equal(left.current, preserved); assert.equal(right.current.value, '20');
 } finally { f.close(); }
});

test('Slider native-slot change cannot bypass RHF: top then slot see live dirty number once', () => {
 const F = api.createForm(); let methods; const events = [];
 const observe = location => event => { events.push([location, methods.getValues('value'), methods.getFieldState('value').isDirty, event.currentTarget.tagName]); };
 function Owner() {
  methods = F.useForm({ defaultValues: { value: 30 } });
  return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.Slider, { name: 'value', min: 20, max: 40, onChange: observe('top'), input: { onChange: observe('slot') } }));
 }
 const f = mount(React.createElement(Owner));
 const writes = [];
 const unsubscribe = methods.subscribe({ name: 'value', exact: true, formState: { values: true }, callback: state => writes.push(state.values.value) });
 try {
  act(() => require('react-dom/test-utils').Simulate.change(f.host.querySelector('input'), { target: { value: '34' } }));
  assert.deepEqual(writes, [34]);
  assert.deepEqual(events, [['top', 34, true, 'INPUT'], ['slot', 34, true, 'INPUT']]);
  assert.equal(methods.getValues('value'), 34); assert.equal(f.host.querySelector('input').value, '34');
 } finally { unsubscribe(); f.close(); }
});

test('SpinButton native-slot change observes presentation editing; numeric commit has its own callback', () => {
 const F = api.createForm(); let methods; const events = [];
 function Owner() {
  methods = F.useForm({ defaultValues: { value: 2 } });
  return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.SpinButton, { name: 'value', onChange: () => events.push(['commit', methods.getValues('value')]), input: { onChange: event => events.push(['draft', methods.getValues('value'), event.currentTarget.tagName]) } }));
 }
 const f = mount(React.createElement(Owner));
 try {
  act(() => require('react-dom/test-utils').Simulate.change(f.host.querySelector('input'), { target: { value: '7.25' } }));
  assert.deepEqual(events, [['draft', 2, 'INPUT']]); assert.equal(methods.getFieldState('value').isDirty, false);
  act(() => require('react-dom/test-utils').Simulate.keyDown(f.host.querySelector('input'), { key: 'Enter' }));
  assert.deepEqual(events, [['draft', 2, 'INPUT'], ['commit', 7.25]]); assert.equal(methods.getFieldState('value').isDirty, true);
 } finally { f.close(); }
});

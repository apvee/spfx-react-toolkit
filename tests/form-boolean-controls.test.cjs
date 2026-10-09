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


for(const member of ['Checkbox','Switch'])test(`${member} binds boolean before native callbacks and protects disabled`,()=>{const F=api.createForm();let methods;const ref=React.createRef();const seen=[];function Owner(){methods=F.useForm({defaultValues:{value:false}});return React.createElement(F.Root,{methods,onSubmit:()=>{}},React.createElement(F[member],{name:'value',ref,id:'boolean-id',label:'Choice',onChange:()=>seen.push(methods.getValues('value')),onBlur:()=>seen.push(methods.getFieldState('value').isTouched)}));}const f=mount(React.createElement(Owner));try{assert.equal(ref.current.tagName,'INPUT');assert.equal(f.host.querySelector('label').htmlFor,'boolean-id');act(()=>{ref.current.checked=true;ref.current.indeterminate=false;require('react-dom/test-utils').Simulate.change(ref.current);});assert.deepEqual(seen,[true]);assert.equal(methods.getFieldState('value').isDirty,true);act(()=>require('react-dom/test-utils').Simulate.blur(ref.current));assert.equal(seen[1],true);f.render(React.createElement(F.Root,{disabled:true,defaultValues:{value:false},onSubmit:()=>{}},React.createElement(F[member],{name:'value',disabled:false,input:{disabled:false}})));assert.equal(f.host.querySelector('input').disabled,true);}finally{f.close();}});
test('Checkbox mixed is retained only with triState',()=>{for(const triState of [false,true]){const F=api.createForm();let methods;function Owner(){methods=F.useForm({defaultValues:{value:triState?'mixed':false}});return React.createElement(F.Root,{methods,onSubmit:()=>{}},React.createElement(F.Checkbox,{name:'value',triState}));}const f=mount(React.createElement(Owner));try{const node=f.host.querySelector('input');assert.equal(node.indeterminate,triState);act(()=>{node.checked=false;node.indeterminate=true;require('react-dom/test-utils').Simulate.change(node);});assert.equal(methods.getValues('value'),triState?'mixed':false);}finally{f.close();}}});

for (const member of ['Checkbox', 'Switch']) test(`${member} updates RHF before top and native-slot callbacks exactly once`, () => {
 const F = api.createForm(); let methods; const events = [];
 const observe = location => event => { events.push([location, methods.getValues('value'), methods.getFieldState('value').isDirty, event.currentTarget.tagName]); };
 function Owner() {
  methods = F.useForm({ defaultValues: { value: false } });
  return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F[member], { name: 'value', onChange: observe('top'), input: { onChange: observe('slot') } }));
 }
 const f = mount(React.createElement(Owner));
 const writes = [];
 const unsubscribe = methods.subscribe({ name: 'value', exact: true, formState: { values: true }, callback: state => writes.push(state.values.value) });
 try {
  const node = f.host.querySelector('input');
  act(() => { node.checked = true; require('react-dom/test-utils').Simulate.change(node); });
  assert.deepEqual(writes, [true]);
  assert.deepEqual(events, [['top', true, true, 'INPUT'], ['slot', true, true, 'INPUT']]);
  assert.equal(methods.getValues('value'), true); assert.equal(node.checked, true);
 } finally { unsubscribe(); f.close(); }
});

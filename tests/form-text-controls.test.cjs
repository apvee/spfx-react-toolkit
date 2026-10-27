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


for (const member of ['Textarea', 'SearchBox']) test(`${member} writes strings before callbacks, preserves ref and Field associations`, async () => {
 const F=api.createForm(); const ref=React.createRef(); let methods; const seen=[];
 function Owner(){ methods=F.useForm({defaultValues:{value:'Initial'}}); return React.createElement(F.Root,{methods,onSubmit:()=>{}},React.createElement(F[member],{name:'value',id:'consumer-id',ref,label:'Value',hint:'Help',rules:{required:'Required'},onChange:()=>seen.push(methods.getValues('value')),onBlur:()=>seen.push(methods.getFieldState('value').isTouched)})); }
 const f=mount(React.createElement(Owner)); try { const node=ref.current; assert.equal(node.tagName,member==='Textarea'?'TEXTAREA':'INPUT'); assert.equal(node.id,'consumer-id'); assert.equal(f.host.querySelector('label').htmlFor,node.id); assert.ok(node.getAttribute('aria-describedby'));
 act(()=>require('react-dom/test-utils').Simulate.change(node,{target:{value:'Edited'}})); assert.deepEqual(seen,['Edited']); assert.equal(methods.getFieldState('value').isDirty,true);
 act(()=>require('react-dom/test-utils').Simulate.blur(node)); assert.equal(seen[1],true);
 act(()=>require('react-dom/test-utils').Simulate.change(node,{target:{value:''}})); await act(async()=>{await methods.trigger('value');}); assert.ok(document.getElementById(node.getAttribute('aria-describedby').split(' ').find(id=>document.getElementById(id)?.textContent==='Required')));
 } finally {f.close();}
});
test('SearchBox dismiss writes an empty string before consumer callback',()=>{ const F=api.createForm();let methods;const seen=[]; function Owner(){methods=F.useForm({defaultValues:{value:'Search'}});return React.createElement(F.Root,{methods,onSubmit:()=>{}},React.createElement(F.SearchBox,{name:'value',onChange:()=>seen.push(methods.getValues('value'))}));}const f=mount(React.createElement(Owner));try{const button=f.host.querySelector('[role=button]');assert.ok(button);act(()=>require('react-dom/test-utils').Simulate.click(button));assert.deepEqual(seen,['']);assert.equal(f.host.querySelector('input').value,'');}finally{f.close();}});
for(const member of ['Textarea','SearchBox'])test(`${member} nested input cannot bypass root disabled`,()=>{const F=api.createForm();const f=mount(React.createElement(F.Root,{disabled:true,defaultValues:{value:''},onSubmit:()=>{}},React.createElement(F[member],{name:'value',disabled:false,...(member==='SearchBox'?{input:{disabled:false}}:{textarea:{disabled:false}})})));try{assert.equal(f.host.querySelector('input,textarea').disabled,true);}finally{f.close();}});

test('adapter types preserve two-way domains, actual native events and refs', () => {
 const ts=require('typescript');const path=require('node:path');const program=ts.createProgram([path.resolve(__dirname, '../packages/spfx-react-toolkit/src/localization/native-resources.internal.d.ts'), path.resolve(__dirname,'fixtures/form-control-types.tsx')],{noEmit:true,strict:true,skipLibCheck:true,target:ts.ScriptTarget.ES2020,module:ts.ModuleKind.ESNext,moduleResolution:ts.ModuleResolutionKind.NodeJs,jsx:ts.JsxEmit.React,lib:['lib.es2020.d.ts','lib.dom.d.ts']});const diagnostics=ts.getPreEmitDiagnostics(program);assert.equal(diagnostics.length,0,ts.formatDiagnosticsWithColorAndContext(diagnostics,{getCanonicalFileName:f=>f,getCurrentDirectory:()=>process.cwd(),getNewLine:()=> '\n'}));
});

for (const member of ['Textarea', 'SearchBox']) test(`${member} semantic change updates RHF before top and native-slot consumers exactly once`, () => {
 const F = api.createForm(); let methods; const events = [];
 const observe = location => event => { events.push([location, methods.getValues('value'), methods.getFieldState('value').isDirty, event.currentTarget.tagName]); };
 function Owner() {
  methods = F.useForm({ defaultValues: { value: 'Initial' } });
  return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F[member], { name: 'value', onChange: observe('top'), [member === 'Textarea' ? 'textarea' : 'input']: { onChange: observe('slot') } }));
 }
 const f = mount(React.createElement(Owner));
 const writes = [];
 const unsubscribe = methods.subscribe({ name: 'value', exact: true, formState: { values: true }, callback: state => writes.push(state.values.value) });
 try {
  act(() => require('react-dom/test-utils').Simulate.change(f.host.querySelector('textarea,input'), { target: { value: 'Changed' } }));
  const tag = member === 'Textarea' ? 'TEXTAREA' : 'INPUT';
  assert.deepEqual(writes, ['Changed']);
  assert.deepEqual(events, [['top', 'Changed', true, tag], ['slot', 'Changed', true, tag]]);
  assert.equal(methods.getValues('value'), 'Changed');
  assert.equal(f.host.querySelector('textarea,input').value, 'Changed');
  if (member === 'SearchBox') {
   act(() => require('react-dom/test-utils').Simulate.click(f.host.querySelector('[role=button]')));
   assert.deepEqual(events.slice(2), [['top', '', true, 'SPAN']]);
   assert.equal(methods.getValues('value'), '');
  }
 } finally { unsubscribe(); f.close(); }
});

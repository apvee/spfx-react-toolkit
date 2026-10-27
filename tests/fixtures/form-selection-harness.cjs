const assert = require('node:assert/strict');
const harness = require('./formatted-text-harness.cjs');
let React, act, api, dom;
async function setup() {
 dom = harness.createSxDom();
 for (const key of ['HTMLElement','HTMLInputElement','HTMLSelectElement','NodeFilter','Event','CustomEvent','MouseEvent','KeyboardEvent','Element','Node','File','FileList','Blob']) global[key] = dom.window[key];
 // JSDOM supplies no layout observer; popup positioning is not under test.
 dom.window.ResizeObserver = class ResizeObserver { observe() {} unobserve() {} disconnect() {} };
 // Fluent's option walker reads ARIAMixin.role. JSDOM 15 lacks that native
 // reflection, so retain the DOM contract here without replacing navigation.
 // Nullable absence belongs to the external reflected DOM attribute contract.
 if (!('role' in dom.window.Element.prototype)) {
  Object.defineProperty(dom.window.Element.prototype, 'role', {
   configurable: true,
   get() { return this.getAttribute('role'); },
   set(value) { if (value === null) this.removeAttribute('role'); else this.setAttribute('role', value); }
  });
 }
 React = require('react'); ({act} = require('react-dom/test-utils'));
 api = await harness.loadModule('tests/fixtures/form-selection-entry.tsx',{reactHookForm:true});
}
async function cleanup(){dom?.window.close();await harness.cleanup();}
function suite(member){const F=api.createForm();assert.ok(F[member],`Form.${member} is implemented in the actual public suite`);return F;}
function mount(element) {
 const ReactDOM=require('react-dom');
 const {RendererProvider}=require('@griffel/react');
 const {FluentProvider,webLightTheme}=require('@fluentui/react-components');
 const renderer=require('@griffel/core').createDOMRenderer(document,{unstable_filterCSSRule:rule=>!rule.startsWith('@container')});
 const host=document.createElement('div');document.body.appendChild(host);
 // Each fixture owns its renderer. Removing the host alone leaves its sheets in
 // JSDOM, making subsequent geometry/style measurements grow with every test.
 function release(){
  host.remove();
  for(const sheet of Object.values(renderer.stylesheets))sheet.element?.remove();
 }
 function render(next=element){
  element=next;
  act(()=>{ReactDOM.render(React.createElement(RendererProvider,{renderer},React.createElement(FluentProvider,{theme:webLightTheme},element)),host);});
 }
 try{render();}catch(error){release();throw error;}
 return {host,render,close(){try{act(()=>{ReactDOM.unmountComponentAtNode(host);});}finally{release();}}};
}
function change(node,value){node.value=value;act(()=>require('react-dom/test-utils').Simulate.change(node,{target:{value}}));}
function click(node){assert.ok(node,'interactive target exists');act(()=>require('react-dom/test-utils').Simulate.click(node));}
function key(node,key){act(()=>require('react-dom/test-utils').Simulate.keyDown(node,{key,code:key}));}
function owner(F,member,defaults,props={},extra){let methods;function Owner(){methods=F.useForm({defaultValues:defaults});return React.createElement(F.Root,{methods,onSubmit:()=>{}},React.createElement(F[member],{name:'value',...props}),extra);}const f=mount(React.createElement(Owner));return {...f,get methods(){return methods;}};}
module.exports={setup,cleanup,suite,mount,change,click,key,owner,get React(){return React;},get act(){return act;},get api(){return api;}};

function rejectsRender(callback,pattern){const original=console.error;const logged=[];console.error=(...args)=>logged.push(args);try{assert.throws(callback,pattern);}finally{console.error=original;}assert.ok(logged.length>0,'React reports the rejected render');}
module.exports.rejectsRender=rejectsRender;

/** Idle labels and an empty query are distinct presentation states of ComboboxMulti. */
module.exports.assertMultiQueryCleared = (node, labels) => {
  // A surviving edited node can still be in query mode; remounted idle nodes show labels.
  assert.ok(node.value === labels || node.value === '', 'cleared query or idle selection must replace stale text');
  act(() => require('react-dom/test-utils').Simulate.focus(node));
  assert.equal(node.value, '', 'focusing exposes the cleared query');
  act(() => require('react-dom/test-utils').Simulate.blur(node));
  assert.equal(node.value, labels, 'blur restores labels without restoring stale query');
};

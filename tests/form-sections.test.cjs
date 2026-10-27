const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const ts = require('typescript');
const h = require('./fixtures/form-selection-harness.cjs');
test.before(h.setup); test.after(h.cleanup);
function owner(member, defaults, props, rootProps={}) {
  const F=h.suite(member); let methods;
  function Owner(){methods=F.useForm({defaultValues:defaults});return h.React.createElement(F.Root,{methods,onSubmit:()=>{},...rootProps},h.React.createElement(F[member],props));}
  return {...h.mount(h.React.createElement(Owner)),get methods(){return methods;}};
}
function array(defaults=[],props={},rootProps={}) {return owner('ArraySection',{rows:defaults},{name:'rows',title:'Rows',createItem:()=>({text:''}),renderItem:({Fields})=>h.React.createElement(Fields.Input,{name:'text','aria-label':'Text'}),...props},rootProps);}
const button=(host,label)=>Array.from(host.querySelectorAll('button')).find(node=>node.getAttribute('aria-label')===label||node.textContent===label);
const items=host=>Array.from(host.querySelectorAll('ol > li'));
// Missing semantic wrappers/associations would leave consumers without a named group or described errors.
test('ObjectSection names a native fieldset and associates description and direct object errors',()=>{
 const f=owner('ObjectSection',{customer:{text:'A'}},{name:'customer',title:'Customer',description:'Contact details',children:({Fields})=>h.React.createElement(Fields.Input,{name:'text',label:'Text'})});
 try {const group=f.host.querySelector('fieldset');assert.ok(group);assert.equal(group.querySelector('legend').textContent,'Customer');assert.equal(group.querySelector('input').name,'customer.text');h.act(()=>f.methods.setError('customer',{type:'manual',message:'Invalid customer'}));const ids=group.getAttribute('aria-describedby').split(' ');assert.deepEqual(ids.map(id=>document.getElementById(id).textContent),['Contact details','Invalid customer']);assert.equal(group.getAttribute('aria-invalid'),'true');}finally{f.close();}
});
test('Sections reject absent accessible names and permit external labels',()=>{
 const F=h.suite('ObjectSection');h.rejectsRender(()=>h.mount(h.React.createElement(F.Root,{defaultValues:{customer:{}},onSubmit:()=>{}},h.React.createElement(F.ObjectSection,{name:'customer',children:()=>undefined}))),/accessible name/);
 const f=array([],{title:undefined,'aria-label':'Localized rows'});try{assert.equal(f.host.querySelector('fieldset').getAttribute('aria-label'),'Localized rows');}finally{f.close();}
});
// Native disabling alone would retain submitted RHF data; scoped adapters must receive disabled state.
test('Disabled object/array sections disable relative adapters and omit their values on submit',async()=>{
 for(const member of ['ObjectSection','ArraySection']) {let submitted;const f=member==='ObjectSection'?owner(member,{customer:{text:'A'}},{name:'customer',title:'Customer',disabled:true,children:({Fields})=>h.React.createElement(Fields.Input,{name:'text',disabled:false})},{onSubmit:value=>{submitted=value;}}):array([{text:'A'}],{disabled:true},{onSubmit:value=>{submitted=value;}});
 try{assert.equal(f.host.querySelector('input').disabled,true);await h.act(async()=>{await f.methods.handleSubmit(value=>{submitted=value;})();});assert.equal(member==='ObjectSection'?submitted.customer?.text:submitted.rows?.[0]?.text,undefined);if(member==='ArraySection'){assert.ok(Array.from(f.host.querySelectorAll('button')).every(node=>node.disabled));}}finally{f.close();}}
});
test('ArraySection renders ordered items with localized names and direct array/root errors',()=>{
 const f=array([{text:'A'},{text:'B'}],{description:'Ordered entries',itemLabel:({index})=>`Riga ${index+1}`,strings:{add:'Aggiungi',remove:label=>`Elimina ${label}`,moveUp:label=>`Su ${label}`,moveDown:label=>`Giù ${label}`},reorderable:true});
 try{assert.deepEqual(items(f.host).map(node=>node.getAttribute('aria-label')),['Riga 1','Riga 2']);assert.ok(button(f.host,'Elimina Riga 1'));assert.ok(button(f.host,'Aggiungi'));assert.equal(button(f.host,'Su Riga 1').disabled,true);assert.equal(button(f.host,'Giù Riga 2').disabled,true);h.act(()=>f.methods.setError('rows.root',{type:'manual',message:'Array error'}));const group=f.host.querySelector('fieldset');assert.ok(group.getAttribute('aria-describedby').split(' ').some(id=>document.getElementById(id).textContent==='Array error'));}finally{f.close();}
});
// Missing bounds guards could still mutate via direct click invocation on a disabled action.
test('Array actions enforce min/max, removable, reorderable and inherited disabled constraints',()=>{
 const f=array([{text:'A'}],{minItems:1,maxItems:2,reorderable:true});try{h.click(button(f.host,'Remove Item 1'));assert.equal(f.methods.getValues('rows').length,1);h.click(button(f.host,'Add item'));assert.equal(f.methods.getValues('rows').length,2);h.click(button(f.host,'Add item'));assert.equal(f.methods.getValues('rows').length,2);assert.equal(button(f.host,'Add item').disabled,true);h.click(button(f.host,'Move Item 2 up'));assert.deepEqual(f.methods.getValues('rows'),[{text:''},{text:'A'}]);}finally{f.close();}
 const g=array([{text:'A'}],{removable:false});try{assert.equal(g.host.querySelectorAll('button').length,1);}finally{g.close();}
 const F=h.suite('ArraySection');const disabled=h.mount(h.React.createElement(F.Root,{disabled:true,defaultValues:{rows:[]},onSubmit:()=>{}},h.React.createElement(F.ArraySection,{name:'rows',title:'Rows',createItem:()=>({text:''}),renderItem:()=>undefined})));try{assert.equal(button(disabled.host,'Add item').disabled,true);}finally{disabled.close();}
});
test('Add focuses the first eligible control, skipping hidden inert disabled aria-disabled and rejecting targets',()=>{
 const f=array([],{renderItem:({Fields})=>h.React.createElement(h.React.Fragment,{},h.React.createElement('div',{hidden:true},h.React.createElement('input',{'aria-label':'hidden'})),h.React.createElement('div',{inert:''},h.React.createElement('input',{'aria-label':'inert'})),h.React.createElement('input',{disabled:true}),h.React.createElement('input',{'aria-disabled':'true'}),h.React.createElement('input',{ref:node=>{if(node)node.focus=()=>{};},'aria-label':'reject'}),h.React.createElement(Fields.Input,{name:'text','aria-label':'eligible'}))});
 try{h.click(button(f.host,'Add item'));assert.equal(document.activeElement.getAttribute('aria-label'),'eligible');}finally{f.close();}
});
test('Remove focuses next then previous then Add as arrays become empty',()=>{
 const f=array([{text:'A'},{text:'B'},{text:'C'}]);try{const before=items(f.host).map(node=>node.querySelector('input'));h.click(button(f.host,'Remove Item 2'));assert.equal(document.activeElement,before[2]);h.click(button(f.host,'Remove Item 2'));assert.equal(document.activeElement,before[0]);h.click(button(f.host,'Remove Item 1'));assert.equal(document.activeElement,button(f.host,'Add item'));assert.ok(f.host.textContent.includes('No items'));}finally{f.close();}
});
test('Move preserves item identity and focuses the same enabled action with boundary fallback',()=>{
 const f=array([{text:'A'},{text:'B'},{text:'C'}],{reorderable:true});try{const node=items(f.host)[1],input=node.querySelector('input');h.click(button(f.host,'Move Item 2 up'));assert.equal(items(f.host)[0],node);assert.equal(node.querySelector('input'),input);assert.equal(document.activeElement,button(f.host,'Move Item 1 down'));h.click(button(f.host,'Move Item 1 down'));assert.equal(items(f.host)[1],node);assert.equal(document.activeElement,button(f.host,'Move Item 2 down'));}finally{f.close();}
});
test('Repeated localized announcements produce fresh live updates and resets do not steal focus',()=>{
 const f=array([],{strings:{empty:'Vuoto',added:()=> 'Aggiornato',removed:()=> 'Aggiornato'}});try{assert.ok(f.host.textContent.includes('Vuoto'));h.click(button(f.host,'Add item'));const live=f.host.querySelector('[aria-live="polite"]');const first=live.firstElementChild;assert.equal(live.textContent,'Aggiornato');h.click(button(f.host,'Remove Item 1'));assert.equal(live.textContent,'Aggiornato');assert.notEqual(live.firstElementChild,first);const add=button(f.host,'Add item');add.focus();h.act(()=>f.methods.reset({rows:[{text:'reset'}]}));assert.equal(document.activeElement,add);}finally{f.close();}
});
test('Section declarations retain relative model inference and reject invalid configuration',()=>{
 const program=ts.createProgram([path.resolve(__dirname, '../packages/spfx-react-toolkit/src/localization/native-resources.internal.d.ts'), path.resolve(__dirname,'fixtures/form-sections-types.tsx')],{noEmit:true,strict:true,skipLibCheck:true,target:ts.ScriptTarget.ES2020,module:ts.ModuleKind.ESNext,moduleResolution:ts.ModuleResolutionKind.NodeJs,jsx:ts.JsxEmit.React,lib:['lib.es2020.d.ts','lib.dom.d.ts']});const diagnostics=ts.getPreEmitDiagnostics(program);assert.equal(diagnostics.length,0,ts.formatDiagnosticsWithColorAndContext(diagnostics,{getCanonicalFileName:file=>file,getCurrentDirectory:()=>process.cwd(),getNewLine:()=> '\n'}));
});
test('Whitespace-only legends are rejected unless an explicit name is supplied',()=>{
 const F=h.suite('ObjectSection');h.rejectsRender(()=>h.mount(h.React.createElement(F.Root,{defaultValues:{customer:{}},onSubmit:()=>{}},h.React.createElement(F.ObjectSection,{name:'customer',title:'   ',children:()=>undefined}))),/accessible name/);
});
test('Remove skips unavailable next rows and falls back to an eligible previous row',()=>{
 const f=array([{text:'A'},{text:'B'},{text:'C'}],{renderItem:({index,Fields})=>h.React.createElement('div',{'aria-disabled':index===2?'true':undefined},h.React.createElement(Fields.Input,{name:'text'}))});
 try{const previous=items(f.host)[0].querySelector('input');const unavailable=items(f.host)[2];unavailable.setAttribute('inert','');h.click(button(f.host,'Remove Item 2'));assert.equal(document.activeElement,previous);}finally{f.close();}
});
test('Add ignores CSS-hidden controls and aria-disabled ancestors and does not focus after record replacement',()=>{
 const f=array([],{renderItem:({Fields})=>h.React.createElement(h.React.Fragment,{},h.React.createElement('input',{className:'section-test-hidden'}),h.React.createElement('div',{'aria-disabled':'true'},h.React.createElement('input')),h.React.createElement(Fields.Input,{name:'text','aria-label':'visible'}))});
 const style=document.createElement('style');style.textContent='.section-test-hidden {visibility:hidden;}';document.head.appendChild(style);
 try{h.click(button(f.host,'Add item'));assert.equal(document.activeElement.getAttribute('aria-label'),'visible');const add=button(f.host,'Add item');add.focus();h.act(()=>f.methods.reset({rows:[{text:'new'},{text:'newer'}]}));assert.equal(document.activeElement,add);}finally{f.close();style.remove();}
});
test('Nested visual sections retain absolute paths and additive disabled state',async()=>{
 const F=h.suite('ObjectSection');let methods;
 function Owner(){
   methods=F.useForm({defaultValues:{customer:{rows:[{detail:{text:'A'}}]}}});
   return h.React.createElement(F.Root,{methods,onSubmit:()=>{}},
     h.React.createElement(F.ObjectSection,{name:'customer',title:'Customer',disabled:true},({Fields})=>
       h.React.createElement(Fields.ArraySection,{
         name:'rows',title:'Rows',disabled:false,createItem:()=>({detail:{text:''}}),
         renderItem:({Fields})=>h.React.createElement(Fields.ObjectSection,{name:'detail',title:'Detail',disabled:false},({Fields})=>h.React.createElement(Fields.Input,{name:'text',disabled:false}))
       })));
 }
 const f=h.mount(h.React.createElement(Owner));try{const input=f.host.querySelector('input');assert.equal(input.name,'customer.rows.0.detail.text');assert.equal(input.disabled,true);let submitted;await h.act(async()=>{await methods.handleSubmit(value=>{submitted=value;})();});assert.equal(submitted.customer?.rows?.[0]?.detail?.text,undefined);assert.equal(methods.getValues('customer.rows.0.detail.text'),'A');}finally{f.close();}
});
test('Array bounds reject invalid values and do not manufacture rows or validate externally reset records',()=>{
 const F=h.suite('ArraySection');for(const bounds of [{minItems:-1},{minItems:0.5},{maxItems:NaN},{minItems:2,maxItems:1}])h.rejectsRender(()=>h.mount(h.React.createElement(F.Root,{defaultValues:{rows:[]},onSubmit:()=>{}},h.React.createElement(F.ArraySection,{name:'rows',title:'Rows',createItem:()=>({text:''}),renderItem:()=>undefined,...bounds}))),/bounds/);
 const f=array([],{minItems:2,maxItems:3});try{assert.deepEqual(f.methods.getValues('rows'),[]);h.act(()=>f.methods.reset({rows:[{text:'A'},{text:'B'},{text:'C'},{text:'D'}]}));assert.equal(items(f.host).length,4);assert.equal(button(f.host,'Add item').disabled,true);}finally{f.close();}
});
// A replacement reset after the internal button handler must supersede the pending visual action.
for (const action of ['Add item','Remove Item 2','Move Item 2 up']) {
 test(`${action} followed by same-handler reset never focuses or announces the replacement record`,()=>{
  const f=array(action==='Add item'?[]:[{text:'A'},{text:'B'},{text:'C'}],{reorderable:true,onClick:event=>{if(event.target.closest('button'))f.methods.reset({rows:[{text:'replacement'}]});}});
  try{const add=button(f.host,'Add item');add.focus();h.click(button(f.host,action));assert.deepEqual(f.methods.getValues('rows'),[{text:'replacement'}]);assert.equal(document.activeElement===add,true,'the surviving Add button retains focus after reset');assert.equal(f.host.querySelector('[aria-live]').textContent,'');}finally{f.close();}
 });
}
test('An exact array reset supersedes pending Add even if a later unrelated reset is the latest event',()=>{
 const f=owner('ArraySection',{rows:[],other:'original'},{name:'rows',title:'Rows',createItem:()=>({text:'added'}),renderItem:({Fields})=>h.React.createElement(Fields.Input,{name:'text'}),onClick:event=>{if(event.target.closest('button')){f.methods.resetField('rows',{defaultValue:[{text:'replacement'}]});f.methods.resetField('other');}}});
 try{f.methods.register('rows');f.methods.register('other');const add=button(f.host,'Add item');add.focus();h.click(add);assert.deepEqual(f.methods.getValues('rows'),[{text:'replacement'}]);assert.equal(document.activeElement===add,true);assert.equal(f.host.querySelector('[aria-live]').textContent,'');}finally{f.close();}
});
test('An unrelated targeted reset preserves the valid pending Add focus and announcement',()=>{
 const f=owner('ArraySection',{rows:[],other:'original'},{name:'rows',title:'Rows',createItem:()=>({text:'added'}),renderItem:({Fields})=>h.React.createElement(Fields.Input,{name:'text','aria-label':'added control'}),onClick:event=>{if(event.target.closest('button'))f.methods.resetField('other');}});
 try{f.methods.register('other');const add=button(f.host,'Add item');add.focus();h.click(add);assert.deepEqual(f.methods.getValues('rows'),[{text:'added'}]);assert.equal(document.activeElement.getAttribute('aria-label'),'added control');assert.equal(f.host.querySelector('[aria-live]').textContent,'Item 1 added');}finally{f.close();}
});
test('An ancestor reset cancels pending actions in a nested scoped section',()=>{
 const F=h.suite('ObjectSection');let methods;
 function Owner(){methods=F.useForm({defaultValues:{customer:{rows:[]}}});return h.React.createElement(F.Root,{methods,onSubmit:()=>{}},h.React.createElement(F.ObjectSection,{name:'customer',title:'Customer'},({Fields})=>h.React.createElement(Fields.ArraySection,{name:'rows',title:'Rows',createItem:()=>({text:'added'}),renderItem:({Fields})=>h.React.createElement(Fields.Input,{name:'text'}),onClick:event=>{if(event.target.closest('button'))methods.resetField('customer',{defaultValue:{rows:[{text:'replacement'}]}});}})));}
 const f=h.mount(h.React.createElement(Owner));try{methods.register('customer');const add=button(f.host,'Add item');add.focus();h.click(add);assert.deepEqual(methods.getValues('customer.rows'),[{text:'replacement'}]);assert.equal(document.activeElement===add,true);assert.equal(f.host.querySelector('[aria-live]').textContent,'');}finally{f.close();}
});

for (const member of ['ObjectSection', 'ArraySection']) test(`${member} presentation objects callbacks and empty descriptions keep fieldset associations`, () => {
 const titleRef = h.React.createRef(), descriptionRef = h.React.createRef(); let calls = 0;
 const props = { title: { id: 'inner-title', ref: titleRef, children: (Component, attrs) => { calls++; return h.React.createElement(Component, attrs, 'Slot section'); } }, description: { id: 'inner-description', ref: descriptionRef, children: 'Slot description' } };
 const f = member === 'ObjectSection' ? owner(member, { customer: {} }, { name: 'customer', children: () => undefined, ...props }) : array([], props);
 try {
  const group = f.host.querySelector('fieldset'), legend = group.querySelector('legend');
  assert.equal(titleRef.current.tagName, 'SPAN'); assert.ok(legend.contains(titleRef.current)); assert.equal(legend.textContent, 'Slot section'); assert.ok(calls > 0);
  assert.equal(descriptionRef.current.tagName, 'DIV'); const described = document.getElementById(group.getAttribute('aria-describedby'));
  assert.notEqual(described.id, 'inner-description'); assert.ok(described.contains(descriptionRef.current));
 } finally { f.close(); }
 const empty = member === 'ObjectSection' ? owner(member, { customer: {} }, { name: 'customer', title: { children: 'Named' }, description: { id: 'empty-description', children: '' }, children: () => undefined }) : array([], { title: { children: 'Named' }, description: { id: 'empty-description', children: '' } });
 try { assert.equal(empty.host.querySelector('fieldset').getAttribute('aria-describedby'), null); assert.ok(empty.host.querySelector('#empty-description')); } finally { empty.close(); }
});
test('empty section title objects need explicit outer names without pre-rendering callbacks', () => {
 const F = h.suite('ObjectSection');
 for (const title of [{ children: '' }, { children: '   ' }, { children: h.React.createElement(h.React.Fragment) }]) {
  h.rejectsRender(() => h.mount(h.React.createElement(F.Root, { defaultValues: { customer: {} }, onSubmit: () => {} }, h.React.createElement(F.ObjectSection, { name: 'customer', title, children: () => undefined }))), /accessible name/);
 }
 const f = owner('ObjectSection', { customer: {} }, { name: 'customer', title: { children: '' }, 'aria-label': 'External', children: () => undefined });
 try { assert.equal(f.host.querySelector('fieldset').getAttribute('aria-label'), 'External'); assert.ok(f.host.querySelector('legend span span')); } finally { f.close(); }
});

test('Form dictionary supplies array names, actions and literal-safe announcements with leaf precedence', () => {
 const F=h.suite('ArraySection'); let methods;
 function Owner({direct=false}) {methods=F.useForm({defaultValues:{rows:[]},strings:{arraySection:{add:'Insert row',empty:'No rows',itemLabel:position=>`Row ${position} $& {label}`,added:label=>`${label} inserted`}}});return h.React.createElement(F.Root,{methods,onSubmit:()=>{}},h.React.createElement(F.ArraySection,{name:'rows',title:'Rows',createItem:()=>({text:''}),renderItem:({Fields})=>h.React.createElement(Fields.Input,{name:'text'}),...(direct?{itemLabel:()=> 'Direct row',strings:{add:'Direct add',remove:label=>`Delete ${label}`}}:{})}));}
 const f=h.mount(h.React.createElement(Owner));
 try {assert.ok(f.host.textContent.includes('No rows'));h.click(button(f.host,'Insert row'));assert.equal(items(f.host)[0].getAttribute('aria-label'),'Row 1 $& {label}');assert.ok(button(f.host,'Remove Row 1 $& {label}'));assert.equal(f.host.querySelector('[aria-live="polite"]').textContent,'Row 1 $& {label} inserted');const input=f.host.querySelector('input');assert.equal(document.activeElement,input);f.render(h.React.createElement(Owner,{direct:true}));assert.ok(button(f.host,'Direct add'));assert.ok(button(f.host,'Delete Direct row'));assert.equal(f.host.querySelector('input'),input);assert.equal(methods.getValues('rows').length,1);}finally{f.close();}
});

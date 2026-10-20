const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, act, api, dom, fluent, renderer;
test.before(async () => {
 dom = harness.createSxDom();
 for (const key of ['HTMLElement','HTMLInputElement','HTMLSelectElement','NodeFilter','Event','CustomEvent','MouseEvent','KeyboardEvent','Element','Node','File','FileList','Blob']) global[key] = dom.window[key];
 dom.window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
 React = require('react'); ({act} = require('react-dom/test-utils')); fluent = require('@fluentui/react-components');
 renderer = require('@griffel/core').createDOMRenderer(document,{unstable_filterCSSRule:rule=>!rule.startsWith('@container')});
 api = await harness.loadModule('tests/fixtures/form-settings-entry.tsx', {reactHookForm:true});
});
test.after(async () => { dom?.window.close(); await harness.cleanup(); });
function mount(element) {
 const ReactDOM = require('react-dom'); const {RendererProvider} = require('@griffel/react');
 const host=document.createElement('div');document.body.appendChild(host);
 function render(next=element){element=next;act(()=>{ReactDOM.render(React.createElement(RendererProvider,{renderer},React.createElement(fluent.FluentProvider,{theme:fluent.webLightTheme},element)),host);});}
 try{render();}catch(error){act(()=>{ReactDOM.unmountComponentAtNode(host);});host.remove();throw error;}
 return {host,render,close(){act(()=>{ReactDOM.unmountComponentAtNode(host);});host.remove();}};
}
const options=[{value:'a',label:'Alpha'},{value:'b',label:'Beta'}];
const native=['Input','Textarea','SearchBox','Checkbox','Switch','Slider','SpinButton','Select','Dropdown','Combobox','DropdownMulti','ComboboxMulti','TagPicker','DatePicker','TimePicker','FileInput'];
const composite=['RadioGroup','CheckboxGroup','Rating','ColorPicker','SwatchPicker','DateTime'];
function value(member){return ['Checkbox','Switch'].includes(member)?false:['Slider','SpinButton','Rating'].includes(member)?2:['CheckboxGroup','DropdownMulti','ComboboxMulti'].includes(member)?[]:member==='TagPicker'?['a']:member==='ColorPicker'?{h:20,s:.4,v:.8}:member==='DateTime'?{date:null,time:null}:['DatePicker','TimePicker','FileInput'].includes(member)?null:'a';}
function extras(member){return ['Select','Dropdown','Combobox','DropdownMulti','ComboboxMulti','TagPicker','RadioGroup','CheckboxGroup'].includes(member)?{options}:member==='ColorPicker'?{children:React.createElement(fluent.ColorSlider,{channel:'hue','aria-label':'Hue'})}:member==='SwatchPicker'?{children:React.createElement(fluent.ColorSwatch,{value:'a',color:'red','aria-label':'Red'})}:member==='DateTime'?{dateLabel:'Date part',timeLabel:'Time part'}:{};}
function fixture(member, control={}, row={}, root={}) {
 const F=api.createForm();let methods;const ref=React.createRef();
 function Owner({disabled=false,required=true}){methods=F.useForm({defaultValues:{value:value(member),other:''}});return React.createElement(F.Root,{methods,onSubmit:()=>{},disabled,...root},React.createElement(api.SettingsGroup.Item,{field:true,title:React.createElement('em',undefined,'Setting title'),description:'Setting help',...row},React.createElement(F[member],{name:'value',ref,required,...extras(member),...control})));}
 const f=mount(React.createElement(Owner));return {...f,F,Owner,ref,get methods(){return methods;}};
}
function ids(node,attribute){return (node.getAttribute(attribute)||'').split(/\s+/).filter(Boolean);}
function texts(node,attribute){return ids(node,attribute).map(id=>document.getElementById(id)?.textContent?.replace(/\*$/, ''));}
for(const member of [...native,...composite])test(`${member} inherits one rich row name and description at its real semantic target`,async()=>{
 const f=fixture(member,{'aria-describedby':'extra-help',rules:{validate:()=> 'Fix setting'}});
 try{
  const target=f.ref.current;assert.ok(target);assert.ok(texts(target,'aria-labelledby').includes('Setting title'),target.outerHTML);
  assert.ok(texts(target,'aria-describedby').includes('Setting help'));assert.ok(ids(target,'aria-describedby').includes('extra-help'));
  assert.equal(f.host.textContent.split('Setting title').length-1,1);assert.equal(f.host.textContent.split('Setting help').length-1,1);
  const label=f.host.querySelector('label[data-settings-field-label]');assert.ok(label);assert.equal(label.getAttribute('data-required'),'true');
  if(native.includes(member))assert.equal(label.htmlFor,target.id);else assert.equal(label.htmlFor,'');
  await act(async()=>{await f.methods.trigger('value');});assert.ok(texts(f.ref.current,'aria-describedby').some(text=>text==='Fix setting'),JSON.stringify({current:f.ref.current.outerHTML,old:target===f.ref.current,body:f.host.textContent}));
  f.render(React.createElement(f.Owner,{disabled:true,required:false}));assert.equal(label.getAttribute('aria-disabled'),'true');assert.equal(label.getAttribute('data-required'),null);
 }finally{f.close();}
});
test('native row label activates checkbox and explicit native ids persist through reset',()=>{
 const f=fixture('Checkbox',{id:'explicit-setting'});try{const label=f.host.querySelector('label[data-settings-field-label]');assert.ok(label);act(()=>label.click());assert.equal(f.methods.getValues('value'),true);assert.equal(f.ref.current.id,'explicit-setting');act(()=>f.methods.reset());assert.equal(f.ref.current.checked,false);assert.equal(label.htmlFor,'explicit-setting');}finally{f.close();}
});
for(const control of [{label:'Repeated'},{hint:'Repeated'},{fieldProps:{hint:'Repeated'}},{fieldProps:{label:'Repeated'}}])test(`field rows reject repeated presentation ${JSON.stringify(control)}`,()=>{
 const original=console.error;console.error=()=>{};try{assert.throws(()=>fixture('Input',control),/SettingsGroup.*(label|hint)/);}finally{console.error=original;}
});
test('field rows require meaningful titles and reject multiple logical adapters',()=>{
 const F=api.createForm();const original=console.error;console.error=()=>{};
 try{assert.throws(()=>fixture('Input',{}, {title:'  '}),/SettingsGroup.*title/);
 assert.throws(()=>mount(React.createElement(F.Root,{defaultValues:{a:'',b:''},onSubmit:()=>{}},React.createElement(api.SettingsGroup.Item,{field:true,title:'Two'},React.createElement(F.Input,{name:'a'}),React.createElement(F.Input,{name:'b'})))),/SettingsGroup.*one/);
 }finally{console.error=original;}
});
test('action, nested opt-out rows and custom useField do not become field owners',()=>{
 const F=api.createForm();function Custom(){const binding=F.useField({name:'custom'});return React.createElement('input',{'data-custom':true,ref:binding.ref,value:binding.value,onChange:event=>binding.onChange(event.target.value)});}
 const f=mount(React.createElement(F.Root,{defaultValues:{primary:'',action:'',nested:'',custom:''},onSubmit:()=>{}},React.createElement(api.SettingsGroup.Item,{field:true,title:'Primary',description:'Primary help',action:React.createElement(F.Input,{name:'action',label:'Action'})},React.createElement(F.Input,{name:'primary'}),React.createElement(api.SettingsGroup.Item,{title:'Nested'},React.createElement(F.Input,{name:'nested',label:'Nested control'})),React.createElement(Custom))));
 try{const inputs=f.host.querySelectorAll('input');assert.equal(texts(inputs[0],'aria-labelledby')[0],'Primary');assert.ok(!texts(inputs[1],'aria-labelledby').includes('Primary'));assert.ok(!inputs[2].getAttribute('aria-labelledby'));assert.ok(!texts(inputs[3],'aria-labelledby').includes('Primary'));}finally{f.close();}
});
test('independent field rows retain isolated ids after reorder and opt-out remains caller labeled',()=>{
 const F=api.createForm();function Owner({reverse=false}){const methods=F.useForm({defaultValues:{a:'',b:''}});return React.createElement(F.Root,{methods,onSubmit:()=>{}},(reverse?['b','a']:['a','b']).map(name=>React.createElement(api.SettingsGroup.Item,{key:name,field:true,title:name,description:`Help ${name}`},React.createElement(F.Input,{name}))));}
 const f=mount(React.createElement(Owner)),g=fixture('Input',{label:'Own label'}, {field:false});try{const before=[...f.host.querySelectorAll('input')].map(input=>[input.name,input.id]);f.render(React.createElement(Owner,{reverse:true}));for(const input of f.host.querySelectorAll('input')){assert.equal(input.id,before.find(pair=>pair[0]===input.name)[1]);assert.deepEqual(texts(input,'aria-labelledby'),[input.name]);}assert.equal(g.host.querySelector('label').textContent.includes('Own label'),true);}finally{f.close();g.close();}
});
test('standalone unnamed TagPicker is rejected and context-named TagPicker is accepted',()=>{
 const F=api.createForm();const original=console.error;console.error=()=>{};try{assert.throws(()=>mount(React.createElement(F.Root,{defaultValues:{tags:[]},onSubmit:()=>{}},React.createElement(F.TagPicker,{name:'tags',options}))),/TagPicker.*(name|label)/);}finally{console.error=original;}
 const f=fixture('TagPicker');try{assert.ok(texts(f.host.querySelector('[role=listbox]'),'aria-labelledby').includes('Setting title'));}finally{f.close();}
});
test('public field composition types permit inherited TagPicker naming',()=>{
 const ts=require('typescript'),path=require('node:path');const program=ts.createProgram([path.resolve(__dirname,'fixtures/form-settings-types.tsx')],{noEmit:true,strict:true,skipLibCheck:true,target:ts.ScriptTarget.ES2020,module:ts.ModuleKind.ESNext,moduleResolution:ts.ModuleResolutionKind.NodeJs,jsx:ts.JsxEmit.React,lib:['lib.es2020.d.ts','lib.dom.d.ts']});const diagnostics=ts.getPreEmitDiagnostics(program);assert.equal(diagnostics.length,0,ts.formatDiagnosticsWithColorAndContext(diagnostics,{getCanonicalFileName:f=>f,getCurrentDirectory:()=>process.cwd(),getNewLine:()=> '\n'}));
});
test('field rows reject absent built-in owners, including custom-only bindings',()=>{
 const F=api.createForm();const original=console.error;console.error=()=>{};
 try{assert.throws(()=>mount(React.createElement(api.SettingsGroup.Item,{field:true,title:'Missing'},'Content')),/SettingsGroup.*one/);}finally{console.error=original;}
});
for(const member of ['Input','Textarea','SearchBox','Checkbox','Switch','Slider','SpinButton','Select','Dropdown','Combobox','DropdownMulti','ComboboxMulti'])test(`${member} preserves an explicit native-slot id and native description`,async()=>{
 const slot=member==='Textarea'?'textarea':member==='Select'?'select':['Dropdown','DropdownMulti'].includes(member)?'button':'input';
 const f=fixture(member,{[slot]:{id:`slot-${member}`,'aria-describedby':'slot-description'}});try{assert.equal(f.ref.current.id,`slot-${member}`);assert.equal(f.host.querySelector('label[data-settings-field-label]').htmlFor,`slot-${member}`);assert.ok(ids(f.ref.current,'aria-describedby').includes('slot-description'));await act(async()=>{await f.methods.trigger('value');});}finally{f.close();}
});
test('description removal, field toggle and owner replacement clean up relationships',()=>{
 const F=api.createForm();function Owner({field=true,description,replacement=false}){return React.createElement(F.Root,{defaultValues:{a:'',b:false},onSubmit:()=>{}},React.createElement(api.SettingsGroup.Item,{field,title:'Mutable',description},replacement?React.createElement(F.Switch,{name:'b',id:'replacement'}):React.createElement(F.Input,{name:'a',id:'original'})));}
 const f=mount(React.createElement(Owner,{description:'Help'}));try{f.render(React.createElement(Owner,{description:undefined,replacement:true}));assert.ok(!texts(f.host.querySelector('input'),'aria-describedby').includes('Help'));assert.equal(f.host.querySelector('label[data-settings-field-label]').htmlFor,'replacement');f.render(React.createElement(Owner,{field:false}));assert.equal(f.host.querySelector('label[data-settings-field-label]'),null);assert.equal(f.host.querySelector('input').getAttribute('aria-labelledby'),null);f.render(React.createElement(Owner));assert.equal(f.host.querySelector('label[data-settings-field-label]').htmlFor,'original');}finally{f.close();}
});
test('DateTime parts keep their own names and inherited group does not label parts',()=>{
 const f=fixture('DateTime');try{for(const input of f.ref.current.querySelectorAll('input')){assert.ok(!texts(input,'aria-labelledby').includes('Setting title'));assert.ok(!texts(input,'aria-describedby').includes('Setting help'));}assert.ok(f.host.textContent.includes('Date part'));assert.ok(f.host.textContent.includes('Time part'));}finally{f.close();}
});
for(const member of ['Input','Checkbox','Textarea'])test(`${member} row label reflects native-slot required and disabled`,()=>{
 const slot=member==='Textarea'?'textarea':'input';const f=fixture(member,{required:false,[slot]:{required:true,disabled:true}});try{const label=f.host.querySelector('label[data-settings-field-label]');assert.equal(label.getAttribute('data-required'),'true');assert.equal(label.getAttribute('aria-disabled'),'true');assert.equal(f.ref.current.disabled,true);}finally{f.close();}
});
for(const member of ['Input','Textarea','DatePicker','TimePicker','FileInput','RadioGroup','CheckboxGroup','Rating','ColorPicker','SwatchPicker','TagPicker','DateTime'])test(`${member} descriptions contain each row, caller and error relationship once`,async()=>{
 const f=fixture(member,{'aria-describedby':'extra-help',rules:{validate:()=> 'Fix setting'}});try{await act(async()=>{await f.methods.trigger('value');});const values=ids(f.ref.current,'aria-describedby');assert.equal(values.length,new Set(values).size,values.join(' '));}finally{f.close();}
});
test('native required false keeps row marker truthful when outer presentation is required',()=>{
 const f=fixture('Input',{required:true,input:{required:false}});try{assert.equal(f.ref.current.required,false);assert.equal(f.host.querySelector('label[data-settings-field-label]').getAttribute('data-required'),null);}finally{f.close();}
});

test('independent forms with the same field name have isolated label and hint identities',()=>{
 const f=fixture('Input'),g=fixture('Input');try{assert.notEqual(f.ref.current.id,g.ref.current.id);for(const mounted of [f,g]){assert.equal(mounted.host.querySelector('label[data-settings-field-label]').htmlFor,mounted.ref.current.id);for(const id of [...ids(mounted.ref.current,'aria-labelledby'),...ids(mounted.ref.current,'aria-describedby')])assert.ok(mounted.host.contains(document.getElementById(id)));}}finally{f.close();g.close();}
});
for(const shape of ['whitespace span','empty span','nested empty markup','empty fragment'])test(`field rows reject statically empty rich title: ${shape}`,()=>{
 const titles={
  'whitespace span':React.createElement('span',undefined,'   '),
  'empty span':React.createElement('span'),
  'nested empty markup':React.createElement('span',undefined,React.createElement(React.Fragment,undefined,React.createElement('em',undefined,' \n '))),
  'empty fragment':React.createElement(React.Fragment,undefined,undefined,false,' ')
 };
 const original=console.error;console.error=()=>{};try{assert.throws(()=>fixture('Input',{}, {title:titles[shape]}),/SettingsGroup.*title/);}finally{console.error=original;}
});
for(const shape of ['zero','nested number','nested text','custom component'])test(`field rows accept meaningful rich title: ${shape}`,()=>{
 function Title(){return React.createElement('strong',undefined,'Custom title');}
 const titles={zero:0,'nested number':React.createElement('span',undefined,React.createElement(React.Fragment,undefined,0)),'nested text':React.createElement(React.Fragment,undefined,' ',React.createElement('em',undefined,React.createElement('strong',undefined,'Named title'))),'custom component':React.createElement(Title)};
 const f=fixture('Input',{required:false},{title:titles[shape]});try{assert.ok(texts(f.ref.current,'aria-labelledby').some(text=>text?.trim().length>0));}finally{f.close();}
});
test('standalone TagPicker preserves caller descriptions alongside the Fluent Field hint',()=>{
 const F=api.createForm(),ref=React.createRef();
 const f=mount(React.createElement(F.Root,{defaultValues:{tags:[]},onSubmit:()=>{}},React.createElement('p',{id:'standalone-caller-help'},'Caller help'),React.createElement(F.TagPicker,{name:'tags',ref,label:'Tags',hint:'Field hint','aria-describedby':'standalone-caller-help',options})));
 try{const descriptions=ids(ref.current,'aria-describedby');assert.ok(descriptions.includes('standalone-caller-help'));assert.ok(texts(ref.current,'aria-describedby').includes('Field hint'));assert.equal(descriptions.length,new Set(descriptions).size);}finally{f.close();}
});

test('field bridge inspects slot children, preserves outer Label and excludes empty description associations', () => {
 const ref=React.createRef(); const f=fixture('Input',{}, {title:{ref,children:'Slot label'},description:{children:''}});
 try{const label=f.host.querySelector('label[data-settings-field-label]');assert.equal(ref.current.tagName,'SPAN');assert.ok(label.contains(ref.current));assert.equal(label.htmlFor,f.ref.current.id);assert.deepEqual(texts(f.ref.current,'aria-labelledby'),['Slot label']);assert.ok(!ids(f.ref.current,'aria-describedby').some(id=>id.startsWith('settings-description-')));}finally{f.close();}
 const original=console.error;console.error=()=>{};try{assert.throws(()=>fixture('Input',{}, {title:{children:React.createElement('em',{},' ')}}),/SettingsGroup.*title/);}finally{console.error=original;}
 let calls=0,renders=0;function Title(){renders++;return 'Callback label';}const g=fixture('Input',{}, {title:{children:(Tag,props)=>{calls++;return React.createElement(Tag,props,React.createElement(Title));}}});try{assert.equal(calls,renders,'validation must not execute title callbacks');assert.ok(texts(g.ref.current,'aria-labelledby').includes('Callback label'));}finally{g.close();}
});

test('field bridge keeps blank object description containers without blank associations and preserves whitespace shorthand', () => {
 for (const description of [{children:' '},{children:React.createElement(React.Fragment,{},' ',false,'\n')}]) {
  const f=fixture('Input',{}, {description});try{assert.ok(f.host.querySelector('[id^="settings-description-"] span span'));assert.ok(!ids(f.ref.current,'aria-describedby').some(id=>id.startsWith('settings-description-')));}finally{f.close();}
 }
 const f=fixture('Input',{}, {description:' '});try{assert.ok(ids(f.ref.current,'aria-describedby').some(id=>id.startsWith('settings-description-')));}finally{f.close();}
});

for (const control of [{ label: undefined, hint: { children: 'Competing hint' } }, { label: { children: 'Competing label' } }, { label: null }, { label: false }, { hint: null }, { hint: false }]) test(`field rows retain defined presentation conflict policy ${JSON.stringify(control)}`, () => {
 const original = console.error; console.error = () => {};
 try { assert.throws(() => fixture('Input', control), /SettingsGroup.*(label|hint)/); }
 finally { console.error = original; }
});

const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const ts = require('typescript');
const h = require('./fixtures/form-selection-harness.cjs');
test.before(h.setup);
test.after(h.cleanup);
const {arrayOwner}=require('./fixtures/form-composition-harness.cjs');
test('Object-in-Array and Array-in-Object bind actual relative model paths without layout', () => {
  const F = h.suite('Object'); let methods;
  function Owner() {
    methods = F.useForm({defaultValues: {customer: {lines: [{detail: {sku: 'first'}}]}}});
    return h.React.createElement(F.Root, {methods,onSubmit:()=>{}},
      h.React.createElement(F.Object,{name:'customer'},({Fields}) =>
        h.React.createElement(Fields.Array,{name:'lines'},({items})=>items.map(({key,Fields})=>
          h.React.createElement(Fields.Object,{key,name:'detail'},({Fields})=>
            h.React.createElement(Fields.Input,{name:'sku','aria-label':'SKU'}))))));
  }
  const f=h.mount(h.React.createElement(Owner));
  try {const input=f.host.querySelector('input'); assert.equal(input.name,'customer.lines.0.detail.sku'); h.change(input,'changed'); assert.equal(methods.getValues('customer.lines.0.detail.sku'),'changed'); assert.equal(input.closest('form').querySelectorAll('input').length,1);} finally {f.close();}
});
test('Captured item Fields follow RHF stable key across move and swap without input remount or focus movement', () => {
  const captured = new Map();
  const f=arrayOwner({rows:[{id:'domain-a',text:'A'},{id:'domain-b',text:'B'},{id:'domain-c',text:'C'}]},({items})=>items.map(item=>{
    if(!captured.has(item.key)) captured.set(item.key,item.Fields);
    const Fields=captured.get(item.key);
    return h.React.createElement(Fields.Input,{key:item.key,name:'text','aria-label':item.key});
  }));
  try {
    const before=f.array.items.map(item=>({key:item.key,Fields:item.Fields,node:f.host.querySelector(`[aria-label="${item.key}"]`)}));
    assert.ok(before.every((item,index)=>item.key!==f.methods.getValues(`rows.${index}.id`)));
    before[0].node.focus(); h.act(()=>f.array.move(0,2));
    assert.equal(f.array.items[2].key,before[0].key); assert.equal(f.array.items[2].Fields,before[0].Fields); assert.equal(f.host.querySelector(`[aria-label="${before[0].key}"]`),before[0].node); assert.equal(document.activeElement,before[0].node);
    h.change(before[0].node,'moved'); assert.equal(f.methods.getValues('rows.2.text'),'moved');
    h.act(()=>f.array.swap(0,2)); h.change(before[0].node,'swapped'); assert.equal(f.methods.getValues('rows.0.text'),'swapped');
    assert.equal(document.activeElement,before[0].node);
  } finally {f.close();}
});
test('All eight structured operations preserve domain values and headless append/prepend/insert never focus', () => {
  const f=arrayOwner({rows:[{text:'A'},{text:'B'}]},({items})=>items.map(({key,Fields})=>h.React.createElement(Fields.Input,{key,name:'text'})));
  const values=()=>f.methods.getValues('rows').map(value=>value.text);
  try {
    const original=f.host.querySelector('input'); original.focus();
    h.act(()=>f.array.append({text:'C'})); assert.deepEqual(values(),['A','B','C']); assert.equal(document.activeElement,original);
    h.act(()=>f.array.prepend({text:'Z'})); assert.deepEqual(values(),['Z','A','B','C']); assert.equal(document.activeElement,original);
    h.act(()=>f.array.insert(2,{text:'X'})); assert.deepEqual(values(),['Z','A','X','B','C']); assert.equal(document.activeElement,original);
    h.act(()=>f.array.update(2,{text:'Y'})); assert.deepEqual(values(),['Z','A','Y','B','C']);
    h.act(()=>f.array.move(0,4)); assert.deepEqual(values(),['A','Y','B','C','Z']);
    h.act(()=>f.array.swap(0,4)); assert.deepEqual(values(),['Z','Y','B','C','A']);
    h.act(()=>f.array.remove([1,3])); assert.deepEqual(values(),['Z','B','A']);
    h.act(()=>f.array.replace([{text:'R'}])); assert.deepEqual(values(),['R']);
    h.act(()=>f.array.remove(undefined)); assert.deepEqual(values(),[]);
  } finally {f.close();}
});
test('Invalid indices, duplicate or empty removals are exact no-ops', () => {
  const f=arrayOwner({rows:[{text:'A'},{text:'B'}]});
  try {const keys=f.array.items.map(item=>item.key);
    for(const op of [a=>a.insert(-1,{text:'X'}),a=>a.insert(3,{text:'X'}),a=>a.insert(.5,{text:'X'}),a=>a.update(2,{text:'X'}),a=>a.update(NaN,{text:'X'}),a=>a.remove([0,0]),a=>a.remove([]),a=>a.remove([0,3]),a=>a.remove(-1),a=>a.move(0,2),a=>a.move(-1,0),a=>a.swap(.5,1),a=>a.swap(0,2)]) h.act(()=>op(f.array));
    assert.deepEqual(f.methods.getValues('rows'),[{text:'A'},{text:'B'}]); assert.deepEqual(f.array.items.map(item=>item.key),keys);
  } finally {f.close();}
});
test('Disabled Root suppresses all eight array operations', () => {
  const f=arrayOwner({rows:[{text:'A'},{text:'B'}]},undefined,{disabled:true});
  try {const keys=f.array.items.map(item=>item.key);
    for(const op of [a=>a.append({text:'X'}),a=>a.prepend({text:'X'}),a=>a.insert(0,{text:'X'}),a=>a.remove(),a=>a.update(0,{text:'X'}),a=>a.replace([]),a=>a.move(0,1),a=>a.swap(0,1)]) h.act(()=>op(f.array));
    assert.deepEqual(f.methods.getValues('rows'),[{text:'A'},{text:'B'}]); assert.deepEqual(f.array.items.map(item=>item.key),keys);
  } finally {f.close();}
});
test('Array-level domain rules reject a resolver instead of silently disappearing', () => {
  const F=h.suite('Array');
  h.rejectsRender(()=>h.mount(h.React.createElement(F.Root,{defaultValues:{rows:[]},resolver:async values=>({values,errors:{}}),onSubmit:()=>{}},h.React.createElement(F.Array,{name:'rows',rules:{minLength:1}},()=>undefined))),/rules cannot accompany.*resolver/);
});
test('Array-root errors are available to headless renderer', async () => {
  const f=arrayOwner({rows:[]},undefined,{}, {rules:{validate:values=>values.length>0||'Add a row'}});
  try {await h.act(async()=>{await f.methods.trigger('rows');}); assert.equal(f.array.error,'Add a row'); assert.deepEqual(f.array.errors,['Add a row']);}finally {f.close();}
});
test('Composition declarations infer actual object/item models and reject primitive arrays', () => {
  const program=ts.createProgram([path.resolve(__dirname,'fixtures/form-composition-types.tsx')],{noEmit:true,strict:true,skipLibCheck:true,target:ts.ScriptTarget.ES2020,module:ts.ModuleKind.ESNext,moduleResolution:ts.ModuleResolutionKind.NodeJs,jsx:ts.JsxEmit.React,lib:['lib.es2020.d.ts','lib.dom.d.ts']});
  const diagnostics=ts.getPreEmitDiagnostics(program);
  assert.equal(diagnostics.length,0,ts.formatDiagnosticsWithColorAndContext(diagnostics,{getCanonicalFileName:file=>file,getCurrentDirectory:()=>process.cwd(),getNewLine:()=> '\n'}));
});
test('Item metadata follows order, while update/replace allocate replacement keys without changing domain IDs', () => {
  const f=arrayOwner({rows:[{id:'a',text:'A'},{id:'b',text:'B'}]},({items})=>items.map(({key,Fields})=>h.React.createElement(Fields.Input,{key,name:'text','aria-label':key})));
  try {
    const old=f.array.items.map(item=>item.key);
    assert.deepEqual(f.array.items.map(({index,total,isFirst,isLast})=>({index,total,isFirst,isLast})),[{index:0,total:2,isFirst:true,isLast:false},{index:1,total:2,isFirst:false,isLast:true}]);
    h.act(()=>f.array.update(0,{id:'a',text:'Updated'}));
    assert.notEqual(f.array.items[0].key,old[0]);assert.equal(f.array.items[1].key,old[1]);assert.equal(f.methods.getValues('rows.0.id'),'a');
    const updated=f.array.items.map(item=>item.key);
    h.act(()=>f.array.replace([{id:'a',text:'Replacement'}]));
    assert.ok(!updated.includes(f.array.items[0].key));assert.equal(f.array.items[0].total,1);assert.equal(f.array.items[0].isFirst,true);assert.equal(f.array.items[0].isLast,true);
    assert.equal(f.methods.getValues('rows.0.id'),'a');
  }finally{f.close();}
});

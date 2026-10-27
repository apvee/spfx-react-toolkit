const test=require('node:test');
const assert=require('node:assert/strict');
const h=require('./fixtures/form-selection-harness.cjs');
const {arrayOwner}=require('./fixtures/form-composition-harness.cjs');
test.before(h.setup);test.after(h.cleanup);
const options=[{value:'a',label:'Alpha'},{value:'b',label:'Beta'}];
function drafts({items}) {return items.map(({key,Fields})=>h.React.createElement(Fields.ComboboxMulti,{key,name:'choices',options,'aria-label':key}));}
function input(f,key) {return f.host.querySelector(`[aria-label="${key}"]`);}
const defaults={rows:[{choices:['a']},{choices:['b']} ]};
test('Real ComboboxMulti transient drafts and nodes survive move/swap then resetField targets current owner',()=>{
 const f=arrayOwner(defaults,drafts);
 try{const [a,b]=f.array.items.map(item=>item.key);const nodeA=input(f,a),nodeB=input(f,b);h.change(nodeA,'draft a');h.change(nodeB,'draft b');
 h.act(()=>f.array.move(0,1));assert.equal(input(f,a),nodeA);assert.equal(nodeA.value,'draft a');assert.equal(nodeB.value,'draft b');
 h.act(()=>f.methods.resetField('rows.1.choices'));assert.equal(nodeA.value,'');assert.equal(nodeB.value,'draft b');
 h.change(nodeA,'again');h.act(()=>f.array.swap(0,1));assert.equal(nodeA.value,'again');assert.equal(nodeB.value,'draft b');
 h.act(()=>f.methods.resetField('rows.0.choices'));assert.equal(nodeA.value,'');assert.equal(nodeB.value,'draft b');
 }finally{f.close();}
});
for(const operation of ['append','prepend','insert','update','replace','remove','move','swap'])test(`Reset immediately after ${operation} clears actual drafts and establishes supplied record`,()=>{
 const f=arrayOwner(defaults,drafts);try{for(const item of f.array.items)h.change(input(f,item.key),'transient');
 h.act(()=>{const a=f.array;switch(operation){case 'append':a.append({choices:['a']});break;case 'prepend':a.prepend({choices:['a']});break;case 'insert':a.insert(1,{choices:['a']});break;case 'update':a.update(0,{choices:['b']});break;case 'replace':a.replace([{choices:['b']}]);break;case 'remove':a.remove(0);break;case 'move':a.move(0,1);break;case 'swap':a.swap(0,1);break;}f.methods.reset({rows:[{choices:['b']},{choices:['a']}]});});
 assert.deepEqual(f.methods.getValues('rows'),[{choices:['b']},{choices:['a']}]);assert.equal(f.array.items.length,2);for (const [index, item] of f.array.items.entries()) h.assertMultiQueryCleared(input(f,item.key), index === 0 ? 'Beta' : 'Alpha');
 }finally{f.close();}
});
test('Targeted reset immediately after move resolves moved owners before their render commits',()=>{
 const f=arrayOwner(defaults,drafts);try{const [a,b]=f.array.items.map(item=>item.key);h.change(input(f,a),'draft a');h.change(input(f,b),'draft b');
 h.act(()=>{f.array.move(0,1);f.methods.resetField('rows.1.choices');});assert.equal(input(f,a).value,'');assert.equal(input(f,b).value,'draft b');
 }finally{f.close();}
});
test('Removal cleans draft validators and resetting sibling arrays/forms preserves independent queries',async()=>{
 const F=h.suite('Array');let methods,left,right;
 function Owner(){methods=F.useForm({defaultValues:{left:[{choice:'a'},{choice:'b'}],right:[{choice:'a'}]}});return h.React.createElement(F.Root,{methods,onSubmit:()=>{}},['left','right'].map(name=>h.React.createElement(F.Array,{key:name,name},a=>{if(name==='left')left=a;else right=a;return a.items.map(({key,Fields})=>h.React.createElement(Fields.Combobox,{key,name:'choice',options,'aria-label':`${name}-${key}`}));})));}
 const f=h.mount(h.React.createElement(Owner));const other=arrayOwner(defaults,drafts);
 try{const key=left.items[0].key;h.change(f.host.querySelector(`[aria-label="left-${key}"]`),'removed invalid query');
 const sibling=f.host.querySelector(`[aria-label="right-${right.items[0].key}"]`);h.change(sibling,'sibling draft');const otherNode=input(other,other.array.items[0].key);h.change(otherNode,'independent draft');
 h.act(()=>left.remove(0));await h.act(async()=>{assert.equal(await methods.trigger('left'),true);});
 h.act(()=>methods.resetField('left.0.choice'));assert.equal(sibling.value,'sibling draft');assert.equal(otherNode.value,'independent draft');
 h.act(()=>methods.reset({left:[{choice:'a'}],right:[{choice:'a'}]}));assert.equal(otherNode.value,'independent draft');
 }finally{f.close();other.close();}
});
test('keepDirtyValues immediately after move preserves the current dirty owner query and clears clean owners',()=>{
 const baseline={rows:[{choices:['a']},{choices:['a']}]};
 const f=arrayOwner(baseline,drafts);try{const [key,cleanKey]=f.array.items.map(item=>item.key);
 h.act(()=>f.methods.setValue('rows.0.choices',['a','b'],{shouldDirty:true}));h.change(input(f,key),'preserved draft');h.change(input(f,cleanKey),'clean draft');
 h.act(()=>{f.array.move(0,1);assert.equal(f.methods.getFieldState('rows.1.choices').isDirty,true);assert.equal(f.methods.getFieldState('rows.0.choices').isDirty,false);f.methods.reset(baseline,{keepDirtyValues:true});});
 assert.equal(input(f,f.array.items[1].key).value,'preserved draft');h.assertMultiQueryCleared(input(f,f.array.items[0].key),'Alpha');assert.deepEqual(f.methods.getValues('rows.1.choices'),['a','b']);
 }finally{f.close();}
});

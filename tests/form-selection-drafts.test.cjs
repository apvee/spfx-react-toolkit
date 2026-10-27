const test=require('node:test');const assert=require('node:assert/strict');const h=require('./fixtures/form-selection-harness.cjs');test.before(h.setup);test.after(h.cleanup);
const options=[{value:'a',label:'Alpha'},{value:'b',label:'Beta'}];
for(const member of ['Combobox','ComboboxMulti','TagPicker'])test(`${member} retains visible query, synchronizes explicit same-value resets and targeted resetField`,()=>{const F=h.suite(member);const initial=member==='Combobox'?'a':['a'];const f=h.owner(F,member,{value:initial},{label:'Choices',options});try{const input=f.host.querySelector('input');h.change(input,'missing query');assert.equal(input.value,'missing query');assert.deepEqual(f.methods.getValues('value'),member==='Combobox'?'':['a']);h.act(()=>f.methods.reset({value:f.methods.getValues('value')}));assert.equal(input.value,'');h.change(input,'another query');h.act(()=>f.methods.resetField('value'));assert.equal(input.value,'');}finally{f.close();}});
test('Combobox clears identity for a nonmatching query before both query callbacks',()=>{const F=h.suite('Combobox');let f;const seen=[];f=h.owner(F,'Combobox',{value:'a'},{options,onChange:()=>seen.push(['top',f.methods.getValues('value')]),input:{onChange:()=>seen.push(['slot',f.methods.getValues('value')])}});try{h.change(f.host.querySelector('input'),'zzz');assert.deepEqual(seen,[['top',''],['slot','']]);assert.equal(f.host.querySelector('input').value,'zzz');assert.equal(f.methods.getFieldState('value').isDirty,true);}finally{f.close();}});
for(const member of ['Combobox','ComboboxMulti','TagPicker'])test(`${member} explicit keepDirtyValues preserves dirty owners, ordinary reset clears query`,()=>{const F=h.suite(member),initial=member==='Combobox'?'a':['a'];const f=h.owner(F,member,{value:initial},{label:'Choices',options});try{h.act(()=>f.methods.setValue('value',member==='Combobox'?'b':['b'],{shouldDirty:true}));h.change(f.host.querySelector('input'),'draft');h.act(()=>f.methods.reset({value:initial},{keepDirtyValues:true}));assert.equal(f.host.querySelector('input').value,'draft');h.act(()=>f.methods.reset({value:initial}));assert.equal(f.host.querySelector('input').value,member==='Combobox'?'Alpha':'');}finally{f.close();}});
for(const member of ['Combobox','ComboboxMulti','TagPicker'])test(`${member} external field updates synchronize display while local query edits remain visible`,()=>{const F=h.suite(member),initial=member==='Combobox'?'a':['a'];const f=h.owner(F,member,{value:initial},{label:'Choices',options});try{h.change(f.host.querySelector('input'),'draft');assert.equal(f.host.querySelector('input').value,'draft');h.act(()=>f.methods.setValue('value',member==='Combobox'?'b':['b']));assert.equal(f.host.querySelector('input').value,member==='Combobox'?'Beta':'');}finally{f.close();}});
for(const member of ['Combobox','ComboboxMulti','TagPicker'])test(`${member} reactive values clear query only after RHF actually resets the record`,()=>{const F=h.suite(member);const initial={value:member==='Combobox'?'a':['a'],other:'first'};let methods;function Owner({values}){methods=F.useForm({values});return h.React.createElement(F.Root,{methods,onSubmit:()=>{}},h.React.createElement(F[member],{name:'value',options,label:'Choices'}));}const f=h.mount(h.React.createElement(Owner,{values:initial}));try{h.change(f.host.querySelector('input'),'draft');f.render(h.React.createElement(Owner,{values:{...initial}}));assert.equal(f.host.querySelector('input').value,'draft');f.render(h.React.createElement(Owner,{values:{...initial,other:'changed'}}));assert.equal(f.host.querySelector('input').value,member==='Combobox'?'Alpha':'');assert.equal(methods.getValues('other'),'changed');}finally{f.close();}});
for(const member of ['Combobox','ComboboxMulti','TagPicker'])test(`${member} targeted reset leaves another owner query and native ref unchanged`,()=>{const F=h.suite(member);let methods;const left=h.React.createRef(),right=h.React.createRef();function Owner(){methods=F.useForm({defaultValues:{left:member==='Combobox'?'a':['a'],right:member==='Combobox'?'a':['a']}});return h.React.createElement(F.Root,{methods,onSubmit:()=>{}},h.React.createElement(F[member],{name:'left',options,label:'Left',ref:left}),h.React.createElement(F[member],{name:'right',options,label:'Right',ref:right}));}const f=h.mount(h.React.createElement(Owner));try{h.change(left.current,'left query');h.change(right.current,'right query');const original=right.current;h.act(()=>{methods.resetField('left');});assert.equal(left.current.value,member==='Combobox'?'Alpha':'');assert.equal(right.current.value,'right query');assert.equal(right.current,original);}finally{f.close();}});
test('Combobox query-clear selection notification follows its single RHF write exactly once',()=>{const F=h.suite('Combobox');let f;const seen=[];f=h.owner(F,'Combobox',{value:'a'},{options,open:true,inlinePopup:true,onOptionSelect:()=>seen.push(f.methods.getValues('value'))});const writes=[];const unsubscribe=f.methods.subscribe({name:'value',exact:true,formState:{values:true},callback:state=>writes.push(state.values.value)});try{h.change(f.host.querySelector('input'),'unmatched');assert.deepEqual(writes,['']);assert.deepEqual(seen,['']);}finally{unsubscribe();f.close();}});

for (const member of ['Combobox', 'ComboboxMulti', 'TagPicker']) test(`${member} dirty array query consumes one reset handoff while clean queries clear`, () => {
  const F = h.suite(member); let methods, array;
  const defaults = { rows: member === 'Combobox' ? [{ value: 'a' }, { value: '' }] : [{ value: ['a'] }, { value: ['a'] }] };
  function Owner() {
    methods = F.useForm({ defaultValues: defaults });
    return h.React.createElement(F.Root, { methods, onSubmit: () => {} }, h.React.createElement(F.Array, { name: 'rows' }, context => {
      array = context;
      return context.items.map(({ key, Fields }) => h.React.createElement(Fields[member], { key, name: 'value', options, 'aria-label': key }));
    }));
  }
  const f = h.mount(h.React.createElement(Owner));
  const nodes = () => array.items.map(({ key }) => f.host.querySelector(`input[aria-label="${key}"]`));
  try {
    if (member !== 'Combobox') h.act(() => methods.setValue('rows.0.value', ['a', 'b'], { shouldDirty: true }));
    const original = nodes(); h.change(original[0], 'dirty draft'); h.change(original[1], 'clean query');
    assert.equal(methods.getFieldState('rows.0.value').isDirty, true); assert.equal(methods.getFieldState('rows.1.value').isDirty, false);
    h.act(() => methods.reset(defaults, { keepDirtyValues: true }));
    const remounted = nodes(); assert.notEqual(remounted[0], original[0], 'RHF regenerated the row key');
    assert.equal(remounted[0].value, 'dirty draft'); assert.equal(remounted[1].value, member === 'ComboboxMulti' ? 'Alpha' : '');
    assert.deepEqual(methods.getValues('rows.0.value'), member === 'Combobox' ? '' : ['a', 'b']);
    h.act(() => methods.reset(defaults, { keepValues: true }));
    assert.equal(nodes()[0], remounted[0]); assert.equal(remounted[0].value, 'dirty draft');
    h.act(() => methods.reset(defaults));
    assert.equal(nodes()[0].value, member === 'TagPicker' ? '' : 'Alpha'); assert.equal(nodes()[1].value, member === 'ComboboxMulti' ? 'Alpha' : '');
  } finally { f.close(); }
});

for (const member of ['Combobox', 'ComboboxMulti', 'TagPicker']) for (const target of ['rows.0.value', 'rows.0']) test(`${member} resetField ${target} supersedes its pending query handoff and retains a dirty sibling`, () => {
  const F = h.suite(member); let methods, array;
  const initial = member === 'Combobox' ? 'a' : ['a'];
  const defaults = { rows: [{ value: initial }, { value: initial }] };
  function Owner() {
    methods = F.useForm({ defaultValues: defaults });
    if (target === 'rows.0') methods.register('rows.0');
    return h.React.createElement(F.Root, { methods, onSubmit: () => {} }, h.React.createElement(F.Array, { name: 'rows' }, context => {
      array = context;
      return context.items.map(({ key, Fields }) => h.React.createElement(Fields[member], { key, name: 'value', options, 'aria-label': key }));
    }));
  }
  const f = h.mount(h.React.createElement(Owner));
  const nodes = () => array.items.map(({ key }) => f.host.querySelector(`input[aria-label="${key}"]`));
  try {
    if (member !== 'Combobox') h.act(() => { methods.setValue('rows.0.value', ['a', 'b'], { shouldDirty: true }); methods.setValue('rows.1.value', ['b'], { shouldDirty: true }); });
    const original = nodes(); h.change(original[0], 'targeted draft'); h.change(original[1], 'sibling draft');
    assert.equal(methods.getFieldState('rows.0.value').isDirty, true); assert.equal(methods.getFieldState('rows.1.value').isDirty, true);
    h.act(() => { methods.reset(defaults, { keepDirtyValues: true }); methods.resetField(target); });
    const remounted = nodes(); assert.notEqual(remounted[0], original[0], 'the new RHF key consumes presentation after the targeted reset');
    assert.deepEqual(methods.getValues('rows.0.value'), initial);
    assert.deepEqual(methods.getValues('rows.1.value'), member === 'Combobox' ? '' : ['b']);
    assert.equal(remounted[0].value, member === 'TagPicker' ? '' : 'Alpha', 'targeted reset discards preserved query text');
    if (member === 'ComboboxMulti') { h.act(() => require('react-dom/test-utils').Simulate.focus(remounted[0])); assert.equal(remounted[0].value, '', 'targeted query is actually cleared'); }
    assert.equal(remounted[1].value, 'sibling draft', 'unrelated dirty query keeps its valid handoff');
  } finally { f.close(); }
});

test('ComboboxMulti shows selected labels at rest and only the query while editing', () => {
 const F=h.suite('ComboboxMulti'); const f=h.owner(F,'ComboboxMulti',{value:['a']},{label:'Choices',options});
 try {
  const input=f.host.querySelector('input');
  assert.equal(input.value,'Alpha');
  h.act(()=>require('react-dom/test-utils').Simulate.focus(input)); assert.equal(input.value,'');
  h.change(input,'query'); assert.equal(input.value,'query'); assert.deepEqual(f.methods.getValues('value'),['a']);
  h.act(()=>require('react-dom/test-utils').Simulate.blur(input)); assert.equal(input.value,'Alpha');
  h.act(()=>f.methods.reset({value:['b']})); assert.equal(input.value,'Beta');
  assert.equal(f.host.querySelector('[data-form-selection-summary]'),null);
  h.act(()=>f.methods.reset({value:[]})); assert.equal(input.value,'');
 } finally {f.close();}
});

const test=require('node:test');
const assert=require('node:assert/strict');
const React=require('react');
const harness=require('./fixtures/formatted-text-harness.cjs');
let familyModule,normalization,rows,state;
test.before(async()=>{
 [familyModule,normalization,rows,state]=await Promise.all(['column-family.internal.tsx','column-normalization.internal.ts','prepared-rows.internal.ts','state-validation.internal.ts'].map(file=>harness.loadModule(`packages/spfx-react-toolkit/src/controls/data-grid/${file}`)));
});
test.after(()=>harness.cleanup());
function family(){return familyModule.createTypedColumnFamily();}
test('exactly one declared collection, including an explicit empty collection',()=>{
 const f=family();
 assert.throws(()=>normalization.normalizeDataGridColumns(f,{}),/exactly one/);
 assert.throws(()=>normalization.normalizeDataGridColumns(f,{children:[],columns:[]}),/exactly one/);
 assert.deepEqual(normalization.normalizeDataGridColumns(f,{columns:[]}),[]);
 assert.deepEqual(normalization.normalizeDataGridColumns(f,{children:[]}),[]);
});
test('nested fragments and arrays preserve same-family marker order',()=>{
 const f=family();const marker=id=>React.createElement(f.TextColumn,{key:id,id,header:id,field:'name'});
 const columns=normalization.normalizeDataGridColumns(f,{children:[marker('a'),React.createElement(React.Fragment,{},[marker('b'),false,undefined])]});
 assert.deepEqual(columns.map(c=>c.id),['a','b']);assert.ok(columns.every(Object.isFrozen));
 assert.throws(()=>normalization.normalizeDataGridColumns(f,{children:React.createElement('div')}),/same.*family|column marker/);
 assert.throws(()=>normalization.normalizeDataGridColumns(f,{children:React.createElement(family().TextColumn,{id:'x',header:'X',field:'name'})}),/family/);
});
test('descriptor ownership and immutability cannot be forged with literals',()=>{
 const f=family();const columns=f.defineColumns(c=>[c.text({id:'x',header:'X',field:'name'})]);
 assert.ok(Object.isFrozen(columns));assert.ok(Object.isFrozen(columns[0]));
 assert.equal(normalization.normalizeDataGridColumns(f,{columns})[0].read({name:'value'}),'value');
 assert.throws(()=>normalization.normalizeDataGridColumns(f,{columns:[{id:'x',kind:'text',field:'name'}]}),/descriptor|family/);
 assert.throws(()=>normalization.normalizeDataGridColumns(family(),{columns}),/family/);
});
test('invalid IDs, sources and options fail at deterministic column boundary',()=>{
 const f=family();const make=(method,props)=>()=>f.defineColumns(c=>[c[method](props)]);
 for(const props of [{id:'',header:'x',field:'name'},{id:'x',field:'name'},{id:'x',header:'x'}, {id:'x',header:'x',field:'name',accessor:()=>''},{id:'x',header:'x',field:'name',width:Infinity},{id:'x',header:'x',field:'name',hideable:true},{id:'x',header:'x',field:'name',locale:'bad_!'}])assert.throws(make('text',props),TypeError);
 assert.throws(make('boolean',{id:'b',header:'B',field:'x',display:'bad'}),/display/);
 assert.throws(make('progress',{id:'p',header:'P',field:'x',max:0}),/max/);
 assert.throws(make('rating',{id:'r',header:'R',field:'x',max:1.5}),/max/);
 assert.throws(make('people',{id:'p',header:'P',field:'x',getName:()=>'',getPersonId:()=>1,maxVisible:-1}),/maxVisible/);
 assert.throws(make('actions',{id:'a',actions:[{id:'',label:'X',placement:'inline',onClick(){}}]}),/action/);
 assert.throws(make('actions',{id:'a',field:'x',actions:[]}),/source/);
 assert.throws(make('tags',{id:'t',header:'T',field:'tags',getLabel:()=>''}),/together|mapping/);
 assert.throws(()=>normalization.normalizeDataGridColumns(f,{columns:f.defineColumns(c=>[c.text({id:'same',header:'S',field:'x'}),c.number({id:'same',header:'S',field:'y'})])}),/duplicate/);
});
test('row IDs preserve original objects and reject finite/type/key collisions',()=>{
 const items=Object.freeze([Object.freeze({id:1,name:'A'}),Object.freeze({id:'2',name:'B'})]);
 const prepared=rows.prepareDataGridRows(items,r=>r.id,[]);
 assert.strictEqual(prepared[0].item,items[0]);assert.equal(prepared[0].rowId,1);
 for(const id of [NaN,Infinity,-Infinity,undefined,{},true])assert.throws(()=>rows.prepareDataGridRows([{id}],r=>r.id,[]),/row ID/);
 for(const ids of [[1,1],[1,'1']])assert.throws(()=>rows.prepareDataGridRows(ids.map(id=>({id})),r=>r.id,[]),/duplicate|collision/);
});
test('one pass caches undefined, shared accessor and classification without mutating inputs',()=>{
 let reads=0,classifies=0;const accessor=item=>{reads++;return item.value;};
 const f=family();const columns=normalization.normalizeDataGridColumns(f,{columns:f.defineColumns(c=>[c.text({id:'a',header:'A',accessor}),c.text({id:'b',header:'B',accessor})])});
 const items=Object.freeze([Object.freeze({id:1,value:undefined})]);
 const classify=(column,raw,item)=>{classifies++;return {columnId:column.id,raw,item};};
 const pass=rows.prepareDataGridRows(items,r=>r.id,columns,classify);
 assert.equal(reads,1);assert.equal(classifies,2);
 assert.strictEqual(pass[0].getValue(columns[0]),pass[0].getValue(columns[0]));
 assert.equal(reads,1);assert.equal(classifies,2);
 rows.prepareDataGridRows(items,r=>r.id,columns,classify);assert.equal(reads,2);
 assert.throws(()=>rows.prepareDataGridRows([{id:1}],r=>r.id,normalization.normalizeDataGridColumns(f,{columns:f.defineColumns(c=>[c.custom({id:'x',header:'X',accessor:()=>{throw Error('observable');},renderCell:()=>''})])})),/observable/);
});
test('state validation ignores stale sort/hidden IDs and rejects activation-selection conflicts',()=>{
 const f=family();const columns=normalization.normalizeDataGridColumns(f,{columns:f.defineColumns(c=>[c.text({id:'a',header:'A',field:'name',sortable:true})])});
 const hidden=new Set(['old']);const sort={sortColumn:'old',sortDirection:'ascending'};
 assert.deepEqual(state.resolveVisibleDataGridColumns(columns,hidden),columns);
 assert.equal(state.resolveDataGridSortState(columns,sort),undefined);assert.equal(hidden.size,1);assert.equal(sort.sortColumn,'old');
 assert.throws(()=>state.validateDataGridInteraction({selectionMode:'multiselect',onRowActivate(){}}),/activation|selection/);
});
test('tags require an explicit mapping or paired callbacks; unmapped IDs stay invalid',()=>{
 const f=family();
 assert.throws(()=>f.defineColumns(c=>[c.tags({id:'t',header:'T',field:'tags'})]),/mapping/);
 const columns=normalization.normalizeDataGridColumns(f,{columns:f.defineColumns(c=>[c.tags({id:'t',header:'T',field:'tags',values:{a:{label:'A'}}})])});
 assert.equal(columns[0].presentTags(['unmapped'],{}),undefined);
 assert.deepEqual(columns[0].presentTags(['a'],{}).map(t=>[t.id,t.label]),[['a','A']]);
});
test('malformed native presentation objects and chart colors are rejected',()=>{
 const f=family();
 for(const [kind,extras]of [['link',{href:'x',linkProps:[]}],['person',{getName:()=>'',personaProps:false}],['progress',{progressBarProps:{value:3}}],['rating',{ratingDisplayProps:{max:3}}],['sparkline',{color:42}]]) {
  assert.throws(()=>f.defineColumns(c=>[c[kind]({id:'x',header:'X',field:'value',...extras})]),TypeError);
 }
});
test('captured typed tag semantics preserve scalar IDs and prefer mapped object labels',()=>{
 const f=family();let calls=0;
 const columns=normalization.normalizeDataGridColumns(f,{columns:f.defineColumns(c=>[c.tags({id:'t',header:'T',field:'tags',values:new Map([[1,{label:'Mapped'}]]),getTagId:tag=>{calls++;return tag.id;},getLabel:tag=>tag.label})])});
 const item={tags:[{id:1,label:'Explicit'}]};
 assert.equal(columns[0].presentTags(item.tags,item)[0].label,'Mapped');
 const before=calls;
 assert.equal(columns[0].presentTags([1],item)[0].id,1);
 assert.equal(calls,before);
 assert.equal(columns[0].presentTags([2],item),undefined);
});
test('capture snapshots maps and native options, retaining original callback data and JSX',()=>{
 const f=family();const element=React.createElement('strong',{},'custom');let received;
 const map=new Map([[1,{label:'Number'}],['1',{label:'Text'}]]);
 const config={id:'n',header:'N',field:'n',values:map};
 const descriptors=f.defineColumns(c=>[c.badge(config),c.custom({id:'c',header:'C',field:'person',renderCell:context=>{received=context;return element;}})]);
 const columns=normalization.normalizeDataGridColumns(f,{columns:descriptors});
 map.set(1,{label:'Changed'});config.field='changed';
 assert.equal(columns[0].mapValue(1).label,'Number');assert.equal(columns[0].mapValue('1').label,'Text');
 const item={id:1,n:1,person:{name:'A'}};
 assert.equal(columns[0].read(item),1);
 assert.strictEqual(columns[1].renderCell(item.person,item,1),element);
 assert.strictEqual(received.item,item);assert.strictEqual(received.value,item.person);assert.equal(received.rowId,1);
});
test('field getters and undefined classifiers cache per pass; fresh row/getRowId/column input invalidates',()=>{
 let reads=0,classifies=0;const item={id:1,get value(){reads++;return undefined;}};
 const f=family();const create=()=>normalization.normalizeDataGridColumns(f,{columns:f.defineColumns(c=>[c.text({id:'a',header:'A',field:'value'}),c.text({id:'b',header:'B',field:'value'})])});
 const columns=create();const classify=()=>{classifies++;return undefined;};
 const first=rows.prepareDataGridRows([item],r=>r.id,columns,classify);
 first[0].getRawValue(columns[0]);first[0].getValue(columns[1]);assert.equal(reads,1);assert.equal(classifies,2);
 const second=rows.prepareDataGridRows([item],()=>2,create(),classify);
 assert.equal(second[0].rowId,2);assert.equal(reads,2);assert.equal(classifies,4);
 const fresh={id:3,value:'fresh'};const third=rows.prepareDataGridRows([fresh],r=>r.id,columns);
 assert.equal(third[0].getValue(columns[0]),'fresh');
 assert.throws(()=>first[0].getValue(create()[0]),/this pass/);
 assert.throws(()=>rows.prepareDataGridRows([item],r=>r.id,columns,()=>{throw Error('classification exception');}),/classification exception/);
});
test('ReadonlyMap facades preserve exact string and numeric mapped identities',()=>{
 const f=family();const backing=new Map([[1,{label:'Numeric'}],['1',{label:'Text'}]]);
 const facade={get:key=>backing.get(key),entries:()=>backing.entries(),has:key=>backing.has(key),keys:()=>backing.keys(),values:()=>backing.values(),forEach:callback=>backing.forEach(callback),size:backing.size,[Symbol.iterator]:()=>backing[Symbol.iterator]()};
 const columns=normalization.normalizeDataGridColumns(f,{columns:f.defineColumns(c=>[c.badge({id:'b',header:'B',field:'value',values:facade})])});
 assert.equal(columns[0].mapValue(1).label,'Numeric');assert.equal(columns[0].mapValue('1').label,'Text');
});
test('paired tag callbacks support scalar inputs when no mapping is configured',()=>{
 const f=family();let ids=0,labels=0;const item={tags:['alpha','beta']};
 const columns=normalization.normalizeDataGridColumns(f,{columns:f.defineColumns(c=>[c.tags({id:'t',header:'T',field:'tags',getTagId:tag=>{ids++;return `id:${tag}`;},getLabel:tag=>{labels++;return tag.toUpperCase();}})])});
 assert.deepEqual(columns[0].presentTags(item.tags,item).map(tag=>[tag.id,tag.label]),[['id:alpha','ALPHA'],['id:beta','BETA']]);
 assert.equal(ids,2);assert.equal(labels,2);
});

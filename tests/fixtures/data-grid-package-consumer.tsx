import * as React from 'react';
import { DataGrid, createDataGrid } from '@apvee/spfx-react-toolkit';
import { DataGrid as DomainGrid, createDataGrid as DomainFactory } from '@apvee/spfx-react-toolkit/controls';
import { DataGrid as ControlGrid, createDataGrid as ControlFactory } from '@apvee/spfx-react-toolkit/controls/data-grid';
const aliases: typeof DataGrid[] = [DomainGrid, ControlGrid];
const factories: typeof createDataGrid[] = [DomainFactory, ControlFactory];
void [aliases, factories];
type Person = { name:string; key:number };
type Row = { id:number; text:string; count:number; active:boolean; date:Date|string; status:'open'|'closed'; person:Person; people:readonly Person[]; tags:readonly ('a'|'b')[]; series:readonly number[] };
const Grid=createDataGrid<Row>();
const rows:readonly Row[]=[];
// @ts-expect-error Stable identity is mandatory; no index fallback is exposed.
const missingRowId = <Grid items={rows} columns={[]} />;
// @ts-expect-error The ref belongs to the actual native grid div.
const wrongRef = <Grid items={rows} getRowId={row => row.id} columns={[]} ref={React.createRef<HTMLSpanElement>()} />;
void [missingRowId, wrongRef];
const columns=Grid.defineColumns(c=>[
 c.text({id:'text',header:'Text',field:'text'}), c.number({id:'n',header:'N',field:'count'}),
 c.boolean({id:'b',header:'B',field:'active'}),c.date({id:'d',header:'D',field:'date'}),c.dateTime({id:'dt',header:'DT',field:'date'}),
 c.badge({id:'badge',header:'Badge',field:'status',values:{open:{label:'Open'},closed:{label:'Closed'}},tooltip:({value})=>value}),
 c.icon({id:'icon',header:'Icon',accessor:r=>r.status,values:{open:{icon:<span/>,label:'Open'}}}),
 c.link({id:'link',header:'Link',field:'text',href:(row,value)=>row.text+value,onNavigate:(event,data)=>{event.currentTarget.href;data.item.id;data.value.toUpperCase();}}),
 c.person({id:'p',header:'Person',field:'person',getName:person=>person.name}),
 c.people({id:'pp',header:'People',accessor:r=>r.people,getName:p=>p.name,getPersonId:p=>p.key}),
 c.tags({id:'tags',header:'Tags',field:'tags',values:{a:{label:'A'}}}),
 c.progress({id:'progress',header:'Progress',field:'count'}),c.rating({id:'rating',header:'Rating',field:'count'}),
 c.sparkline({id:'sparkline',header:'Series',field:'series',color:({item,summary})=>item.status+summary.last}),
 c.actions({id:'actions',actions:[{id:'edit',label:'Edit',placement:'inline',onClick:(event,data)=>{event.currentTarget.id;data.item.id;}}]}),
 c.custom({id:'custom',header:'Custom',field:'person',renderCell:({value,item,rowId})=><span>{value?.name}{item.id}{rowId}</span>}),
]);
const valid=<Grid items={rows} getRowId={r=>r.id} columns={columns} ref={React.createRef<HTMLDivElement>()} onClick={e=>{e.currentTarget.tabIndex;}}/>;
const jsx=<Grid items={rows} getRowId={r=>r.id}><Grid.PersonColumn id="p" header="Person" field="person" getName={p=>p.name}/></Grid>;
const empty=<Grid items={rows} getRowId={r=>r.id} columns={[]}/>;
const dynamic=<DataGrid items={[{id:1}]} getRowId={r=>String(r.id)}><DataGrid.TextColumn id="t" header="T" field="anything" tooltip={({value})=>typeof value==='string'?value:undefined}/></DataGrid>;
// @ts-expect-error number field cannot read text
Grid.defineColumns(c=>[c.text({id:'bad',header:'Bad',field:'count'})]);
// @ts-expect-error accessor must supply numeric data
Grid.defineColumns(c=>[c.number({id:'bad',header:'Bad',accessor:r=>r.text})]);
// @ts-expect-error two sources
Grid.defineColumns(c=>[c.text({id:'bad',header:'Bad',field:'text',accessor:r=>r.text})]);
// @ts-expect-error source required
Grid.defineColumns(c=>[c.text({id:'bad',header:'Bad'})]);
// @ts-expect-error dictionary literal domain remains narrow
Grid.defineColumns(c=>[c.badge({id:'bad',header:'Bad',field:'status',values:{other:{label:'Other'}}})]);
// @ts-expect-error wrong inferred person callback
Grid.defineColumns(c=>[c.person({id:'bad',header:'Bad',field:'person',getName:p=>p.missing})]);
// @ts-expect-error actions have no source
Grid.defineColumns(c=>[c.actions({id:'bad',field:'text',actions:[]})]);
// @ts-expect-error both collections invalid
const both=<Grid items={rows} getRowId={r=>r.id} columns={columns}><Grid.TextColumn id="t" header="T" field="text"/></Grid>;
// @ts-expect-error one declared collection required
const neither=<Grid items={rows} getRowId={r=>r.id}/>;
// @ts-expect-error no query engine
const query=<Grid items={rows} getRowId={r=>r.id} columns={columns} query={{}}/>;
// @ts-expect-error raw HTML is owned by grid
const html=<Grid items={rows} getRowId={r=>r.id} columns={columns} dangerouslySetInnerHTML={{__html:'x'}}/>;
// @ts-expect-error no View API
Grid.View;
// @ts-expect-error dynamic callback value is unknown
const weak=<DataGrid.TextColumn id="t" header="T" field="text" tooltip={({value})=>value.toUpperCase()}/>;
void [valid,jsx,empty,dynamic,both,neither,query,html,weak];
const native=<Grid items={rows} getRowId={r=>r.id} columns={columns} onKeyDown={event=>event.currentTarget.tabIndex} onFocus={event=>event.currentTarget.tabIndex} style={{width:'100%'}} size="small"/>;
type NullableRow={id:number; text:string|null; person:Person|undefined};
const Nullable=createDataGrid<NullableRow>();
Nullable.defineColumns(c=>[c.text({id:'nullable',header:'Nullable',accessor:row=>row.text,tooltip:({value})=>value??undefined}),c.person({id:'p',header:'P',field:'person',getName:p=>p.name})]);
type IndexedRow={id:string; [key:string]:string};
const Indexed=createDataGrid<IndexedRow>();
Indexed.defineColumns(c=>[c.text({id:'i',header:'I',field:'anything',tooltip:({value})=>value?.toUpperCase()})]);
Grid.defineColumns(c=>[c.custom({id:'a',header:'A',accessor:row=>row.person,renderCell:({value})=>value?.name})]);
// @ts-expect-error inferred custom callback has an object value
Grid.defineColumns(c=>[c.custom({id:'a',header:'A',accessor:row=>row.person,renderCell:({value})=>value?.toUpperCase()})]);
// @ts-expect-error mapped callback cannot narrow a declared wider source domain
Grid.defineColumns(c=>[c.badge({id:'a',header:'A',field:'status',values:{open:{label:'Open'}},tooltip:({value}:{value:'open'})=>value})]);
// @ts-expect-error numeric mapped keys require Map, preserving 1 versus string 1
Grid.defineColumns(c=>[c.badge({id:'a',header:'A',field:'count',values:{1:{label:'One'}}})]);
Grid.defineColumns(c=>[c.badge({id:'a',header:'A',field:'count',values:new Map([[1,{label:'One'}]])})]);
// @ts-expect-error callback type stays unknown in dynamic mode
const dynamicPerson=<DataGrid.PersonColumn id="p" header="P" field="person" getName={person=>person.name}/>;
void [native,dynamicPerson];
// @ts-expect-error tags require a mapping or paired identity/label callbacks
Grid.defineColumns(c=>[c.tags({id:'t',header:'T',field:'tags'})]);
// @ts-expect-error callbacks must be provided together even with a mapping
Grid.defineColumns(c=>[c.tags({id:'t',header:'T',field:'tags',values:{a:{label:'A'}},getLabel:tag=>tag})]);
Grid.defineColumns(c=>[c.tags({id:'objects',header:'Object tags',accessor:row=>row.people,getTagId:p=>p.key,getLabel:p=>p.name,values:new Map([[1,{label:'Mapped person'}]])})]);
// @ts-expect-error object-tag mapping presentations keep their type
Grid.defineColumns(c=>[c.tags({id:'o',header:'O',field:'people',getTagId:p=>p.key,getLabel:p=>p.name,values:new Map([[1,'invalid presentation']])})]);
// @ts-expect-error owned navigation uses an anchor, not a button event target
Grid.defineColumns(c=>[c.link({id:'l',header:'L',field:'text',href:'/',linkProps:{onCopy:(event:React.ClipboardEvent<HTMLButtonElement>)=>event.currentTarget.disabled}})]);
// @ts-expect-error the grid owns the anchor element for link navigation
Grid.defineColumns(c=>[c.link({id:'l',header:'L',field:'text',href:'/',linkProps:{as:'button'}})]);
// @ts-expect-error native grid root and ref are always a div
const tableRoot=<Grid items={rows} getRowId={r=>r.id} columns={columns} as="table"/>;
void tableRoot;

const nativeSizing = <Grid items={rows} getRowId={row => row.id} columns={columns} columnSizingOptions={{ text: { minWidth: 90, defaultWidth: 180 } }} containerWidthOffset={100} />;
// @ts-expect-error Sortability belongs to individual toolkit columns.
const ownedSortability = <Grid items={rows} getRowId={row => row.id} columns={columns} sortable />;
void [nativeSizing, ownedSortability];

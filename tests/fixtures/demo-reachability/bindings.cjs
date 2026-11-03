const app = 'apps/spfx-react-toolkit-test/';
const { [app + 'src/catalogue.ts']: catalogue } = require('./files.cjs');
function scenario(code) { return { [app + 'src/scenario.ts']: undefined, [app + 'src/scenario.tsx']: code, [app + 'src/catalogue.ts']: catalogue.replace('() => number', '() => unknown') }; }
function form(extra = '', lookalike = false) {
  return scenario(`import * as React from 'react';
${lookalike ? `function createForm<T>() { return { useForm: (_options: {defaultValues: T}) => ({}), Root: (props: {methods: object; onSubmit: () => void; children?: React.ReactNode}) => <div>{props.children}</div>, Input: (props: {name: keyof T}) => <span>{String(props.name)}</span> }; }` : "import { createForm as makeForm } from '@apvee/spfx-react-toolkit/controls/form';"}
interface Fields { live: string; ${extra ? 'orphan: string;' : ''} }
const defaults: Fields = { live: 'kept'${extra} };
const Editor = ${lookalike ? 'createForm' : 'makeForm'}<Fields>();
export function Demo(): React.ReactElement { const methods = Editor.useForm({defaultValues: defaults}); return <Editor.Root methods={methods} onSubmit={() => undefined}><Editor.Input name="live" /></Editor.Root>; }
`);
}
function grid({ extra = '', mapExtra = '', presentationExtra = '', invalid = false, descriptor = false, lookalike = false } = {}) {
  return scenario(`import * as React from 'react';
${lookalike ? `function DataGrid(props: {items: readonly Record<string, unknown>[]; getRowId: (row: Record<string, unknown>) => string; children?: React.ReactNode}) { return <div>{props.children}</div>; } DataGrid.TextColumn = (props: {id: string; field: string; header: string}) => <span>{props.field}</span>;` : "import { DataGrid, createDataGrid as createGrid } from '@apvee/spfx-react-toolkit/controls/data-grid';"}
${descriptor ? "interface Row { id: string; live: string; orphan: string; }\nconst Grid = createGrid<Row>();\nconst columns = Grid.defineColumns(c => [c.text({id:'live',header:'Live',field:'live'})]);" : ''}
const initialRows: readonly ${descriptor ? 'Row' : 'Record<string, unknown>'}[] = [{ id: 'a', live: ${invalid ? "{ invalid: 'object classifier input' }" : "'ready'"}${extra} }];
const getRowId = (row: ${descriptor ? 'Row' : 'Record<string, unknown>'}): string => String(row.id);
${!lookalike && !descriptor ? `const Grid = createGrid<{id: string; status: 'ready' | 'pending'}>();\nconst presentations = {ready: {label:'Ready'${presentationExtra}}, pending: {label:'Pending'}${mapExtra}};` : ''}
export function Demo(): React.ReactElement { const [rows] = React.useState(initialRows); return ${descriptor ? '<Grid items={rows} getRowId={getRowId} columns={columns} />' : `<><DataGrid items={rows} getRowId={getRowId}><DataGrid.TextColumn id="live" header="Live" field="live" /></DataGrid>${lookalike ? '' : '<Grid items={[{id:"mapped", status:"ready"}]} getRowId={row => row.id}><Grid.BadgeColumn id="status" header="Status" field="status" values={presentations} /></Grid>'}</>`}; }
`);
}
function react(extra = false, lookalike = false) {
  return scenario(`import * as React from 'react';
${lookalike ? "class Local<P,S> { public props: P; public state: S; constructor(props:P,state:S){this.props=props;this.state=state;} }" : ''}
class Boundary extends ${lookalike ? "Local<{}, {error?:Error}>" : "React.Component<{}, {error?:Error}>"} {
 public state: {error?:Error} = {};
 ${lookalike ? 'constructor(){super({},{});}' : ''}
 public static getDerivedStateFromError(error: Error): {error: Error${extra ? '; orphan: number' : ''}} { return {error${extra ? ', orphan: 1' : ''}}; }
 public render(): React.ReactNode { return this.state.error?.message; }
}
export function Demo(): React.ReactElement { ${lookalike ? 'return <p>{new Boundary().render()}</p>;' : 'return <Boundary />;'} }
`);
}
module.exports = { form, grid, react };

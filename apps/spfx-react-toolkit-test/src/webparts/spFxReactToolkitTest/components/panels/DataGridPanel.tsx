import * as React from 'react';
import { Button, Checkbox, Field, Select, FluentProvider, makeStyles } from '@fluentui/react-components';
import { DemoCard, useDemoStyles } from '../shared';
import { CheckmarkCircle20Regular } from '@fluentui/react-icons';
import { DataGrid, createDataGrid, useSx, flex, gap, width } from '@apvee/spfx-react-toolkit';
import type { DataGridProps } from '@apvee/spfx-react-toolkit';

interface DemoPerson { readonly id: string; readonly name: string; }
interface DemoRow {
  readonly id: string;
  readonly title: string;
  readonly total: number;
  readonly enabled: boolean;
  readonly day: string;
  readonly instant: string;
  readonly status: 'ready' | 'pending';
  readonly owner: DemoPerson;
  readonly people: readonly DemoPerson[];
  readonly tags: readonly ('design' | 'review' | 'release')[];
  readonly progress: number;
  readonly rating: number;
  readonly trend: readonly number[];
}

// A stable family outside rendering preserves descriptor and marker ownership.
const Grid = createDataGrid<DemoRow>();
const statusValues = { ready: { label: 'Ready', color: 'success' as const }, pending: { label: 'Pending', color: 'warning' as const } };
const iconValues = {
  ready: { label: 'Ready status', icon: <CheckmarkCircle20Regular /> },
  pending: { label: 'Pending status', icon: <CheckmarkCircle20Regular /> },
};
const tagValues = { design: { label: 'Design' }, review: { label: 'Review' }, release: { label: 'Release' } };
const initialRows: readonly DemoRow[] = [
  { id: 'a', title: 'Quarterly report', total: 12.5, enabled: true, day: '2024-02-29', instant: '2024-02-29T10:30:00Z', status: 'ready', owner: { id: 'alice', name: 'Alice Example' }, people: [{ id: 'alice', name: 'Alice Example' }, { id: 'bob', name: 'Bob Example' }, { id: 'carol', name: 'Carol Example' }], tags: ['design', 'review', 'release'], progress: 75, rating: 4, trend: [2, 4, 3, 7] },
  { id: 'b', title: 'Operations plan', total: 0, enabled: false, day: '0007-01-02', instant: '2024-03-01T09:00:00+01:00', status: 'pending', owner: { id: 'bob', name: 'Bob Example' }, people: [{ id: 'bob', name: 'Bob Example' }], tags: ['review'], progress: 0, rating: 2, trend: [-4, -2, -3, -1] },
];
const descriptorColumns = Grid.defineColumns(c => [
  c.text({ id: 'title', header: 'Title', field: 'title', sortable: true, width: 180 }),
  c.number({ id: 'total', header: 'Total', field: 'total', width: 120 }),
  c.boolean({ id: 'enabled', header: 'Enabled', field: 'enabled', width: 100 }),
  c.date({ id: 'day', header: 'Date', field: 'day', width: 140 }),
]);
const dynamicRows: readonly Record<string, unknown>[] = [
  { id: 'valid', text: '', number: 0, enabled: false },
  { id: 'missing' },
  { id: 'invalid', text: { unsafe: 'object' }, number: '42', enabled: 'yes' },
];
const useGridStyles = makeStyles({ frame: { maxInlineSize: '100%', minInlineSize: 0 } });

const CustomCell: React.FC<{ rowId: string }> = ({ rowId }) => {
  const [clicks, setClicks] = React.useState(0);
  return <Button data-custom-row={rowId} onClick={() => setClicks(value => value + 1)}>Custom {clicks}</Button>;
};

const DataGridPanel: React.FC = () => {
  const sx = useSx();
  const styles = useDemoStyles();
  const gridStyles = useGridStyles();
  const [dir, setDir] = React.useState<'host' | 'ltr' | 'rtl'>('host');
  const [size, setSize] = React.useState(640);
  const [api, setApi] = React.useState<'typed' | 'descriptors' | 'dynamic'>('typed');
  const [preset, setPreset] = React.useState<'overview' | 'dates' | 'team' | 'metrics' | 'custom'>('overview');
  const [gridSize, setGridSize] = React.useState<'small' | 'medium' | 'extra-small'>('medium');
  const [selectionMode, setSelectionMode] = React.useState<'multiselect' | 'single' | undefined>('multiselect');
  const [rows, setRows] = React.useState(initialRows);
  const [currentDynamicRows, setDynamicRows] = React.useState(dynamicRows);
  const [selected, setSelected] = React.useState<ReadonlySet<string | number>>(() => new Set(['off-page']));
  const [sort, setSort] = React.useState<DataGridProps<DemoRow>['sortState']>();
  const [page, setPage] = React.useState(false);
  const [hidden, setHidden] = React.useState(false);
  const [customStrings, setCustomStrings] = React.useState(false);
  const gridStrings = customStrings ? { selectAllRows: 'Select every document', selectRow: (id: string | number) => `Select document ${id}`, empty: 'No team documents', loading: 'Fetching documents' } : undefined;
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState(false);
  const [empty, setEmpty] = React.useState(false);
  const [emptyState, setEmptyState] = React.useState<'empty' | 'noResults'>('empty');
  const [resizable, setResizable] = React.useState(false);
  const [alternateColor, setAlternateColor] = React.useState(false);
  const [outcome, setOutcome] = React.useState('Interact with the grid to observe selection, sorting and actions.');
  const [action, setAction] = React.useState('none');
  const [navigation, setNavigation] = React.useState('none');
  const [activation, setActivation] = React.useState('none');
  const [retries, setRetries] = React.useState(0);
  const [observedRef, setObservedRef] = React.useState('not inspected');
  const [slotClicks, setSlotClicks] = React.useState(0);
  const [slotRefs, setSlotRefs] = React.useState('not inspected');
  const [suppressSlots, setSuppressSlots] = React.useState(false);
  const [customStatusSlots, setCustomStatusSlots] = React.useState(false);
  const headerSlotRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
  const errorSlotRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const placeholderSlotRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
  const gridRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const visibleRows = empty ? [] : page ? rows.filter(row => row.id === 'a') : rows;
  const visibleDynamicRows = empty ? [] : page ? currentDynamicRows.filter(row => row.id === 'valid') : currentDynamicRows;
  const activeRowCount = api === 'dynamic' ? visibleDynamicRows.length : visibleRows.length;
  const hiddenIds = hidden ? new Set(['text', 'number', 'title', 'total', 'enabled', 'day', 'instant', 'status', 'icon', 'link', 'owner', 'people', 'tags', 'progress', 'rating', 'trend', 'actions', 'custom']) : undefined;
  const applySort = (): void => {
    if (sort?.sortColumn !== 'title') return;
    const direction = sort.sortDirection === 'ascending' ? 1 : -1;
    setRows(previous => [...previous].sort((left, right) => direction * left.title.localeCompare(right.title)));
  };
  return <FluentProvider dir={dir === 'host' ? undefined : dir}>
    <section id="data-grid-demo" className={sx(flex.column, gap.medium)}>
      <DemoCard title="DataGrid" iconName="Table" description="Explore team documents and choose how to present the columns."
        controls={<>
          <Field label="API"><Select data-demo-variant="API" id="dg-api" value={api} onChange={(_, data) => { if (data.value === 'typed' || data.value === 'descriptors' || data.value === 'dynamic') setApi(data.value); }}><option value="typed">Typed columns</option><option value="descriptors">Descriptors</option><option value="dynamic">Dynamic data</option></Select></Field>
          <Field label="Columns"><Select data-demo-variant="Columns" id="dg-preset" value={preset} disabled={api !== 'typed'} onChange={(_, data) => { if (data.value === 'overview' || data.value === 'dates' || data.value === 'team' || data.value === 'metrics' || data.value === 'custom') setPreset(data.value); }}><option value="overview">Documents</option><option value="dates">Dates and status</option><option value="team">People and tags</option><option value="metrics">Progress</option><option value="custom">Custom cells</option></Select></Field>
          <Field label="Width"><Select id="dg-width" value={size} onChange={event => setSize(Number(event.target.value))}>{[320, 640, 1100].map(value => <option key={value} value={value}>{value}px</option>)}</Select></Field>
          <Field label="size"><Select id="dg-size" value={gridSize} onChange={event => { const value = event.target.value; if (value === 'small' || value === 'medium' || value === 'extra-small') setGridSize(value); }}><option value="extra-small">Extra small</option><option value="small">Small</option><option value="medium">Medium</option></Select></Field>
          <Field label="selectionMode"><Select id="dg-selection-mode" value={selectionMode ?? 'none'} onChange={event => setSelectionMode(event.target.value === 'single' ? 'single' : event.target.value === 'multiselect' ? 'multiselect' : undefined)}><option value="none">None</option><option value="single">Single</option><option value="multiselect">Multiselect</option></Select></Field>
        </>} commands={<>
          <div className={styles.actions} role="group" aria-label="Consumer simulation commands">
            {api !== 'dynamic' && <Button id="dg-apply-sort" onClick={applySort}>Apply sort intent</Button>}
            <Button id="dg-reverse" onClick={() => { if (api === 'dynamic') setDynamicRows(previous => [...previous].reverse()); else setRows(previous => [...previous].reverse()); }}>Reverse supplied rows</Button>
            {api === 'typed' && preset === 'metrics' && <Button id="dg-update-sparkline" onClick={() => setRows(previous => previous.map(row => ({ ...row, trend: row.trend.map((value, index) => value + index * 3 + 1) })))}>Update chart observations</Button>}
          </div>
        </>} status={<>
          <p role="status">{outcome} Rows: {activeRowCount}; selected: {selected.size}; retries: <output id="dg-retries">{retries}</output></p>
        </>} diagnostics={<>
          <Field label="Direction"><Select id="dg-dir" value={dir} onChange={event => setDir(event.target.value === 'rtl' ? 'rtl' : event.target.value === 'ltr' ? 'ltr' : 'host')}><option value="host">Host</option><option value="ltr">LTR</option><option value="rtl">RTL</option></Select></Field>
          <Checkbox id="dg-page" label="Supply first row only" checked={page} onChange={(_event, data) => setPage(data.checked === true)} />
          <Checkbox id="dg-hidden" label="Hide all columns" checked={hidden} onChange={(_event, data) => setHidden(data.checked === true)} />
          <Checkbox id="dg-strings" label="Custom UI strings (number/date locale stays en-US)" checked={customStrings} onChange={(_event, data) => setCustomStrings(data.checked === true)} />
          <Checkbox id="dg-loading" label="Loading" checked={loading} onChange={(_event, data) => setLoading(data.checked === true)} />
          <Checkbox id="dg-error" label="Error" checked={error} onChange={(_event, data) => setError(data.checked === true)} />
          <Checkbox id="dg-empty" label="Empty rows" checked={empty} onChange={(_event, data) => setEmpty(data.checked === true)} />
          <Field label="Empty presentation"><Select id="dg-empty-state" value={emptyState} onChange={event => setEmptyState(event.target.value === 'noResults' ? 'noResults' : 'empty')}><option value="empty">Empty</option><option value="noResults">No results</option></Select></Field>
          <Checkbox id="dg-resizable" label="Native resizing" checked={resizable} onChange={(_event, data) => setResizable(data.checked === true)} />
          {api === 'typed' && preset === 'metrics' && <Checkbox id="dg-color" label="Alternate chart color" checked={alternateColor} onChange={(_event, data) => setAlternateColor(data.checked === true)} />}
          <Checkbox id="dg-custom-status-slots" label="Custom empty and no-results slots" checked={customStatusSlots} onChange={(_event, data) => setCustomStatusSlots(data.checked === true)} />
          <Checkbox id="dg-slot-suppress" label="Suppress status and placeholder text" checked={suppressSlots} onChange={(_event, data) => setSuppressSlots(data.checked === true)} />

            <Button id="dg-inspect-slot-refs" onClick={() => setSlotRefs(`header=${headerSlotRef.current?.tagName ?? 'absent'} error=${errorSlotRef.current?.tagName ?? 'absent'} missing=${placeholderSlotRef.current?.tagName ?? 'absent'}`)}>Inspect inner slot refs</Button>
            <Button id="dg-inspect-ref" onClick={() => setObservedRef(gridRef.current?.tagName.toLowerCase() ?? 'absent')}>Inspect native grid ref</Button>
          <div>
            <p>Inner slot refs: <output id="dg-slot-refs">{slotRefs}</output></p>
            <p>Header slot callback clicks: <output id="dg-slot-clicks">{slotClicks}</output></p>
            <p>Selected IDs: <output id="dg-selection">{JSON.stringify(Array.from(selected))}</output></p>
            <p>Sort intent: <output id="dg-sort">{sort ? JSON.stringify({ columnId: sort.sortColumn, direction: sort.sortDirection }) : 'none'}</output></p>
            <p>Action: <output id="dg-action">{action}</output></p>
            <p>Navigation: <output id="dg-navigation">{navigation}</output></p>
            <p>Activation: <output id="dg-activation-output">{activation}</output></p>
            <p>Native ref: <output id="dg-ref">{observedRef}</output></p>
          </div>
        </>}>
        <div className={styles.scroll} role="region" aria-label="DataGrid preview scroll" tabIndex={0}>
          <div id="dg-frame" className={sx(flex.column, gap.large, width.px(size), gridStyles.frame)}>
            {api === 'dynamic' ? <DataGrid id="data-grid-main" strings={gridStrings} aria-label="Dynamic data" ref={gridRef} items={visibleDynamicRows} getRowId={row => String(row.id)} size={gridSize}
              selectionMode={selectionMode} selectedItems={selected} onSelectionChange={(_event, data) => { setSelected(new Set(data.selectedItems)); setOutcome(`Selection changed: ${data.selectedItems.size} selected`); }}
              hiddenColumnIds={hiddenIds} resizableColumns={resizable} loading={loading} emptyState={emptyState}
              error={error ? { ref: errorSlotRef, children: <p>Consumer-owned loading error</p> } : undefined}
              emptyContent={suppressSlots ? false : customStatusSlots ? { children: 'Consumer-owned empty slot' } : undefined}
              noResultsContent={suppressSlots ? false : customStatusSlots ? { children: (Component, nativeProps) => <Component {...nativeProps}>Consumer-owned no-results slot</Component> } : undefined}
              onRowActivate={selectionMode === undefined ? (_event, data) => setActivation(JSON.stringify({ rowId: data.rowId })) : undefined}
              onRetry={() => { setRetries(value => value + 1); setError(false); }}>
              <DataGrid.TextColumn id="text" header="Title" field="text" placeholder={suppressSlots ? false : { ref: placeholderSlotRef, children: 'Missing text' }} invalidPlaceholder="Invalid text" />
              <DataGrid.NumberColumn id="number" header="Total" field="number" placeholder={suppressSlots ? false : 'Missing number'} invalidPlaceholder="Invalid number" />
              <DataGrid.BooleanColumn id="enabled" header="Enabled" field="enabled" placeholder={suppressSlots ? false : undefined} />
              <DataGrid.TextColumn id="status" header="Status" field="status" placeholder={suppressSlots ? false : 'Missing status'} />
            </DataGrid> : api === 'descriptors' ? <Grid id="data-grid-main" strings={gridStrings} ref={gridRef} aria-label="Team documents" items={visibleRows} getRowId={row => row.id} locale="en-US"
              size={gridSize} selectionMode={selectionMode} selectedItems={selected} onSelectionChange={(_event, data) => { setSelected(new Set(data.selectedItems)); setOutcome(`Selection changed: ${data.selectedItems.size} selected`); }}
              sortState={sort} onSortChange={(_event, next) => { setSort(next); setOutcome(`Sort requested: ${next.sortColumn}, ${next.sortDirection}`); }} resizableColumns={resizable} hiddenColumnIds={hiddenIds}
              loading={loading} error={error ? { ref: errorSlotRef, children: <p>Consumer-owned loading error</p> } : undefined} emptyState={emptyState}
              emptyContent={suppressSlots ? false : customStatusSlots ? { children: 'Consumer-owned empty slot' } : undefined}
              noResultsContent={suppressSlots ? false : customStatusSlots ? { children: (Component, nativeProps) => <Component {...nativeProps}>Consumer-owned no-results slot</Component> } : undefined}
              onRowActivate={selectionMode === undefined ? (_event, data) => setActivation(JSON.stringify({ rowId: data.rowId })) : undefined} onRetry={() => { setRetries(value => value + 1); setError(false); }} columns={descriptorColumns} /> : <Grid id="data-grid-main" strings={gridStrings} ref={gridRef} aria-label="Team documents" items={visibleRows} getRowId={row => row.id} locale="en-US"
              size={gridSize} selectionMode={selectionMode} selectedItems={selected} onSelectionChange={(_event, data) => { setSelected(new Set(data.selectedItems)); setOutcome(`Selection changed: ${data.selectedItems.size} selected`); }}
              sortState={sort} onSortChange={(_event, next) => { setSort(next); setOutcome(`Sort requested: ${next.sortColumn}, ${next.sortDirection}`); }} resizableColumns={resizable} hiddenColumnIds={hiddenIds}
              loading={loading} error={error ? { ref: errorSlotRef, children: <p>Consumer-owned loading error</p> } : undefined} emptyState={emptyState}
              emptyContent={suppressSlots ? false : customStatusSlots ? { children: 'Consumer-owned empty slot' } : undefined}
              noResultsContent={suppressSlots ? false : customStatusSlots ? { children: (Component, nativeProps) => <Component {...nativeProps}>Consumer-owned no-results slot</Component> } : undefined}
              onRowActivate={selectionMode === undefined ? (_event, data) => setActivation(JSON.stringify({ rowId: data.rowId })) : undefined} onRetry={() => { setRetries(value => value + 1); setError(false); }}>
              {api === 'typed' && ["overview","dates","team","metrics","custom"].includes(preset) && <Grid.TextColumn id="title" header={{ ref: headerSlotRef, onClick: () => setSlotClicks(value => value + 1), children: (Component, nativeProps) => <Component {...nativeProps} data-grid-slot-callback="header">Title</Component> }} headerAriaLabel="Title" field="title" sortable width={160} minWidth={100} />}
              {api === 'typed' && ["overview"].includes(preset) && <Grid.NumberColumn id="total" header="Total" field="total" format={{ minimumFractionDigits: 2 }} width={120} />}
              {api === 'typed' && ["dates","custom"].includes(preset) && <Grid.BooleanColumn id="enabled" header="Enabled" field="enabled" display="icon" trueLabel="Enabled yes" falseLabel="Enabled no" width={110} />}
              {api === 'typed' && ["dates"].includes(preset) && <Grid.DateColumn id="day" header="Calendar day" field="day" width={150} />}
              {api === 'typed' && ["dates"].includes(preset) && <Grid.DateTimeColumn id="instant" header="Instant UTC" field="instant" width={180} />}
              {api === 'typed' && ["overview"].includes(preset) && <Grid.BadgeColumn id="status" header="Status" field="status" values={statusValues} width={120} />}
              {api === 'typed' && ["custom"].includes(preset) && <Grid.IconColumn id="icon" header="Status icon" field="status" values={iconValues} width={110} />}
              {api === 'typed' && ["team"].includes(preset) && <Grid.LinkColumn id="link" header="Details" field="title" href={(row) => `#details-${row.id}`} onNavigate={(event, data) => { event.preventDefault(); setNavigation(JSON.stringify({ rowId: data.rowId, value: data.value, href: data.href })); }} width={140} />}
              {api === 'typed' && ["overview","team"].includes(preset) && <Grid.PersonColumn id="owner" header="Owner" field="owner" getName={person => person.name} width={200} />}
              {api === 'typed' && ["team"].includes(preset) && <Grid.PeopleColumn id="people" header="People" field="people" getName={person => person.name} getPersonId={person => person.id} maxVisible={1} width={160} />}
              {api === 'typed' && ["team"].includes(preset) && <Grid.TagsColumn id="tags" header="Tags" field="tags" values={tagValues} maxVisible={1} width={160} />}
              {api === 'typed' && ["metrics"].includes(preset) && <Grid.ProgressColumn id="progress" header="Progress" field="progress" max={100} format={{ style: 'percent' }} showValue width={150} />}
              {api === 'typed' && ["metrics"].includes(preset) && <Grid.RatingColumn id="rating" header="Rating" field="rating" max={5} width={180} />}
              {api === 'typed' && ["metrics"].includes(preset) && <Grid.SparklineColumn id="trend" header="Trend" field="trend" showValue color={alternateColor ? '#d13438' : '#107c10'} width={160} />}
              {api === 'typed' && ["overview"].includes(preset) && <Grid.ActionsColumn id="actions" header="Actions" width={150} minWidth={120} actions={[
                { id: 'review', label: 'Open report', placement: 'inline', onClick: (_event, data) => { setAction(JSON.stringify({ rowId: data.rowId, actionId: data.actionId })); setOutcome(`${data.actionId}: ${data.rowId}`); } },
                { id: 'archive', label: 'Archive record', placement: 'menu', onClick: (_event, data) => { setAction(JSON.stringify({ rowId: data.rowId, actionId: data.actionId })); setOutcome(`${data.actionId}: ${data.rowId}`); } },
              ]} />}
              {api === 'typed' && ["custom"].includes(preset) && <Grid.Column id="custom" header="Custom" field="title" renderCell={({ item }) => <CustomCell rowId={item.id} />} width={130} />}
            </Grid>}


          </div></div>
      </DemoCard>
    </section>
  </FluentProvider>;
};

export default DataGridPanel;

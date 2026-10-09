import * as React from 'react';
import { Button, Checkbox, Select, FluentProvider, makeStyles, tokens, webDarkTheme, webLightTheme, teamsHighContrastTheme } from '@fluentui/react-components';
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
  { id: 'a', title: 'Alpha', total: 12.5, enabled: true, day: '2024-02-29', instant: '2024-02-29T10:30:00Z', status: 'ready', owner: { id: 'alice', name: 'Alice Example' }, people: [{ id: 'alice', name: 'Alice Example' }, { id: 'bob', name: 'Bob Example' }, { id: 'carol', name: 'Carol Example' }], tags: ['design', 'review', 'release'], progress: 75, rating: 4, trend: [2, 4, 3, 7] },
  { id: 'b', title: 'Beta', total: 0, enabled: false, day: '0007-01-02', instant: '2024-03-01T09:00:00+01:00', status: 'pending', owner: { id: 'bob', name: 'Bob Example' }, people: [{ id: 'bob', name: 'Bob Example' }], tags: ['review'], progress: 0, rating: 2, trend: [-4, -2, -3, -1] },
];
const descriptorColumns = Grid.defineColumns(c => [
  c.text({ id: 'title', header: 'Descriptor title', field: 'title', width: 180 }),
  c.number({ id: 'total', header: 'Descriptor total', field: 'total', width: 140 }),
]);
const dynamicRows: readonly Record<string, unknown>[] = [
  { id: 'valid', text: '', number: 0, enabled: false },
  { id: 'missing' },
  { id: 'invalid', text: { unsafe: 'object' }, number: '42', enabled: 'yes' },
];
const useStyles = makeStyles({
  frame: { maxInlineSize: '100%', minInlineSize: 0 },
  controls: { padding: tokens.spacingVerticalM, border: `1px solid ${tokens.colorNeutralStroke2}` },
});

const CustomCell: React.FC<{ rowId: string }> = ({ rowId }) => {
  const [clicks, setClicks] = React.useState(0);
  return <Button data-custom-row={rowId} onClick={() => setClicks(value => value + 1)}>Custom {clicks}</Button>;
};

const DataGridPanel: React.FC = () => {
  const sx = useSx();
  const styles = useStyles();
  const [theme, setTheme] = React.useState('light');
  const [dir, setDir] = React.useState<'ltr' | 'rtl'>('ltr');
  const [size, setSize] = React.useState(1100);
  const [rows, setRows] = React.useState(initialRows);
  const [selected, setSelected] = React.useState<ReadonlySet<string | number>>(() => new Set(['off-page']));
  const [sort, setSort] = React.useState<DataGridProps<DemoRow>['sortState']>();
  const [page, setPage] = React.useState(false);
  const [hidden, setHidden] = React.useState(false);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState(false);
  const [empty, setEmpty] = React.useState(false);
  const [emptyState, setEmptyState] = React.useState<'empty' | 'noResults'>('empty');
  const [resizable, setResizable] = React.useState(false);
  const [alternateColor, setAlternateColor] = React.useState(false);
  const [action, setAction] = React.useState('none');
  const [navigation, setNavigation] = React.useState('none');
  const [activation, setActivation] = React.useState('none');
  const [retries, setRetries] = React.useState(0);
  const [observedRef, setObservedRef] = React.useState('not inspected');
  const gridRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const visibleRows = empty ? [] : page ? rows.filter(row => row.id === 'a') : rows;
  const hiddenIds = hidden ? new Set(['title', 'total', 'enabled', 'day', 'instant', 'status', 'icon', 'link', 'owner', 'people', 'tags', 'progress', 'rating', 'trend', 'actions', 'custom']) : undefined;
  const applySort = (): void => {
    if (sort?.sortColumn !== 'title') return;
    const direction = sort.sortDirection === 'ascending' ? 1 : -1;
    setRows(previous => [...previous].sort((left, right) => direction * left.title.localeCompare(right.title)));
  };
  return <FluentProvider theme={theme === 'dark' ? webDarkTheme : theme === 'highContrast' ? teamsHighContrastTheme : webLightTheme} dir={dir}>
  <section id="data-grid-demo" className={sx(flex.column, gap.medium)}>
    <h3>DataGrid</h3>
    <p>All sixteen native cell presentations. The consumer owns rows, sorting intent and selection.</p>
    <div className={sx(styles.controls, flex.row, flex.wrap, gap.medium)}>
      <label>Theme <Select id="dg-theme" value={theme} onChange={event => setTheme(event.target.value)}><option value="light">Light</option><option value="dark">Dark</option><option value="highContrast">High contrast</option></Select></label>
      <label>Direction <Select id="dg-dir" value={dir} onChange={event => setDir(event.target.value === 'rtl' ? 'rtl' : 'ltr')}><option value="ltr">LTR</option><option value="rtl">RTL</option></Select></label>
      <label>Width <Select id="dg-width" value={size} onChange={event => setSize(Number(event.target.value))}>{[320, 640, 1100].map(value => <option key={value} value={value}>{value}px</option>)}</Select></label>
      <Checkbox id="dg-page" label="Supply Alpha only" checked={page} onChange={(_event, data) => setPage(data.checked === true)} />
      <Checkbox id="dg-hidden" label="Hide all columns" checked={hidden} onChange={(_event, data) => setHidden(data.checked === true)} />
      <Checkbox id="dg-loading" label="Loading" checked={loading} onChange={(_event, data) => setLoading(data.checked === true)} />
      <Checkbox id="dg-error" label="Error" checked={error} onChange={(_event, data) => setError(data.checked === true)} />
      <Checkbox id="dg-empty" label="Empty rows" checked={empty} onChange={(_event, data) => setEmpty(data.checked === true)} />
      <label>Empty presentation <Select id="dg-empty-state" value={emptyState} onChange={event => setEmptyState(event.target.value === 'noResults' ? 'noResults' : 'empty')}><option value="empty">Empty</option><option value="noResults">No results</option></Select></label>
      <Checkbox id="dg-resizable" label="Native resizing" checked={resizable} onChange={(_event, data) => setResizable(data.checked === true)} />
      <Checkbox id="dg-color" label="Alternate chart color" checked={alternateColor} onChange={(_event, data) => setAlternateColor(data.checked === true)} />
      <Button id="dg-apply-sort" onClick={applySort}>Apply sort intent</Button>
      <Button id="dg-reverse" onClick={() => setRows(previous => [...previous].reverse())}>Reverse supplied rows</Button>
      <Button id="dg-update-sparkline" onClick={() => setRows(previous => previous.map(row => ({ ...row, trend: row.trend.map((value, index) => value + index * 3 + 1) })))}>Update chart observations</Button>
      <Button id="dg-inspect-ref" onClick={() => setObservedRef(gridRef.current?.tagName.toLowerCase() ?? 'absent')}>Inspect native grid ref</Button>
    </div>
      <div id="dg-frame" className={sx(styles.frame, flex.column, gap.large, width.px(size))}>
        <Grid id="data-grid-main" ref={gridRef} aria-label="Sixteen column grid" items={visibleRows} getRowId={row => row.id} locale="en-US"
          selectionMode="multiselect" selectedItems={selected} onSelectionChange={(_event, data) => setSelected(new Set(data.selectedItems))}
          sortState={sort} onSortChange={(_event, next) => setSort(next)} resizableColumns={resizable} hiddenColumnIds={hiddenIds}
          loading={loading} error={error ? 'Consumer-owned loading error' : undefined} emptyState={emptyState}
          onRetry={() => { setRetries(value => value + 1); setError(false); }}>
          <Grid.TextColumn id="title" header="Title" field="title" sortable width={160} minWidth={100} />
          <Grid.NumberColumn id="total" header="Total" field="total" format={{ minimumFractionDigits: 2 }} width={120} />
          <Grid.BooleanColumn id="enabled" header="Enabled" field="enabled" display="icon" trueLabel="Enabled yes" falseLabel="Enabled no" width={110} />
          <Grid.DateColumn id="day" header="Calendar day" field="day" width={150} />
          <Grid.DateTimeColumn id="instant" header="Instant UTC" field="instant" width={180} />
          <Grid.BadgeColumn id="status" header="Status" field="status" values={statusValues} width={120} />
          <Grid.IconColumn id="icon" header="Status icon" field="status" values={iconValues} width={110} />
          <Grid.LinkColumn id="link" header="Details" field="title" href={(row) => `#details-${row.id}`} onNavigate={(event, data) => { event.preventDefault(); setNavigation(JSON.stringify({ rowId: data.rowId, value: data.value, href: data.href })); }} width={140} />
          <Grid.PersonColumn id="owner" header="Owner" field="owner" getName={person => person.name} width={200} />
          <Grid.PeopleColumn id="people" header="People" field="people" getName={person => person.name} getPersonId={person => person.id} maxVisible={1} width={160} />
          <Grid.TagsColumn id="tags" header="Tags" field="tags" values={tagValues} maxVisible={1} width={160} />
          <Grid.ProgressColumn id="progress" header="Progress" field="progress" max={100} format={{ style: 'percent' }} showValue width={150} />
          <Grid.RatingColumn id="rating" header="Rating" field="rating" max={5} width={180} />
          <Grid.SparklineColumn id="trend" header="Trend" field="trend" showValue color={alternateColor ? '#d13438' : '#107c10'} width={160} />
          <Grid.ActionsColumn id="actions" header="Actions" width={480} minWidth={400} actions={[
            { id: 'review', label: 'Review this localized record in detail', placement: 'inline', onClick: (_event, data) => setAction(JSON.stringify({ rowId: data.rowId, actionId: data.actionId })) },
            { id: 'disabled', label: 'Disabled action', icon: <CheckmarkCircle20Regular />, placement: 'inline', disabled: true, onClick: (_event, data) => setAction(JSON.stringify({ rowId: data.rowId, actionId: data.actionId })) },
            { id: 'archive', label: 'Archive record', placement: 'menu', onClick: (_event, data) => setAction(JSON.stringify({ rowId: data.rowId, actionId: data.actionId })) },
          ]} />
          <Grid.Column id="custom" header="Custom" field="title" renderCell={({ item }) => <CustomCell rowId={item.id} />} width={130} />
        </Grid>
        <h4>Immutable descriptor columns</h4>
        <Grid id="data-grid-descriptors" aria-label="Descriptor grid" items={initialRows} getRowId={row => row.id} columns={descriptorColumns} />
        <h4>Dynamic values: valid, missing and invalid</h4>
        <DataGrid id="data-grid-dynamic" aria-label="Dynamic grid" items={dynamicRows} getRowId={row => String(row.id)}>
          <DataGrid.TextColumn id="text" header="Dynamic text" field="text" placeholder="Missing text" invalidPlaceholder="Invalid text" />
          <DataGrid.NumberColumn id="number" header="Dynamic number" field="number" placeholder="Missing number" invalidPlaceholder="Invalid number" />
          <DataGrid.BooleanColumn id="enabled" header="Dynamic boolean" field="enabled" />
        </DataGrid>
        <h4>Row activation without selection</h4>
        <Grid id="data-grid-activation" aria-label="Activation grid" items={initialRows} getRowId={row => row.id}
          onRowActivate={(_event, data) => setActivation(JSON.stringify({ rowId: data.rowId }))}>
          <Grid.TextColumn id="title" header="Activate record" field="title" />
          <Grid.Column id="custom" header="Independent custom button" field="title" renderCell={({ item }) => <CustomCell rowId={item.id} />} />
        </Grid>
      </div>
    <div aria-live="polite">
      <p>Selected IDs: <output id="dg-selection">{JSON.stringify(Array.from(selected))}</output></p>
      <p>Sort intent: <output id="dg-sort">{sort ? JSON.stringify({ columnId: sort.sortColumn, direction: sort.sortDirection }) : 'none'}</output></p>
      <p>Action: <output id="dg-action">{action}</output></p>
      <p>Navigation: <output id="dg-navigation">{navigation}</output></p>
      <p>Activation: <output id="dg-activation-output">{activation}</output></p>
      <p>Retries: <output id="dg-retries">{retries}</output></p>
      <p>Native ref: <output id="dg-ref">{observedRef}</output></p>
    </div>
  </section>
  </FluentProvider>;
};

export default DataGridPanel;

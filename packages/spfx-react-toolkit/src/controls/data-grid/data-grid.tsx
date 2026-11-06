/**
 * Controlled row presentation with suite-owned typed columns and native Fluent grid interactions.
 *
 * @packageDocumentation
 * @docPage controls/data-grid
 */
import * as React from 'react';
import { renderContentSlot } from '../content-slot.internal';
import {
  createTableColumn,
  DataGrid as FluentDataGrid,
  DataGridBody,
  DataGridCell,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridRow,
  Tooltip,
  MessageBar,
  MessageBarBody,
  MessageBarActions,
  Button,
} from '@fluentui/react-components';
import type { TableColumnDefinition, TableColumnSizingOptions } from '@fluentui/react-components';
import { useStableCallback } from '../../hooks/useStableCallback';
import { useSx } from '../../hooks/useSx';
import { AutoSkeleton } from '../auto-skeleton/auto-skeleton';
import * as flex from '../../helpers/styles/flex';
import * as flexItem from '../../helpers/styles/flex-item';
import * as gap from '../../helpers/styles/gap';
import * as width from '../../helpers/styles/width';
import * as minWidth from '../../helpers/styles/min-width';
import * as paddingBlock from '../../helpers/styles/padding-block';
import * as textAlign from '../../helpers/styles/text-align';
import * as horizontalOverflow from '../../helpers/styles/overflow-horizontal';
import { createDynamicColumnFamily, getPublicDataGridColumnFamily } from './column-family.internal';
import type { FamilyResolution } from './column-family.internal';
import { normalizeDataGridColumns } from './column-normalization.internal';
import type { ResolvedDataGridColumn } from './column-normalization.internal';
import { prepareDataGridRows } from './prepared-rows.internal';
import type { PreparedDataGridRow } from './prepared-rows.internal';
import { classifyDataGridValue } from './data-grid-values.internal';
import type { DataGridCellRenderContext, DataGridCellValue } from './data-grid-cell.types';
import { getDataGridCellTooltip, isInteractiveDataGridColumn, renderDataGridCell } from './cells/data-grid-cells.internal';
import { resolveDataGridSortState, resolveDataGridStrings, resolveVisibleDataGridColumns, validateDataGridRootConfiguration } from './state-validation.internal';
import { useDataGridRuntimeStyles } from './data-grid.styles.internal';
import { DataGridLoadingCell } from './data-grid-loading-cell.internal';
import type { DataGridProps, DynamicDataGridSuite } from './data-grid.types';
import type { DataGridRowId } from './data-grid-column.types';

interface LoadingGridRow {
  readonly kind: 'loading';
  readonly rowId: string;
}
type GridRow<T extends object> =
  | { readonly kind: 'data'; readonly row: PreparedDataGridRow<T, DataGridCellValue> }
  | LoadingGridRow;
const emptySort: NonNullable<DataGridProps<object>['sortState']> = {
  sortColumn: undefined,
  sortDirection: 'ascending',
};
const interactiveSelector = [
  'a[href]', 'button', 'input', 'select', 'textarea', 'summary',
  '[contenteditable]:not([contenteditable="false"])',
  ...[
    'button', 'checkbox', 'combobox', 'link', 'listbox', 'menuitem',
    'menuitemcheckbox', 'menuitemradio', 'option', 'radio', 'scrollbar',
    'searchbox', 'slider', 'spinbutton', 'switch', 'tab', 'textbox',
    'treeitem', 'grid', 'tree',
  ].map(role => `[role="${role}"]`),
].join(',');
function canActivate(event: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>): boolean {
  if (event.defaultPrevented) return false;
  if (!(event.target instanceof Element)) return false;
  const interactive = event.target.closest(interactiveSelector);
  // The external DOM closest API uses null for absence. The owning grid contains
  // this row and must not suppress plain row/cell activation. Portal
  // controls do not contain their React owning row and remain independent actions.
  return interactive === null || interactive.contains(event.currentTarget);
}
function rowIdentity<T extends object>(row: GridRow<T>): DataGridRowId {
  return row.kind === 'data' ? row.row.rowId : row.rowId;
}
function neutralSortableCompare<T>(_left: T, _right: T): number {
  // Fluent 9.74 determines header sortability with compare.length > 0. This arity
  // bridge always returns equality, so its native body preserves supplied order.
  return 0;
}
function sizingOptions<T extends object>(columns: readonly ResolvedDataGridColumn<T>[], native: TableColumnSizingOptions | undefined): TableColumnSizingOptions {
  const options: TableColumnSizingOptions = {};
  for (const column of columns) {
    const { width: preferred, minWidth: minimum } = column.options;
    options[column.id] = {
      ...native?.[column.id],
      ...(preferred === undefined ? {} : { defaultWidth: preferred, idealWidth: preferred }),
      ...(minimum === undefined ? {} : { minWidth: minimum }),
    };
  }
  return options;
}
const ownedRootPropNames = [
  'items', 'getRowId', 'columns', 'children', 'loading', 'error', 'onRetry',
  'emptyState', 'emptyContent', 'noResultsContent', 'hiddenColumnIds', 'locale',
  'strings', 'onRowActivate', 'selectedItems', 'onSelectionChange', 'sortState',
  'onSortChange', 'className', 'style', 'selectionMode', 'resizableColumns',
  'resizableColumnsOptions', 'size', 'dangerouslySetInnerHTML', 'defaultSortState',
  'defaultSelectedItems', 'as', 'noNativeElements', 'sortable', 'columnSizingOptions',
] as const;
/**
 * Strip owned runtime paths from untyped callers before forwarding native attributes.
 */
function nativeRootProps<T extends object>(props: DataGridProps<T>): Omit<
  DataGridProps<T>, (typeof ownedRootPropNames)[number]
> {
  const result = { ...props };
  for (const key of ownedRootPropNames) Reflect.deleteProperty(result, key);
  return result;
}
function DataGridRuntime<T extends object>({ family, props, rootRef }: {
  family: FamilyResolution<T>; props: DataGridProps<T>; rootRef: React.ForwardedRef<HTMLDivElement>;
}): React.ReactElement {
  validateDataGridRootConfiguration(props);
  const sx = useSx(), styles = useDataGridRuntimeStyles();
  const declaredColumns = Object.prototype.hasOwnProperty.call(props, 'columns');
  const declaredChildren = Object.prototype.hasOwnProperty.call(props, 'children');
  const allColumns = React.useMemo(() => normalizeDataGridColumns(family, props), [family, props.columns, props.children, declaredColumns, declaredChildren]);
  const columns = React.useMemo(() => resolveVisibleDataGridColumns(allColumns, props.hiddenColumnIds), [allColumns, props.hiddenColumnIds]);
  const prepared = React.useMemo(() => prepareDataGridRows(props.items, props.getRowId, columns, classifyDataGridValue), [props.items, props.getRowId, columns]);
  const strings = React.useMemo(() => resolveDataGridStrings(props.strings), [props.strings]);
  const suppliedIds = React.useMemo(() => new Set(prepared.map(row => row.rowId)), [prepared]);
  const rows = React.useMemo((): GridRow<T>[] => {
    const result: GridRow<T>[] = prepared.map(row => ({ kind: 'data', row }));
    if (props.loading) {
      let index = 0;
      while (result.length < 4) {
        const rowId = `data-grid-loading-${index++}`;
        if (!suppliedIds.has(rowId)) result.push({ kind: 'loading', rowId });
      }
    }
    return result;
  }, [prepared, suppliedIds, props.loading]);
  const selected = React.useMemo(() => new Set(props.selectedItems), [props.selectedItems]);
  const selectionHandler: NonNullable<DataGridProps<T>['onSelectionChange']> = (event, data) => {
    const next = new Set(data.selectedItems);
    if (props.selectionMode === 'multiselect') for (const id of selected) if (!suppliedIds.has(id)) next.add(id);
    props.onSelectionChange?.(event, { selectedItems: next });
  };
  const onSelectionChange = useStableCallback(selectionHandler);
  const sortHandler: NonNullable<DataGridProps<T>['onSortChange']> = (event, sort) => props.onSortChange?.(event, sort);
  const onSortChange = useStableCallback(sortHandler);
  const activate = useStableCallback((event: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, row: GridRow<T>): void => {
    if (!props.loading && row.kind === 'data' && canActivate(event)) props.onRowActivate?.(event, { item: row.row.item, rowId: row.row.rowId });
  });
  const nativeColumns = React.useMemo(() => columns.map(column => createTableColumn<GridRow<T>>({
    columnId: column.id,
    compare: column.kind !== 'actions' && column.options.sortable === true ? neutralSortableCompare : () => 0,
  })), [columns]);
  // Fluent's context selectors can transiently supply the prior rows/definitions
  // to new render callbacks. Resolve native IDs against this render's prepared
  // pass so source readers/classifications always retain their owning descriptors.
  const rowsById = React.useMemo(() => new Map(rows.map(row => [rowIdentity(row), row])), [rows]);
  const byId = React.useMemo(() => new Map(columns.map(column => [column.id, column])), [columns]);
  const widths = React.useMemo(() => sizingOptions(columns, props.columnSizingOptions), [columns, props.columnSizingOptions]);
  const columnClasses = new Map(columns.map(column => {
    const preferred = column.options.width, minimum = column.options.minWidth ?? Math.min(preferred ?? 100, 100);
    return [column.id, props.resizableColumns ? undefined : sx(
      minWidth.px(minimum),
      preferred !== undefined && width.px(preferred),
      preferred !== undefined && styles.fixedColumn,
      preferred !== undefined && flexItem.noGrow,
      preferred !== undefined && flexItem.noShrink,
    )];
  }));
  // ReactNode accepts React's legacy null/false empty-child contract at this
  // external content boundary; internal absence and defaults remain undefined.
  const hasError = props.error !== undefined && props.error !== null && props.error !== false;
  const emptyContent = props.emptyState === 'noResults' ? props.noResultsContent : props.emptyContent;
  const status = columns.length === 0 ? strings.noVisibleColumns : prepared.length === 0 && !props.loading && !hasError
    ? emptyContent ?? (props.emptyState === 'noResults' ? strings.noResults : strings.empty) : undefined;
  const renderBodyCell = (definition: TableColumnDefinition<GridRow<T>>, row: GridRow<T>): React.ReactElement => {
    const column = byId.get(String(definition.columnId));
    if (column === undefined) return <></>; // Ignore a retained native definition while its context updates.
    const context: DataGridCellRenderContext<T> | undefined = row.kind === 'data'
      ? {
        column,
        item: row.row.item,
        rowId: row.row.rowId,
        value: row.row.getValue(column),
        locale: props.locale,
        size: props.size,
        strings,
      } : undefined;
    const tooltip = context === undefined ? undefined : getDataGridCellTooltip(context);
    const cell = <DataGridCell
      data-grid-column-id={column.id}
      className={columnClasses.get(column.id)}
      focusMode={props.loading ? 'none' : isInteractiveDataGridColumn(column) ? 'group' : 'cell'}>
      {context !== undefined && <div data-skeleton-ignore={props.loading ? '' : undefined}
        className={styles.cellContent}>
        {renderDataGridCell(context)}
      </div>}
      {props.loading && <DataGridLoadingCell column={column} rowId={rowIdentity(row)} />}
    </DataGridCell>;
    return tooltip === undefined || tooltip.trim().length === 0 ? cell : <Tooltip content={tooltip} relationship="description">{cell}</Tooltip>;
  };
  return <div className={sx(flex.column, gap.small, width.full, minWidth.zero)}>
    {hasError && <MessageBar intent="error">
      <MessageBarBody>{renderContentSlot(props.error, 'div', sx, { shorthand: 'preserve' })}</MessageBarBody>
      {props.onRetry !== undefined && <MessageBarActions>
        <Button appearance="transparent" onClick={props.onRetry}>{strings.retry}</Button>
      </MessageBarActions>}
    </MessageBar>}
    {status !== undefined ? <div role="status" className={columns.length > 0 && (emptyContent ?? undefined) === undefined
      ? sx(textAlign.center, paddingBlock.medium) : undefined}>{renderContentSlot(status, 'div', sx, { shorthand: 'preserve' })}</div> : columns.length === 0 || (hasError && rows.length === 0) ? <></> : <>
      <div data-grid-scroll="" className={sx(horizontalOverflow.auto, width.full, minWidth.zero)}>
        <FluentDataGrid {...nativeRootProps(props)} ref={rootRef} data-grid-root="" items={rows} getRowId={rowIdentity} columns={nativeColumns}
          className={sx(styles.grid, width.full, props.className)} style={props.style} size={props.size}
          aria-busy={props.loading || undefined} selectionMode={props.loading ? undefined : props.selectionMode}
          selectedItems={selected} onSelectionChange={onSelectionChange} sortState={resolveDataGridSortState(columns, props.sortState) ?? emptySort} onSortChange={onSortChange}
          sortable={columns.some(column => column.kind !== 'actions' && column.options.sortable === true)}
          resizableColumns={props.resizableColumns} columnSizingOptions={widths} resizableColumnsOptions={{ autoFitColumns: false, ...props.resizableColumnsOptions }}>
          <DataGridHeader><DataGridRow<GridRow<T>> selectionCell={props.loading || props.selectionMode === undefined ? undefined : { checkboxIndicator: { 'aria-label': strings.selectAllRows } }}>
            {definition => {
              const column = byId.get(String(definition.columnId));
              if (column === undefined) return <></>;
              return <DataGridHeaderCell
                data-grid-column-id={definition.columnId}
                className={columnClasses.get(String(definition.columnId))}
                title={column?.options.headerTooltip}
                aria-label={column?.options.headerAriaLabel}>
                {renderContentSlot(column.kind === 'actions' && column.options.header === undefined ? strings.actions : column.options.header, 'span', sx, { shorthand: 'preserve' })}
              </DataGridHeaderCell>;
            }}
          </DataGridRow></DataGridHeader>
          <AutoSkeleton loading={props.loading === true} className={sx(width.full)}>
          <DataGridBody<GridRow<T>>>{({ rowId }) => {
            const item = rowsById.get(rowId);
            if (item === undefined) return <></>;
            return <DataGridRow<GridRow<T>>
              data-grid-row-id={item.kind === 'data' ? item.row.rowId : undefined}
              data-grid-loading-row={item.kind === 'loading' ? '' : undefined}
              selectionCell={props.loading || props.selectionMode === undefined ? undefined : {
                checkboxIndicator: { 'aria-label': strings.selectRow(rowIdentity(item)) },
                radioIndicator: { 'aria-label': strings.selectRow(rowIdentity(item)) },
              }}
              onClick={(event: React.MouseEvent<HTMLDivElement>) => activate(event, item)}
              onKeyDown={(event: React.KeyboardEvent<HTMLDivElement>) => { if (event.key === 'Enter') activate(event, item); }}>
              {definition => renderBodyCell(definition, item)}
            </DataGridRow>;
          }}</DataGridBody>
          </AutoSkeleton>
        </FluentDataGrid>
      </div>
    </>}
  </div>;
}
/**
 * Create the native root around a feature-local owning family.
 */
export function createDataGridRuntimeComponent<T extends object>(family: FamilyResolution<T>): React.ForwardRefExoticComponent<DataGridProps<T> & React.RefAttributes<HTMLDivElement>> {
  const Root = React.forwardRef<HTMLDivElement, DataGridProps<T>>((props, ref) => <DataGridRuntime<T> family={family} props={props} rootRef={ref} />);
  Root.displayName = 'DataGrid';
  return Root;
}
const dynamicFamily = createDynamicColumnFamily();
/**
 * Native controlled grid for dictionary rows. Values remain unknown until classified.
 * Requires immutable items, stable getRowId and one same-suite column collection.
 * Native div ref is unavailable while a status replaces the grid. Fluent supplies
 * navigation, resize, selection and cell widgets; no SPFx provider is required.
 * @example Render a dynamic title column
 * ```tsx
 * import * as React from 'react';
 * import { DataGrid } from '@apvee/spfx-react-toolkit';
 *
 * const rows = [{ id: 1, title: 'Plan' }];
 * <DataGrid items={rows} getRowId={row => String(row.id)}>
 *   <DataGrid.TextColumn id="title" header="Title" field="title" />
 * </DataGrid>;
 * ```
 *
 * Executable examples mount beneath the host FluentProvider/theme and shared Griffel
 * renderer with the Fluent Charts peer installed. Optional render maps the same owned
 * regions to host presentation; omission renders controls, commands, status, subject
 * and diagnostics in order. Optional resetRevision defaults to zero; only later
 * changes restore the baseline while preserving Demo identity. Selecting another
 * scenario mounts that scenario's baseline. All changes are local consumer state.
 *
 * @scenario Classify dynamic row values and preserve controlled selection
 * ```tsx
 * import * as React from 'react';
 * import { Button, Checkbox, Field, Select, FluentProvider, makeStyles } from '@fluentui/react-components';
 * import { DataGrid, useSx, flex, gap, width } from '@apvee/spfx-react-toolkit';
 *
 * const initialRows: readonly Record<string, unknown>[] = [
 *   { id: 'valid', text: '', number: 0, enabled: false },
 *   { id: 'missing' },
 *   { id: 'invalid', text: { unsafe: 'object' }, number: '42', enabled: 'yes' },
 * ];
 * const getRowId = (row: Record<string, unknown>): string => String(row.id);
 * type DemoRegions = { readonly controls?: React.ReactNode; readonly commands?: React.ReactNode; readonly status?: React.ReactNode; readonly diagnostics?: React.ReactNode; readonly children?: React.ReactNode };
 * type DemoProps = { readonly render?: (regions: DemoRegions) => React.ReactElement; readonly resetRevision?: number };
 * const useStyles = makeStyles({
 *   root: { display: 'flex', flexDirection: 'column', gap: '12px' },
 *   actions: { display: 'flex', flexWrap: 'wrap', gap: '8px' },
 *   scroll: { overflowX: 'auto', minWidth: 0 },
 *   frame: { maxInlineSize: '100%', minInlineSize: 0 },
 * });
 * export function Demo({ render, resetRevision = 0 }: DemoProps = {}): React.ReactElement {
 *   const sx = useSx(), styles = useStyles();
 *   const [dir, setDir] = React.useState<'host' | 'ltr' | 'rtl'>('host');
 *   const [size, setSize] = React.useState(640);
 *   const [gridSize, setGridSize] = React.useState<'small' | 'medium' | 'extra-small'>('medium');
 *   const [selectionMode, setSelectionMode] = React.useState<'multiselect' | 'single' | undefined>('multiselect');
 *   const [rows, setRows] = React.useState(initialRows);
 *   const [selected, setSelected] = React.useState<ReadonlySet<string | number>>(() => new Set(['off-page']));
 *   const [page, setPage] = React.useState(false);
 *   const [hidden, setHidden] = React.useState(false);
 *   const [customStrings, setCustomStrings] = React.useState(false);
 *   const gridStrings = customStrings ? { selectAllRows: 'Select every document', selectRow: (id: string | number) => `Select document ${id}`, empty: 'No team documents', loading: 'Fetching documents' } : undefined;
 *   const [loading, setLoading] = React.useState(false);
 *   const [error, setError] = React.useState(false);
 *   const [empty, setEmpty] = React.useState(false);
 *   const [emptyState, setEmptyState] = React.useState<'empty' | 'noResults'>('empty');
 *   const [resizable, setResizable] = React.useState(false);
 *   const [outcome, setOutcome] = React.useState('Interact with the grid to observe selection and row identity.');
 *   const [activation, setActivation] = React.useState('none');
 *   const [retries, setRetries] = React.useState(0);
 *   const [observedRef, setObservedRef] = React.useState('not inspected');
 *   const [slotRefs, setSlotRefs] = React.useState('not inspected');
 *   const [suppressSlots, setSuppressSlots] = React.useState(false);
 *   const [customStatusSlots, setCustomStatusSlots] = React.useState(false);
 *   const errorSlotRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
 *   const gridRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
 *   const placeholderSlotRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
 *   const previousReset = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (previousReset.current === resetRevision) return;
 *     previousReset.current = resetRevision;
 *     setDir('host'); setSize(640); setGridSize('medium'); setSelectionMode('multiselect');
 *     setRows(initialRows); setSelected(new Set(['off-page'])); setPage(false); setHidden(false); setCustomStrings(false);
 *     setLoading(false); setError(false); setEmpty(false); setEmptyState('empty'); setResizable(false);
 *     setOutcome('Interact with the grid to observe selection and row identity.'); setActivation('none'); setRetries(0);
 *     setObservedRef('not inspected'); setSlotRefs('not inspected'); setSuppressSlots(false); setCustomStatusSlots(false);
 *   }, [resetRevision]);
 *   const visibleRows = empty ? [] : page ? rows.filter(row => row.id === 'valid') : rows;
 *   const hiddenIds = hidden ? new Set(['text', 'number', 'enabled', 'status']) : undefined;
 *   const regions: DemoRegions = {
 *     controls: <>
 *       <Field label="Width"><Select id="dg-width" value={size} onChange={event => setSize(Number(event.target.value))}>{[320, 640, 1100].map(value => <option key={value} value={value}>{value}px</option>)}</Select></Field>
 *       <Field label="size"><Select id="dg-size" value={gridSize} onChange={event => { const value = event.target.value; if (value === 'small' || value === 'medium' || value === 'extra-small') setGridSize(value); }}><option value="extra-small">Extra small</option><option value="small">Small</option><option value="medium">Medium</option></Select></Field>
 *       <Field label="selectionMode"><Select id="dg-selection-mode" value={selectionMode ?? 'none'} onChange={event => setSelectionMode(event.target.value === 'single' ? 'single' : event.target.value === 'multiselect' ? 'multiselect' : undefined)}><option value="none">None</option><option value="single">Single</option><option value="multiselect">Multiselect</option></Select></Field>
 *     </>,
 *     commands: <div className={styles.actions} role="group" aria-label="Consumer simulation commands">
 *       <Button id="dg-reverse" onClick={() => setRows(previous => [...previous].reverse())}>Reverse supplied rows</Button>
 *     </div>,
 *     status: <p role="status">{outcome} Rows: {visibleRows.length}; selected: {selected.size}; retries: <output id="dg-retries">{retries}</output></p>,
 *     diagnostics: <>
 *       <Field label="Direction"><Select id="dg-dir" value={dir} onChange={event => setDir(event.target.value === 'rtl' ? 'rtl' : event.target.value === 'ltr' ? 'ltr' : 'host')}><option value="host">Host</option><option value="ltr">LTR</option><option value="rtl">RTL</option></Select></Field>
 *       <Checkbox id="dg-page" label="Supply first row only" checked={page} onChange={(_event, data) => setPage(data.checked === true)} />
 *       <Checkbox id="dg-hidden" label="Hide all columns" checked={hidden} onChange={(_event, data) => setHidden(data.checked === true)} />
 *       <Checkbox id="dg-strings" label="Custom UI strings (number/date locale stays en-US)" checked={customStrings} onChange={(_event, data) => setCustomStrings(data.checked === true)} />
 *       <Checkbox id="dg-loading" label="Loading" checked={loading} onChange={(_event, data) => setLoading(data.checked === true)} />
 *       <Checkbox id="dg-error" label="Error" checked={error} onChange={(_event, data) => setError(data.checked === true)} />
 *       <Checkbox id="dg-empty" label="Empty rows" checked={empty} onChange={(_event, data) => setEmpty(data.checked === true)} />
 *       <Field label="Empty presentation"><Select id="dg-empty-state" value={emptyState} onChange={event => setEmptyState(event.target.value === 'noResults' ? 'noResults' : 'empty')}><option value="empty">Empty</option><option value="noResults">No results</option></Select></Field>
 *       <Checkbox id="dg-resizable" label="Native resizing" checked={resizable} onChange={(_event, data) => setResizable(data.checked === true)} />
 *       <Checkbox id="dg-custom-status-slots" label="Custom empty and no-results slots" checked={customStatusSlots} onChange={(_event, data) => setCustomStatusSlots(data.checked === true)} />
 *       <Checkbox id="dg-slot-suppress" label="Suppress status and placeholder text" checked={suppressSlots} onChange={(_event, data) => setSuppressSlots(data.checked === true)} />
 *       <Button id="dg-inspect-slot-refs" onClick={() => setSlotRefs(`error=${errorSlotRef.current?.tagName ?? 'absent'} missing=${placeholderSlotRef.current?.tagName ?? 'absent'}`)}>Inspect inner slot refs</Button>
 *       <Button id="dg-inspect-ref" onClick={() => setObservedRef(gridRef.current?.tagName.toLowerCase() ?? 'absent')}>Inspect native grid ref</Button>
 *       <p>Inner slot refs: <output id="dg-slot-refs">{slotRefs}</output></p>
 *       <p>Selected IDs: <output id="dg-selection">{JSON.stringify(Array.from(selected))}</output></p>
 *       <p>Activation: <output id="dg-activation-output">{activation}</output></p>
 *       <p>Native ref: <output id="dg-ref">{observedRef}</output></p>
 *     </>,
 *     children: <div className={styles.scroll} role="region" aria-label="DataGrid preview scroll" tabIndex={0}>
 *       <div id="dg-frame" className={sx(flex.column, gap.large, width.px(size), styles.frame)}>
 *         <DataGrid id="data-grid-main" strings={gridStrings} ref={gridRef} aria-label="Dynamic data" items={visibleRows} getRowId={getRowId} locale="en-US"
 *           size={gridSize} selectionMode={selectionMode} selectedItems={selected} onSelectionChange={(_event, data) => { setSelected(new Set(data.selectedItems)); setOutcome(`Selection changed: ${data.selectedItems.size} selected`); }}
 *           resizableColumns={resizable} hiddenColumnIds={hiddenIds} loading={loading} emptyState={emptyState}
 *           error={error ? { ref: errorSlotRef, children: <p>Consumer-owned loading error</p> } : undefined}
 *           emptyContent={suppressSlots ? false : customStatusSlots ? { children: 'Consumer-owned empty slot' } : undefined}
 *           noResultsContent={suppressSlots ? false : customStatusSlots ? { children: (Component, nativeProps) => <Component {...nativeProps}>Consumer-owned no-results slot</Component> } : undefined}
 *           onRowActivate={selectionMode === undefined ? (_event, data) => setActivation(JSON.stringify({ rowId: data.rowId })) : undefined}
 *           onRetry={() => { setRetries(value => value + 1); setError(false); }}>
 *           <DataGrid.TextColumn id="text" header="Title" field="text" placeholder={suppressSlots ? false : { ref: placeholderSlotRef, children: 'Missing text' }} invalidPlaceholder="Invalid text" />
 *           <DataGrid.NumberColumn id="number" header="Total" field="number" placeholder={suppressSlots ? false : 'Missing number'} invalidPlaceholder="Invalid number" />
 *           <DataGrid.BooleanColumn id="enabled" header="Enabled" field="enabled" placeholder={suppressSlots ? false : undefined} />
 *           <DataGrid.TextColumn id="status" header="Status" field="status" placeholder={suppressSlots ? false : 'Missing status'} />
 *         </DataGrid>
 *       </div>
 *     </div>,
 *   };
 *   return <FluentProvider dir={dir === 'host' ? undefined : dir}>
 *     <section id="data-grid-demo" className={styles.root}>
 *       {render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>}
 *     </section>
 *   </FluentProvider>;
 * }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'Compare valid empty text, zero and false with missing and invalid values, then select or activate original IDs and inspect native status and placeholder refs.' };
 * ```
 *
 * @useWhen Rows are dynamic records and column callbacks can explicitly handle unknown values.
 * @avoidWhen Use createDataGrid for a known row model to preserve inferred source domains.
 * @param props - Controlled row data, same-suite columns and native grid presentation props.
 * @returns Native Fluent grid output, or the configured empty/no-visible-column status instead of the grid root.
 */
export const DataGrid: DynamicDataGridSuite = Object.assign(createDataGridRuntimeComponent(dynamicFamily), getPublicDataGridColumnFamily(dynamicFamily));

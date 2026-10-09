import * as React from 'react';
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
/** Strip owned runtime paths from untyped callers before forwarding native attributes. */
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
      <MessageBarBody>{props.error}</MessageBarBody>
      {props.onRetry !== undefined && <MessageBarActions>
        <Button appearance="transparent" onClick={props.onRetry}>{strings.retry}</Button>
      </MessageBarActions>}
    </MessageBar>}
    {status !== undefined ? <div role="status" className={columns.length > 0 && (emptyContent ?? undefined) === undefined
      ? sx(textAlign.center, paddingBlock.medium) : undefined}>{status}</div> : columns.length === 0 || (hasError && rows.length === 0) ? <></> : <>
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
                {column.kind === 'actions' && column.options.header === undefined ? strings.actions : column.options.header}
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
/** Create the native root around a feature-local owning family. */
export function createDataGridRuntimeComponent<T extends object>(family: FamilyResolution<T>): React.ForwardRefExoticComponent<DataGridProps<T> & React.RefAttributes<HTMLDivElement>> {
  const Root = React.forwardRef<HTMLDivElement, DataGridProps<T>>((props, ref) => <DataGridRuntime family={family} props={props} rootRef={ref} />);
  Root.displayName = 'DataGrid';
  return Root;
}
const dynamicFamily = createDynamicColumnFamily();
/** Native controlled grid for dictionary rows. Values remain unknown until classified.
 * Requires immutable items, stable getRowId and one same-suite column collection.
 * Native div ref is unavailable while a status replaces the grid. Fluent supplies
 * navigation, resize, selection and cell widgets; no SPFx provider is required.
 * @example <DataGrid items={rows} getRowId={row => String(row.id)}><DataGrid.TextColumn id="title" header="Title" field="title" /></DataGrid>
 */
export const DataGrid: DynamicDataGridSuite = Object.assign(createDataGridRuntimeComponent(dynamicFamily), getPublicDataGridColumnFamily(dynamicFamily));

import type { DataGridProps, DataGridStrings, DataGridSparklineSummaryText } from './data-grid.types';
import type { ResolvedDataGridColumn } from './column-normalization.internal';
import { nativeDataGridResource } from './localization/data-grid-native.internal';
import { formatResourcePlural, formatResourceTemplate } from '../../localization/resource-formatters.internal';
export function resolveVisibleDataGridColumns<T extends object>(columns: readonly ResolvedDataGridColumn<T>[], hidden?: ReadonlySet<string>): readonly ResolvedDataGridColumn<T>[] {
  return hidden === undefined ? columns : Object.freeze(columns.filter(column => !hidden.has(column.id)));
}
export function resolveDataGridSortState<T extends object>(columns: readonly ResolvedDataGridColumn<T>[], sort: DataGridProps<T>['sortState']): DataGridProps<T>['sortState'] {
  return sort !== undefined && columns.some(column => column.id === sort.sortColumn && column.kind !== 'actions' && column.options.sortable === true) ? sort : undefined;
}
export function validateDataGridInteraction<T extends object>(props: Pick<DataGridProps<T>, 'selectionMode' | 'onRowActivate'>): void {
  if (props.selectionMode !== undefined && props.onRowActivate !== undefined)
    throw new TypeError('DataGrid row activation and selection are mutually exclusive');
}

/** Native UI grammar is independent of column/root Intl data formatting. */
function nativeDataGridStrings(): DataGridStrings {
  const resource = nativeDataGridResource;
  return {
    yes: resource.yes, no: resource.no, noTags: resource.noTags,
    tagOverflow: (count: number): string => formatResourcePlural(resource.tagOverflow, resource.__locale, count),
    ratingCount: (count: number, formattedCount: string): string =>
      formatResourcePlural(resource.ratingCount, resource.__locale, count, { formattedCount }),
    noObservations: resource.noObservations, noVisibleColumns: resource.noVisibleColumns,
    actions: resource.actions, moreActions: resource.moreActions, loading: resource.loading,
    empty: resource.empty, noResults: resource.noResults, invalidValue: resource.invalidValue,
    retry: resource.retry, selectAllRows: resource.selectAllRows,
    selectRow: (rowId: string | number): string => formatResourceTemplate(resource.selectRowTemplate, { rowId }),
    sparklineSummary: (values: DataGridSparklineSummaryText): string => formatResourceTemplate(resource.sparklineSummaryTemplate, { ...values }),
  };
}
export function resolveDataGridStrings(overrides?: Partial<DataGridStrings>): DataGridStrings {
  const result = nativeDataGridStrings();
  if (overrides === undefined) return result;
  // Nullable caller input is an untyped external configuration boundary.
  if (typeof overrides !== 'object' || overrides === null || Array.isArray(overrides)) throw new TypeError('DataGrid strings must be an override object');
  for (const key of Object.keys(result) as (keyof DataGridStrings)[]) {
    const value = overrides[key];
    if (value === undefined) continue;
    if (typeof value !== typeof result[key]) throw new TypeError(`DataGrid strings.${key} has an invalid type`);
    if (typeof value === 'string' && value.trim().length === 0) continue;
    Object.assign(result, { [key]: value });
  }
  return Object.freeze(result);
}
/** Validate root configuration even when no values/cells are rendered. */
export function validateDataGridRootConfiguration<T extends object>(props: DataGridProps<T>): void {
  if (!Array.isArray(props.items)) throw new TypeError('DataGrid items must be a readonly array');
  if (typeof props.getRowId !== 'function') throw new TypeError('DataGrid getRowId must be a callback');
  if (props.selectionMode !== undefined && props.selectionMode !== 'single' && props.selectionMode !== 'multiselect') throw new TypeError('DataGrid selectionMode must be single or multiselect');
  for (const ids of [props.selectedItems, props.hiddenColumnIds]) {
    if (ids === undefined) continue;
    if (typeof ids.has !== 'function' || typeof ids[Symbol.iterator] !== 'function') throw new TypeError('DataGrid controlled IDs must be a ReadonlySet');
  }
  if (props.hiddenColumnIds !== undefined) for (const id of props.hiddenColumnIds) {
    if (typeof id !== 'string') throw new TypeError('DataGrid hidden column IDs must be strings');
  }
  if (props.selectedItems !== undefined) for (const id of props.selectedItems) {
    if (typeof id !== 'string' && (typeof id !== 'number' || !Number.isFinite(id))) throw new TypeError('DataGrid selection IDs must be strings or finite numbers');
  }
  const configured = props.locale;
  if (configured !== undefined) {
    if (typeof configured !== 'string' && (!Array.isArray(configured) || !configured.every(locale => typeof locale === 'string'))) throw new TypeError('DataGrid locale must be text or a locale list');
    try { new Intl.Collator(typeof configured === 'string' ? configured : [...configured]).resolvedOptions(); }
    catch { throw new TypeError('DataGrid locale is invalid'); }
  }
  validateDataGridInteraction(props);
}

import type { DataGridProps, DataGridStrings, DataGridSparklineSummaryText } from './data-grid.types';
import type { ResolvedDataGridColumn } from './column-normalization.internal';
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

/** English fallbacks are presentation, independently of the Intl data locale. */
const defaultStrings: DataGridStrings = Object.freeze({
  yes: 'Yes', no: 'No', noTags: 'No tags',
  tagOverflow: (count: number): string => count === 1 ? 'Show 1 more tag' : `Show ${count} more tags`,
  ratingCount: (count: number, formattedCount: string): string => `${formattedCount} ${count === 1 ? 'vote' : 'votes'}`,
  noObservations: 'No observations', noVisibleColumns: 'No visible columns',
  actions: 'Actions', moreActions: 'More actions', loading: 'Loading',
  empty: 'No data available', noResults: 'No results found', invalidValue: 'Invalid value',
  retry: 'Retry', selectAllRows: 'Select all rows',
  selectRow: (rowId: string | number): string => `Select row ${rowId}`,
  sparklineSummary: (values: DataGridSparklineSummaryText): string =>
    `First ${values.first}; last ${values.last}; minimum ${values.minimum}; maximum ${values.maximum}; average ${values.average}; change ${values.change}; change percent ${values.changePercent}`,
});
export function resolveDataGridStrings(overrides?: Partial<DataGridStrings>): DataGridStrings {
  const result = { ...defaultStrings };
  if (overrides === undefined) return result;
  // Nullable caller input is an untyped external configuration boundary.
  if (typeof overrides !== 'object' || overrides === null || Array.isArray(overrides)) throw new TypeError('DataGrid strings must be an override object');
  for (const key of Object.keys(defaultStrings) as (keyof typeof defaultStrings)[]) {
    const value = overrides[key];
    if (value === undefined) continue;
    if (typeof value !== typeof defaultStrings[key]) throw new TypeError(`DataGrid strings.${key} has an invalid type`);
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

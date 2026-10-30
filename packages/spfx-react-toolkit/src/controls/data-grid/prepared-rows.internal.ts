import type { DataGridRowId } from './data-grid-column.types';
import type { ResolvedDataGridColumn } from './column-normalization.internal';
/** One immutable source row and pass-local value cache. */
export interface PreparedDataGridRow<T extends object, V> {
    readonly item: T;
    readonly rowId: DataGridRowId;
    readonly getRawValue: (column: ResolvedDataGridColumn<T>) => unknown;
    readonly getValue: (column: ResolvedDataGridColumn<T>) => V;
}
/** Source reads are cached by field or accessor identity, including undefined results. */
export function prepareDataGridRows<T extends object>(
    items: readonly T[], getRowId: (item: T) => DataGridRowId,
    columns: readonly ResolvedDataGridColumn<T>[],
): readonly PreparedDataGridRow<T, unknown>[];
export function prepareDataGridRows<T extends object, V>(
    items: readonly T[], getRowId: (item: T) => DataGridRowId,
    columns: readonly ResolvedDataGridColumn<T>[],
    classify: (column: ResolvedDataGridColumn<T>, raw: unknown, item: T) => V,
): readonly PreparedDataGridRow<T, V>[];
export function prepareDataGridRows<T extends object>(
    items: readonly T[], getRowId: (item: T) => DataGridRowId,
    columns: readonly ResolvedDataGridColumn<T>[],
    classify?: (column: ResolvedDataGridColumn<T>, raw: unknown, item: T) => unknown,
): readonly PreparedDataGridRow<T, unknown>[] {
    const ids = new Set<DataGridRowId>();
    const keys = new Set<string>();
    const declared = new Set(columns);
    return Object.freeze(items.map(item => {
        const rowId = getRowId(item);
        if (typeof rowId !== 'string' && (typeof rowId !== 'number' || !Number.isFinite(rowId)))
            throw new TypeError('DataGrid row ID must be a string or finite number');
        if (ids.has(rowId))
            throw new TypeError(`DataGrid duplicate row ID "${rowId}"`);
        const key = String(rowId);
        if (keys.has(key))
            throw new TypeError(`DataGrid row ID serialization collision "${key}"`);
        ids.add(rowId);
        keys.add(key);
        const reads = new Map<string | ((item: T) => unknown), unknown>();
        const values = new Map<ResolvedDataGridColumn<T>, unknown>();
        const getRawValue = (column: ResolvedDataGridColumn<T>): unknown => {
            if (!declared.has(column))
                throw new TypeError('DataGrid prepared row requires a column from this pass');
            if (column.source === undefined)
                return undefined;
            if (!reads.has(column.source))
                reads.set(column.source, column.read(item));
            return reads.get(column.source);
        };
        const getValue = (column: ResolvedDataGridColumn<T>): unknown => {
            if (!values.has(column)) {
                const raw = getRawValue(column);
                values.set(column, classify === undefined ? raw : classify(column, raw, item));
            }
            return values.get(column); // Map.has above distinguishes cached undefined.
        };
        columns.forEach(getValue);
        return Object.freeze({ item, rowId, getRawValue, getValue });
    }));
}

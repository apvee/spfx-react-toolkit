import type { DataGridProps, DataGridStrings } from './data-grid.types';
import type { DataGridBadgeValue, DataGridIconValue, DataGridLocale, DataGridRowId, DataGridSparklinePoint, DataGridSparklineSummary } from './data-grid-column.types';
import type { ResolvedDataGridColumn, ResolvedDataGridPerson, ResolvedDataGridTag } from './column-normalization.internal';
export interface DataGridSparklineValue {
    readonly points: readonly DataGridSparklinePoint[];
    readonly summary: DataGridSparklineSummary;
}
type ValueDomains = {
    text: string;
    number: number;
    boolean: boolean;
    date: number;
    dateTime: number;
    badge: DataGridBadgeValue;
    icon: DataGridIconValue;
    link: string;
    person: ResolvedDataGridPerson;
    people: readonly ResolvedDataGridPerson[];
    tags: readonly ResolvedDataGridTag[];
    progress: number;
    rating: number;
    sparkline: DataGridSparklineValue;
    custom: unknown;
    actions: undefined;
};
export type DataGridCellValue = {
    readonly status: 'missing';
    readonly raw: unknown;
} | {
    readonly status: 'invalid';
    readonly raw: unknown;
} | {
    [K in keyof ValueDomains]: {
        readonly status: 'valid';
        readonly kind: K;
        readonly raw: unknown;
        readonly value: ValueDomains[K];
    };
}[keyof ValueDomains];
export interface DataGridCellRenderContext<T extends object> {
    readonly column: ResolvedDataGridColumn<T>;
    readonly item: T;
    readonly rowId: DataGridRowId;
    readonly value: DataGridCellValue;
    readonly locale: DataGridLocale | undefined;
    readonly size: DataGridProps<T>['size'];
    readonly strings: DataGridStrings;
}

import type * as React from 'react';
import type { DataGridProps as FluentDataGridProps, SelectionMode } from '@fluentui/react-components';
import type {
  ActionsColumnProps,
  DataGridColumnBuilder,
  DataGridColumnDefinition,
  DataGridColumnMethod,
  DataGridLocale,
  DataGridRowId,
  DynamicDataGridColumnBuilder,
  DynamicDataGridColumnProps,
} from './data-grid-column.types';
/** Locale-formatted chart metrics supplied to the summary string callback. */
export interface DataGridSparklineSummaryText {
    /** First observation. */
    first: string;
    /** Latest observation. */
    last: string;
    /** Minimum observation. */
    minimum: string;
    /** Maximum observation. */
    maximum: string;
    /** Overflow-safe mean. */
    average: string;
    /** Formatted change or placeholder when independently unrepresentable. */
    change: string;
    /** Formatted percent change or placeholder for zero denominator/unrepresentable result. */
    changePercent: string;
}
/** Localized status and cell strings. */
export interface DataGridStrings {
    /** Empty input status. */
    empty: string;
    /** No matches status. */
    noResults: string;
    /** No visible columns status. */
    noVisibleColumns: string;
    /** Retained loading label for source compatibility; the grid shows skeletons
     * and aria-busy rather than a visible loading message. */
    loading: string;
    /** Default actions heading. */
    actions: string;
    /** Overflow actions trigger. */
    moreActions: string;
    /** True label shared by text, icon and tooltip. */
    yes: string;
    /** False label shared by text, icon and tooltip. */
    no: string;
    /** Empty tag list label. */
    noTags: string;
    /** Accessible trigger label for hidden tags. */
    tagOverflow: (count: number) => string;
    /** Accessible vote count when a rating column explicitly supplies count. Defaults to "1 vote" or "N votes".
     * @param count Original non-negative whole number of votes, including zero.
     * @param formattedCount Count formatted with the column or grid locale.
     * @returns Localized count text, distinct from the rating value and maximum.
     */
    ratingCount: (count: number, formattedCount: string) => string;
    /** Empty chart series label. */
    noObservations: string;
    /** Invalid present cell data label. */
    invalidValue: string;
    /** Select all supplied rows. */
    selectAllRows: string;
    /** Accessible selection label for an original row ID. */
    selectRow: (rowId: DataGridRowId) => string;
    /** Accessible chart summary from locale-formatted original observations.
     * Defaults to English metric labels. Derived metrics independently use an em dash
     * when unavailable; showValue=false does not suppress this summary.
     */
    sparklineSummary: (values: DataGridSparklineSummaryText) => string;
    /** Retry label. */
    retry: string;
}
/** Exactly one declared column collection, including an explicitly empty collection. */
export type DataGridColumnCollection<TItem extends object> = {
    /** Same-suite descriptors, including an explicitly empty list. */
    columns: readonly DataGridColumnDefinition<TItem>[];
    /** Excluded when descriptors are provided. */
    children?: never;
} | {
    /** Same-suite JSX markers, including an explicitly empty collection. */
    children: React.ReactNode;
    /** Excluded when JSX is declared. */
    columns?: never;
};
/** Controlled grid configuration. Input order is preserved.
 * @example <Grid items={rows} getRowId={row => row.id} columns={columns} />
 */
export type DataGridProps<TItem extends object> = Omit<
    FluentDataGridProps,
    keyof React.HTMLAttributes<HTMLDivElement>
    | 'items' | 'getRowId' | 'columns' | 'children'
    | 'sortState' | 'defaultSortState' | 'onSortChange'
    | 'selectedItems' | 'defaultSelectedItems' | 'onSelectionChange'
    | 'selectionMode' | 'dangerouslySetInnerHTML' | 'onClick'
    | 'as' | 'noNativeElements' | 'sortable' | 'columnSizingOptions' | 'containerWidthOffset'
> & Omit<React.HTMLAttributes<HTMLDivElement>, 'children' | 'dangerouslySetInnerHTML'>
& DataGridColumnCollection<TItem> & {
    /** Immutable source rows; objects are passed to callbacks unchanged. */
    items: readonly TItem[];
    /** Native div click event. */
    onClick?: React.MouseEventHandler<HTMLDivElement>;
    /** Stable finite number or string ID. Never inferred from position. */
    getRowId: (item: TItem) => DataGridRowId;
    /** Controlled sort intent; rows are not sorted internally. */
    sortState?: FluentDataGridProps['sortState'];
    /** Native controlled sort intent callback. */
    onSortChange?: FluentDataGridProps['onSortChange'];
    /** Controlled selection mode. Mutually exclusive with row activation. */
    selectionMode?: SelectionMode;
    /** Controlled IDs, retaining off-supplied selections. */
    selectedItems?: ReadonlySet<DataGridRowId>;
    /** Receives a fresh native selection Set. */
    onSelectionChange?: FluentDataGridProps['onSelectionChange'];
    /** Controlled hidden IDs; stale IDs are ignored. */
    hiddenColumnIds?: ReadonlySet<string>;
    /** Native resizing options keyed by declared column IDs. Descriptor width and
     * minWidth override matching native entries; other native options are preserved.
     * Active when resizableColumns is enabled. */
    columnSizingOptions?: FluentDataGridProps['columnSizingOptions'];
    /** Native offset applied to measured resize container width. */
    containerWidthOffset?: FluentDataGridProps['containerWidthOffset'];
    /** Root Intl locale fallback. */
    locale?: DataGridLocale;
    /** Overrides selected localized strings. */
    strings?: Partial<DataGridStrings>;
    /** Initial or refresh loading status. Reuses AutoSkeleton over type-aware,
     * rounded rectangle templates for every visible column and at least four rows.
     * Simulated geometry varies deterministically by row/column identity, without
     * fabricating consumer items or reading absent values. Real cell state remains
     * mounted; selection and activation are disabled until loading completes.
     * No separate loading label or progress indicator is rendered. */
    loading?: boolean;
    /** Consumer-owned error content. */
    error?: React.ReactNode;
    /** Retry intent without fetching. */
    onRetry?: () => void;
    /** Selects the empty or no-results presentation. The default localized text
     * is horizontally centered with Fluent medium block padding and natural height. */
    emptyState?: 'empty' | 'noResults';
    /** Custom empty input content, preserving its own presentation.
     * Can compose StateMessage for richer copy, artwork or consumer-owned actions. */
    emptyContent?: React.ReactNode;
    /** Custom no-results content, preserving its own presentation.
     * Can compose StateMessage without adding search/filter behavior to the grid. */
    noResultsContent?: React.ReactNode;
    /** Original row activation from unhandled click or Enter. */
    onRowActivate?: (event: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, data: {
        /** Original row. */
        item: TItem;
        /** Stable row identity. */
        rowId: DataGridRowId;
    }) => void;
};
/** Same-suite JSX column markers; runtime additionally validates erased JSX ownership. */
export interface DataGridColumnMarkers<TItem extends object> {
    /** Declares a text column with callbacks inferred from its source. */
    TextColumn: DataGridColumnMethod<TItem, 'text', React.ReactElement>;
    /** Declares a number column with callbacks inferred from its source. */
    NumberColumn: DataGridColumnMethod<TItem, 'number', React.ReactElement>;
    /** Declares a boolean column with callbacks inferred from its source. */
    BooleanColumn: DataGridColumnMethod<TItem, 'boolean', React.ReactElement>;
    /** Declares a date column with callbacks inferred from its source. */
    DateColumn: DataGridColumnMethod<TItem, 'date', React.ReactElement>;
    /** Declares a dateTime column with callbacks inferred from its source. */
    DateTimeColumn: DataGridColumnMethod<TItem, 'dateTime', React.ReactElement>;
    /** Declares a badge column with callbacks inferred from its source. */
    BadgeColumn: DataGridColumnMethod<TItem, 'badge', React.ReactElement>;
    /** Declares a icon column with callbacks inferred from its source. */
    IconColumn: DataGridColumnMethod<TItem, 'icon', React.ReactElement>;
    /** Declares a link column with callbacks inferred from its source. */
    LinkColumn: DataGridColumnMethod<TItem, 'link', React.ReactElement>;
    /** Declares a person column with callbacks inferred from its source. */
    PersonColumn: DataGridColumnMethod<TItem, 'person', React.ReactElement>;
    /** Declares a people column with callbacks inferred from its source. */
    PeopleColumn: DataGridColumnMethod<TItem, 'people', React.ReactElement>;
    /** Declares a tags column with callbacks inferred from its source. */
    TagsColumn: DataGridColumnMethod<TItem, 'tags', React.ReactElement>;
    /** Declares a progress column with callbacks inferred from its source. */
    ProgressColumn: DataGridColumnMethod<TItem, 'progress', React.ReactElement>;
    /** Declares a rating column with callbacks inferred from its source. */
    RatingColumn: DataGridColumnMethod<TItem, 'rating', React.ReactElement>;
    /** Declares a sparkline column with callbacks inferred from its source. */
    SparklineColumn: DataGridColumnMethod<TItem, 'sparkline', React.ReactElement>;
    /** Declares a custom column with callbacks inferred from its source. */
    Column: DataGridColumnMethod<TItem, 'custom', React.ReactElement>;
    /** Declares source-free row actions. */
    ActionsColumn: (props: ActionsColumnProps<TItem>) => React.ReactElement;
}
/** Typed grid and its suite-owned marker/builder family. Construct outside rendering. */
export type DataGridSuite<TItem extends object> = React.ForwardRefExoticComponent<DataGridProps<TItem> & React.RefAttributes<HTMLDivElement>> & DataGridColumnMarkers<TItem> & {
    /** Freezes descriptors from this suite's builder. */
    defineColumns: (factory: (builder: DataGridColumnBuilder<TItem>) => readonly DataGridColumnDefinition<TItem>[]) => readonly DataGridColumnDefinition<TItem>[];
};
/** Explicit dynamic JSX markers; generic source callbacks receive unknown. */
export interface DynamicDataGridColumnMarkers {
    /** Declares a guarded dynamic text column. */
    TextColumn: (props: DynamicDataGridColumnProps<'text'>) => React.ReactElement;
    /** Declares a guarded dynamic number column. */
    NumberColumn: (props: DynamicDataGridColumnProps<'number'>) => React.ReactElement;
    /** Declares a guarded dynamic boolean column. */
    BooleanColumn: (props: DynamicDataGridColumnProps<'boolean'>) => React.ReactElement;
    /** Declares a guarded dynamic date column. */
    DateColumn: (props: DynamicDataGridColumnProps<'date'>) => React.ReactElement;
    /** Declares a guarded dynamic dateTime column. */
    DateTimeColumn: (props: DynamicDataGridColumnProps<'dateTime'>) => React.ReactElement;
    /** Declares a guarded dynamic badge column. */
    BadgeColumn: (props: DynamicDataGridColumnProps<'badge'>) => React.ReactElement;
    /** Declares a guarded dynamic icon column. */
    IconColumn: (props: DynamicDataGridColumnProps<'icon'>) => React.ReactElement;
    /** Declares a guarded dynamic link column. */
    LinkColumn: (props: DynamicDataGridColumnProps<'link'>) => React.ReactElement;
    /** Declares a guarded dynamic person column. */
    PersonColumn: (props: DynamicDataGridColumnProps<'person'>) => React.ReactElement;
    /** Declares a guarded dynamic people column. */
    PeopleColumn: (props: DynamicDataGridColumnProps<'people'>) => React.ReactElement;
    /** Declares a guarded dynamic tags column. */
    TagsColumn: (props: DynamicDataGridColumnProps<'tags'>) => React.ReactElement;
    /** Declares a guarded dynamic progress column. */
    ProgressColumn: (props: DynamicDataGridColumnProps<'progress'>) => React.ReactElement;
    /** Declares a guarded dynamic rating column. */
    RatingColumn: (props: DynamicDataGridColumnProps<'rating'>) => React.ReactElement;
    /** Declares a guarded dynamic sparkline column. */
    SparklineColumn: (props: DynamicDataGridColumnProps<'sparkline'>) => React.ReactElement;
    /** Declares a guarded dynamic custom column. */
    Column: (props: DynamicDataGridColumnProps<'custom'>) => React.ReactElement;
    /** Declares source-free dynamic row actions. */
    ActionsColumn: (props: ActionsColumnProps<Record<string, unknown>>) => React.ReactElement;
}
/** Dynamic grid for dictionary rows. Typed factories remain typed even for index signatures. */
export type DynamicDataGridSuite = React.ForwardRefExoticComponent<DataGridProps<Record<string, unknown>> & React.RefAttributes<HTMLDivElement>> & DynamicDataGridColumnMarkers & {
    /** Freezes dynamic descriptors belonging to this family. */
    defineColumns: (factory: (builder: DynamicDataGridColumnBuilder) => readonly DataGridColumnDefinition<Record<string, unknown>>[]) => readonly DataGridColumnDefinition<Record<string, unknown>>[];
};

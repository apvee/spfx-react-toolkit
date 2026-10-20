import type * as React from 'react';
import type { ContentSlot } from '../content-slot.types.internal';
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
/** Locale-formatted chart metrics supplied to the summary string callback.
 * @example
 * ```ts
 * const describe = (values: DataGridSparklineSummaryText) => `Latest ${values.last}; change ${values.change}`;
 * ```
 */
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
/** Localized status and cell strings. Root strings accepts partial overrides; omitted
 * members and blank static strings retain English fallbacks, independently of locale.
 * @example
 * ```tsx
 * const strings: Partial<DataGridStrings> = { empty: 'Nessun dato', noResults: 'Nessun risultato' };
 * <Grid items={rows} getRowId={row => row.id} columns={columns} strings={strings} />;
 * ```
 */
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
/** Exactly one declared column collection, including an explicitly empty collection.
 * @example
 * ```ts
 * const Grid = createDataGrid<{ id: number }>();
 * const collection: DataGridColumnCollection<{ id: number }> = { columns: Grid.defineColumns(() => []) };
 * ```
 */
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
    /** Synchronously returns a stable finite number or string ID unique across supplied rows. Never inferred from position; duplicate/invalid IDs throw. */
    getRowId: (item: TItem) => DataGridRowId;
    /** Controlled sort intent; rows are not sorted internally. Omitted, stale, hidden or non-sortable column intent is inactive. */
    sortState?: FluentDataGridProps['sortState'];
    /** Native Fluent sort intent notification with event and proposed sort state. Update sortState and item ordering externally; return values are ignored. */
    onSortChange?: FluentDataGridProps['onSortChange'];
    /** Enables native single or multiselect intent; omission disables selection. Mutually exclusive with onRowActivate (invalid combinations throw). */
    selectionMode?: SelectionMode;
    /** Controlled selected IDs; omission uses an empty selection. In multiselect mode, IDs outside supplied rows remain in change notifications; single mode proposes the newly selected ID. Update this set externally. */
    selectedItems?: ReadonlySet<DataGridRowId>;
    /** Native Fluent selection notification with event and a fresh proposed selection Set.
     * Update selectedItems externally. Suppressed during loading; return values are ignored. */
    onSelectionChange?: FluentDataGridProps['onSelectionChange'];
    /** Controlled hidden column IDs; omission shows every declared column and stale IDs are ignored. */
    hiddenColumnIds?: ReadonlySet<string>;
    /** Native resizing options keyed by declared column IDs. Descriptor width and
     * minWidth override matching native entries; other native options are preserved.
     * Active when resizableColumns is enabled. */
    columnSizingOptions?: FluentDataGridProps['columnSizingOptions'];
    /** Native offset applied to measured resize container width. */
    containerWidthOffset?: FluentDataGridProps['containerWidthOffset'];
    /** Root Intl locale fallback; omission uses the host Intl default. Column locale overrides it. Invalid Intl locales throw. */
    locale?: DataGridLocale;
    /** Overrides selected localized strings; omitted members and blank static strings retain English defaults. locale does not translate these strings. */
    strings?: Partial<DataGridStrings>;
    /** Initial or refresh loading status. Reuses AutoSkeleton over type-aware,
     * rounded rectangle templates for every visible column and at least four rows.
     * Simulated geometry varies deterministically by row/column identity, without
     * fabricating consumer items or reading absent values. Real cell state remains
     * mounted; selection and activation are disabled until loading completes.
     * No separate loading label or progress indicator is rendered. */
    loading?: boolean;
    /** Literal error content inside the owned MessageBarBody. Object slots use a
     * fixed inner div; refs/events target it. Block content is supported. Undefined,
     * legacy null and false omit the error bar; true requests an empty bar.
     * @example error={{ children: 'Request failed', className: 'error-copy' }}
     */
    error?: ContentSlot<'div'>;
    /** Called by the localized retry button when an error is displayed. Omission hides that button. No fetching or promise handling is performed. */
    onRetry?: () => void;
    /** Selects the empty or no-results presentation. The default localized text
     * is horizontally centered with Fluent medium block padding and natural height. */
    emptyState?: 'empty' | 'noResults';
    /** Custom empty input content, preserving its own presentation.
     * Can compose StateMessage for richer copy, artwork or consumer-owned actions.
     * Object slots create a fixed inner div; refs target it, not the status shell.
     * Nullish values use localized defaults; boolean/empty content suppresses text.
     * Strings are literal. @example emptyContent={{ children: 'No records' }} */
    emptyContent?: ContentSlot<'div'>;
    /** Custom no-results content, preserving its own presentation.
     * Can compose StateMessage without adding search/filter behavior to the grid.
     * Fixed inner div for objects/callbacks; refs/events belong to that div.
     * Nullish values use localized defaults; boolean/empty content suppresses text.
     * @example noResultsContent={{ children: 'Try another query' }} */
    noResultsContent?: ContentSlot<'div'>;
    /** Called with the original row/ID after an unhandled row click or Enter.
     * Prevented events, interactive descendants and loading suppress activation.
     * Mutually exclusive with selectionMode; return values/promises are ignored. */
    onRowActivate?: (event: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, data: {
        /** Original row. */
        item: TItem;
        /** Stable row identity. */
        rowId: DataGridRowId;
    }) => void;
};
/** Same-suite JSX column markers; runtime additionally validates erased JSX ownership.
 * @example
 * ```tsx
 * <Grid items={rows} getRowId={row => row.id}>
 *   <Grid.TextColumn id="title" header="Title" field="title" />
 * </Grid>
 * ```
 */
export interface DataGridColumnMarkers<TItem extends object> {
    /** Declares a text column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.TextColumn id="text" header="text" field="text" />
     * ```
     */
    TextColumn: DataGridColumnMethod<TItem, 'text', React.ReactElement>;
    /** Declares a number column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.NumberColumn id="number" header="number" field="number" />
     * ```
     */
    NumberColumn: DataGridColumnMethod<TItem, 'number', React.ReactElement>;
    /** Declares a boolean column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.BooleanColumn id="boolean" header="boolean" field="boolean" />
     * ```
     */
    BooleanColumn: DataGridColumnMethod<TItem, 'boolean', React.ReactElement>;
    /** Declares a date column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.DateColumn id="date" header="date" field="date" />
     * ```
     */
    DateColumn: DataGridColumnMethod<TItem, 'date', React.ReactElement>;
    /** Declares a dateTime column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.DateTimeColumn id="dateTime" header="dateTime" field="dateTime" />
     * ```
     */
    DateTimeColumn: DataGridColumnMethod<TItem, 'dateTime', React.ReactElement>;
    /** Declares a badge column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.BadgeColumn id="badge" header="badge" field="badge" values={{ open: { label: "Open" } }} />
     * ```
     */
    BadgeColumn: DataGridColumnMethod<TItem, 'badge', React.ReactElement>;
    /** Declares a icon column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.IconColumn id="icon" header="icon" field="icon" values={{ open: { icon: <CheckmarkCircle20Regular />, label: "Open" } }} />
     * ```
     */
    IconColumn: DataGridColumnMethod<TItem, 'icon', React.ReactElement>;
    /** Declares a link column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.LinkColumn id="link" header="link" field="link" href="/details" />
     * ```
     */
    LinkColumn: DataGridColumnMethod<TItem, 'link', React.ReactElement>;
    /** Declares a person column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.PersonColumn id="person" header="person" field="person" getName={person => String(person)} />
     * ```
     */
    PersonColumn: DataGridColumnMethod<TItem, 'person', React.ReactElement>;
    /** Declares a people column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.PeopleColumn id="people" header="people" field="people" getName={person => String(person)} getPersonId={person => String(person)} />
     * ```
     */
    PeopleColumn: DataGridColumnMethod<TItem, 'people', React.ReactElement>;
    /** Declares a tags column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.TagsColumn id="tags" header="tags" field="tags" getTagId={tag => String(tag)} getLabel={tag => String(tag)} />
     * ```
     */
    TagsColumn: DataGridColumnMethod<TItem, 'tags', React.ReactElement>;
    /** Declares a progress column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.ProgressColumn id="progress" header="progress" field="progress" />
     * ```
     */
    ProgressColumn: DataGridColumnMethod<TItem, 'progress', React.ReactElement>;
    /** Declares a rating column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.RatingColumn id="rating" header="rating" field="rating" />
     * ```
     */
    RatingColumn: DataGridColumnMethod<TItem, 'rating', React.ReactElement>;
    /** Declares a sparkline column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.SparklineColumn id="sparkline" header="sparkline" field="sparkline" />
     * ```
     */
    SparklineColumn: DataGridColumnMethod<TItem, 'sparkline', React.ReactElement>;
    /** Declares a custom column with callbacks inferred from its source.
     * @example
     * ```tsx
     * <Grid.Column id="custom" header="custom" field="custom" renderCell={({ value }) => <span>{String(value)}</span>} />
     * ```
     */
    Column: DataGridColumnMethod<TItem, 'custom', React.ReactElement>;
    /** Declares source-free row actions.
     * @example
     * ```tsx
     * <Grid.ActionsColumn id="actions" actions={rowActions} />
     * ```
     */
    ActionsColumn: (props: ActionsColumnProps<TItem>) => React.ReactElement;
}
/** Typed grid and its suite-owned marker/builder family. Construct outside rendering.
 * @example
 * ```ts
 * const Grid: DataGridSuite<{ id: number; title: string }> = createDataGrid();
 * const columns = Grid.defineColumns(c => [c.text({ id: 'title', header: 'Title', field: 'title' })]);
 * ```
 */
export type DataGridSuite<TItem extends object> = React.ForwardRefExoticComponent<DataGridProps<TItem> & React.RefAttributes<HTMLDivElement>> & DataGridColumnMarkers<TItem> & {
    /** Synchronously calls factory with this suite's builder and returns a frozen same-suite descriptor list.
     * Foreign descriptors are rejected here; duplicate IDs are rejected when the grid resolves columns. Use the result as columns on this grid.
     * @example
     * ```ts
     * const columns = Grid.defineColumns(c => [c.text({ id: "title", header: "Title", field: "title" })]);
     * ```
     */
    defineColumns: (factory: (builder: DataGridColumnBuilder<TItem>) => readonly DataGridColumnDefinition<TItem>[]) => readonly DataGridColumnDefinition<TItem>[];
};
/** Explicit dynamic JSX markers; generic source callbacks receive unknown.
 * @example
 * ```tsx
 * <DataGrid items={[{ id: 1, title: 'Plan' }]} getRowId={row => String(row.id)}>
 *   <DataGrid.TextColumn id="title" header="Title" field="title" />
 * </DataGrid>
 * ```
 */
export interface DynamicDataGridColumnMarkers {
    /** Declares a guarded dynamic text column.
     * @example
     * ```tsx
     * <DataGrid.TextColumn id="text" header="text" field="text" />
     * ```
     */
    TextColumn: (props: DynamicDataGridColumnProps<'text'>) => React.ReactElement;
    /** Declares a guarded dynamic number column.
     * @example
     * ```tsx
     * <DataGrid.NumberColumn id="number" header="number" field="number" />
     * ```
     */
    NumberColumn: (props: DynamicDataGridColumnProps<'number'>) => React.ReactElement;
    /** Declares a guarded dynamic boolean column.
     * @example
     * ```tsx
     * <DataGrid.BooleanColumn id="boolean" header="boolean" field="boolean" />
     * ```
     */
    BooleanColumn: (props: DynamicDataGridColumnProps<'boolean'>) => React.ReactElement;
    /** Declares a guarded dynamic date column.
     * @example
     * ```tsx
     * <DataGrid.DateColumn id="date" header="date" field="date" />
     * ```
     */
    DateColumn: (props: DynamicDataGridColumnProps<'date'>) => React.ReactElement;
    /** Declares a guarded dynamic dateTime column.
     * @example
     * ```tsx
     * <DataGrid.DateTimeColumn id="dateTime" header="dateTime" field="dateTime" />
     * ```
     */
    DateTimeColumn: (props: DynamicDataGridColumnProps<'dateTime'>) => React.ReactElement;
    /** Declares a guarded dynamic badge column.
     * @example
     * ```tsx
     * <DataGrid.BadgeColumn id="badge" header="badge" field="badge" values={{ open: { label: "Open" } }} />
     * ```
     */
    BadgeColumn: (props: DynamicDataGridColumnProps<'badge'>) => React.ReactElement;
    /** Declares a guarded dynamic icon column.
     * @example
     * ```tsx
     * <DataGrid.IconColumn id="icon" header="icon" field="icon" values={{ open: { icon: <CheckmarkCircle20Regular />, label: "Open" } }} />
     * ```
     */
    IconColumn: (props: DynamicDataGridColumnProps<'icon'>) => React.ReactElement;
    /** Declares a guarded dynamic link column.
     * @example
     * ```tsx
     * <DataGrid.LinkColumn id="link" header="link" field="link" href="/details" />
     * ```
     */
    LinkColumn: (props: DynamicDataGridColumnProps<'link'>) => React.ReactElement;
    /** Declares a guarded dynamic person column.
     * @example
     * ```tsx
     * <DataGrid.PersonColumn id="person" header="person" field="person" getName={person => String(person)} />
     * ```
     */
    PersonColumn: (props: DynamicDataGridColumnProps<'person'>) => React.ReactElement;
    /** Declares a guarded dynamic people column.
     * @example
     * ```tsx
     * <DataGrid.PeopleColumn id="people" header="people" field="people" getName={person => String(person)} getPersonId={person => String(person)} />
     * ```
     */
    PeopleColumn: (props: DynamicDataGridColumnProps<'people'>) => React.ReactElement;
    /** Declares a guarded dynamic tags column.
     * @example
     * ```tsx
     * <DataGrid.TagsColumn id="tags" header="tags" field="tags" getTagId={tag => String(tag)} getLabel={tag => String(tag)} />
     * ```
     */
    TagsColumn: (props: DynamicDataGridColumnProps<'tags'>) => React.ReactElement;
    /** Declares a guarded dynamic progress column.
     * @example
     * ```tsx
     * <DataGrid.ProgressColumn id="progress" header="progress" field="progress" />
     * ```
     */
    ProgressColumn: (props: DynamicDataGridColumnProps<'progress'>) => React.ReactElement;
    /** Declares a guarded dynamic rating column.
     * @example
     * ```tsx
     * <DataGrid.RatingColumn id="rating" header="rating" field="rating" />
     * ```
     */
    RatingColumn: (props: DynamicDataGridColumnProps<'rating'>) => React.ReactElement;
    /** Declares a guarded dynamic sparkline column.
     * @example
     * ```tsx
     * <DataGrid.SparklineColumn id="sparkline" header="sparkline" field="sparkline" />
     * ```
     */
    SparklineColumn: (props: DynamicDataGridColumnProps<'sparkline'>) => React.ReactElement;
    /** Declares a guarded dynamic custom column.
     * @example
     * ```tsx
     * <DataGrid.Column id="custom" header="custom" field="custom" renderCell={({ value }) => <span>{String(value)}</span>} />
     * ```
     */
    Column: (props: DynamicDataGridColumnProps<'custom'>) => React.ReactElement;
    /** Declares source-free dynamic row actions.
     * @example
     * ```tsx
     * <DataGrid.ActionsColumn id="actions" actions={rowActions} />
     * ```
     */
    ActionsColumn: (props: ActionsColumnProps<Record<string, unknown>>) => React.ReactElement;
}
/** Dynamic grid for dictionary rows. Typed factories remain typed even for index signatures.
 * @example
 * ```tsx
 * const columns = DataGrid.defineColumns(c => [c.text({ id: 'title', header: 'Title', field: 'title' })]);
 * <DataGrid items={[{ id: 1, title: 'Plan' }]} getRowId={row => String(row.id)} columns={columns} />;
 * ```
 */
export type DynamicDataGridSuite = React.ForwardRefExoticComponent<DataGridProps<Record<string, unknown>> & React.RefAttributes<HTMLDivElement>> & DynamicDataGridColumnMarkers & {
    /** Synchronously calls factory with the dynamic builder and returns a frozen descriptor list owned by DataGrid. Foreign descriptors are rejected here; duplicate IDs are rejected when the grid resolves columns.
     * @example
     * ```ts
     * const columns = DataGrid.defineColumns(c => [c.text({ id: "title", header: "Title", field: "title" })]);
     * ```
     */
    defineColumns: (factory: (builder: DynamicDataGridColumnBuilder) => readonly DataGridColumnDefinition<Record<string, unknown>>[]) => readonly DataGridColumnDefinition<Record<string, unknown>>[];
};

/**
 * @packageDocumentation
 * @docPage controls/data-grid
 */
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
/**
 * Locale-formatted chart metrics supplied to the summary string callback.
 * @example
 * ```ts
 * import { DataGridSparklineSummaryText } from '@apvee/spfx-react-toolkit';
 * const describe = (values: DataGridSparklineSummaryText) => `Latest ${values.last}; change ${values.change}`;
 * ```
 */
export interface DataGridSparklineSummaryText {
    /**
     * First observation.
     */
    first: string;
    /**
     * Latest observation.
     */
    last: string;
    /**
     * Minimum observation.
     */
    minimum: string;
    /**
     * Maximum observation.
     */
    maximum: string;
    /**
     * Overflow-safe mean.
     */
    average: string;
    /**
     * Formatted change or placeholder when independently unrepresentable.
     */
    change: string;
    /**
     * Formatted percent change or placeholder for zero denominator/unrepresentable result.
     */
    changePercent: string;
}
/**
 * Localized status and cell strings. Root strings accepts partial overrides; omitted
 * members and blank static strings retain host-selected native resources, independently of the data locale. Nonblank static
 * strings retain whitespace; invalid object/member types throw. No toolkit provider is required.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { DataGridStrings, createDataGrid } from '@apvee/spfx-react-toolkit';
 * const Grid = createDataGrid<{ id: number; title: string }>();
 * const rows = [{ id: 1, title: 'Plan' }];
 * const columns = Grid.defineColumns(c => [c.text({ id: 'title', header: 'Title', field: 'title' })]);
 * const strings: Partial<DataGridStrings> = { empty: 'Nessun dato', noResults: 'Nessun risultato' };
 * <Grid items={rows} getRowId={row => row.id} columns={columns} strings={strings} />;
 * ```
 */
export interface DataGridStrings {
    /**
     * Empty input status.
     */
    empty: string;
    /**
     * No matches status.
     */
    noResults: string;
    /**
     * No visible columns status.
     */
    noVisibleColumns: string;
    /**
     * Retained loading label for source compatibility; the grid shows skeletons
     * and aria-busy rather than a visible loading message.
     */
    loading: string;
    /**
     * Default actions heading.
     */
    actions: string;
    /**
     * Overflow actions trigger.
     */
    moreActions: string;
    /**
     * True label shared by text, icon and tooltip.
     */
    yes: string;
    /**
     * False label shared by text, icon and tooltip.
     */
    no: string;
    /**
     * Empty tag list label.
     */
    noTags: string;
    /**
     * Accessible trigger label for hidden tags.
     * @param count - Number of hidden tags available through the overflow trigger.
     * @returns Localized plain text for the corresponding accessible label or chart summary.
     */
    tagOverflow: (count: number) => string;
    /**
     * Accessible vote count when a rating column explicitly supplies count. Defaults to the native resource vote-count forms.
     * @param count - Original non-negative whole number of votes, including zero.
     * @param formattedCount - Count formatted with the column or grid locale.
     * @returns Localized count text, distinct from the rating value and maximum.
     */
    ratingCount: (count: number, formattedCount: string) => string;
    /**
     * Empty chart series label.
     */
    noObservations: string;
    /**
     * Invalid present cell data label.
     */
    invalidValue: string;
    /**
     * Select all supplied rows.
     */
    selectAllRows: string;
    /**
     * Accessible selection label for an original row ID.
     * @param rowId - Original stable row identifier, without display-position substitution.
     * @returns Localized plain text for the corresponding accessible label or chart summary.
     */
    selectRow: (rowId: DataGridRowId) => string;
    /**
     * Accessible chart summary from locale-formatted original observations.
     * Defaults to host-selected native metric labels. Derived metrics independently use an em dash
     * when unavailable; showValue=false does not suppress this summary.
     * @param values - Locale-formatted original observations and independently available derived metrics.
     * @returns Localized plain text for the corresponding accessible label or chart summary.
     */
    sparklineSummary: (values: DataGridSparklineSummaryText) => string;
    /**
     * Retry label.
     */
    retry: string;
}
/**
 * Exactly one declared column collection, including an explicitly empty collection.
 * @example
 * ```ts
 * import { DataGridColumnCollection, createDataGrid } from '@apvee/spfx-react-toolkit';
 * const Grid = createDataGrid<{ id: number }>();
 * const collection: DataGridColumnCollection<{ id: number }> = { columns: Grid.defineColumns(() => []) };
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 */
export type DataGridColumnCollection<TItem extends object> = {
    /**
     * Same-suite descriptors, including an explicitly empty list.
     */
    columns: readonly DataGridColumnDefinition<TItem>[];
    /**
     * Excluded when descriptors are provided.
     * Must be omitted entirely in this branch. An own children key, including
     * children: undefined, conflicts with columns and throws at runtime.
     */
    children?: never;
} | {
    /**
     * Same-suite JSX markers, including an explicitly empty collection.
     */
    children: React.ReactNode;
    /**
     * Excluded when JSX is declared.
     * Must be omitted entirely in this branch. An own columns key, including
     * columns: undefined, conflicts with children and throws at runtime.
     */
    columns?: never;
};
/**
 * Controlled grid configuration. Input order is preserved.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createDataGrid } from '@apvee/spfx-react-toolkit';
 * const Grid = createDataGrid<{ id: number; title: string }>();
 * const rows = [{ id: 1, title: 'Plan' }];
 * const columns = Grid.defineColumns(c => [c.text({ id: 'title', header: 'Title', field: 'title' })]);
 * <Grid items={rows} getRowId={row => row.id} columns={columns} />
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
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
    /**
     * Immutable source rows; objects are passed to callbacks unchanged.
     */
    items: readonly TItem[];
    /**
     * Native div click event.
     * When omitted, no consumer root-click handler is attached.
     */
    onClick?: React.MouseEventHandler<HTMLDivElement>;
    /**
     * Synchronously returns a stable finite number or string ID unique across supplied rows. Never inferred from position; duplicate/invalid IDs throw.
     * @param item - Original row object, unchanged from the supplied items array.
     * @returns Stable string or finite numeric identity, unique within the relevant row/element collection.
     */
    getRowId: (item: TItem) => DataGridRowId;
    /**
     * Controlled sort intent; rows are not sorted internally. Omitted, stale, hidden or non-sortable column intent is inactive.
     */
    sortState?: FluentDataGridProps['sortState'];
    /**
     * Native Fluent sort intent notification with event and proposed sort state. Update sortState and item ordering externally; return values are ignored.
     * When omitted, no consumer receives sort requests; supplied rows and controlled sort state remain unchanged.
     */
    onSortChange?: FluentDataGridProps['onSortChange'];
    /**
     * Enables native single or multiselect intent; when omitted, disables selection. Mutually exclusive with onRowActivate (invalid combinations throw).
     */
    selectionMode?: SelectionMode;
    /**
     * Controlled selected IDs; when omitted, uses an empty selection. In multiselect mode, IDs outside supplied rows remain in change notifications; single mode proposes the newly selected ID. Update this set externally.
     */
    selectedItems?: ReadonlySet<DataGridRowId>;
    /**
     * Native Fluent selection notification with event and a fresh proposed selection Set.
     * Update selectedItems externally. Suppressed during loading; return values are ignored.
     * When omitted, no consumer receives selection requests; selectedItems remains authoritative.
     */
    onSelectionChange?: FluentDataGridProps['onSelectionChange'];
    /**
     * Controlled hidden column IDs; when omitted, shows every declared column and stale IDs are ignored.
     */
    hiddenColumnIds?: ReadonlySet<string>;
    /**
     * Native resizing options keyed by declared column IDs. Descriptor width and
     * minWidth override matching native entries; other native options are preserved.
     * Active when resizableColumns is enabled.
     * When omitted, only declared column width/minWidth overrides and native resizing defaults are used.
     */
    columnSizingOptions?: FluentDataGridProps['columnSizingOptions'];
    /**
     * Native offset applied to measured resize container width.
     * When omitted, the native Fluent grid offset applies.
     */
    containerWidthOffset?: FluentDataGridProps['containerWidthOffset'];
    /**
     * Root Intl locale fallback; when omitted, uses the host Intl default. Column locale overrides it. Invalid Intl locales throw.
     */
    locale?: DataGridLocale;
    /**
     * Overrides selected localized strings; omitted members and blank static strings retain the host-selected native baseline. Nonblank whitespace is preserved; invalid types throw. locale governs data formatting independently.
     */
    strings?: Partial<DataGridStrings>;
    /**
     * Initial or refresh loading status. Reuses AutoSkeleton over type-aware,
     * rounded rectangle templates for every visible column and at least four rows.
     * Simulated geometry varies deterministically by row/column identity, without
     * fabricating consumer items or reading absent values. Real cell state remains
     * mounted; selection and activation are disabled until loading completes.
     * No separate loading label or progress indicator is rendered.
     * When omitted, the grid is not loading and interaction is available.
     * @defaultValue false
     */
    loading?: boolean;
    /**
     * Literal error content inside the owned MessageBarBody. Object slots use a
     * fixed inner div; refs/events target it. Block content is supported. Undefined,
     * legacy null and false omit the error bar; true requests an empty bar.
     * @example
     * ```ts
     * import { DataGridProps } from '@apvee/spfx-react-toolkit';
     * const error: DataGridProps<{ id: number }>['error'] = { children: 'Request failed', className: 'error-copy' };
     * ```
     */
    error?: ContentSlot<'div'>;
    /**
     * Called by the localized retry button when an error is displayed. When omitted, hides that button. No fetching or promise handling is performed.
     */
    onRetry?: () => void;
    /**
     * Selects the empty or no-results presentation; when omitted, selects empty. The default localized text
     * is horizontally centered with Fluent medium block padding and natural height.
     * @defaultValue 'empty'
     */
    emptyState?: 'empty' | 'noResults';
    /**
     * Custom empty input content, preserving its own presentation.
     * Can compose StateMessage for richer copy, artwork or consumer-owned actions.
     * Object slots create a fixed inner div; refs target it, not the status shell.
     * Nullish values use localized defaults; boolean/empty content suppresses text.
     * Strings are literal.
     *
     * When omitted or undefined, uses the localized strings.empty content.
     * @example
     * ```ts
     * import { DataGridProps } from '@apvee/spfx-react-toolkit';
     * const emptyContent: DataGridProps<{ id: number }>['emptyContent'] = { children: 'No records' };
     * ```
     */
    emptyContent?: ContentSlot<'div'>;
    /**
     * Custom no-results content, preserving its own presentation.
     * Can compose StateMessage without adding search/filter behavior to the grid.
     * Fixed inner div for objects/callbacks; refs/events belong to that div.
     * Nullish values use localized defaults; boolean/empty content suppresses text.
     *
     * When omitted or undefined, uses the localized strings.noResults content.
     * @example
     * ```ts
     * import { DataGridProps } from '@apvee/spfx-react-toolkit';
     * const noResultsContent: DataGridProps<{ id: number }>['noResultsContent'] = { children: 'Try another query' };
     * ```
     */
    noResultsContent?: ContentSlot<'div'>;
    /**
     * Called with the original row/ID after an unhandled row click or Enter.
     * Prevented events, interactive descendants and loading suppress activation.
     * Mutually exclusive with selectionMode; return values/promises are ignored.
     * When omitted, rows have no toolkit activation callback or activation mode.
     * @param data - Original row object and its stable row ID.
     * @param event - Unhandled row click or Enter keyboard event.
     */
    onRowActivate?: (event: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, /** Activated original row and its stable identity. */ data: {
        /**
         * Original row.
         */
        item: TItem;
        /**
         * Stable row identity.
         */
        rowId: DataGridRowId;
    }) => void;
};
/**
 * Same-suite JSX column markers; runtime additionally validates erased JSX ownership.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createDataGrid } from '@apvee/spfx-react-toolkit';
 * const Grid = createDataGrid<{ id: number; title: string }>();
 * const rows = [{ id: 1, title: 'Plan' }];
 * <Grid items={rows} getRowId={row => row.id}>
 *   <Grid.TextColumn id="title" header="Title" field="title" />
 * </Grid>
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 */
export interface DataGridColumnMarkers<TItem extends object> {
    /**
     * Declares a text column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; text: string }>();
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.TextColumn id="text" header="text" field="text" />
     * </Grid>
     * ```
     */
    TextColumn: DataGridColumnMethod<TItem, 'text', React.ReactElement>;
    /**
     * Declares a number column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; number: number }>();
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.NumberColumn id="number" header="number" field="number" />
     * </Grid>
     * ```
     */
    NumberColumn: DataGridColumnMethod<TItem, 'number', React.ReactElement>;
    /**
     * Declares a boolean column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; boolean: boolean }>();
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.BooleanColumn id="boolean" header="boolean" field="boolean" />
     * </Grid>
     * ```
     */
    BooleanColumn: DataGridColumnMethod<TItem, 'boolean', React.ReactElement>;
    /**
     * Declares a date column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; date: string }>();
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.DateColumn id="date" header="date" field="date" />
     * </Grid>
     * ```
     */
    DateColumn: DataGridColumnMethod<TItem, 'date', React.ReactElement>;
    /**
     * Declares a dateTime column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; dateTime: string }>();
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.DateTimeColumn id="dateTime" header="dateTime" field="dateTime" />
     * </Grid>
     * ```
     */
    DateTimeColumn: DataGridColumnMethod<TItem, 'dateTime', React.ReactElement>;
    /**
     * Declares a badge column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; badge: 'open' }>();
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.BadgeColumn id="badge" header="badge" field="badge" values={{ open: { label: "Open" } }} />
     * </Grid>
     * ```
     */
    BadgeColumn: DataGridColumnMethod<TItem, 'badge', React.ReactElement>;
    /**
     * Declares a icon column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * import { CheckmarkCircle20Regular } from '@fluentui/react-icons';
     * const Grid = createDataGrid<{ id: number; icon: 'open' }>();
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.IconColumn id="icon" header="icon" field="icon" values={{ open: { icon: <CheckmarkCircle20Regular />, label: "Open" } }} />
     * </Grid>
     * ```
     */
    IconColumn: DataGridColumnMethod<TItem, 'icon', React.ReactElement>;
    /**
     * Declares a link column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; link: string }>();
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.LinkColumn id="link" header="link" field="link" href="/details" />
     * </Grid>
     * ```
     */
    LinkColumn: DataGridColumnMethod<TItem, 'link', React.ReactElement>;
    /**
     * Declares a person column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; person: { id: string; name: string } }>();
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.PersonColumn id="person" header="person" field="person" getName={person => String(person)} />
     * </Grid>
     * ```
     */
    PersonColumn: DataGridColumnMethod<TItem, 'person', React.ReactElement>;
    /**
     * Declares a people column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; people: readonly { id: string; name: string }[] }>();
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.PeopleColumn id="people" header="people" field="people" getName={person => String(person)} getPersonId={person => String(person)} />
     * </Grid>
     * ```
     */
    PeopleColumn: DataGridColumnMethod<TItem, 'people', React.ReactElement>;
    /**
     * Declares a tags column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; tags: readonly string[] }>();
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.TagsColumn id="tags" header="tags" field="tags" getTagId={tag => String(tag)} getLabel={tag => String(tag)} />
     * </Grid>
     * ```
     */
    TagsColumn: DataGridColumnMethod<TItem, 'tags', React.ReactElement>;
    /**
     * Declares a progress column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; progress: number }>();
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.ProgressColumn id="progress" header="progress" field="progress" />
     * </Grid>
     * ```
     */
    ProgressColumn: DataGridColumnMethod<TItem, 'progress', React.ReactElement>;
    /**
     * Declares a rating column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; rating: number }>();
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.RatingColumn id="rating" header="rating" field="rating" />
     * </Grid>
     * ```
     */
    RatingColumn: DataGridColumnMethod<TItem, 'rating', React.ReactElement>;
    /**
     * Declares a sparkline column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; sparkline: readonly number[] }>();
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.SparklineColumn id="sparkline" header="sparkline" field="sparkline" />
     * </Grid>
     * ```
     */
    SparklineColumn: DataGridColumnMethod<TItem, 'sparkline', React.ReactElement>;
    /**
     * Declares a custom column with callbacks inferred from its source.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; custom: string }>();
     * <Grid.Column id="custom" header="custom" field="custom" renderCell={({ value }) => <span>{String(value)}</span>} />
     * ```
     */
    Column: DataGridColumnMethod<TItem, 'custom', React.ReactElement>;
    /**
     * Declares source-free row actions.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGridAction, createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number }>();
     * const rowActions: DataGridAction<{ id: number }>[] = [{ id: 'open', label: 'Open', placement: 'inline', onClick: () => console.log('Open requested') }];
     * <Grid items={[]} getRowId={row => row.id}>
     *   <Grid.ActionsColumn id="actions" actions={rowActions} />
     * </Grid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    ActionsColumn: (props: ActionsColumnProps<TItem>) => React.ReactElement;
}
/**
 * Typed grid and its suite-owned marker/builder family. Construct outside rendering.
 * @example
 * ```ts
 * import { DataGridSuite, createDataGrid } from '@apvee/spfx-react-toolkit';
 * const Grid: DataGridSuite<{ id: number; title: string }> = createDataGrid();
 * const columns = Grid.defineColumns(c => [c.text({ id: 'title', header: 'Title', field: 'title' })]);
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 * @param props - Controlled row data, same-suite columns and native grid presentation props.
 * @returns Native Fluent grid output, or the configured empty/no-visible-column status instead of the grid root.
 */
export type DataGridSuite<TItem extends object> = React.ForwardRefExoticComponent<DataGridProps<TItem> & React.RefAttributes<HTMLDivElement>> & DataGridColumnMarkers<TItem> & {
    /**
     * Synchronously calls factory with this suite's builder and returns a frozen same-suite descriptor list.
     * Foreign descriptors are rejected here; duplicate IDs are rejected when the grid resolves columns. Use the result as columns on this grid.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; title: string }>();
     * const columns = Grid.defineColumns(c => [c.text({ id: "title", header: "Title", field: "title" })]);
     * ```
     *
     * @param factory - Pure synchronous factory receiving this suite’s column builder; return only descriptors from that builder.
     * @returns Ordered immutable descriptors from this suite, validated without sorting or transforming the row data.
     */
    defineColumns: (/** Pure factory called once with this suite’s builder.
     * @param builder - Builder creating descriptors owned by this grid suite.
     * @returns Ordered same-suite descriptors; sources and callbacks remain caller-owned.
     */ factory: (builder: DataGridColumnBuilder<TItem>) => readonly DataGridColumnDefinition<TItem>[]) => readonly DataGridColumnDefinition<TItem>[];
};
/**
 * Explicit dynamic JSX markers; generic source callbacks receive unknown.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { DataGrid } from '@apvee/spfx-react-toolkit';
 * <DataGrid items={[{ id: 1, title: 'Plan' }]} getRowId={row => String(row.id)}>
 *   <DataGrid.TextColumn id="title" header="Title" field="title" />
 * </DataGrid>
 * ```
 */
export interface DynamicDataGridColumnMarkers {
    /**
     * Declares a guarded dynamic text column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.TextColumn id="text" header="text" field="text" />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    TextColumn: (props: DynamicDataGridColumnProps<'text'>) => React.ReactElement;
    /**
     * Declares a guarded dynamic number column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.NumberColumn id="number" header="number" field="number" />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    NumberColumn: (props: DynamicDataGridColumnProps<'number'>) => React.ReactElement;
    /**
     * Declares a guarded dynamic boolean column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.BooleanColumn id="boolean" header="boolean" field="boolean" />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    BooleanColumn: (props: DynamicDataGridColumnProps<'boolean'>) => React.ReactElement;
    /**
     * Declares a guarded dynamic date column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.DateColumn id="date" header="date" field="date" />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    DateColumn: (props: DynamicDataGridColumnProps<'date'>) => React.ReactElement;
    /**
     * Declares a guarded dynamic dateTime column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.DateTimeColumn id="dateTime" header="dateTime" field="dateTime" />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    DateTimeColumn: (props: DynamicDataGridColumnProps<'dateTime'>) => React.ReactElement;
    /**
     * Declares a guarded dynamic badge column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.BadgeColumn id="badge" header="badge" field="badge" values={{ open: { label: "Open" } }} />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    BadgeColumn: (props: DynamicDataGridColumnProps<'badge'>) => React.ReactElement;
    /**
     * Declares a guarded dynamic icon column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * import { CheckmarkCircle20Regular } from '@fluentui/react-icons';
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.IconColumn id="icon" header="icon" field="icon" values={{ open: { icon: <CheckmarkCircle20Regular />, label: "Open" } }} />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    IconColumn: (props: DynamicDataGridColumnProps<'icon'>) => React.ReactElement;
    /**
     * Declares a guarded dynamic link column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.LinkColumn id="link" header="link" field="link" href="/details" />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    LinkColumn: (props: DynamicDataGridColumnProps<'link'>) => React.ReactElement;
    /**
     * Declares a guarded dynamic person column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.PersonColumn id="person" header="person" field="person" getName={person => String(person)} />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    PersonColumn: (props: DynamicDataGridColumnProps<'person'>) => React.ReactElement;
    /**
     * Declares a guarded dynamic people column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.PeopleColumn id="people" header="people" field="people" getName={person => String(person)} getPersonId={person => String(person)} />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    PeopleColumn: (props: DynamicDataGridColumnProps<'people'>) => React.ReactElement;
    /**
     * Declares a guarded dynamic tags column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.TagsColumn id="tags" header="tags" field="tags" getTagId={tag => String(tag)} getLabel={tag => String(tag)} />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    TagsColumn: (props: DynamicDataGridColumnProps<'tags'>) => React.ReactElement;
    /**
     * Declares a guarded dynamic progress column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.ProgressColumn id="progress" header="progress" field="progress" />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    ProgressColumn: (props: DynamicDataGridColumnProps<'progress'>) => React.ReactElement;
    /**
     * Declares a guarded dynamic rating column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.RatingColumn id="rating" header="rating" field="rating" />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    RatingColumn: (props: DynamicDataGridColumnProps<'rating'>) => React.ReactElement;
    /**
     * Declares a guarded dynamic sparkline column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.SparklineColumn id="sparkline" header="sparkline" field="sparkline" />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    SparklineColumn: (props: DynamicDataGridColumnProps<'sparkline'>) => React.ReactElement;
    /**
     * Declares a guarded dynamic custom column.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * <DataGrid.Column id="custom" header="custom" field="custom" renderCell={({ value }) => <span>{String(value)}</span>} />
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    Column: (props: DynamicDataGridColumnProps<'custom'>) => React.ReactElement;
    /**
     * Declares source-free dynamic row actions.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid, DataGridAction } from '@apvee/spfx-react-toolkit';
     * const rowActions: DataGridAction<Record<string, unknown>>[] = [{ id: 'open', label: 'Open', placement: 'inline', onClick: () => console.log('Open requested') }];
     * <DataGrid items={[]} getRowId={row => String(row.id)}>
     *   <DataGrid.ActionsColumn id="actions" actions={rowActions} />
     * </DataGrid>
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Same-suite declarative JSX marker consumed by the grid; do not render it independently.
     */
    ActionsColumn: (props: ActionsColumnProps<Record<string, unknown>>) => React.ReactElement;
}
/**
 * Dynamic grid for dictionary rows. Typed factories remain typed even for index signatures.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { DataGrid } from '@apvee/spfx-react-toolkit';
 * const columns = DataGrid.defineColumns(c => [c.text({ id: 'title', header: 'Title', field: 'title' })]);
 * <DataGrid items={[{ id: 1, title: 'Plan' }]} getRowId={row => String(row.id)} columns={columns} />;
 * ```
 *
 * @param props - Controlled row data, same-suite columns and native grid presentation props.
 * @returns Native Fluent grid output, or the configured empty/no-visible-column status instead of the grid root.
 */
export type DynamicDataGridSuite = React.ForwardRefExoticComponent<DataGridProps<Record<string, unknown>> & React.RefAttributes<HTMLDivElement>> & DynamicDataGridColumnMarkers & {
    /**
     * Synchronously calls factory with the dynamic builder and returns a frozen descriptor list owned by DataGrid. Foreign descriptors are rejected here; duplicate IDs are rejected when the grid resolves columns.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.text({ id: "title", header: "Title", field: "title" })]);
     * ```
     *
     * @param factory - Pure synchronous factory receiving this suite’s column builder; return only descriptors from that builder.
     * @returns Ordered immutable descriptors from this suite, validated without sorting or transforming the row data.
     */
    defineColumns: (/** Pure factory called once with this suite’s builder.
     * @param builder - Builder creating descriptors owned by this grid suite.
     * @returns Ordered same-suite descriptors; sources and callbacks remain caller-owned.
     */ factory: (builder: DynamicDataGridColumnBuilder) => readonly DataGridColumnDefinition<Record<string, unknown>>[]) => readonly DataGridColumnDefinition<Record<string, unknown>>[];
};

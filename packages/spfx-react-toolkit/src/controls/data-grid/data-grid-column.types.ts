import type * as React from 'react';
import type {
  AvatarProps,
  BadgeProps,
  LinkProps,
  PersonaProps,
  ProgressBarProps,
  RatingDisplayProps,
  TagProps,
} from '@fluentui/react-components';
/** One supported cell presentation. */
export type DataGridColumnKind =
    | 'text' | 'number' | 'boolean' | 'date' | 'dateTime'
    | 'badge' | 'icon' | 'link' | 'person' | 'people' | 'tags'
    | 'progress' | 'rating' | 'sparkline' | 'actions' | 'custom';
/** Intl locale or ordered fallback locales. */
export type DataGridLocale = string | readonly string[];
/** Stable original row identity. Numbers must be finite. */
export type DataGridRowId = string | number;
/** Cell callback context. External nullable data is supported at the source boundary. */
export interface DataGridValueContext<TItem, TValue> {
    /** Original row object. */
    item: TItem;
    /** Original source value; legacy nullable source fields remain nullable. */
    value: TValue | undefined;
}
/** Field names whose non-nullish domain is compatible with a column. */
export type CompatibleDataGridField<TItem, TValue> = {
    [K in Extract<keyof TItem, string>]:
        [NonNullable<TItem[K]>] extends [never] ? never :
        NonNullable<TItem[K]> extends TValue ? K : never;
}[Extract<keyof TItem, string>];
/** Synchronous result accepting legacy nullable data; new absence should use undefined. */
// Legacy source rows/accessors may contain null; never used as an internal default.
// eslint-disable-next-line @rushstack/no-new-null
export type DataGridSourceResult<V> = V | null | undefined;
/** Exactly one synchronous source. Nullable fields/accessors interoperate with legacy row models. */
export type DataGridValueSource<TItem, TValue> = {
    /** Compatible top-level field; mutually exclusive with accessor. */
    field: CompatibleDataGridField<TItem, TValue>;
    /** Excluded for field-backed sources. */
    accessor?: never;
} | {
    /** Excluded for computed sources. */
    field?: never;
    /** Pure synchronous source returning original values. */
    accessor: (item: TItem) => DataGridSourceResult<TValue>;
};
/** Shared column configuration. Sources and callbacks must be pure. */
export interface DataGridColumnCommonProps<TItem, TValue> {
    /** Unique non-empty column ID. */
    id: string;
    /** Visible header. */
    header: React.ReactNode;
    /** Accessible name for a non-text header. */
    headerAriaLabel?: string;
    /** Plain-text header description. */
    headerTooltip?: string;
    /** Plain-text description of valid original data. Missing values are normalized to
     * undefined for this callback, including legacy null; invalid values skip the
     * callback and use a static tooltip or the localized invalid label. */
    tooltip?: string | ((context: DataGridValueContext<TItem, TValue>) => string | undefined);
    /** Missing source presentation; defaults to an em dash. */
    placeholder?: React.ReactNode;
    /** Present but invalid source presentation. */
    invalidPlaceholder?: React.ReactNode;
    /** Enables controlled sorting intent, without changing item order. */
    sortable?: boolean;
    /** Preferred width in pixels. */
    width?: number;
    /** Minimum width in pixels. */
    minWidth?: number;
    /** Overrides the root locale. */
    locale?: DataGridLocale;
}
/** Badge presentation mapped from an exact scalar value. */
export interface DataGridBadgeValue {
    /** Display and accessible label. */
    label: string;
    /** Fluent color. */
    color?: BadgeProps['color'];
    /** Fluent appearance; defaults to tint. */
    appearance?: BadgeProps['appearance'];
    /** Fluent shape. */
    shape?: BadgeProps['shape'];
    /** Fluent size. */
    size?: BadgeProps['size'];
    /** Optional icon. */
    icon?: React.ReactNode;
}
/** Informative icon and its text alternative. */
export interface DataGridIconValue {
    /** Visible icon. */
    icon: React.ReactNode;
    /** Accessible label. */
    label: string;
}
/** Tag presentation mapped from an exact scalar ID. */
export interface DataGridTagValue {
    /** Display and accessible label. */
    label: string;
    /** Fluent appearance. */
    appearance?: TagProps['appearance'];
    /** Fluent shape. */
    shape?: TagProps['shape'];
    /** Fluent size. */
    size?: TagProps['size'];
    /** Optional icon. */
    icon?: React.ReactNode;
}
/** Exact dictionaries for string domains; numeric keys require a Map. */
export type DataGridMappedValues<TValue, TPresentation> = unknown extends TValue ? Readonly<Partial<Record<string, TPresentation>>> | ReadonlyMap<string | number, TPresentation> : [
    Extract<TValue, number>
] extends [
    never
] ? Readonly<Partial<Record<Extract<TValue, string>, TPresentation>>> | ReadonlyMap<Extract<TValue, string>, TPresentation> : ReadonlyMap<Extract<TValue, string | number>, TPresentation>;
/** Original data supplied to link navigation. */
export interface DataGridLinkNavigateData<TItem> {
    /** Original row. */
    item: TItem;
    /** Original stable ID. */
    rowId: DataGridRowId;
    /** Valid original text. */
    value: string;
    /** Resolved destination. */
    href: string;
}
/** Person display callbacks retain their inferred person type. */
export interface DataGridPersonOptions<TItem, TPerson> {
    /** Required person name. */
    getName: (person: TPerson, item: TItem) => string;
    /** Optional secondary text. */
    getSecondaryText?: (person: TPerson, item: TItem) => string | undefined;
    /** Avatar options except its owned name. */
    getAvatarProps?: (person: TPerson, item: TItem) => Omit<AvatarProps, 'name'>;
    /** Native presence configuration. */
    getPresenceProps?: (person: TPerson, item: TItem) => NonNullable<PersonaProps['presence']>;
    /** Native Persona presentation options; textAlignment defaults to center. */
    personaProps?: Omit<PersonaProps, 'name' | 'primaryText' | 'secondaryText' | 'avatar' | 'presence'>;
}
/** A source-order chart observation. */
export interface DataGridSparklinePoint {
    /** Numeric or date X coordinate. */
    x: number | Date;
    /** Finite original observation. */
    y: number;
}
/** Metric available to consumers reading a chart summary. */
export type DataGridSparklineMetric = 'latest' | 'minimum' | 'maximum' | 'average' | 'change' | 'changePercent';
/** Immutable chart summary; each derived change is independently optional. */
export interface DataGridSparklineSummary {
    /** First observation. */
    readonly first: number;
    /** Last observation. */
    readonly last: number;
    /** Minimum observation. */
    readonly minimum: number;
    /** Maximum observation. */
    readonly maximum: number;
    /** Overflow-safe arithmetic mean. */
    readonly average: number;
    /** Last minus first, absent when not representable. */
    readonly change: number | undefined;
    /** Percent change, absent for zero denominator or unrepresentable result. */
    readonly changePercent: number | undefined;
}
/** Original row and action identities. */
export interface DataGridActionData<TItem> {
    /** Original row. */
    item: TItem;
    /** Stable row ID. */
    rowId: DataGridRowId;
    /** Stable action ID. */
    actionId: string;
}
/** One synchronous row action. */
export interface DataGridAction<TItem> {
    /** Non-empty unique action ID. */
    id: string;
    /** Visible and accessible label. */
    label: string;
    /** Optional icon. */
    icon?: React.ReactNode;
    /** Placement in the row. */
    placement: 'inline' | 'menu';
    /** Static or row-dependent visibility. */
    visible?: boolean | ((item: TItem) => boolean);
    /** Static or row-dependent disabled state. */
    disabled?: boolean | ((item: TItem) => boolean);
    /** Static or row-dependent busy state. */
    loading?: boolean | ((item: TItem) => boolean);
    /** Receives the original item and IDs. */
    onClick: (event: React.MouseEvent<HTMLElement>, data: DataGridActionData<TItem>) => void;
}
/** Supported source domains; dynamic mode explicitly uses unknown instead. */
export interface DataGridColumnDomains {
    /** Text source. */
    text: string;
    /** Numeric source. */
    number: number;
    /** Boolean source. */
    boolean: boolean;
    /** Calendar source. */
    date: Date | string;
    /** Instant source. */
    dateTime: Date | string;
    /** Mapped scalar source. */
    badge: string | number;
    /** Mapped scalar source. */
    icon: string | number;
    /** Link text source. */
    link: string;
    /** Arbitrary person source. */
    person: unknown;
    /** Immutable people source. */
    people: readonly unknown[];
    /** Immutable tags source. */
    tags: readonly unknown[];
    /** Numeric progress. */
    progress: number;
    /** Numeric rating. */
    rating: number;
    /** Immutable chart observations. */
    sparkline: readonly number[] | readonly DataGridSparklinePoint[];
    /** Custom source. */
    custom: unknown;
}
type ElementOf<T> = T extends readonly (infer E)[] ? E : unknown;
/** Intl number presentation. */
export interface DataGridNumberOptions {
    /** Formatting changes presentation only. */
    format?: Intl.NumberFormatOptions;
}
/** Readonly boolean presentation. */
export interface DataGridBooleanOptions {
    /** Text or a vertically centered native 20px informative icon; defaults to text. */
    display?: 'text' | 'icon';
    /** Visible and accessible true label. */
    trueLabel?: string;
    /** Visible and accessible false label. */
    falseLabel?: string;
}
/** Absolute calendar or instant presentation. */
export interface DataGridDateOptions {
    /** Explicit IANA zone for Date objects and dateTime instants; defaults to UTC.
     * Calendar-only YYYY-MM-DD strings always display their UTC calendar day. */
    timeZone?: string;
    /** Date formatting; configure the zone with timeZone. */
    format?: Omit<Intl.DateTimeFormatOptions, 'timeZone'>;
}
/** Badge mapping keeps the exact source scalar domain. */
export interface DataGridBadgeOptions<V> {
    /** Exact scalar values mapped to Fluent badge presentations. */
    values: DataGridMappedValues<V, DataGridBadgeValue>;
}
/** Informative icon mapping keeps exact source keys. */
export interface DataGridIconOptions<V> {
    /** Exact scalar values mapped to icons and labels. */
    values: DataGridMappedValues<V, DataGridIconValue>;
}
/** Link destination and navigation intent. */
export interface DataGridLinkOptions<TItem> {
    /** Static or row-aware destination for valid text. */
    href: string | ((item: TItem, value: string) => string | undefined);
    /** Native anchor presentation; the grid owns its element, content and navigation. */
    linkProps?: Omit<Extract<LinkProps, { as?: 'a' }>, 'as' | 'children' | 'href' | 'onClick' | 'dangerouslySetInnerHTML'>;
    /** Receives original item, row ID, value and destination. */
    onNavigate?: (event: React.MouseEvent<HTMLAnchorElement>, data: DataGridLinkNavigateData<TItem>) => void;
}
/** Multiple-person identity and native stack layout with recoverable overflow. */
export interface DataGridPeopleOptions<TItem, TPerson> extends DataGridPersonOptions<TItem, TPerson> {
    /** Returns a stable ID unique within this cell. */
    getPersonId: (person: TPerson, item: TItem) => DataGridRowId;
    /** Maximum visible avatars before the recoverable overflow affordance. */
    maxVisible?: number;
}
/** Exact scalar tag maps; object-tag callback IDs use a guarded string/number map. */
export type DataGridTagValues<TTag> = [TTag] extends [string | number]
    ? DataGridMappedValues<TTag, DataGridTagValue>
    : DataGridMappedValues<unknown, DataGridTagValue>;
/** Exact tag mappings or paired callbacks for scalar or object tags.
 * Missing configuration and incomplete callback pairs are rejected by types and runtime.
 */
export type DataGridTagsOptions<TItem, TTag> = {
    /** Maximum visible tags before recoverable overflow. */
    maxVisible?: number;
} & ({
    /** Exact scalar IDs mapped to Fluent presentations; numeric keys require a Map. */
    values: DataGridTagValues<TTag>;
    /** Excluded when using scalar map identities. */
    getTagId?: never;
    /** Labels are supplied by the map. */
    getLabel?: never;
} | {
    /** Optional mapped presentation for IDs returned by getTagId. */
    values?: DataGridTagValues<TTag>;
    /** Stable scalar ID for each tag. Supply with getLabel; without a mapping,
     * callbacks also read scalar tags. With a mapping, scalar IDs stay original. */
    getTagId: (tag: TTag, item: TItem) => DataGridRowId;
    /** Explicit tag label. Supply with getTagId. Mapped labels take precedence;
     * scalar tags use these callbacks when no mapping is declared. */
    getLabel: (tag: TTag, item: TItem) => string;
});
/** Determinate readonly progress presentation. */
export interface DataGridProgressOptions extends DataGridNumberOptions {
    /** Positive finite upper bound; defaults to one. */
    max?: number;
    /** Display the localized value next to the indicator. */
    showValue?: boolean;
    /** Native visual options, excluding owned value and accessibility. */
    progressBarProps?: Pick<ProgressBarProps, 'color' | 'shape' | 'thickness'>;
}
/** Readonly filled marigold stars with an accessible value/maximum and optional vote count. */
export interface DataGridRatingOptions<TItem> {
    /** Whole number greater than one; defaults to five. */
    max?: number;
    /** Optional static or row-aware non-negative whole number of votes for the
     * accessible description only. Zero is included; undefined omits the count.
     * Localize its text with strings.ratingCount. */
    count?: number | ((context: DataGridValueContext<TItem, number>) => number | undefined);
    /** Native visual options, excluding owned value and range.
     * Defaults to marigold, Star20Filled and native large (20px);
     * explicit consumer overrides are retained. Numeric slots are hidden. */
    ratingDisplayProps?: Pick<RatingDisplayProps, 'color' | 'compact' | 'icon' | 'size'>;
}
/** Official Charts Sparkline presentation. */
export interface DataGridSparklineOptions<TItem> extends DataGridNumberOptions {
    /** Show the latest observation; accessible summary remains when false. */
    showValue?: boolean;
    /** Static or summary-aware Fluent color string. */
    color?: string | ((context: {
        /** Original row. */
        item: TItem;
        /** Immutable summary of valid observations. */
        summary: DataGridSparklineSummary;
    }) => string);
}
/** Custom content preserves original source and JSX identity. */
export interface DataGridCustomOptions<TItem, V> {
    /** Renders one valid value with original row identity. */
    renderCell: (context: DataGridValueContext<TItem, V> & {
        /** Stable original row ID. */
        rowId: DataGridRowId;
    }) => React.ReactNode;
}
/** Specific options selected by presentation kind and inferred source domain. */
export type DataGridColumnOptions<TItem, K extends keyof DataGridColumnDomains, V> =
    K extends 'number' ? DataGridNumberOptions :
    K extends 'boolean' ? DataGridBooleanOptions :
    K extends 'date' | 'dateTime' ? DataGridDateOptions :
    K extends 'badge' ? DataGridBadgeOptions<V> :
    K extends 'icon' ? DataGridIconOptions<V> :
    K extends 'link' ? DataGridLinkOptions<TItem> :
    K extends 'person' ? DataGridPersonOptions<TItem, V> :
    K extends 'people' ? DataGridPeopleOptions<TItem, ElementOf<V>> :
    K extends 'tags' ? DataGridTagsOptions<TItem, ElementOf<V>> :
    K extends 'progress' ? DataGridProgressOptions :
    K extends 'rating' ? DataGridRatingOptions<TItem> :
    K extends 'sparkline' ? DataGridSparklineOptions<TItem> :
    K extends 'custom' ? DataGridCustomOptions<TItem, V> : object;
/** Props for one strongly typed value column. */
export type DataGridValueColumnProps<TItem, K extends keyof DataGridColumnDomains, V = DataGridColumnDomains[K]> = DataGridColumnCommonProps<TItem, V> & DataGridValueSource<TItem, V> & DataGridColumnOptions<TItem, K, V>;
/** Actions deliberately have no source or sorting option. */
export type ActionsColumnProps<TItem> = Omit<DataGridColumnCommonProps<TItem, undefined>, 'header' | 'sortable'> & {
    /** Optional actions heading; defaults to the localized actions label. */
    header?: React.ReactNode;
    /** Ordered synchronous actions evaluated for the original row. */
    actions: readonly DataGridAction<TItem>[];
    /** Actions cannot read a field. */
    field?: never;
    /** Actions cannot read an accessor. */
    accessor?: never;
};
declare const descriptorType: unique symbol;
/** Immutable opaque descriptor owned by one suite. Create it with defineColumns.
 * @example const columns = Grid.defineColumns(c => [c.text({ id:'name', header:'Name', field:'name' })]);
 */
export interface DataGridColumnDefinition<TItem extends object> {
    /** Display kind, without exposing captured callbacks. */
    readonly kind: DataGridColumnKind;
    /** Stable unique column ID. */
    readonly id: string;
    /** Compile-time invariant row ownership; never constructed by consumers. */
    readonly [descriptorType]: (item: TItem) => TItem;
}
// TS 5.3 inference fence: only a source selects V, not its mapping or callbacks.
type SourceOnly<T> = [T][T extends unknown ? 0 : never];
type FieldProps<TItem, K extends keyof DataGridColumnDomains, F extends Extract<keyof TItem, string>> = DataGridColumnCommonProps<TItem, TItem[F]> & {
    field: F;
    accessor?: never;
} & DataGridColumnOptions<TItem, K, NonNullable<TItem[F]>>;
/** Field/accessor overloads infer callbacks from their own source, without domain widening. */
export type DataGridColumnMethod<TItem extends object, K extends keyof DataGridColumnDomains, R> = K extends 'badge' | 'icon' ? MappedColumnMethod<TItem, K, R> : SourceColumnMethod<TItem, K, R>;
interface SourceColumnMethod<TItem extends object, K extends keyof DataGridColumnDomains, R> {
    /** Configure a field; callbacks infer its original domain. @param props Column configuration. @returns Same-suite descriptor or JSX marker. */
    <F extends CompatibleDataGridField<TItem, DataGridColumnDomains[K]>>(props: FieldProps<TItem, K, F>): R;
    /** Configure a computed source. @param props Pure accessor and presentation. @returns Same-suite descriptor or JSX marker. */
    <A extends (item: TItem) => DataGridSourceResult<DataGridColumnDomains[K]>>(props: DataGridColumnCommonProps<TItem, ReturnType<A>> & {
        field?: never;
        accessor: A;
    } & DataGridColumnOptions<TItem, K, SourceOnly<NonNullable<ReturnType<A>>>>): R;
}
interface MappedColumnMethod<TItem extends object, K extends 'badge' | 'icon', R> {
    /** Configure a field; callbacks infer its original domain. @param props Column configuration. @returns Same-suite descriptor or JSX marker. */
    <F extends CompatibleDataGridField<TItem, string | number>>(props: FieldProps<TItem, K, F>): R;
    /** Configure an exact mapped accessor. @param props Source and exact map. @returns Same-suite descriptor or JSX marker. */
    <V extends string>(props: DataGridColumnCommonProps<TItem, V> & {
        field?: never;
        accessor: (item: TItem) => DataGridSourceResult<V>;
    } & DataGridColumnOptions<TItem, K, SourceOnly<V>>): R;
    /** Configure an exact mapped accessor. @param props Source and exact map. @returns Same-suite descriptor or JSX marker. */
    <V extends string | number>(props: DataGridColumnCommonProps<TItem, V> & {
        field?: never;
        accessor: (item: TItem) => DataGridSourceResult<V>;
    } & DataGridColumnOptions<TItem, K, SourceOnly<V>>): R;
}
/** Builder for all sixteen presentation kinds.
 * @example const columns = Grid.defineColumns(c => [c.number({ id:"total", header:"Total", field:"total" })]);
 */
export interface DataGridColumnBuilder<TItem extends object> {
    /** Captures a text source and creates a same-suite immutable descriptor. */
    text: DataGridColumnMethod<TItem, 'text', DataGridColumnDefinition<TItem>>;
    /** Captures a number source and creates a same-suite immutable descriptor. */
    number: DataGridColumnMethod<TItem, 'number', DataGridColumnDefinition<TItem>>;
    /** Captures a boolean source and creates a same-suite immutable descriptor. */
    boolean: DataGridColumnMethod<TItem, 'boolean', DataGridColumnDefinition<TItem>>;
    /** Captures a date source and creates a same-suite immutable descriptor. */
    date: DataGridColumnMethod<TItem, 'date', DataGridColumnDefinition<TItem>>;
    /** Captures a dateTime source and creates a same-suite immutable descriptor. */
    dateTime: DataGridColumnMethod<TItem, 'dateTime', DataGridColumnDefinition<TItem>>;
    /** Captures a badge source and creates a same-suite immutable descriptor. */
    badge: DataGridColumnMethod<TItem, 'badge', DataGridColumnDefinition<TItem>>;
    /** Captures a icon source and creates a same-suite immutable descriptor. */
    icon: DataGridColumnMethod<TItem, 'icon', DataGridColumnDefinition<TItem>>;
    /** Captures a link source and creates a same-suite immutable descriptor. */
    link: DataGridColumnMethod<TItem, 'link', DataGridColumnDefinition<TItem>>;
    /** Captures a person source and creates a same-suite immutable descriptor. */
    person: DataGridColumnMethod<TItem, 'person', DataGridColumnDefinition<TItem>>;
    /** Captures a people source and creates a same-suite immutable descriptor. */
    people: DataGridColumnMethod<TItem, 'people', DataGridColumnDefinition<TItem>>;
    /** Captures a tags source and creates a same-suite immutable descriptor. */
    tags: DataGridColumnMethod<TItem, 'tags', DataGridColumnDefinition<TItem>>;
    /** Captures a progress source and creates a same-suite immutable descriptor. */
    progress: DataGridColumnMethod<TItem, 'progress', DataGridColumnDefinition<TItem>>;
    /** Captures a rating source and creates a same-suite immutable descriptor. */
    rating: DataGridColumnMethod<TItem, 'rating', DataGridColumnDefinition<TItem>>;
    /** Captures a sparkline source and creates a same-suite immutable descriptor. */
    sparkline: DataGridColumnMethod<TItem, 'sparkline', DataGridColumnDefinition<TItem>>;
    /** Captures a custom source and creates a same-suite immutable descriptor. */
    custom: DataGridColumnMethod<TItem, 'custom', DataGridColumnDefinition<TItem>>;
    /** Creates a source-free action descriptor. */
    actions: (props: ActionsColumnProps<TItem>) => DataGridColumnDefinition<TItem>;
}
/** Explicit dynamic source props. Generic callbacks receive unknown. */
export type DynamicDataGridColumnProps<K extends keyof DataGridColumnDomains> = DataGridColumnCommonProps<Record<string, unknown>, unknown> & ({
    /** Explicit dynamic top-level field. */
    field: string;
    /** Excluded for field sources. */
    accessor?: never;
} | {
    /** Excluded for computed sources. */
    field?: never;
    /** Pure synchronous source; its return remains unknown. */
    accessor: (item: Record<string, unknown>) => unknown;
}) & DataGridColumnOptions<Record<string, unknown>, K, unknown>;
/** Explicit dynamic builder; generic source callbacks receive unknown. */
export interface DynamicDataGridColumnBuilder {
    /** Creates a guarded dynamic text descriptor. */
    text: (props: DynamicDataGridColumnProps<'text'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a guarded dynamic number descriptor. */
    number: (props: DynamicDataGridColumnProps<'number'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a guarded dynamic boolean descriptor. */
    boolean: (props: DynamicDataGridColumnProps<'boolean'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a guarded dynamic date descriptor. */
    date: (props: DynamicDataGridColumnProps<'date'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a guarded dynamic dateTime descriptor. */
    dateTime: (props: DynamicDataGridColumnProps<'dateTime'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a guarded dynamic badge descriptor. */
    badge: (props: DynamicDataGridColumnProps<'badge'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a guarded dynamic icon descriptor. */
    icon: (props: DynamicDataGridColumnProps<'icon'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a guarded dynamic link descriptor. */
    link: (props: DynamicDataGridColumnProps<'link'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a guarded dynamic person descriptor. */
    person: (props: DynamicDataGridColumnProps<'person'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a guarded dynamic people descriptor. */
    people: (props: DynamicDataGridColumnProps<'people'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a guarded dynamic tags descriptor. */
    tags: (props: DynamicDataGridColumnProps<'tags'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a guarded dynamic progress descriptor. */
    progress: (props: DynamicDataGridColumnProps<'progress'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a guarded dynamic rating descriptor. */
    rating: (props: DynamicDataGridColumnProps<'rating'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a guarded dynamic sparkline descriptor. */
    sparkline: (props: DynamicDataGridColumnProps<'sparkline'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a guarded dynamic custom descriptor. */
    custom: (props: DynamicDataGridColumnProps<'custom'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /** Creates a source-free dynamic action descriptor. */
    actions: (props: ActionsColumnProps<Record<string, unknown>>) => DataGridColumnDefinition<Record<string, unknown>>;
}

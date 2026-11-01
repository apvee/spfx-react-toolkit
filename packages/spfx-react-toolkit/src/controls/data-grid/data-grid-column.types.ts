/**
 * @packageDocumentation
 * @docPage controls/data-grid
 */
import type { ContentSlot } from '../content-slot.types.internal';
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
/**
 * One supported cell presentation.
 * @example
 * ```ts
 * import { DataGridColumnKind } from '@apvee/spfx-react-toolkit';
 * const kind: DataGridColumnKind = 'number';
 * ```
 */
export type DataGridColumnKind =
    | 'text' | 'number' | 'boolean' | 'date' | 'dateTime'
    | 'badge' | 'icon' | 'link' | 'person' | 'people' | 'tags'
    | 'progress' | 'rating' | 'sparkline' | 'actions' | 'custom';
/**
 * Intl locale or ordered fallback locales.
 * @example
 * ```ts
 * import { DataGridLocale } from '@apvee/spfx-react-toolkit';
 * const locale: DataGridLocale = ['it-IT', 'en-US'];
 * ```
 */
export type DataGridLocale = string | readonly string[];
/**
 * Stable original row identity. Numbers must be finite.
 * @example
 * ```ts
 * import { DataGridRowId } from '@apvee/spfx-react-toolkit';
 * const id: DataGridRowId = 'project-42';
 * ```
 */
export type DataGridRowId = string | number;
/**
 * Cell callback context. External nullable data is supported at the source boundary.
 * @example
 * ```ts
 * import { DataGridValueContext } from '@apvee/spfx-react-toolkit';
 * const tooltip = ({ item, value }: DataGridValueContext<{ id: number }, string>) =>
 *   value === undefined ? undefined : `${item.id}: ${value}`;
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 * @typeParam TValue - Source value domain used to restrict fields and preserve callback value types.
 */
export interface DataGridValueContext<TItem, TValue> {
    /**
     * Original row object.
     */
    item: TItem;
    /**
     * Original source value, with undefined for missing data where the callback supports it.
     * Tooltip callbacks normalize legacy null to undefined and skip invalid sources; the declared generic domain may still include legacy nullable types.
     */
    value: TValue | undefined;
}
/**
 * Field names whose non-nullish domain is compatible with a column.
 * @example
 * ```ts
 * import { DataGridValueSource } from '@apvee/spfx-react-toolkit';
 * type NumericField = DataGridValueSource<{ name: string; total: number }, number>['field'];
 * const field: NumericField = 'total';
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 * @typeParam TValue - Source value domain used to restrict fields and preserve callback value types.
 */
export type CompatibleDataGridField<TItem, TValue> = {
    [K in Extract<keyof TItem, string>]:
        [NonNullable<TItem[K]>] extends [never] ? never :
        NonNullable<TItem[K]> extends TValue ? K : never;
}[Extract<keyof TItem, string>];
/**
 * Synchronous result accepting legacy nullable data; new absence should use undefined.
 * @example
 * ```ts
 * import { DataGridValueSource } from '@apvee/spfx-react-toolkit';
 * const source: DataGridValueSource<{ total: number }, number> = { accessor: () => undefined };
 * ```
 *
 * @typeParam V - Source domain retained by the column presentation and callbacks.
 */
// Legacy source rows/accessors may contain null; never used as an internal default.
// eslint-disable-next-line @rushstack/no-new-null
export type DataGridSourceResult<V> = V | null | undefined;
/**
 * Exactly one synchronous source. Nullable fields/accessors interoperate with legacy row models.
 * @example
 * ```ts
 * import { DataGridValueSource } from '@apvee/spfx-react-toolkit';
 * const source: DataGridValueSource<{ total: number }, number> = {
 *   accessor: item => item.total * 2,
 * };
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 * @typeParam TValue - Source value domain used to restrict fields and preserve callback value types.
 */
export type DataGridValueSource<TItem, TValue> = {
    /**
     * Compatible top-level field; mutually exclusive with accessor.
     */
    field: CompatibleDataGridField<TItem, TValue>;
    /**
     * Excluded for field-backed sources.
     * Must be omitted or undefined in this branch.
     */
    accessor?: never;
} | {
    /**
     * Excluded for computed sources.
     * Must be omitted or undefined in this branch.
     */
    field?: never;
    /**
     * Pure synchronous source returning original values.
     * @param item - Original row object, unchanged from the supplied items array.
     * @returns Original synchronous source value, or undefined/legacy null for a missing cell.
     */
    accessor: (item: TItem) => DataGridSourceResult<TValue>;
};
/**
 * Shared column configuration. Sources and callbacks must be pure.
 * @example
 * ```ts
 * import { DataGridColumnCommonProps } from '@apvee/spfx-react-toolkit';
 * const common: DataGridColumnCommonProps<{ name: string }, string> = {
 *   id: 'name', header: 'Name', sortable: true, width: 180,
 * };
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 * @typeParam TValue - Source value domain used to restrict fields and preserve callback value types.
 */
export interface DataGridColumnCommonProps<TItem, TValue> {
    /**
     * Unique non-empty column ID.
     */
    id: string;
    /**
     * Literal visible header content. Object/callback slots use a fixed inner
     * span; refs target it rather than the native header/sort button. Supply phrasing,
     * noninteractive content; headerAriaLabel owns the native accessible name.
     * @example
     * ```ts
     * import { DataGridColumnCommonProps } from '@apvee/spfx-react-toolkit';
     * const header: DataGridColumnCommonProps<{ title: string }, string>['header'] = { children: 'Name', className: 'heading-copy' };
     * ```
     */
    header: ContentSlot<'span'>;
    /**
     * Accessible name for a non-text header.
     * When omitted, no explicit accessible-name override is added; provide meaningful header content.
     */
    headerAriaLabel?: string;
    /**
     * Plain-text header description.
     * When omitted, no separate header description is configured.
     */
    headerTooltip?: string;
    /**
     * Plain-text description of valid original data. Missing values are normalized to
     * undefined for this callback, including legacy null; invalid values skip the
     * callback and use a static tooltip or the localized invalid label.
     * @param context - Original row and source value; missing tooltip values are normalized to undefined.
     * @returns Plain tooltip text, or undefined to use the grid’s normal formatted-value tooltip.
     */
    tooltip?: string | ((context: DataGridValueContext<TItem, TValue>) => string | undefined);
    /**
     * Literal missing value inside the existing text span; when omitted, defaults to an em dash
     * (sparkline uses noObservations). Object/callback slots create an inner span;
     * refs/events target it. Supply phrasing content. Nullish values retain defaults;
     * boolean/empty text suppresses them.
     * @example
     * ```ts
     * import { DataGridColumnCommonProps } from '@apvee/spfx-react-toolkit';
     * const placeholder: DataGridColumnCommonProps<{ title: string }, string>['placeholder'] = { children: 'Missing' };
     * ```
     */
    placeholder?: ContentSlot<'span'>;
    /**
     * Literal invalid value inside the existing text span. Object/callback slots
     * create a fixed inner span; refs/events target it. Supply phrasing content.
     * Nullish values use invalidValue; boolean/empty text suppresses the default.
     *
     * When omitted or undefined, uses strings.invalidValue.
     * @example
     * ```ts
     * import { DataGridColumnCommonProps } from '@apvee/spfx-react-toolkit';
     * const invalidPlaceholder: DataGridColumnCommonProps<{ title: string }, string>['invalidPlaceholder'] = { children: 'Invalid', className: 'invalid-copy' };
     * ```
     */
    invalidPlaceholder?: ContentSlot<'span'>;
    /**
     * Enables controlled sorting intent when true; omitted/false disables this column's sort affordance. Supplied items retain their order.
     */
    sortable?: boolean;
    /**
     * Preferred finite, non-negative width in pixels; must be at least minWidth. Used by native resizing when enabled. When omitted, retains native sizing.
     */
    width?: number;
    /**
     * Finite, non-negative minimum width in pixels. Used by native resizing when enabled.
     * When omitted, nonresizable columns use min(width ?? 100, 100) pixels; native resizing retains its peer default.
     */
    minWidth?: number;
    /**
     * Overrides the root Intl locale for this column. When omitted, uses the root locale, then the host Intl default.
     */
    locale?: DataGridLocale;
}
/**
 * Badge presentation mapped from an exact scalar value.
 * @example
 * ```ts
 * import { DataGridBadgeValue } from '@apvee/spfx-react-toolkit';
 * const active: DataGridBadgeValue = { label: 'Active', color: 'success' };
 * ```
 */
export interface DataGridBadgeValue {
    /**
     * Display and accessible label.
     */
    label: string;
    /**
     * Fluent color.
     * When omitted, the corresponding native Fluent Badge default applies.
     */
    color?: BadgeProps['color'];
    /**
     * Fluent appearance; when omitted, defaults to tint.
     * @defaultValue 'tint'
     */
    appearance?: BadgeProps['appearance'];
    /**
     * Fluent shape.
     * When omitted, the corresponding native Fluent Badge default applies.
     */
    shape?: BadgeProps['shape'];
    /**
     * Fluent size.
     * When omitted, the corresponding native Fluent Badge default applies.
     */
    size?: BadgeProps['size'];
    /**
     * Optional icon.
     * When omitted, no decorative icon is rendered.
     */
    icon?: React.ReactNode;
}
/**
 * Informative icon and its text alternative.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { DataGridIconValue } from '@apvee/spfx-react-toolkit';
 * import { CheckmarkCircle20Regular } from '@fluentui/react-icons';
 * const icon: DataGridIconValue = { icon: <CheckmarkCircle20Regular />, label: 'Approved' };
 * ```
 */
export interface DataGridIconValue {
    /**
     * Visible icon.
     */
    icon: React.ReactNode;
    /**
     * Accessible label.
     */
    label: string;
}
/**
 * Tag presentation mapped from an exact scalar ID.
 * @example
 * ```ts
 * import { DataGridTagValue } from '@apvee/spfx-react-toolkit';
 * const tag: DataGridTagValue = { label: 'Priority', appearance: 'brand' };
 * ```
 */
export interface DataGridTagValue {
    /**
     * Display and accessible label.
     */
    label: string;
    /**
     * Fluent appearance.
     * When omitted, the corresponding native Fluent Tag default applies.
     */
    appearance?: TagProps['appearance'];
    /**
     * Fluent shape.
     * When omitted, the corresponding native Fluent Tag default applies.
     */
    shape?: TagProps['shape'];
    /**
     * Fluent size.
     * When omitted, the corresponding native Fluent Tag default applies.
     */
    size?: TagProps['size'];
    /**
     * Optional icon.
     * When omitted, no decorative icon is rendered.
     */
    icon?: React.ReactNode;
}
/**
 * Exact dictionaries for string domains; numeric keys require a Map.
 * @example
 * ```ts
 * import { DataGridBadgeValue, DataGridMappedValues } from '@apvee/spfx-react-toolkit';
 * const values: DataGridMappedValues<1 | 2, DataGridBadgeValue> = new Map([
 *   [1, { label: 'Open' }], [2, { label: 'Closed' }],
 * ]);
 * ```
 *
 * @typeParam TPresentation - Presentation object stored for each matching scalar key.
 * @typeParam TValue - Source value domain used to restrict fields and preserve callback value types.
 */
export type DataGridMappedValues<TValue, TPresentation> = unknown extends TValue ? Readonly<Partial<Record<string, TPresentation>>> | ReadonlyMap<string | number, TPresentation> : [
    Extract<TValue, number>
] extends [
    never
] ? Readonly<Partial<Record<Extract<TValue, string>, TPresentation>>> | ReadonlyMap<Extract<TValue, string>, TPresentation> : ReadonlyMap<Extract<TValue, string | number>, TPresentation>;
/**
 * Original data supplied to link navigation.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { DataGridLinkNavigateData } from '@apvee/spfx-react-toolkit';
 * const onNavigate = (event: React.MouseEvent<HTMLAnchorElement>, data: DataGridLinkNavigateData<{ id: number }>) => {
 *   event.preventDefault();
 *   console.log(data.href);
 * };
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 */
export interface DataGridLinkNavigateData<TItem> {
    /**
     * Original row.
     */
    item: TItem;
    /**
     * Original stable ID.
     */
    rowId: DataGridRowId;
    /**
     * Valid original text.
     */
    value: string;
    /**
     * Resolved destination.
     */
    href: string;
}
/**
 * Person display callbacks retain their inferred person type.
 * @example
 * ```ts
 * import { DataGridPersonOptions } from '@apvee/spfx-react-toolkit';
 * const person: DataGridPersonOptions<{ id: number }, { name: string; email: string }> = {
 *   getName: value => value.name, getSecondaryText: value => value.email,
 * };
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 * @typeParam TPerson - Original person element passed to name, avatar and presence callbacks.
 */
export interface DataGridPersonOptions<TItem, TPerson> {
    /**
     * Synchronously returns the person's display and accessible name from the original person and row.
     * @param item - Original row object, unchanged from the supplied items array.
     * @param person - Original person source or array element for this row.
     * @returns Plain display and accessible name for the original source element.
     */
    getName: (person: TPerson, item: TItem) => string;
    /**
     * Synchronously returns secondary display text for the original person and row. When omitted or undefined, omits it.
     * @param item - Original row object, unchanged from the supplied items array.
     * @param person - Original person source or array element for this row.
     * @returns Supporting person text, or undefined to omit that text.
     */
    getSecondaryText?: (person: TPerson, item: TItem) => string | undefined;
    /**
     * Synchronously returns Fluent avatar options for the original person and row; the grid owns name. When omitted, uses native avatar defaults.
     * @param item - Original row object, unchanged from the supplied items array.
     * @param person - Original person source or array element for this row.
     * @returns Native Fluent avatar options; the grid supplies the person name separately.
     */
    getAvatarProps?: (person: TPerson, item: TItem) => Omit<AvatarProps, 'name'>;
    /**
     * Synchronously returns Fluent presence configuration for the original person and row. When omitted, renders no configured presence.
     * @param item - Original row object, unchanged from the supplied items array.
     * @param person - Original person source or array element for this row.
     * @returns Native Fluent Persona presence-slot configuration.
     */
    getPresenceProps?: (person: TPerson, item: TItem) => NonNullable<PersonaProps['presence']>;
    /**
     * Native Persona presentation options; textAlignment when omitted, defaults to center.
     * @defaultValue Native Persona options with textAlignment set to center.
     */
    personaProps?: Omit<PersonaProps, 'name' | 'primaryText' | 'secondaryText' | 'avatar' | 'presence'>;
}
/**
 * A source-order chart observation.
 * @example
 * ```ts
 * import { DataGridSparklinePoint } from '@apvee/spfx-react-toolkit';
 * const point: DataGridSparklinePoint = { x: new Date('2026-01-01T00:00:00Z'), y: 42 };
 * ```
 */
export interface DataGridSparklinePoint {
    /**
     * Finite numeric or valid Date X coordinate; observations retain source order.
     */
    x: number | Date;
    /**
     * Finite original observation.
     */
    y: number;
}
/**
 * Metric available to consumers reading a chart summary.
 * @example
 * ```ts
 * import { DataGridSparklineMetric } from '@apvee/spfx-react-toolkit';
 * const metric: DataGridSparklineMetric = 'changePercent';
 * ```
 */
export type DataGridSparklineMetric = 'latest' | 'minimum' | 'maximum' | 'average' | 'change' | 'changePercent';
/**
 * Immutable chart summary; each derived change is independently optional.
 * @example
 * ```ts
 * import { DataGridSparklineSummary } from '@apvee/spfx-react-toolkit';
 * const color = ({ summary }: { summary: DataGridSparklineSummary }) =>
 *   summary.change !== undefined && summary.change < 0 ? 'red' : 'green';
 * ```
 */
export interface DataGridSparklineSummary {
    /**
     * First observation.
     */
    readonly first: number;
    /**
     * Last observation.
     */
    readonly last: number;
    /**
     * Minimum observation.
     */
    readonly minimum: number;
    /**
     * Maximum observation.
     */
    readonly maximum: number;
    /**
     * Overflow-safe arithmetic mean.
     */
    readonly average: number;
    /**
     * Last minus first, absent when not representable.
     */
    readonly change: number | undefined;
    /**
     * (last - first) / abs(first) as a percentage, absent for a zero first value or an unrepresentable result.
     */
    readonly changePercent: number | undefined;
}
/**
 * Original row and action identities.
 * @example
 * ```ts
 * import { DataGridActionData } from '@apvee/spfx-react-toolkit';
 * const open = (data: DataGridActionData<{ title: string }>) => {
 *   console.log(data.rowId, data.item.title);
 * };
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 */
export interface DataGridActionData<TItem> {
    /**
     * Original row.
     */
    item: TItem;
    /**
     * Stable row ID.
     */
    rowId: DataGridRowId;
    /**
     * Stable action ID.
     */
    actionId: string;
}
/**
 * One synchronous row action.
 * @example
 * ```ts
 * import { DataGridAction } from '@apvee/spfx-react-toolkit';
 * const action: DataGridAction<{ id: number }> = {
 *   id: 'open', label: 'Open', placement: 'inline',
 *   onClick: (_event, { item }) => console.log(item.id),
 * };
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 */
export interface DataGridAction<TItem> {
    /**
     * Non-empty unique action ID.
     */
    id: string;
    /**
     * Visible and accessible label.
     */
    label: string;
    /**
     * Optional icon.
     * When omitted, no decorative icon is rendered.
     */
    icon?: React.ReactNode;
    /**
     * Required placement: inline renders a visible button; menu renders an item in the row's overflow menu.
     */
    placement: 'inline' | 'menu';
    /**
     * Static or synchronous original-row visibility; omitted means visible. False omits the action.
     * @param item - Original row object, unchanged from the supplied items array.
     * @returns True to include this action; false to omit it.
     */
    visible?: boolean | ((item: TItem) => boolean);
    /**
     * Static or synchronous original-row disabled state; omitted means enabled unless loading is true.
     * @param item - Original row object, unchanged from the supplied items array.
     * @returns True to disable action activation; loading also disables activation.
     */
    disabled?: boolean | ((item: TItem) => boolean);
    /**
     * Static or synchronous original-row busy state; omitted means idle. True shows a spinner and disables activation.
     * @param item - Original row object, unchanged from the supplied items array.
     * @returns True to show busy presentation and disable action activation.
     */
    loading?: boolean | ((item: TItem) => boolean);
    /**
     * Called for an enabled, idle action with its click event and original row/IDs.
     * Row propagation is stopped before invocation. Return values/promises are not observed; manage async loading and errors externally.
     * @param data - Original row, stable row ID and activated action ID.
     * @param event - Native enabled-action click event after row propagation is stopped.
     */
    onClick: (event: React.MouseEvent<HTMLElement>, data: DataGridActionData<TItem>) => void;
}
/**
 * Source domains used to select typed column overloads; dynamic mode explicitly uses unknown instead.
 * Calendar strings require YYYY-MM-DD; instant strings require an ISO timestamp with Z/offset.
 * @example
 * ```ts
 * import { DataGridValueSource } from '@apvee/spfx-react-toolkit';
 * type DateSource = ReturnType<NonNullable<DataGridValueSource<{ date: string }, Date | string>['accessor']>>;
 * const day: DateSource = '2026-01-01';
 * ```
 */
export interface DataGridColumnDomains {
    /**
     * Text source.
     */
    text: string;
    /**
     * Numeric source.
     */
    number: number;
    /**
     * Boolean source.
     */
    boolean: boolean;
    /**
     * Calendar source.
     */
    date: Date | string;
    /**
     * Instant source.
     */
    dateTime: Date | string;
    /**
     * Mapped scalar source.
     */
    badge: string | number;
    /**
     * Mapped scalar source.
     */
    icon: string | number;
    /**
     * Link text source.
     */
    link: string;
    /**
     * Arbitrary person source.
     */
    person: unknown;
    /**
     * Immutable people source.
     */
    people: readonly unknown[];
    /**
     * Immutable tags source.
     */
    tags: readonly unknown[];
    /**
     * Finite progress value from zero through the configured max (default one).
     */
    progress: number;
    /**
     * Finite rating value from zero through the configured max (default five); fractional values are supported.
     */
    rating: number;
    /**
     * Immutable chart observations.
     */
    sparkline: readonly number[] | readonly DataGridSparklinePoint[];
    /**
     * Custom source.
     */
    custom: unknown;
}
/**
 * Infers an immutable array element for person/tag option callbacks; other values resolve to unknown.
 * @typeParam T - Array source whose element type is projected into person or tag callbacks.
 * @example
 * ```ts
 * import { createDataGrid } from '@apvee/spfx-react-toolkit';
 * const Grid = createDataGrid<{ tags: readonly { id: string; label: string }[] }>();
 * const columns = Grid.defineColumns(c => [c.tags({ id: 'tags', header: 'Tags', field: 'tags', getTagId: tag => tag.id, getLabel: tag => tag.label })]);
 * ```
 */
type ElementOf<T> = T extends readonly (infer E)[] ? E : unknown;
/**
 * Intl number presentation.
 * @example
 * ```ts
 * import { DataGridNumberOptions } from '@apvee/spfx-react-toolkit';
 * const options: DataGridNumberOptions = { format: { style: 'currency', currency: 'EUR' } };
 * ```
 */
export interface DataGridNumberOptions {
    /**
     * Formatting changes presentation only.
     * When omitted, uses Intl.NumberFormat defaults for the resolved column/grid locale.
     */
    format?: Intl.NumberFormatOptions;
}
/**
 * Readonly boolean presentation.
 * @example
 * ```ts
 * import { DataGridBooleanOptions } from '@apvee/spfx-react-toolkit';
 * const options: DataGridBooleanOptions = { display: 'icon', trueLabel: 'Enabled', falseLabel: 'Disabled' };
 * ```
 */
export interface DataGridBooleanOptions {
    /**
     * Text or a vertically centered native 20px informative icon; when omitted, defaults to text.
     * @defaultValue 'text'
     */
    display?: 'text' | 'icon';
    /**
     * Visible and accessible true label; when omitted, uses strings.yes from the host-selected native resource or grid override.
     */
    trueLabel?: string;
    /**
     * Visible and accessible false label; when omitted, uses strings.no from the host-selected native resource or grid override.
     */
    falseLabel?: string;
}
/**
 * Absolute calendar or instant presentation.
 * @example
 * ```ts
 * import { DataGridDateOptions } from '@apvee/spfx-react-toolkit';
 * const options: DataGridDateOptions = { timeZone: 'Europe/Rome', format: { dateStyle: 'medium' } };
 * ```
 */
export interface DataGridDateOptions {
    /**
     * Explicit IANA zone for Date objects and dateTime instants; when omitted, defaults to UTC.
     * Calendar-only YYYY-MM-DD strings always display their UTC calendar day.
     * @defaultValue 'UTC'
     */
    timeZone?: string;
    /**
     * Date formatting; when omitted, defaults to medium date, plus short time for dateTime. Configure the zone with timeZone.
     * @defaultValue Medium date, plus short time for dateTime columns.
     */
    format?: Omit<Intl.DateTimeFormatOptions, 'timeZone'>;
}
/**
 * Badge mapping keeps the exact source scalar domain.
 * @example
 * ```ts
 * import { DataGridBadgeOptions } from '@apvee/spfx-react-toolkit';
 * const options: DataGridBadgeOptions<'open' | 'closed'> = {
 *   values: { open: { label: 'Open', color: 'success' }, closed: { label: 'Closed' } },
 * };
 * ```
 *
 * @typeParam V - Source domain retained by the column presentation and callbacks.
 */
export interface DataGridBadgeOptions<V> {
    /**
     * Exact scalar values mapped to Fluent badge presentations.
     */
    values: DataGridMappedValues<V, DataGridBadgeValue>;
}
/**
 * Informative icon mapping keeps exact source keys.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { DataGridIconOptions } from '@apvee/spfx-react-toolkit';
 * import { CheckmarkCircle20Regular } from '@fluentui/react-icons';
 * const options: DataGridIconOptions<'approved'> = {
 *   values: { approved: { icon: <CheckmarkCircle20Regular />, label: 'Approved' } },
 * };
 * ```
 *
 * @typeParam V - Source domain retained by the column presentation and callbacks.
 */
export interface DataGridIconOptions<V> {
    /**
     * Exact scalar values mapped to icons and labels.
     */
    values: DataGridMappedValues<V, DataGridIconValue>;
}
/**
 * Link destination and navigation intent.
 * @example
 * ```ts
 * import { DataGridLinkOptions } from '@apvee/spfx-react-toolkit';
 * const options: DataGridLinkOptions<{ id: number }> = { href: item => `/projects/${item.id}` };
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 */
export interface DataGridLinkOptions<TItem> {
    /**
     * Static destination or synchronous callback receiving the original row and valid link text. Returning undefined renders plain text without navigation.
     * @param item - Original row object, unchanged from the supplied items array.
     * @param value - Valid original link text from the column source.
     * @returns Destination string, or undefined to render non-navigating text.
     */
    href: string | ((item: TItem, value: string) => string | undefined);
    /**
     * Native anchor presentation; the grid owns its element, content and navigation.
     * When omitted, the native Link presentation is used with the owned destination and text.
     */
    linkProps?: Omit<Extract<LinkProps, {
        /**
         * Selects Fluent's anchor branch for typing; as is omitted from linkProps because the grid owns the anchor element.
         */
        as?: 'a' }>, 'as' | 'children' | 'href' | 'onClick' | 'dangerouslySetInnerHTML'>;
    /**
     * Called on anchor activation with the original row, ID, valid text and resolved destination.
     * Row propagation is stopped; call event.preventDefault() to replace native navigation. Return values/promises are not observed.
     * When omitted, activation follows native anchor navigation without a consumer callback.
     * @param data - Original row, stable ID, valid link text and resolved href.
     * @param event - Native anchor click event; preventDefault replaces browser navigation.
     */
    onNavigate?: (event: React.MouseEvent<HTMLAnchorElement>, data: DataGridLinkNavigateData<TItem>) => void;
}
/**
 * Multiple-person identity and native stack layout with recoverable overflow.
 * @example
 * ```ts
 * import { DataGridPeopleOptions } from '@apvee/spfx-react-toolkit';
 * const options: DataGridPeopleOptions<{ id: number }, { id: string; name: string }> = {
 *   getName: person => person.name, getPersonId: person => person.id, maxVisible: 2,
 * };
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 * @typeParam TPerson - Original person element passed to name, avatar and presence callbacks.
 */
export interface DataGridPeopleOptions<TItem, TPerson> extends DataGridPersonOptions<TItem, TPerson> {
    /**
     * Synchronously returns a string or finite numeric ID unique within this cell from the original person and row.
     * @param item - Original row object, unchanged from the supplied items array.
     * @param person - Original person source or array element for this row.
     * @returns Stable string or finite numeric identity, unique within the relevant row/element collection.
     */
    getPersonId: (person: TPerson, item: TItem) => DataGridRowId;
    /**
     * Positive whole-number maximum visible avatars; when omitted, defaults to three. Additional people remain available in the native overflow popover.
     * @defaultValue 3
     */
    maxVisible?: number;
}
/**
 * Exact scalar tag maps; object-tag callback IDs use a guarded string/number map.
 * @example
 * ```ts
 * import { DataGridTagValues } from '@apvee/spfx-react-toolkit';
 * const values: DataGridTagValues<'urgent' | 'normal'> = {
 *   urgent: { label: 'Urgent', appearance: 'brand' }, normal: { label: 'Normal' },
 * };
 * ```
 *
 * @typeParam TTag - Original tag element passed to identity and label callbacks.
 */
export type DataGridTagValues<TTag> = [TTag] extends [string | number]
    ? DataGridMappedValues<TTag, DataGridTagValue>
    : DataGridMappedValues<unknown, DataGridTagValue>;
/**
 * Exact tag mappings or paired callbacks for scalar or object tags.
 * Missing configuration and incomplete callback pairs are rejected by types and runtime.
 * @example
 * ```ts
 * import { DataGridTagsOptions } from '@apvee/spfx-react-toolkit';
 * const options: DataGridTagsOptions<{ id: number }, { id: string; label: string }> = {
 *   getTagId: tag => tag.id, getLabel: tag => tag.label, maxVisible: 3,
 * };
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 * @typeParam TTag - Original tag element passed to identity and label callbacks.
 */
export type DataGridTagsOptions<TItem, TTag> = {
    /**
     * Positive whole-number maximum visible tags before recoverable overflow. When omitted, shows every tag.
     */
    maxVisible?: number;
} & ({
    /**
     * Exact scalar IDs mapped to Fluent presentations; numeric keys require a Map.
     */
    values: DataGridTagValues<TTag>;
    /**
     * Excluded when using scalar map identities.
     * Must be omitted or undefined in this branch.
     */
    getTagId?: never;
    /**
     * Labels are supplied by the map.
     * Must be omitted or undefined in this branch.
     */
    getLabel?: never;
} | {
    /**
     * Optional mapped presentation for IDs returned by getTagId.
     * When omitted in the callback branch, getLabel supplies the tag text and native tag presentation is retained.
     */
    values?: DataGridTagValues<TTag>;
    /**
     * Stable scalar ID for each tag. Supply with getLabel; without a mapping,
     * callbacks also read scalar tags. With a mapping, scalar IDs stay original.
     * @param item - Original row object, unchanged from the supplied items array.
     * @param tag - Original tag element from this row’s source array.
     * @returns Stable string or finite numeric identity, unique within the relevant row/element collection.
     */
    getTagId: (tag: TTag, item: TItem) => DataGridRowId;
    /**
     * Explicit tag label. Supply with getTagId. Mapped labels take precedence;
     * scalar tags use these callbacks when no mapping is declared.
     * @param item - Original row object, unchanged from the supplied items array.
     * @param tag - Original tag element from this row’s source array.
     * @returns Plain display and accessible name for the original source element.
     */
    getLabel: (tag: TTag, item: TItem) => string;
});
/**
 * Determinate readonly progress presentation.
 * @example
 * ```ts
 * import { DataGridProgressOptions } from '@apvee/spfx-react-toolkit';
 * const options: DataGridProgressOptions = { max: 100, showValue: true };
 * ```
 */
export interface DataGridProgressOptions extends DataGridNumberOptions {
    /**
     * Positive finite upper bound; when omitted, defaults to one.
     * @defaultValue 1
     */
    max?: number;
    /**
     * Displays the localized value next to the indicator when true; omitted/false hides visible text while keeping its accessible value.
     * @defaultValue false
     */
    showValue?: boolean;
    /**
     * Native visual options, excluding owned value and accessibility.
     * When omitted, native ProgressBar visual defaults apply; the grid still owns range and accessible label.
     */
    progressBarProps?: Pick<ProgressBarProps, 'color' | 'shape' | 'thickness'>;
}
/**
 * Readonly filled marigold stars with an accessible value/maximum and optional vote count.
 * @example
 * ```ts
 * import { DataGridRatingOptions } from '@apvee/spfx-react-toolkit';
 * const options: DataGridRatingOptions<{ votes: number }> = { max: 5, count: ({ item }) => item.votes };
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 */
export interface DataGridRatingOptions<TItem> {
    /**
     * Whole number greater than one; when omitted, defaults to five.
     * @defaultValue 5
     */
    max?: number;
    /**
     * Optional static or row-aware non-negative whole number of votes for the
     * accessible description only. Zero is included; undefined omits the count.
     * Localize its text with strings.ratingCount.
     * @param context - Original row and valid numeric rating value.
     * @returns Non-negative whole-number vote count, or undefined to omit vote-count presentation.
     */
    count?: number | ((context: DataGridValueContext<TItem, number>) => number | undefined);
    /**
     * Native visual options, excluding owned value and range.
     * when omitted, defaults to marigold, Star20Filled and native large (20px);
     * explicit consumer overrides are retained. Numeric slots are hidden.
     * @defaultValue Marigold color, Star20Filled icon and native large size.
     */
    ratingDisplayProps?: Pick<RatingDisplayProps, 'color' | 'compact' | 'icon' | 'size'>;
}
/**
 * Official `@fluentui/react-charts` Sparkline presentation; the Charts host peer is required.
 * @example
 * ```ts
 * import { DataGridSparklineOptions } from '@apvee/spfx-react-toolkit';
 * const options: DataGridSparklineOptions<{ id: number }> = {
 *   showValue: false, color: ({ summary }) => summary.last >= summary.first ? 'green' : 'red',
 * };
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 */
export interface DataGridSparklineOptions<TItem> extends DataGridNumberOptions {
    /**
     * Shows the latest observation by default. False hides visible value text while retaining the accessible chart summary.
     * When omitted, shows the latest observation beside the chart.
     * @defaultValue true
     */
    showValue?: boolean;
    /**
     * Static or summary-aware Fluent color string.
     * When omitted, uses the chart’s default Fluent theme color.
     * @param context - Original row and the computed source-order observation summary.
     * @returns Fluent chart color string for the observation summary.
     * @defaultValue The Fluent colorBrandForeground1 token.
     */
    color?: string | ((/** Original row and the source-order chart summary. */ context: {
        /**
         * Original row.
         */
        item: TItem;
        /**
         * Immutable summary of valid observations.
         */
        summary: DataGridSparklineSummary;
    }) => string);
}
/**
 * Custom content preserves original source and JSX identity.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { DataGridCustomOptions } from '@apvee/spfx-react-toolkit';
 * const options: DataGridCustomOptions<{ id: number }, string> = {
 *   renderCell: ({ value, rowId }) => <span data-row-id={rowId}>{value}</span>,
 * };
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 * @typeParam V - Source domain retained by the column presentation and callbacks.
 */
export interface DataGridCustomOptions<TItem, V> {
    /**
     * Synchronously renders a non-missing custom source with original row identity and source value.
     * Returned React content is preserved; missing sources use placeholder instead. No async rendering is supported.
     * @param context - Original row/source value together with its stable row ID.
     * @returns Caller-owned React content, preserved without an additional data transformation.
     */
    renderCell: (/** Original row and valid source value with the stable row ID. */ context: DataGridValueContext<TItem, V> & {
        /**
         * Stable original row ID.
         */
        rowId: DataGridRowId;
    }) => React.ReactNode;
}
/**
 * Specific options selected by presentation kind and inferred source domain.
 * @example
 * ```ts
 * import { DataGridNumberOptions } from '@apvee/spfx-react-toolkit';
 * const options: DataGridNumberOptions = {
 *   format: { maximumFractionDigits: 2 },
 * };
 * ```
 *
 * @typeParam K - Column kind selecting the compatible source domain and presentation options.
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 * @typeParam V - Source domain retained by the column presentation and callbacks.
 */
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
/**
 * Props for one strongly typed value column.
 * @example
 * ```ts
 * import { DataGridValueColumnProps } from '@apvee/spfx-react-toolkit';
 * const column: DataGridValueColumnProps<{ total: number }, 'number'> = {
 *   id: 'total', header: 'Total', field: 'total', format: { maximumFractionDigits: 2 },
 * };
 * ```
 *
 * @typeParam K - Column kind selecting the compatible source domain and presentation options.
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 * @typeParam V - Source domain retained by the column presentation and callbacks.
 */
export type DataGridValueColumnProps<TItem, K extends keyof DataGridColumnDomains, V = DataGridColumnDomains[K]> = DataGridColumnCommonProps<TItem, V> & DataGridValueSource<TItem, V> & DataGridColumnOptions<TItem, K, V>;
/**
 * Actions deliberately have no source or sorting option.
 * @example
 * ```ts
 * import { ActionsColumnProps } from '@apvee/spfx-react-toolkit';
 * const props: ActionsColumnProps<{ id: number }> = {
 *   id: 'actions', actions: [{ id: 'open', label: 'Open', placement: 'menu',
 *     onClick: (_event, { item }) => console.log(item.id) }],
 * };
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 */
export type ActionsColumnProps<TItem> = Omit<DataGridColumnCommonProps<TItem, undefined>, 'header' | 'sortable'> & {
    /**
     * Optional literal actions heading; defaults to the localized label only when
     * undefined. Object/callback slots use a fixed inner span; refs target it.
     * Supply phrasing, noninteractive content.
     * @example
     * ```ts
     * import { DataGridColumnCommonProps } from '@apvee/spfx-react-toolkit';
     * const header: DataGridColumnCommonProps<{ title: string }, string>['header'] = { children: 'Actions' };
     * ```
     */
    header?: ContentSlot<'span'>;
    /**
     * Ordered synchronous actions evaluated for the original row.
     */
    actions: readonly DataGridAction<TItem>[];
    /**
     * Actions cannot read a field.
     * Must be omitted or undefined in this branch.
     */
    field?: never;
    /**
     * Actions cannot read an accessor.
     * Must be omitted or undefined in this branch.
     */
    accessor?: never;
};
declare const descriptorType: unique symbol;
/**
 * Immutable opaque descriptor owned by one suite. Create it with defineColumns.
 * @example
 * ```ts
 * import { createDataGrid } from '@apvee/spfx-react-toolkit';
 * const Grid = createDataGrid<{ id: number; name: string; title: string }>();
 * const columns = Grid.defineColumns(c => [c.text({ id:'name', header:'Name', field:'name' })]);
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 */
export interface DataGridColumnDefinition<TItem extends object> {
    /**
     * Display kind, without exposing captured callbacks.
     */
    readonly kind: DataGridColumnKind;
    /**
     * Stable unique column ID.
     */
    readonly id: string;
    /**
     * Compile-time invariant row ownership; never constructed by consumers.
     * @param item - Phantom row-model input preserving compile-time invariance; no consumer runtime call is supported.
     * @returns The same row model type; this opaque invariant marker is not constructed or invoked by consumers.
     */
    readonly [descriptorType]: (item: TItem) => TItem;
}
/**
 * TS 5.3 inference fence: only the source selects the value domain, not mappings or presentation callbacks.
 * @typeParam T - Source-inferred type prevented from widening through presentation callbacks.
 * @example
 * ```ts
 * import { createDataGrid } from '@apvee/spfx-react-toolkit';
 * const Grid = createDataGrid<{ total: number }>();
 * const columns = Grid.defineColumns(c => [c.number({ id: 'total', header: 'Total', field: 'total' })]);
 * ```
 */
type SourceOnly<T> = [T][T extends unknown ? 0 : never];
/**
 * Field-backed column configuration retaining the exact selected property domain.
 * TItem is the row, K selects presentation, and F is the compatible top-level field.
 * @example
 * ```ts
 * import { createDataGrid } from '@apvee/spfx-react-toolkit';
 * const Grid = createDataGrid<{ total: number }>();
 * const columns = Grid.defineColumns(c => [
 *   c.number({ id: 'total', header: 'Total', field: 'total' }),
 * ]);
 * ```
 *
 * @typeParam F - Compatible top-level field name whose property type drives callback inference.
 * @typeParam K - Column kind selecting the compatible source domain and presentation options.
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 */
type FieldProps<TItem, K extends keyof DataGridColumnDomains, F extends Extract<keyof TItem, string>> = DataGridColumnCommonProps<TItem, TItem[F]> & {
    /**
     * Compatible original row field used as the source.
     */
    field: F;
    /**
     * Excluded for field-backed columns.
     * Must be omitted or undefined in this branch.
     */
    accessor?: never;
} & DataGridColumnOptions<TItem, K, NonNullable<TItem[F]>>;
/**
 * Field/accessor overloads infer callbacks from their own source, without domain widening.
 * @example
 * ```ts
 * import { createDataGrid } from '@apvee/spfx-react-toolkit';
 * const Grid = createDataGrid<{ total: number }>();
 * const columns = Grid.defineColumns(c => [c.number({ id: 'total', header: 'Total', field: 'total' })]);
 * ```
 */
export type DataGridColumnMethod<TItem extends object, K extends keyof DataGridColumnDomains, R> = K extends 'badge' | 'icon' ? MappedColumnMethod<TItem, K, R> : SourceColumnMethod<TItem, K, R>;
/**
 * Field/accessor call signatures whose source determines callback types.
 * @example
 * ```ts
 * import { createDataGrid } from '@apvee/spfx-react-toolkit';
 * const Grid = createDataGrid<{ total: number }>();
 * const columns = Grid.defineColumns(c => [c.number({ id: 'double', header: 'Double', accessor: row => row.total * 2 })]);
 * ```
 *
 * @typeParam K - Column kind selecting the compatible source domain and presentation options.
 * @typeParam R - Same-suite descriptor or JSX marker returned by this column method.
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 */
interface SourceColumnMethod<TItem extends object, K extends keyof DataGridColumnDomains, R> {
    /**
     * Configure a field; callbacks infer its original domain.
     * @param props - Column configuration with a compatible original field.
     * @returns Same-suite descriptor or JSX marker.
     * @typeParam F - Compatible top-level field name whose property type drives callback inference.
     */
    <F extends CompatibleDataGridField<TItem, DataGridColumnDomains[K]>>(props: FieldProps<TItem, K, F>): R;
    /**
     * Configure a computed source.
     * @param props - Pure synchronous accessor and presentation.
     * @returns Same-suite descriptor or JSX marker.
     * @typeParam A - Pure synchronous accessor whose return type drives callback inference.
     */
    <A extends (item: TItem) => DataGridSourceResult<DataGridColumnDomains[K]>>(props: DataGridColumnCommonProps<TItem, ReturnType<A>> & {
        /**
         * Excluded for computed sources.
         * Must be omitted or undefined in this branch.
         */
        field?: never;
        /**
         * Pure synchronous source whose return determines callback types.
         */
        accessor: A;
    } & DataGridColumnOptions<TItem, K, SourceOnly<NonNullable<ReturnType<A>>>>): R;
}
/**
 * Field/accessor call signatures preserving exact badge and icon mapping domains.
 * @example
 * ```ts
 * import { createDataGrid } from '@apvee/spfx-react-toolkit';
 * const Grid = createDataGrid<{ status: 'open' | 'closed' }>();
 * const columns = Grid.defineColumns(c => [c.badge({ id: 'status', header: 'Status', field: 'status',
 *   values: { open: { label: 'Open' }, closed: { label: 'Closed' } } })]);
 * ```
 *
 * @typeParam K - Column kind selecting the compatible source domain and presentation options.
 * @typeParam R - Same-suite descriptor or JSX marker returned by this column method.
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 */
interface MappedColumnMethod<TItem extends object, K extends 'badge' | 'icon', R> {
    /**
     * Configure a field; callbacks infer its original domain.
     * @param props - Column configuration with a compatible original field.
     * @returns Same-suite descriptor or JSX marker.
     * @typeParam F - Compatible top-level field name whose property type drives callback inference.
     */
    <F extends CompatibleDataGridField<TItem, string | number>>(props: FieldProps<TItem, K, F>): R;
    /**
     * Configure an exact mapped accessor.
     * @param props - Pure synchronous source and exact presentation map.
     * @returns Same-suite descriptor or JSX marker.
     * @typeParam V - Source domain retained by the column presentation and callbacks.
     */
    <V extends string>(props: DataGridColumnCommonProps<TItem, V> & {
        /**
         * Excluded for computed sources.
         * Must be omitted or undefined in this branch.
         */
        field?: never;
        /**
         * Pure synchronous source whose exact scalar domain determines mapping keys.
         * @param item - Original row object, unchanged from the supplied items array.
         * @returns Original synchronous source value, or undefined/legacy null for a missing cell.
         */
        accessor: (item: TItem) => DataGridSourceResult<V>;
    } & DataGridColumnOptions<TItem, K, SourceOnly<V>>): R;
    /**
     * Configure an exact mapped accessor.
     * @param props - Pure synchronous source and exact presentation map.
     * @returns Same-suite descriptor or JSX marker.
     * @typeParam V - Source domain retained by the column presentation and callbacks.
     */
    <V extends string | number>(props: DataGridColumnCommonProps<TItem, V> & {
        /**
         * Excluded for computed sources.
         * Must be omitted or undefined in this branch.
         */
        field?: never;
        /**
         * Pure synchronous source whose exact scalar domain determines mapping keys.
         */
        accessor: (item: TItem) => DataGridSourceResult<V>;
    } & DataGridColumnOptions<TItem, K, SourceOnly<V>>): R;
}
/**
 * Builder for all sixteen presentation kinds.
 * @example
 * ```ts
 * import { createDataGrid } from '@apvee/spfx-react-toolkit';
 * const Grid = createDataGrid<{ id: number; total: number; title: string }>();
 * const columns = Grid.defineColumns(c => [c.number({ id:"total", header:"Total", field:"total" })]);
 * ```
 *
 * @typeParam TItem - Original row model passed unchanged to sources and callbacks.
 */
export interface DataGridColumnBuilder<TItem extends object> {
    /**
     * Captures a text source and creates a same-suite immutable descriptor.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; title: string }>();
     * const columns = Grid.defineColumns(c => [c.text({ id: 'title', header: 'Title', field: 'title' })]);
     * ```
     */
    text: DataGridColumnMethod<TItem, 'text', DataGridColumnDefinition<TItem>>;
    /**
     * Captures a number source and creates a same-suite immutable descriptor.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; total: number; title: string }>();
     * const columns = Grid.defineColumns(c => [c.number({ id: 'total', header: 'Total', field: 'total' })]);
     * ```
     */
    number: DataGridColumnMethod<TItem, 'number', DataGridColumnDefinition<TItem>>;
    /**
     * Captures a boolean source and creates a same-suite immutable descriptor.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; enabled: boolean; title: string }>();
     * const columns = Grid.defineColumns(c => [c.boolean({ id: 'enabled', header: 'Enabled', field: 'enabled' })]);
     * ```
     */
    boolean: DataGridColumnMethod<TItem, 'boolean', DataGridColumnDefinition<TItem>>;
    /**
     * Captures a date source and creates a same-suite immutable descriptor.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; date: string; title: string }>();
     * const columns = Grid.defineColumns(c => [c.date({ id: 'date', header: 'Date', field: 'date' })]);
     * ```
     */
    date: DataGridColumnMethod<TItem, 'date', DataGridColumnDefinition<TItem>>;
    /**
     * Captures a dateTime source and creates a same-suite immutable descriptor.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; updated: string; title: string }>();
     * const columns = Grid.defineColumns(c => [c.dateTime({ id: 'updated', header: 'Updated', field: 'updated' })]);
     * ```
     */
    dateTime: DataGridColumnMethod<TItem, 'dateTime', DataGridColumnDefinition<TItem>>;
    /**
     * Captures a badge source and creates a same-suite immutable descriptor.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; status: 'open' | 'closed'; title: string }>();
     * const columns = Grid.defineColumns(c => [c.badge({ id: 'status', header: 'Status', field: 'status', values: { open: { label: 'Open' } } })]);
     * ```
     */
    badge: DataGridColumnMethod<TItem, 'badge', DataGridColumnDefinition<TItem>>;
    /**
     * Captures a icon source and creates a same-suite immutable descriptor.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * import { CheckmarkCircle20Regular } from '@fluentui/react-icons';
     * const Grid = createDataGrid<{ id: number; status: 'open' | 'closed'; title: string }>();
     * const columns = Grid.defineColumns(c => [c.icon({ id: 'status', header: 'Status', field: 'status', values: { open: { icon: <CheckmarkCircle20Regular />, label: 'Open' } } })]);
     * ```
     */
    icon: DataGridColumnMethod<TItem, 'icon', DataGridColumnDefinition<TItem>>;
    /**
     * Captures a link source and creates a same-suite immutable descriptor.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; title: string }>();
     * const columns = Grid.defineColumns(c => [c.link({ id: 'title', header: 'Title', field: 'title', href: row => `/projects/${row.id}` })]);
     * ```
     */
    link: DataGridColumnMethod<TItem, 'link', DataGridColumnDefinition<TItem>>;
    /**
     * Captures a person source and creates a same-suite immutable descriptor.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; owner: { id: string; name: string }; title: string }>();
     * const columns = Grid.defineColumns(c => [c.person({ id: 'owner', header: 'Owner', field: 'owner', getName: person => person.name })]);
     * ```
     */
    person: DataGridColumnMethod<TItem, 'person', DataGridColumnDefinition<TItem>>;
    /**
     * Captures a people source and creates a same-suite immutable descriptor.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; team: readonly { id: string; name: string }[]; title: string }>();
     * const columns = Grid.defineColumns(c => [c.people({ id: 'team', header: 'Team', field: 'team', getName: person => person.name, getPersonId: person => person.id })]);
     * ```
     */
    people: DataGridColumnMethod<TItem, 'people', DataGridColumnDefinition<TItem>>;
    /**
     * Captures a tags source and creates a same-suite immutable descriptor.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; tags: readonly string[]; title: string }>();
     * const columns = Grid.defineColumns(c => [c.tags({ id: 'tags', header: 'Tags', field: 'tags', getTagId: tag => tag, getLabel: tag => tag })]);
     * ```
     */
    tags: DataGridColumnMethod<TItem, 'tags', DataGridColumnDefinition<TItem>>;
    /**
     * Captures a progress source and creates a same-suite immutable descriptor.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; progress: number; title: string }>();
     * const columns = Grid.defineColumns(c => [c.progress({ id: 'progress', header: 'Progress', field: 'progress', max: 100 })]);
     * ```
     */
    progress: DataGridColumnMethod<TItem, 'progress', DataGridColumnDefinition<TItem>>;
    /**
     * Captures a rating source and creates a same-suite immutable descriptor.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; score: number; title: string }>();
     * const columns = Grid.defineColumns(c => [c.rating({ id: 'score', header: 'Score', field: 'score', max: 5 })]);
     * ```
     */
    rating: DataGridColumnMethod<TItem, 'rating', DataGridColumnDefinition<TItem>>;
    /**
     * Captures a sparkline source and creates a same-suite immutable descriptor.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; trend: readonly number[]; title: string }>();
     * const columns = Grid.defineColumns(c => [c.sparkline({ id: 'trend', header: 'Trend', field: 'trend' })]);
     * ```
     */
    sparkline: DataGridColumnMethod<TItem, 'sparkline', DataGridColumnDefinition<TItem>>;
    /**
     * Captures a custom source and creates a same-suite immutable descriptor.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; title: string }>();
     * const columns = Grid.defineColumns(c => [c.custom({ id: 'title', header: 'Title', field: 'title', renderCell: ({ value }) => <strong>{String(value)}</strong> })]);
     * ```
     */
    custom: DataGridColumnMethod<TItem, 'custom', DataGridColumnDefinition<TItem>>;
    /**
     * Creates a source-free action descriptor.
     * @example
     * ```ts
     * import { createDataGrid } from '@apvee/spfx-react-toolkit';
     * const Grid = createDataGrid<{ id: number; title: string }>();
     * const columns = Grid.defineColumns(c => [c.actions({ id: 'actions', actions: [{ id: 'open', label: 'Open', placement: 'inline', onClick: (_event, { item }) => console.log(item.id) }] })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    actions: (props: ActionsColumnProps<TItem>) => DataGridColumnDefinition<TItem>;
}
/**
 * Explicit dynamic source props. Generic callbacks receive unknown.
 * @example
 * ```ts
 * import { DynamicDataGridColumnProps } from '@apvee/spfx-react-toolkit';
 * const props: DynamicDataGridColumnProps<'text'> = { id: 'title', header: 'Title', field: 'title' };
 * ```
 *
 * @typeParam K - Column kind selecting the compatible source domain and presentation options.
 */
export type DynamicDataGridColumnProps<K extends keyof DataGridColumnDomains> = DataGridColumnCommonProps<Record<string, unknown>, unknown> & ({
    /**
     * Explicit dynamic top-level field.
     */
    field: string;
    /**
     * Excluded for field sources.
     * Must be omitted or undefined in this branch.
     */
    accessor?: never;
} | {
    /**
     * Excluded for computed sources.
     * Must be omitted or undefined in this branch.
     */
    field?: never;
    /**
     * Pure synchronous source; its return remains unknown.
     * @param item - Original row object, unchanged from the supplied items array.
     * @returns Original synchronous source value, or undefined/legacy null for a missing cell.
     */
    accessor: (item: Record<string, unknown>) => unknown;
}) & DataGridColumnOptions<Record<string, unknown>, K, unknown>;
/**
 * Explicit dynamic builder; generic source callbacks receive unknown.
 * @example
 * ```ts
 * import { DataGrid } from '@apvee/spfx-react-toolkit';
 * const columns = DataGrid.defineColumns(c => [c.text({ id: 'title', header: 'Title', field: 'title' })]);
 * ```
 */
export interface DynamicDataGridColumnBuilder {
    /**
     * Creates a guarded dynamic text descriptor.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.text({ id: 'title', header: 'Title', field: 'title' })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    text: (props: DynamicDataGridColumnProps<'text'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a guarded dynamic number descriptor.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.number({ id: 'total', header: 'Total', field: 'total' })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    number: (props: DynamicDataGridColumnProps<'number'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a guarded dynamic boolean descriptor.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.boolean({ id: 'enabled', header: 'Enabled', field: 'enabled' })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    boolean: (props: DynamicDataGridColumnProps<'boolean'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a guarded dynamic date descriptor.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.date({ id: 'date', header: 'Date', field: 'date' })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    date: (props: DynamicDataGridColumnProps<'date'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a guarded dynamic dateTime descriptor.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.dateTime({ id: 'updated', header: 'Updated', field: 'updated' })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    dateTime: (props: DynamicDataGridColumnProps<'dateTime'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a guarded dynamic badge descriptor.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.badge({ id: 'status', header: 'Status', field: 'status', values: { open: { label: 'Open' } } })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    badge: (props: DynamicDataGridColumnProps<'badge'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a guarded dynamic icon descriptor.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * import { CheckmarkCircle20Regular } from '@fluentui/react-icons';
     * const columns = DataGrid.defineColumns(c => [c.icon({ id: 'status', header: 'Status', field: 'status', values: { open: { icon: <CheckmarkCircle20Regular />, label: 'Open' } } })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    icon: (props: DynamicDataGridColumnProps<'icon'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a guarded dynamic link descriptor.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.link({ id: 'title', header: 'Title', field: 'title', href: row => `/projects/${row.id}` })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    link: (props: DynamicDataGridColumnProps<'link'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a guarded dynamic person descriptor.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.person({ id: 'owner', header: 'Owner', field: 'owner', getName: person => String(person) })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    person: (props: DynamicDataGridColumnProps<'person'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a guarded dynamic people descriptor.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.people({ id: 'team', header: 'Team', field: 'team', getName: person => String(person), getPersonId: person => String(person) })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    people: (props: DynamicDataGridColumnProps<'people'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a guarded dynamic tags descriptor.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.tags({ id: 'tags', header: 'Tags', field: 'tags', getTagId: tag => String(tag), getLabel: tag => String(tag) })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    tags: (props: DynamicDataGridColumnProps<'tags'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a guarded dynamic progress descriptor.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.progress({ id: 'progress', header: 'Progress', field: 'progress', max: 100 })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    progress: (props: DynamicDataGridColumnProps<'progress'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a guarded dynamic rating descriptor.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.rating({ id: 'score', header: 'Score', field: 'score', max: 5 })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    rating: (props: DynamicDataGridColumnProps<'rating'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a guarded dynamic sparkline descriptor.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.sparkline({ id: 'trend', header: 'Trend', field: 'trend' })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    sparkline: (props: DynamicDataGridColumnProps<'sparkline'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a guarded dynamic custom descriptor.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.custom({ id: 'title', header: 'Title', field: 'title', renderCell: ({ value }) => <strong>{String(value)}</strong> })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    custom: (props: DynamicDataGridColumnProps<'custom'>) => DataGridColumnDefinition<Record<string, unknown>>;
    /**
     * Creates a source-free dynamic action descriptor.
     * @example
     * ```ts
     * import { DataGrid } from '@apvee/spfx-react-toolkit';
     * const columns = DataGrid.defineColumns(c => [c.actions({ id: 'actions', actions: [{ id: 'open', label: 'Open', placement: 'inline', onClick: (_event, { item }) => console.log(item.id) }] })]);
     * ```
     *
     * @param props - Column configuration for this source/presentation kind; sources and callbacks must remain synchronous and pure.
     * @returns Immutable descriptor owned by this suite for the configured column kind.
     */
    actions: (props: ActionsColumnProps<Record<string, unknown>>) => DataGridColumnDefinition<Record<string, unknown>>;
}

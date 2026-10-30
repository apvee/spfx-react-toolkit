import * as React from 'react';
import type { AvatarProps, PersonaProps } from '@fluentui/react-components';
import type {
  ActionsColumnProps,
  DataGridBadgeValue,
  DataGridColumnCommonProps,
  DataGridColumnDomains,
  DataGridColumnOptions,
  DataGridIconValue,
  DataGridLocale,
  DataGridRowId,
  DataGridTagValue,
} from './data-grid-column.types';
import type { FamilyResolution } from './column-family.internal';
type CommonOptions<T extends object> = Omit<DataGridColumnCommonProps<T, unknown>, 'tooltip'>;
interface ResolvedBase<T extends object> {
    readonly id: string;
    readonly read: (item: T) => unknown;
    readonly source: string | ((item: T) => unknown) | undefined;
    readonly staticTooltip: string | undefined;
    readonly tooltip: (raw: unknown, item: T) => string | undefined;
}
export interface ResolvedDataGridPerson {
    readonly id?: DataGridRowId;
    readonly name: string;
    readonly secondaryText?: string;
    readonly avatarProps?: Omit<AvatarProps, 'name'>;
    readonly presenceProps?: NonNullable<PersonaProps['presence']>;
}
export interface ResolvedDataGridTag {
    readonly id: DataGridRowId;
    readonly label: string;
    readonly presentation?: DataGridTagValue;
}
type OrdinaryKind = Exclude<keyof DataGridColumnDomains, 'badge' | 'icon' | 'person' | 'people' | 'tags' | 'custom'>;
type OrdinaryColumn<T extends object> = {
    [K in OrdinaryKind]: ResolvedBase<T> & {
        readonly kind: K;
        readonly options: Readonly<CommonOptions<T> & DataGridColumnOptions<T, K, DataGridColumnDomains[K]>>;
    };
}[OrdinaryKind];
/** Private discriminated renderer handoff. Generic callbacks are captured in adapters. */
export type ResolvedDataGridColumn<T extends object> = OrdinaryColumn<T> | (ResolvedBase<T> & {
    readonly kind: 'badge';
    readonly options: Readonly<CommonOptions<T>>;
    readonly mapValue: (value: unknown) => DataGridBadgeValue | undefined;
}) | (ResolvedBase<T> & {
    readonly kind: 'icon';
    readonly options: Readonly<CommonOptions<T>>;
    readonly mapValue: (value: unknown) => DataGridIconValue | undefined;
}) | (ResolvedBase<T> & {
    readonly kind: 'person';
    readonly options: Readonly<CommonOptions<T> & Pick<DataGridColumnOptions<T, 'person', unknown>, 'personaProps'>>;
    readonly presentPerson: (value: unknown, item: T) => ResolvedDataGridPerson;
}) | (ResolvedBase<T> & {
    readonly kind: 'people';
    readonly options: Readonly<CommonOptions<T> & Pick<DataGridColumnOptions<T, 'people', readonly unknown[]>, 'personaProps' | 'maxVisible'>>;
    readonly presentPeople: (value: readonly unknown[], item: T) => readonly ResolvedDataGridPerson[];
}) | (ResolvedBase<T> & {
    readonly kind: 'tags';
    readonly options: Readonly<CommonOptions<T> & {
        maxVisible?: number;
    }>;
    readonly presentTags: (value: readonly unknown[], item: T) => readonly ResolvedDataGridTag[] | undefined;
}) | (ResolvedBase<T> & {
    readonly kind: 'custom';
    readonly options: Readonly<CommonOptions<T>>;
    readonly renderCell: (value: unknown, item: T, rowId: DataGridRowId) => React.ReactNode;
}) | (ResolvedBase<T> & {
    readonly kind: 'actions';
    readonly options: Readonly<ActionsColumnProps<T>>;
});
export function assertDataGridConfiguration(condition: unknown, message: string): asserts condition {
    if (!condition)
        throw new TypeError(`DataGrid column configuration: ${message}`);
}
/** Checks unknown JSX configuration before any trusted captured source is invoked. */
export function validateDataGridColumnOptions(kind: string, value: unknown): asserts value is Record<string, unknown> {
    // React and caller row objects may contain legacy null; this is a boundary guard only.
    assertDataGridConfiguration(typeof value === 'object' && value !== null && !Array.isArray(value), 'column options must be an object');
    const props = value as Record<string, unknown>;
    const assert: typeof assertDataGridConfiguration = assertDataGridConfiguration;
    assert(typeof props.id === 'string' && props.id.trim().length > 0, 'column IDs must be non-empty strings');
    for (const excluded of ['sortValue', 'compare', 'hideable', 'filter', 'searchable', 'searchValue'])
        assert(!(excluded in props), `${excluded} is not a supported column option`);
    if (kind === 'actions')
        assert(props.field === undefined && props.accessor === undefined, 'actions cannot define a value source');
    else {
        assert((props.field !== undefined) !== (props.accessor !== undefined), 'value columns require exactly one field or accessor source');
        assert(props.field === undefined || (typeof props.field === 'string' && props.field.length > 0), 'field must be non-empty text');
        assert(props.accessor === undefined || typeof props.accessor === 'function', 'accessor must be a function');
        assert(props.header !== undefined, 'value columns require a header');
    }
    for (const option of ['sortable', 'showValue'])
        assert(props[option] === undefined || typeof props[option] === 'boolean', `${option} must be boolean`);
    for (const option of ['headerAriaLabel', 'headerTooltip', 'trueLabel', 'falseLabel'])
        assert(props[option] === undefined || typeof props[option] === 'string', `${option} must be text`);
    assert(props.tooltip === undefined || typeof props.tooltip === 'string' || typeof props.tooltip === 'function', 'tooltip must be text or a callback');
    for (const option of ['width', 'minWidth'])
        assert(props[option] === undefined || (typeof props[option] === 'number' && Number.isFinite(props[option]) && Number(props[option]) >= 0), `${option} must be finite and non-negative`);
    if (props.width !== undefined && props.minWidth !== undefined)
        assert(Number(props.width) >= Number(props.minWidth), 'width must not be less than minWidth');
    assert(props.maxVisible === undefined || (typeof props.maxVisible === 'number' && Number.isInteger(props.maxVisible) && props.maxVisible >= 1), 'maxVisible must be a positive whole number');
    assert(props.display === undefined || props.display === 'text' || props.display === 'icon', 'display must be text or icon');
    if (kind === 'progress')
        assert(props.max === undefined || (typeof props.max === 'number' && Number.isFinite(props.max) && props.max > 0), 'progress max must be positive and finite');
    if (kind === 'rating') {
        assert(props.max === undefined || (typeof props.max === 'number' && Number.isInteger(props.max) && props.max > 1), 'rating max must be a whole number greater than one');
        assert(props.count === undefined || typeof props.count === 'function' || (typeof props.count === 'number' && Number.isInteger(props.count) && props.count >= 0), 'rating count must be a non-negative whole number or callback');
    }
    if (kind === 'custom')
        assert(typeof props.renderCell === 'function', 'custom renderCell must be a callback');
    if (kind === 'link')
        assert(typeof props.href === 'string' || typeof props.href === 'function', 'link href must be text or a callback');
    if (kind === 'person' || kind === 'people') {
        assert(typeof props.getName === 'function', 'person getName must be a callback');
        for (const option of ['getSecondaryText', 'getAvatarProps', 'getPresenceProps'])
            assert(props[option] === undefined || typeof props[option] === 'function', `${option} must be a callback`);
    }
    if (kind === 'people')
        assert(typeof props.getPersonId === 'function', 'people getPersonId must be a callback');
    if (kind === 'badge' || kind === 'icon' || (kind === 'tags' && props.values !== undefined))
        validateMapping(props.values, kind);
    if (kind === 'tags') {
        assert((props.getTagId === undefined && props.getLabel === undefined) || (typeof props.getTagId === 'function' && typeof props.getLabel === 'function'), 'tag getTagId and getLabel must be supplied together');
        assert(props.values !== undefined || (typeof props.getTagId === 'function' && typeof props.getLabel === 'function'), 'tags require a mapping or getTagId and getLabel');
    }
    if (kind === 'actions') {
        assert(props.sortable === undefined, 'actions do not support sortable');
        assert(Array.isArray(props.actions), 'actions must be an array');
        const ids = new Set<string>();
        for (const value of props.actions) {
            assert(typeof value === 'object' && value !== null, 'action must be an object');
            const action = value as Record<string, unknown>;
            assert(typeof action.id === 'string' && action.id.trim().length > 0 && !ids.has(action.id), 'action IDs must be non-empty and unique');
            ids.add(action.id);
            assert(typeof action.label === 'string', 'action label must be text');
            assert(action.placement === 'inline' || action.placement === 'menu', 'action placement must be inline or menu');
            assert(typeof action.onClick === 'function', 'action onClick must be a callback');
            for (const key of ['visible', 'disabled', 'loading'])
                assert(action[key] === undefined || typeof action[key] === 'boolean' || typeof action[key] === 'function', `action ${key} must be boolean or callback`);
        }
    }
    assert(props.color === undefined || typeof props.color === 'string' || typeof props.color === 'function', 'color must be text or a callback');
    assert(props.onNavigate === undefined || typeof props.onNavigate === 'function', 'onNavigate must be a callback');
    for (const key of ['linkProps', 'personaProps', 'progressBarProps', 'ratingDisplayProps']) {
        const native = props[key];
        if (native === undefined) continue;
        assert(typeof native === 'object' && native !== null && !Array.isArray(native), `${key} must be a presentation object`);
        const owned: Record<string, readonly string[]> = {
            linkProps: ['as', 'children', 'href', 'onClick', 'dangerouslySetInnerHTML'],
            personaProps: ['name', 'primaryText', 'secondaryText', 'avatar', 'presence', 'dangerouslySetInnerHTML'],
            progressBarProps: ['value', 'max'],
            ratingDisplayProps: ['value', 'max', 'count'],
        };
        for (const property of owned[key]) assert(!(property in native), `${key}.${property} is owned by the grid`);
    }
    validateIntl(kind, props);
}
/** Exact map entry identity, including structural ReadonlyMap implementations. */
export function getDataGridMappingEntries(value: unknown): readonly (readonly [unknown, unknown])[] {
    assertDataGridConfiguration(typeof value === 'object' && value !== null && !Array.isArray(value), 'values must be a mapping object');
    const candidate = value as Partial<ReadonlyMap<unknown, unknown>>;
    if (typeof candidate.get === 'function' && typeof candidate.entries === 'function') {
        return Array.from(candidate.entries());
    }
    return Object.entries(value);
}
function validateMapping(value: unknown, kind: string): void {
    const assert: typeof assertDataGridConfiguration = assertDataGridConfiguration;
    assert(typeof value === 'object' && value !== null && !Array.isArray(value), `${kind} requires a values mapping`);
    const entries = getDataGridMappingEntries(value);
    for (const [key, entry] of entries) {
        assert(typeof key === 'string' || (typeof key === 'number' && Number.isFinite(key)), 'mapped keys must be strings or finite numbers');
        if (entry === undefined)
            continue;
        assert(typeof entry === 'object' && entry !== null && !Array.isArray(entry), 'mapped presentation must be an object');
        const presentation = entry as Record<string, unknown>;
        assert(typeof presentation.label === 'string', 'mapped presentation requires a label');
        if (kind === 'icon')
            assert(presentation.icon !== undefined, 'mapped icon requires an icon');
    }
}
function validateIntl(kind: string, props: Record<string, unknown>): void {
    const assert: typeof assertDataGridConfiguration = assertDataGridConfiguration;
    assert(props.locale === undefined || typeof props.locale === 'string' || (Array.isArray(props.locale) && props.locale.every(v => typeof v === 'string')), 'locale must be text or a locale list');
    assert(props.format === undefined || (typeof props.format === 'object' && props.format !== null && !Array.isArray(props.format)), 'format must be an options object');
    assert(props.timeZone === undefined || typeof props.timeZone === 'string', 'timeZone must be text');
    if (props.format !== undefined)
        assert(!Object.prototype.hasOwnProperty.call(props.format, 'timeZone'), 'format.timeZone is excluded; use timeZone');
    try {
        const configuredLocale = props.locale as DataGridLocale | undefined;
        const locale = typeof configuredLocale === 'string' ? configuredLocale : configuredLocale === undefined ? undefined : [...configuredLocale];
        if (locale !== undefined)
            new Intl.Collator(locale).resolvedOptions();
        if (kind === 'date' || kind === 'dateTime')
            new Intl.DateTimeFormat(locale, { ...(props.format as Intl.DateTimeFormatOptions | undefined), timeZone: (props.timeZone as string | undefined) ?? 'UTC' }).resolvedOptions();
        else if (['number', 'progress', 'rating', 'sparkline'].includes(kind))
            new Intl.NumberFormat(locale, props.format as Intl.NumberFormatOptions | undefined).resolvedOptions();
    }
    catch {
        throw new TypeError('DataGrid column configuration: invalid locale, timeZone or format');
    }
}
/** Flatten fragments and arrays without executing column markers or arbitrary components. */
export function normalizeDataGridColumns<T extends object>(family: FamilyResolution<T>, input: {
    children?: React.ReactNode;
    columns?: readonly unknown[];
}): readonly ResolvedDataGridColumn<T>[] {
    const hasChildren = Object.prototype.hasOwnProperty.call(input, 'children');
    const hasColumns = Object.prototype.hasOwnProperty.call(input, 'columns');
    assertDataGridConfiguration(hasChildren !== hasColumns, 'declare exactly one children or columns collection');
    const result: ResolvedDataGridColumn<T>[] = [];
    if (hasColumns) {
        assertDataGridConfiguration(Array.isArray(input.columns), 'columns must be a descriptor array');
        for (const descriptor of input.columns)
            result.push(family.resolveDescriptor(descriptor));
    }
    else {
        const visit = (child: React.ReactNode): void => {
            if (child === undefined || child === null || typeof child === 'boolean')
                return; // React's legacy empty-child contract.
            if (Array.isArray(child)) {
                child.forEach(visit);
                return;
            }
            assertDataGridConfiguration(React.isValidElement(child), 'children must be same-family column markers');
            if (child.type === React.Fragment) {
                visit((child.props as {
                    children?: React.ReactNode;
                }).children);
                return;
            }
            result.push(family.resolveMarker(child));
        };
        visit(input.children);
    }
    const ids = new Set<string>();
    for (const column of result) {
        assertDataGridConfiguration(!ids.has(column.id), `duplicate column ID "${column.id}"`);
        ids.add(column.id);
    }
    return Object.freeze(result);
}

import * as React from 'react';
import type {
  ActionsColumnProps,
  DataGridBadgeValue,
  DataGridColumnBuilder,
  DataGridColumnCommonProps,
  DataGridColumnDefinition,
  DataGridColumnDomains,
  DataGridColumnKind,
  DataGridColumnMethod,
  DataGridColumnOptions,
  DataGridIconValue,
  DataGridPersonOptions,
  DataGridRowId,
  DataGridTagValue,
  DynamicDataGridColumnBuilder,
} from './data-grid-column.types';
import type { DataGridColumnMarkers, DynamicDataGridColumnMarkers } from './data-grid.types';
import { assertDataGridConfiguration, getDataGridMappingEntries, validateDataGridColumnOptions } from './column-normalization.internal';
import type {
  ResolvedDataGridColumn,
  ResolvedDataGridPerson,
  ResolvedDataGridTag,
} from './column-normalization.internal';
type ValueKind = keyof DataGridColumnDomains;
type CaptureProps<T, V, K extends ValueKind> = DataGridColumnCommonProps<T, V> & {
    field?: string;
    accessor?: (item: T) => V | undefined;
} & DataGridColumnOptions<T, K, V>;
type ElementOf<V> = V extends readonly (infer E)[] ? E : unknown;
export interface FamilyResolution<T extends object> {
    readonly resolveDescriptor: (descriptor: unknown) => ResolvedDataGridColumn<T>;
    readonly resolveMarker: (element: React.ReactElement) => ResolvedDataGridColumn<T>;
}
export type ColumnFamily<T extends object> = FamilyResolution<T> & DataGridColumnMarkers<T> & {
    readonly builder: DataGridColumnBuilder<T>;
    readonly defineColumns: (factory: (builder: DataGridColumnBuilder<T>) => readonly DataGridColumnDefinition<T>[]) => readonly DataGridColumnDefinition<T>[];
};
export type DynamicColumnFamily = FamilyResolution<Record<string, unknown>> & DynamicDataGridColumnMarkers & {
    readonly builder: DynamicDataGridColumnBuilder;
    readonly defineColumns: (factory: (builder: DynamicDataGridColumnBuilder) => readonly DataGridColumnDefinition<Record<string, unknown>>[]) => readonly DataGridColumnDefinition<Record<string, unknown>>[];
};
const names = { text: 'TextColumn', number: 'NumberColumn', boolean: 'BooleanColumn', date: 'DateColumn', dateTime: 'DateTimeColumn', badge: 'BadgeColumn', icon: 'IconColumn', link: 'LinkColumn', person: 'PersonColumn', people: 'PeopleColumn', tags: 'TagsColumn', progress: 'ProgressColumn', rating: 'RatingColumn', sparkline: 'SparklineColumn', actions: 'ActionsColumn', custom: 'Column' } as const;
/** Shallow snapshots isolate configuration; caller rows and callback identities are untouched. */
function snapshotOptions<P extends object>(props: P): Readonly<P> {
    const copy = { ...props };
    for (const key of ['format', 'linkProps', 'personaProps', 'progressBarProps', 'ratingDisplayProps']) {
        const record = copy as Record<string, unknown>;
        const value = record[key];
        if (value !== undefined)
            record[key] = Object.freeze({ ...value as object });
    }
    return Object.freeze(copy);
}
function captureMap<P extends object>(values: Readonly<Partial<Record<string, P>>> | ReadonlyMap<string | number, P>): (raw: unknown) => P | undefined {
    // Snapshot exact entries, preserving numeric versus string keys. No string coercion.
    const map = new Map<string | number, P>();
    for (const [key, value] of getDataGridMappingEntries(values)) {
        if (value === undefined) continue;
        // Configuration validation established finite scalar keys and presentation
        // objects; P is retained from the captured generic mapping, never a callback.
        if (typeof key === 'string' || typeof key === 'number') {
            map.set(key, Object.freeze({ ...value as P }));
        }
    }
    return raw => typeof raw === 'string' || typeof raw === 'number' ? map.get(raw) : undefined;
}
function personPresentation<T, P>(options: DataGridPersonOptions<T, P>, person: P, item: T, id?: DataGridRowId): ResolvedDataGridPerson {
    return Object.freeze({ id, name: options.getName(person, item), secondaryText: options.getSecondaryText?.(person, item), avatarProps: options.getAvatarProps?.(person, item), presenceProps: options.getPresenceProps?.(person, item) });
}
/** Capture a generic source and its callbacks before exposing heterogeneous metadata. */
function captureValueColumn<T extends object, K extends ValueKind, V>(kind: K, props: CaptureProps<T, V, K>): ResolvedDataGridColumn<T> {
    validateDataGridColumnOptions(kind, props);
    const options = snapshotOptions(props);
    const accessor = options.accessor;
    const field = options.field;
    const read: (item: T) => V | undefined = accessor ?? (item => {
        // Field membership/domain is checked by the public overload. Dynamic V is unknown.
        // JSX is an erased React boundary, validated before this owning-family reader exists.
        return (item as Record<string, unknown>)[field as string] as V | undefined;
    });
    const tooltip = options.tooltip;
    const base = { id: options.id, kind, read, source: accessor ?? field, options,
        staticTooltip: typeof tooltip === 'string' ? tooltip : undefined,
        tooltip: (raw: unknown, item: T) => typeof tooltip === 'function' ? tooltip({ item, value: raw as V | undefined }) : tooltip };
    // The only existential boundary is private. A same-family resolved column owns the
    // captured reader; adapters receive that reader's original value after classification.
    // Known scalar domains are guarded by the cell classifier before adapters are used.
    // Public callbacks retain their original generic V, including dynamic unknown.
    if (kind === 'badge') {
        const mapped = options as CaptureProps<T, V, 'badge'>;
        return Object.freeze({ ...base, kind: 'badge', mapValue: captureMap<DataGridBadgeValue>(mapped.values) }) as ResolvedDataGridColumn<T>;
    }
    if (kind === 'icon') {
        const mapped = options as CaptureProps<T, V, 'icon'>;
        return Object.freeze({ ...base, kind: 'icon', mapValue: captureMap<DataGridIconValue>(mapped.values) }) as ResolvedDataGridColumn<T>;
    }
    if (kind === 'person') {
        const person = options as CaptureProps<T, V, 'person'>;
        return Object.freeze({ ...base, kind: 'person', presentPerson: (raw: unknown, item: T) => personPresentation(person, raw as V, item) }) as ResolvedDataGridColumn<T>;
    }
    if (kind === 'people') {
        const people = options as CaptureProps<T, V, 'people'>;
        return Object.freeze({ ...base, kind: 'people', presentPeople: (raw: readonly unknown[], item: T) => Object.freeze(raw.map(value => {
                const person = value as ElementOf<V>;
                return personPresentation(people, person, item, people.getPersonId(person, item));
            })) }) as ResolvedDataGridColumn<T>;
    }
    if (kind === 'tags') {
        const tags = options as CaptureProps<T, V, 'tags'>;
        const mapValue = tags.values === undefined ? undefined : captureMap<DataGridTagValue>(tags.values);
        const presentTags = (raw: readonly unknown[], item: T): readonly ResolvedDataGridTag[] | undefined => {
            const result: ResolvedDataGridTag[] = [];
            for (const value of raw) {
                const tag = value as ElementOf<V>;
                const primitive = mapValue !== undefined && (typeof value === 'string' || (typeof value === 'number' && Number.isFinite(value)));
                const id = primitive ? value : tags.getTagId?.(tag, item);
                if (typeof id !== 'string' && (typeof id !== 'number' || !Number.isFinite(id))) return undefined;
                const presentation = mapValue?.(id);
                if (primitive && presentation === undefined) return undefined;
                const explicitLabel = primitive ? undefined : tags.getLabel?.(tag, item);
                if (!primitive && typeof explicitLabel !== 'string') return undefined;
                const label = presentation?.label ?? explicitLabel;
                if (label === undefined) return undefined;
                result.push(Object.freeze({ id, label, presentation }));
            }
            return Object.freeze(result);
        };
        return Object.freeze({ ...base, kind: 'tags', presentTags }) as ResolvedDataGridColumn<T>;
    }
    if (kind === 'custom') {
        const custom = options as CaptureProps<T, V, 'custom'>;
        return Object.freeze({ ...base, kind: 'custom', renderCell: (raw: unknown, item: T, rowId: DataGridRowId) => custom.renderCell({ item, value: raw as V, rowId }) }) as ResolvedDataGridColumn<T>;
    }
    return Object.freeze(base) as ResolvedDataGridColumn<T>;
}
function captureActions<T extends object>(props: ActionsColumnProps<T>): ResolvedDataGridColumn<T> {
    validateDataGridColumnOptions('actions', props);
    const options = Object.freeze({ ...snapshotOptions(props), actions: Object.freeze(props.actions.map(action => Object.freeze({ ...action }))) });
    const tooltip = props.tooltip;
    return Object.freeze({ kind: 'actions', id: props.id, options, source: undefined, read: () => undefined, staticTooltip: typeof tooltip === 'string' ? tooltip : undefined, tooltip: (_raw: unknown, item: T) => typeof tooltip === 'function' ? tooltip({ item, value: undefined }) : tooltip });
}
/** Feature-local ownership; no module-level registry or external registration. */
function createFamily<T extends object>(): ColumnFamily<T> {
    const descriptors = new WeakMap<object, ResolvedDataGridColumn<T>>();
    const markers = new WeakMap<object, {
        kind: DataGridColumnKind;
        capture: (props: unknown) => ResolvedDataGridColumn<T>;
    }>();
    const define = (column: ResolvedDataGridColumn<T>): DataGridColumnDefinition<T> => {
        // Opaque compile-time invariant brand has no public runtime symbol or callback.
        const descriptor = Object.freeze({ id: column.id, kind: column.kind }) as DataGridColumnDefinition<T>;
        descriptors.set(descriptor, column);
        return descriptor;
    };
    const builder: Partial<DataGridColumnBuilder<T>> = {};
    const members: Partial<DataGridColumnMarkers<T>> = {};
    const register = <K extends ValueKind>(kind: K): void => {
        const capture = <V,>(props: CaptureProps<T, V, K>): ResolvedDataGridColumn<T> => captureValueColumn(kind, props);
        // Each public overload constrains its source before the generic capture closure.
        // The assertion changes call shape only; no callback is widened or re-authored.
        const method = (<V,>(props: CaptureProps<T, V, K>) => define(capture(props))) as DataGridColumnMethod<T, K, DataGridColumnDefinition<T>>;
        Object.assign(builder, { [kind]: method });
        const marker = (): React.ReactElement => <></>;
        markers.set(marker, { kind, capture: props => { validateDataGridColumnOptions(kind, props); return capture(props as CaptureProps<T, unknown, K>); } });
        Object.assign(members, { [names[kind]]: marker });
    };
    for (const kind of Object.keys(names) as DataGridColumnKind[])
        if (kind !== 'actions')
            register(kind);
    builder.actions = props => define(captureActions(props));
    const actionsMarker = (): React.ReactElement => <></>;
    markers.set(actionsMarker, { kind: 'actions', capture: props => { validateDataGridColumnOptions('actions', props); return captureActions(props as ActionsColumnProps<T>); } });
    members.ActionsColumn = actionsMarker;
    const completeBuilder = Object.freeze(builder) as DataGridColumnBuilder<T>;
    const resolveDescriptor = (value: unknown): ResolvedDataGridColumn<T> => {
        assertDataGridConfiguration(typeof value === 'object' && value !== null, 'expected a same-family descriptor');
        const column = descriptors.get(value);
        assertDataGridConfiguration(column !== undefined, 'descriptor belongs to a different family or was not created by defineColumns');
        return column;
    };
    return Object.freeze({ ...members, builder: completeBuilder, resolveDescriptor,
        resolveMarker: (element: React.ReactElement) => {
            assertDataGridConfiguration(typeof element.type === 'function', 'children must be same-family column markers');
            const marker = markers.get(element.type);
            assertDataGridConfiguration(marker !== undefined, 'column marker belongs to a different family');
            return marker.capture(element.props);
        },
        defineColumns: (factory: (builder: DataGridColumnBuilder<T>) => readonly DataGridColumnDefinition<T>[]) => {
            const columns = factory(completeBuilder);
            assertDataGridConfiguration(Array.isArray(columns), 'defineColumns callback must return an array');
            columns.forEach(resolveDescriptor);
            return Object.freeze([...columns]);
        },
    }) as ColumnFamily<T>;
}
/** Explicit typed construction, including typed string-indexed row models. */
export function createTypedColumnFamily<T extends object>(): ColumnFamily<T> { return createFamily<T>(); }
/** Explicit dynamic construction uses unknown sources; no index-signature heuristic. */
export function createDynamicColumnFamily(): DynamicColumnFamily {
    // Same runtime source/capture rules; this explicit constructor selects dynamic public
    // call signatures. Dynamic callbacks were declared unknown, never recast to a domain.
    return createFamily<Record<string, unknown>>() as DynamicColumnFamily;
}

/** Curated runtime suite surface without private reader/descriptor resolution methods. */
export function getPublicDataGridColumnFamily<F extends object>(family: F): Omit<F, 'builder' | 'resolveDescriptor' | 'resolveMarker'> {
    const publicMembers = { ...family };
    for (const key of ['builder', 'resolveDescriptor', 'resolveMarker']) Reflect.deleteProperty(publicMembers, key);
    return publicMembers;
}

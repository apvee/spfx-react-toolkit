import { parseInstant } from '../relative-time/relative-time.internal';
import type { ResolvedDataGridColumn, ResolvedDataGridPerson, ResolvedDataGridTag } from './column-normalization.internal';
import type { DataGridCellValue, DataGridSparklineValue } from './data-grid-cell.types';
import type { DataGridLocale, DataGridSparklinePoint } from './data-grid-column.types';
function calendarTimestamp(raw: unknown): number {
    if (raw instanceof Date)
        return raw.getTime();
    if (typeof raw !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(raw))
        return NaN;
    return parseInstant(`${raw}T00:00:00Z`);
}
function instantTimestamp(raw: unknown): number {
    if (raw instanceof Date)
        return parseInstant(raw);
    if (typeof raw !== 'string' || !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d{1,3})?)?(?:Z|[+-]\d{2}:\d{2})$/.test(raw))
        return NaN;
    return parseInstant(raw);
}
function sparklineValue(raw: readonly unknown[]): DataGridSparklineValue | undefined {
    const points: DataGridSparklinePoint[] = [];
    const numericSeries = typeof raw[0] === 'number';
    let minimum = Infinity, maximum = -Infinity, scale = 0;
    for (let i = 0; i < raw.length; i++) {
        const observation = raw[i];
        if ((typeof observation === 'number') !== numericSeries)
            return undefined;
        let x: number | Date, y: number;
        if (typeof observation === 'number') {
            x = i;
            y = observation;
        }
        else if (typeof observation === 'object' && observation !== null && !Array.isArray(observation) && 'x' in observation && 'y' in observation) {
            const candidate = observation;
            if (!(typeof candidate.x === 'number' && Number.isFinite(candidate.x)) && !(candidate.x instanceof Date && Number.isFinite(candidate.x.getTime())))
                return undefined;
            if (typeof candidate.y !== 'number')
                return undefined;
            x = candidate.x;
            y = candidate.y;
        }
        else
            return undefined;
        if (!Number.isFinite(y))
            return undefined;
        points.push(Object.freeze({ x, y }));
        minimum = Math.min(minimum, y);
        maximum = Math.max(maximum, y);
        scale = Math.max(scale, Math.abs(y));
    }
    const first = points[0].y, last = points[points.length - 1].y;
    // Scaling bounds every addend. Online averaging cannot overflow even for a large series.
    let scaledAverage = 0;
    for (let i = 0; i < points.length; i++)
        scaledAverage += ((scale === 0 ? 0 : points[i].y / scale) - scaledAverage) / (i + 1);
    const difference = last - first;
    const percentage = first === 0 ? NaN : ((last / Math.abs(first)) - (first / Math.abs(first))) * 100;
    return Object.freeze({ points: Object.freeze(points), summary: Object.freeze({ first, last, minimum, maximum,
            average: Math.max(-1, Math.min(1, scaledAverage)) * scale,
            change: Number.isFinite(difference) ? difference : undefined,
            changePercent: Number.isFinite(percentage) ? percentage : undefined }) });
}
function validPerson(person: ResolvedDataGridPerson): boolean {
    return typeof person.name === 'string' && (person.secondaryText === undefined || typeof person.secondaryText === 'string');
}
function validIdentifiedValues(values: readonly (ResolvedDataGridPerson | ResolvedDataGridTag)[]): boolean {
    const ids = new Set<string>();
    for (const value of values) {
        const id = value.id;
        if (typeof id !== 'string' && !(typeof id === 'number' && Number.isFinite(id)))
            return false;
        const key = `${typeof id}:${id}`;
        if (ids.has(key))
            return false;
        ids.add(key);
        if ('name' in value ? !validPerson(value) : typeof value.label !== 'string')
            return false;
    }
    return true;
}
/** Classifies a cached original source once; reader and callback exceptions intentionally surface. */
export function classifyDataGridValue<T extends object>(column: ResolvedDataGridColumn<T>, raw: unknown, item: T): DataGridCellValue {
    const invalid: DataGridCellValue = { status: 'invalid', raw };
    if (column.kind === 'actions')
        return { status: 'valid', kind: 'actions', raw, value: undefined };
    // Legacy external sources may contain null; internally absence uses undefined.
    if (raw === undefined || raw === null)
        return { status: 'missing', raw };
    switch (column.kind) {
        case 'text':
        case 'link': return typeof raw === 'string' ? { status: 'valid', kind: column.kind, raw, value: raw } : invalid;
        case 'boolean': return typeof raw === 'boolean' ? { status: 'valid', kind: 'boolean', raw, value: raw } : invalid;
        case 'number':
        case 'progress':
        case 'rating': {
            if (typeof raw !== 'number' || !Number.isFinite(raw))
                return invalid;
            if (column.kind === 'progress' && (raw < 0 || raw > (column.options.max ?? 1)))
                return invalid;
            if (column.kind === 'rating' && (raw < 0 || raw > (column.options.max ?? 5)))
                return invalid;
            return { status: 'valid', kind: column.kind, raw, value: raw };
        }
        case 'date':
        case 'dateTime': {
            const value = column.kind === 'date' ? calendarTimestamp(raw) : instantTimestamp(raw);
            return Number.isFinite(value) ? { status: 'valid', kind: column.kind, raw, value } : invalid;
        }
        case 'badge':
        case 'icon': {
            if (typeof raw !== 'string' && !(typeof raw === 'number' && Number.isFinite(raw)))
                return invalid;
            if (column.kind === 'badge') {
                const value = column.mapValue(raw);
                return value === undefined ? invalid : { status: 'valid', kind: 'badge', raw, value };
            }
            const value = column.mapValue(raw);
            return value === undefined ? invalid : { status: 'valid', kind: 'icon', raw, value };
        }
        case 'person': {
            const value = column.presentPerson(raw, item);
            return validPerson(value) ? { status: 'valid', kind: 'person', raw, value } : invalid;
        }
        case 'people': {
            if (!Array.isArray(raw))
                return invalid;
            const value = column.presentPeople(raw, item);
            return validIdentifiedValues(value) ? { status: 'valid', kind: 'people', raw, value } : invalid;
        }
        case 'tags': {
            if (!Array.isArray(raw))
                return invalid;
            const value = column.presentTags(raw, item);
            return value === undefined || !validIdentifiedValues(value) ? invalid : { status: 'valid', kind: 'tags', raw, value };
        }
        case 'sparkline': {
            if (!Array.isArray(raw))
                return invalid;
            if (raw.length === 0)
                return { status: 'missing', raw };
            const value = sparklineValue(raw);
            return value === undefined ? invalid : { status: 'valid', kind: 'sparkline', raw, value };
        }
        case 'custom': return { status: 'valid', kind: 'custom', raw, value: raw };
    }
}
export function dataGridIntlLocale(locale: DataGridLocale | undefined): string | string[] | undefined {
    return typeof locale === 'string' || locale === undefined ? locale : [...locale];
}
/** Deterministic text formatting; invalid Intl configuration is never treated as invalid row data. */
export function formatDataGridValue<T extends object>(column: ResolvedDataGridColumn<T>, value: DataGridCellValue, locale: DataGridLocale | undefined, _item: T): string {
    if (value.status !== 'valid')
        return '';
    const resolvedLocale = dataGridIntlLocale(column.options.locale ?? locale);
    switch (value.kind) {
        case 'text':
        case 'link': return value.value;
        case 'boolean': return value.value ? (column.kind === 'boolean' ? column.options.trueLabel : undefined) ?? 'Yes' : (column.kind === 'boolean' ? column.options.falseLabel : undefined) ?? 'No';
        case 'badge':
        case 'icon': return value.value.label;
        case 'person': return [value.value.name, value.value.secondaryText].filter(text => text !== undefined).join(', ');
        case 'people': return value.value.map(person => [person.name, person.secondaryText].filter(text => text !== undefined).join(', ')).join(', ');
        case 'tags': return value.value.map(tag => tag.label).join(', ');
        case 'date':
        case 'dateTime': {
            if (column.kind !== 'date' && column.kind !== 'dateTime')
                return '';
            const defaults: Intl.DateTimeFormatOptions = column.kind === 'dateTime' ? { dateStyle: 'medium', timeStyle: 'short' } : { dateStyle: 'medium' };
            const timeZone = column.kind === 'date' && typeof value.raw === 'string' ? 'UTC' : column.options.timeZone ?? 'UTC';
            return new Intl.DateTimeFormat(resolvedLocale, { ...(column.options.format ?? defaults), timeZone }).format(value.value);
        }
        case 'number':
        case 'progress':
        case 'rating':
        case 'sparkline': {
            const format = column.kind === 'number' || column.kind === 'progress' || column.kind === 'sparkline' ? column.options.format : undefined;
            const numeric = value.kind === 'sparkline' ? value.value.summary.last : value.value;
            const formattedValue = value.kind === 'progress' && column.kind === 'progress' && format?.style === 'percent' ? numeric / (column.options.max ?? 1) : numeric;
            return new Intl.NumberFormat(resolvedLocale, format).format(formattedValue);
        }
        case 'custom': return typeof value.value === 'string' ? value.value : '';
        case 'actions': return '';
    }
}

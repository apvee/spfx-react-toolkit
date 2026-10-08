const minute = 60_000;
const hour = 60 * minute;
const day = 24 * hour;
const dateLimit = 8_640_000_000_000_000;

function daysInMonth(year: number, month: number): number {
  if (month === 1) {
    return year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0) ? 29 : 28;
  }
  return month === 3 || month === 5 || month === 8 || month === 10 ? 30 : 31;
}

/** Parses an explicit instant; ambiguous or impossible input returns NaN. */
export function parseInstant(value: Date | string | number | undefined): number {
  if (value instanceof Date) return value.getTime();
  if (typeof value === 'number') return new Date(value).getTime();
  if (typeof value !== 'string') return NaN;
  const match =
    /^(\d{4}|[+-]\d{6})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?(Z|[+-]\d{2}:\d{2})$/i.exec(value);
  if (!match) return NaN;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const date = Number(match[3]);
  const zone = match[8];
  if (
    month < 1 || month > 12 || date < 1 || date > daysInMonth(year, month - 1) ||
    Number(match[4]) > 23 || Number(match[5]) > 59 || Number(match[6] ?? 0) > 59 ||
    (zone.toUpperCase() !== 'Z' && (Number(zone.slice(1, 3)) > 23 || Number(zone.slice(4)) > 59))
  ) return NaN;
  return Date.parse(value);
}

/** Keeps each clamped anniversary anchored to the original event's UTC date. */
function addMonths(timestamp: number, months: number): number {
  const event = new Date(timestamp);
  const totalMonth = event.getUTCFullYear() * 12 + event.getUTCMonth() + months;
  const year = Math.floor(totalMonth / 12);
  const month = totalMonth - year * 12;
  const date = Math.min(event.getUTCDate(), daysInMonth(year, month));
  // Set the final date in one operation: the first day of an extreme month can
  // lie outside Date's range even when the final clamped anniversary is valid.
  event.setUTCFullYear(year, month, date);
  const result = event.getTime();
  // Signed overflow still permits comparison when counting completed months.
  return Number.isNaN(result) ? (months < 0 ? -Infinity : Infinity) : result;
}

function nextBoundary(timestamp: number): number {
  return Number.isFinite(timestamp) && Math.abs(timestamp) <= dateLimit ? timestamp : Infinity;
}

/** Calculates a label and its first future transition for two valid instants. */
export function getRelativeTime(timestamp: number, reference: number): {
  value: number;
  unit: Intl.RelativeTimeFormatUnit;
  nextUpdate: number;
} {
  const distance = Math.abs(reference - timestamp);
  const future = timestamp > reference;
  const direction = future ? -1 : 1;
  const event = new Date(timestamp);
  const now = new Date(reference);
  let months = Math.abs(
    (now.getUTCFullYear() - event.getUTCFullYear()) * 12 + now.getUTCMonth() - event.getUTCMonth()
  );
  const anniversary = addMonths(timestamp, direction * months);
  if (months > 0 && (future ? anniversary < reference : anniversary > reference)) months--;

  if (months >= 1) {
    const years = months >= 12;
    const scale = years ? 12 : 1;
    const value = Math.floor(months / scale);
    return {
      value: future ? value : -value,
      unit: years ? 'year' : 'month',
      nextUpdate: nextBoundary(future
        ? addMonths(timestamp, -value * scale) + 1
        : addMonths(timestamp, (value + 1) * scale))
    };
  }

  for (const [unit, duration] of [
    ['week', 7 * day], ['day', day], ['hour', hour], ['minute', minute]
  ] as const) {
    if (distance < duration) continue;
    const value = Math.floor(distance / duration);
    return {
      value: future ? value : -value,
      unit,
      nextUpdate: nextBoundary(future
        ? timestamp - value * duration + 1
        : Math.min(timestamp + (value + 1) * duration, addMonths(timestamp, 1)))
    };
  }
  return { value: 0, unit: 'second', nextUpdate: nextBoundary(timestamp + minute) };
}

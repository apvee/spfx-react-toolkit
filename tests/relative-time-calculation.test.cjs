const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const filename = path.resolve(__dirname, '../packages/spfx-react-toolkit/src/controls/relative-time/relative-time.internal.ts');
const output = ts.transpileModule(fs.readFileSync(filename, 'utf8'), {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2019 }, fileName: filename
}).outputText;
const loaded = { exports: {} };
new Function('module', 'exports', output)(loaded, loaded.exports);
const { parseInstant, getRelativeTime } = loaded.exports;
const minute = 60000, hour = 3600000, day = 86400000, limit = 8640000000000000;
const origin = 1769860800000; // 2026-01-31T12:00:00Z, independently fixed epoch.
const label = result => [result.value, result.unit];
function expectLabel(event, reference, value, unit) {
  assert.deepEqual(label(getRelativeTime(event, reference)), [value, unit]);
}
function expectTransition(event, reference, expectedNext) {
  const current = getRelativeTime(event, reference);
  if (expectedNext !== undefined) assert.equal(current.nextUpdate, expectedNext);
  assert.ok(current.nextUpdate > reference);
  if (current.nextUpdate === Infinity) return;
  assert.ok(current.nextUpdate <= limit);
  assert.deepEqual(label(getRelativeTime(event, current.nextUpdate - 1)), label(current));
  assert.notDeepEqual(label(getRelativeTime(event, current.nextUpdate)), label(current));
  assert.deepEqual(label(getRelativeTime(event, current.nextUpdate + 1)), label(getRelativeTime(event, current.nextUpdate)));
}
test('parses explicit instants, fractions, leap dates and the native Date range', () => {
  for (const value of [origin, new Date(origin), '2026-01-31T12:00:00Z', '2026-01-31T13:00:00+01:00', '2026-01-31T06:30:00-05:30']) {
    assert.equal(parseInstant(value), origin);
  }
  assert.equal(parseInstant('2026-01-31T12:00:00.123456789Z'), origin + 123);
  assert.equal(parseInstant('2024-02-29T00:00:00Z'), 1709164800000);
  assert.equal(parseInstant('0000-02-29T00:00:00Z'), -62162121600000);
  for (const value of [-limit, limit]) {
    assert.equal(parseInstant(value), value);
    assert.equal(parseInstant(new Date(value).toISOString()), value);
  }
});
test('rejects ambiguous, impossible and out-of-range instant inputs', () => {
  for (const value of [null, undefined, '', 'invalid', '2026-01-31', '2026-01-31T12:00:00',
    '2026-02-30T12:00:00Z', '2025-02-29T12:00:00Z', '2026-00-01T00:00:00Z',
    '2026-13-01T00:00:00Z', '2026-01-00T00:00:00Z', '2026-04-31T00:00:00Z',
    '2026-01-01T24:00:00Z', '2026-01-01T12:60:00Z', '2026-01-01T12:00:60Z',
    '2026-01-01T12:00:00+25:00', '2026-01-01T12:00:00+01:60',
    '+275760-09-13T00:00:00.001Z', '-271821-04-19T23:59:59.999Z',
    '-000000-01-01T00:00:00Z', NaN, Infinity, -Infinity, limit + 1, -limit - 1, new Date(NaN)]) {
    assert.ok(Number.isNaN(parseInstant(value)), String(value));
  }
});
test('completed absolute units change at both sides of each exact threshold', () => {
  for (const [distance, value, unit] of [[0, 0, 'second'], [minute - 1, 0, 'second'],
    [minute, 1, 'minute'], [hour - 1, 59, 'minute'], [hour, 1, 'hour'],
    [day - 1, 23, 'hour'], [day, 1, 'day'], [7 * day - 1, 6, 'day'], [7 * day, 1, 'week']]) {
    for (const direction of [-1, 1]) {
      expectLabel(origin, origin + direction * distance, value === 0 ? 0 : -direction * value, unit);
      expectTransition(origin, origin + direction * distance);
    }
  }
});
test('UTC calendar anniversaries clamp from the event without accumulating end-of-month drift', () => {
  for (const [event, reference, value, unit] of [
    ['2026-01-31T12:00:00Z', '2026-02-28T11:59:59.999Z', -3, 'week'],
    ['2026-01-31T12:00:00Z', '2026-02-28T12:00:00Z', -1, 'month'],
    ['2026-01-31T12:00:00Z', '2026-03-30T12:00:00Z', -1, 'month'],
    ['2026-01-31T12:00:00Z', '2026-03-31T12:00:00Z', -2, 'month'],
    ['2024-02-29T12:00:00Z', '2025-02-28T12:00:00Z', -1, 'year'],
    ['2024-02-29T12:00:00Z', '2028-02-28T12:00:00Z', -3, 'year'],
    ['2026-03-31T12:00:00Z', '2026-02-28T12:00:00Z', 1, 'month'],
    ['2026-03-31T12:00:00Z', '2026-02-28T12:00:00.001Z', 4, 'week'],
    ['2028-02-29T12:00:00Z', '2027-02-28T12:00:00Z', 1, 'year'],
    ['2028-02-29T12:00:00Z', '2027-02-28T12:00:00.001Z', 11, 'month']]) {
    const timestamp = Date.parse(event), now = Date.parse(reference);
    expectLabel(timestamp, now, value, unit);
    expectTransition(timestamp, now);
  }
});
test('next updates follow future year/month/week and minute/now/past-minute transitions', () => {
  const event = Date.parse('2028-02-29T12:00:00Z');
  expectTransition(event, Date.parse('2027-02-28T12:00:00Z'), Date.parse('2027-02-28T12:00:00.001Z'));
  expectTransition(event, Date.parse('2028-01-29T12:00:00Z'), Date.parse('2028-01-29T12:00:00.001Z'));
  expectTransition(origin, origin - minute, origin - minute + 1);
  expectTransition(origin, origin - minute + 1, origin + minute);
  expectTransition(origin, origin + minute, origin + 2 * minute);
});
test('DST offsets affect instants while elapsed hours remain absolute', () => {
  expectLabel(parseInstant('2026-03-29T01:30:00+01:00'), parseInstant('2026-03-29T03:30:00+02:00'), -1, 'hour');
  expectLabel(parseInstant('2026-10-25T01:30:00+02:00'), parseInstant('2026-10-25T03:30:00+01:00'), -3, 'hour');
});
test('extreme Date values preserve finite labels and suppress unrepresentable next boundaries', () => {
  for (const [event, reference] of [[-limit, limit], [limit, -limit], [-limit, -limit], [limit, limit],
    [limit - day, limit], [limit, limit - minute], [0, limit], [0, -limit],
    [Date.parse('+275760-08-31T12:00:00Z'), limit],
    [Date.parse('-271821-05-31T12:00:00Z'), -limit]]) {
    const result = getRelativeTime(event, reference);
    assert.ok(Number.isFinite(result.value), `${event}/${reference}`);
    assert.ok(['second', 'minute', 'hour', 'day', 'week', 'month', 'year'].includes(result.unit));
    expectTransition(event, reference);
  }
  assert.equal(getRelativeTime(limit, limit).nextUpdate, Infinity);
  assert.equal(getRelativeTime(0, limit).nextUpdate, Infinity);
  expectLabel(-limit, limit, -547581, 'year');
  expectLabel(limit, -limit, 547581, 'year');
});

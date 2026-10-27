const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let api;
test.before(async () => { api = await harness.loadModule('packages/spfx-react-toolkit/src/controls/data-grid/data-grid-values.internal.ts'); });
test.after(() => harness.cleanup());
const column = (kind, options = {}) => ({ kind, id: kind, options, read: item => item.value, tooltip: () => undefined });
const classify = (kind, raw, options) => api.classifyDataGridValue(column(kind, options), raw, {});
test('missing differs from invalid and false zero empty text remain valid', () => {
    for (const value of [undefined, null])
        assert.equal(classify('text', value).status, 'missing');
    for (const [kind, value] of [['text', ''], ['number', 0], ['boolean', false]])
        assert.equal(classify(kind, value).status, 'valid');
    for (const value of [NaN, Infinity, -Infinity, '1', {}])
        assert.equal(classify('number', value).status, 'invalid');
    assert.equal(classify('text', {}).status, 'invalid');
});
test('strict calendar and instant domains preserve low years and reject impossible dates', () => {
    for (const date of ['2024-02-29', '0000-02-29', '0099-01-01'])
        assert.equal(classify('date', date).status, 'valid');
    for (const date of ['2023-02-29', '2024-04-31', '2024-01-01T00:00:00Z', 0])
        assert.equal(classify('date', date).status, 'invalid');
    for (const instant of ['2024-02-29T23:59:59.123Z', '0099-01-01T00:00:00+02:00'])
        assert.equal(classify('dateTime', instant).status, 'valid');
    for (const instant of ['2024-01-01T00:00:00', '2024-01-01t00:00:00z', '+002024-01-01T00:00:00Z', '2024-01-01T00:00:00.1234Z', '2023-02-29T00:00:00Z', 0])
        assert.equal(classify('dateTime', instant).status, 'invalid');
    const value = classify('date', '0099-01-02');
    assert.equal(new Date(value.value).getUTCFullYear(), 99);
    assert.match(api.formatDataGridValue(column('date', { format: { year: 'numeric', month: 'numeric', day: 'numeric' } }), value, 'en-US', {}), /1\/2\/99/);
});
test('format configuration errors throw and column locale overrides root', () => {
    const c = column('number', { locale: 'de-DE' });
    const value = api.classifyDataGridValue(c, 1234.5, {});
    assert.equal(api.formatDataGridValue(c, value, 'en-US', {}), '1.234,5');
    assert.throws(() => api.formatDataGridValue(column('number', { locale: 'bad_!' }), value, 'en-US', {}), RangeError);
    assert.throws(() => api.formatDataGridValue(column('date', { timeZone: 'bad-zone' }), classify('date', new Date('2024-01-01T00:00:00Z')), 'en-US', {}), RangeError);
});
test('range violations are invalid and mapped callbacks preserve raw identity', () => {
    assert.equal(classify('progress', 11, { max: 10 }).status, 'invalid');
    assert.equal(classify('rating', -1).status, 'invalid');
    const raw = {};
    let received;
    const c = { ...column('badge'), mapValue: value => { received = value; return { label: 'Mapped' }; } };
    assert.equal(api.classifyDataGridValue(c, raw, {}).status, 'invalid');
    const valid = api.classifyDataGridValue(c, 'x', {});
    assert.equal(received, 'x');
    assert.equal(api.formatDataGridValue(c, valid, 'en-US', {}), 'Mapped');
});
test('sparkline summary avoids overflow and computes metrics independently without mutations', () => {
    const points = Object.freeze([Number.MAX_VALUE, -Number.MAX_VALUE]);
    const v = classify('sparkline', points);
    assert.equal(v.status, 'valid');
    assert.equal(v.value.summary.change, undefined);
    assert.equal(v.value.summary.changePercent, -200);
    assert.equal(v.value.summary.average, 0);
    assert.equal(classify('sparkline', []).status, 'missing');
    assert.equal(classify('sparkline', [1, NaN]).status, 'invalid');
    const many = Array.from({ length: 100000 }, () => Number.MAX_VALUE);
    assert.equal(classify('sparkline', many).value.summary.average, Number.MAX_VALUE);
    assert.equal(classify('sparkline', [0, 1]).value.summary.changePercent, undefined);
    assert.deepEqual(points, [Number.MAX_VALUE, -Number.MAX_VALUE]);
});
test('people and tags reject invalid or duplicate identities and retain original callback objects', () => {
    const item = { value: 'row' };
    const person = { id: 'a', name: 'Alpha' };
    let received;
    const peopleColumn = { ...column('people'), presentPeople: (raw, row) => { received = { raw, row }; return raw; } };
    const raw = Object.freeze([person]);
    const valid = api.classifyDataGridValue(peopleColumn, raw, item);
    assert.equal(valid.status, 'valid');
    assert.strictEqual(received.raw, raw);
    assert.strictEqual(received.row, item);
    assert.strictEqual(valid.value[0], person);
    for (const invalid of [[person, person], [{ id: NaN, name: 'Invalid' }], [{ name: 'Missing ID' }], [{ id: 'a', name: {} }]])
        assert.equal(api.classifyDataGridValue(peopleColumn, invalid, item).status, 'invalid');
    const tagsColumn = { ...column('tags'), presentTags: raw => raw };
    assert.equal(api.classifyDataGridValue(tagsColumn, [{ id: 'a', label: 'Alpha' }, { id: 'a', label: 'Duplicate' }], item).status, 'invalid');
    assert.equal(classify('sparkline', [1, { x: 1, y: 2 }]).status, 'invalid');
});
test('calendar strings stay UTC across configured zones while Date values retain their instant', () => {
    const c = column('date', { timeZone: 'America/Los_Angeles', format: { year: 'numeric', month: '2-digit', day: '2-digit' } });
    assert.equal(api.formatDataGridValue(c, api.classifyDataGridValue(c, '2024-01-02', {}), 'en-US', {}), '01/02/2024');
    const date = new Date('2024-01-02T23:00:00Z');
    const value = api.classifyDataGridValue(c, date, {});
    assert.equal(value.value, date.getTime());
    assert.equal(api.formatDataGridValue(c, value, 'en-US', {}), '01/02/2024');
});
test('progress percent formatting uses value-to-configured-max ratio and source date defaults remain medium', () => {
    const c = column('progress', { max: 200, format: { style: 'percent', maximumFractionDigits: 1 } });
    assert.equal(api.formatDataGridValue(c, api.classifyDataGridValue(c, 50, {}), 'en-US', {}), '25%');
    assert.equal(api.formatDataGridValue(column('date'), classify('date', '2024-02-29'), 'en-US', {}), 'Feb 29, 2024');
});
test('chart X coordinates validate numeric or Date domains without coercion', () => {
    for (const points of [[{ x: Infinity, y: 1 }], [{ x: new Date('invalid'), y: 1 }], [{ x: '2024-01-01', y: 1 }], [{ x: 0, y: '1' }]])
        assert.equal(classify('sparkline', points).status, 'invalid');
    const raw = Object.freeze([{ x: -Number.MAX_VALUE, y: 0 }, { x: Number.MAX_VALUE, y: 2 }]);
    const valid = classify('sparkline', raw);
    assert.equal(valid.status, 'valid');
    assert.equal(valid.value.summary.average, 1);
    assert.deepEqual(raw, [{ x: -Number.MAX_VALUE, y: 0 }, { x: Number.MAX_VALUE, y: 2 }]);
});

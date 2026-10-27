const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/data-grid-cells-harness.cjs');
let api, values, React, ReactDOM, act, host, renderer, fluent;
const strings = { retry: 'Retry', invalidValue: 'Invalid value', noObservations: 'No observations', sparklineSummary: v => `First ${v.first}; Latest ${v.last}; Minimum ${v.minimum}; Maximum ${v.maximum}; Average ${v.average}; Change ${v.change}; Percent ${v.changePercent}` };
test.before(async () => {
    const dom = harness.createSxDom();
    for (const name of ['NodeFilter', 'Node', 'HTMLElement', 'Element', 'Event', 'CustomEvent'])
        global[name] = dom.window[name];
    dom.window.ResizeObserver = class {
        observe() { }
        unobserve() { }
        disconnect() { }
    };
    Object.defineProperty(dom.window, 'crypto', { value: require('node:crypto').webcrypto });
    [api, values] = await Promise.all(['cells/data-grid-sparkline.internal.tsx', 'data-grid-values.internal.ts'].map(f => harness.loadModule(`packages/spfx-react-toolkit/src/controls/data-grid/${f}`)));
    React = require('react');
    ReactDOM = require('react-dom');
    ({ act } = require('react-dom/test-utils'));
    fluent = require('@fluentui/react-components');
    renderer = require('@griffel/core').createDOMRenderer(document);
    host = document.createElement('div');
    document.body.appendChild(host);
});
test.after(async () => { if (host) {
    act(() => { ReactDOM.unmountComponentAtNode(host); });
    host.remove();
} await harness.cleanup(); });
function render(raw, { color = '#0067b8', showValue = true, dir = 'ltr', theme = fluent.webLightTheme, adapter = api.DataGridSparklineCell, locale = 'en-US', cellStrings = strings } = {}) { const { RendererProvider } = require('@griffel/react'); const item = { raw }; const column = { kind: 'sparkline', id: 'trend', options: { color, showValue }, tooltip: () => undefined }; const classified = values.classifyDataGridValue(column, raw, item); assert.equal(classified.status, 'valid'); const context = { column, item, rowId: 'original', value: classified, locale, strings: cellStrings }; act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer }, React.createElement(fluent.FluentProvider, { theme, dir }, React.createElement(adapter, { column, context, value: classified.value }))), host); }); return host.querySelector('svg'); }
test('official chart remounts when plot data change and preserves SVG on color theme RTL-only updates', () => {
    const first = render([1, 2, 1]);
    const firstPath = first.querySelector('path.line').getAttribute('d');
    const recolored = render([1, 2, 1], { color: '#d13438', dir: 'rtl', theme: fluent.webDarkTheme });
    assert.strictEqual(recolored, first);
    assert.equal(recolored.querySelector('path.line').getAttribute('stroke'), '#d13438');
    const changed = render([1, 3, 2]);
    assert.notStrictEqual(changed, first);
    assert.notEqual(changed.querySelector('path.line').getAttribute('d'), firstPath);
});
test('negative constant extreme and Date-X series yield finite paths without mutating inputs', () => {
    for (const raw of [Object.freeze([-3, -2, -1]), Object.freeze([7, 7, 7]), Object.freeze([Number.MAX_VALUE, -Number.MAX_VALUE]), Object.freeze([{ x: new Date('2024-01-01'), y: 1 }, { x: new Date('2024-01-02'), y: 2 }])]) {
        const svg = render(raw);
        for (const path of svg.querySelectorAll('path'))
            assert.doesNotMatch(path.getAttribute('d'), /NaN|Infinity/);
    }
});
test('summary remains available when visible value hidden and decorative SVG is inert', () => {
    render([1, 2, 3], { showValue: false });
    const label = host.querySelector('[role=img]').getAttribute('aria-label');
    assert.match(label, /First 1.*Latest 3.*Minimum 1.*Maximum 3.*Average 2.*Change 2.*Percent 200/);
    assert.equal(host.querySelector('[data-grid-sparkline-value]'), null);
    const svg = host.querySelector('svg');
    const decoration = svg.closest('[inert]');
    assert.ok(decoration);
    assert.equal(decoration.getAttribute('aria-hidden'), 'true');
    assert.equal(decoration.hasAttribute('style'), false);
});
test('color callback exceptions surface instead of being swallowed by chart boundary', () => {
    const original = console.error;
    const errors = [];
    console.error = (...args) => errors.push(args.map(String).join(' '));
    try {
        assert.throws(() => render([1, 2], { color: () => { throw Error('consumer callback'); } }), /consumer callback/);
        assert.ok(errors.some(e => e.includes('consumer callback')));
    }
    finally {
        console.error = original;
    }
});
test('isolated chart failure retries and resets on changed inputs, then renders official Charts again', async () => {
    const faultModule = await harness.loadChartFaultModule('packages/spfx-react-toolkit/src/controls/data-grid/cells/data-grid-sparkline.internal.tsx');
    const adapter = faultModule.DataGridSparklineCell;
    const original = console.error;
    const errors = [];
    console.error = (...args) => errors.push(args.map(String).join(' '));
    try {
        harness.setChartFailure(true);
        render([1, 2, 3], { adapter });
        assert.equal(host.querySelector('svg'), null);
        assert.equal(host.querySelector('button').textContent, 'Retry');
        assert.match(host.querySelector('[role=img]').getAttribute('aria-label'), /Latest 3/);
        harness.setChartFailure(false);
        act(() => host.querySelector('button').click());
        assert.ok(host.querySelector('path.line'));
        assert.equal(host.querySelector('button'), null);
        harness.setChartFailure(true);
        render([1, 3, 2], { adapter });
        assert.ok(host.querySelector('button'));
        const attempts = harness.chartAttempts();
        render([1, 3, 2], { adapter });
        assert.equal(harness.chartAttempts(), attempts, 'unchanged inputs do not loop retry');
        harness.setChartFailure(false);
        render([1, 4, 2], { adapter });
        assert.ok(host.querySelector('path.line'));
        assert.equal(host.querySelector('button'), null);
        assert.ok(errors.some(e => e.includes('deliberate chart rendering failure')));
    }
    finally {
        harness.setChartFailure(false);
        console.error = original;
    }
});
test('large official chart series has finite coordinates and keeps original summary values', () => {
    const raw = Object.freeze(Array.from({ length: 20000 }, (_, i) => ({ x: i, y: (i % 19) - 10 })));
    const svg = render(raw);
    const path = svg.querySelector('path.line').getAttribute('d');
    assert.ok(path.length > 20000);
    assert.doesNotMatch(path, /NaN|Infinity/);
    const summary = host.querySelector('[role=img]').getAttribute('aria-label');
    assert.match(summary, /Minimum -10.*Maximum 8/);
    assert.equal(raw.length, 20000);
});


test('custom English chart summaries keep Italian formatted metrics and original numeric callback data', async () => {
    const state = await harness.loadModule('packages/spfx-react-toolkit/src/controls/data-grid/state-validation.internal.ts', { nativeCulture: 'it-IT' });
    let formatted, received;
    const cellStrings = state.resolveDataGridStrings({ sparklineSummary: values => { formatted = values; return `English latest ${values.last}`; } });
    render([10000.5, 20000.5], { locale: 'it-IT', cellStrings, color: data => { received = data; return '#0067b8'; } });
    assert.equal(host.querySelector('[role=img]').getAttribute('aria-label'), 'English latest 20.000,5');
    assert.equal(formatted.first, '10.000,5'); assert.equal(formatted.last, '20.000,5');
    assert.equal(received.summary.first, 10000.5); assert.equal(received.summary.last, 20000.5);
    assert.deepEqual(received.item.raw, [10000.5, 20000.5]);
    render([10000.5, 20000.5], { locale: 'it-IT', cellStrings: state.resolveDataGridStrings() });
    assert.match(host.querySelector('[role=img]').getAttribute('aria-label'), /Primo 10[.]000,5; Ultimo 20[.]000,5/);
});

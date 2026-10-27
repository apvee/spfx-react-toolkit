const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let cells, values, family, normalization, stateValidation, React, ReactDOM, act, host, renderer, fluent;
const strings = { yes: 'Oui', no: 'Non', moreActions: 'More actions', noTags: 'No tags', tagOverflow: n => `${n} more tags`, noObservations: 'No observations', invalidValue: 'Invalid value', retry: 'Retry' };
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
    [cells, values, family, normalization, stateValidation] = await Promise.all(['cells/data-grid-cells.internal.tsx', 'data-grid-values.internal.ts', 'column-family.internal.tsx', 'column-normalization.internal.ts', 'state-validation.internal.ts'].map(file => harness.loadModule(`packages/spfx-react-toolkit/src/controls/data-grid/${file}`)));
    strings.ratingCount = stateValidation.resolveDataGridStrings().ratingCount;
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
function col(kind, options = {}) { const f = family.createTypedColumnFamily(); const descriptors = f.defineColumns(c => [c[kind]({ id: kind, header: kind, ...(kind === 'actions' ? {} : { field: 'value' }), ...options })]); return normalization.normalizeDataGridColumns(f, { columns: descriptors })[0]; }
function render(column, item, rowId = 7, cellStrings = strings) {
    const { RendererProvider } = require('@griffel/react');
    const context = { column, item, rowId, value: values.classifyDataGridValue(column, item.value, item), locale: 'en-US', size: 'medium', strings: cellStrings };
    act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer }, React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme }, cells.renderDataGridCell(context))), host); });
    return host;
}
test('rating hides numeric slots but preserves an accessible scale and optional vote count', () => {
    const rating = (options, value = 4) => render(col('rating', options), { value }).querySelector('.fui-RatingDisplay');
    let root = rating({ max: 5, count: 12 });
    assert.equal(root.querySelector('.fui-RatingDisplay__valueText') === null, true);
    assert.equal(root.querySelector('.fui-RatingDisplay__countText') === null, true);
    assert.equal(root.getAttribute('aria-label'), '4/5, 12 votes');
    root = rating({ max: 5, count: 0 }, 0);
    assert.equal(root.getAttribute('aria-label'), '0/5, 0 votes');
    root = rating({ max: 5, count: 1 });
    assert.equal(root.getAttribute('aria-label'), '4/5, 1 vote');
    root = rating({ max: 5 });
    assert.equal(root.querySelector('.fui-RatingDisplay__countText') === null, true);
    assert.equal(root.getAttribute('aria-label'), '4/5');
});
test('rating localizes fractional values and count, preserving callback item identity and compact props', () => {
    const item = Object.freeze({ value: 4.5 });
    let received;
    const localized = stateValidation.resolveDataGridStrings({ ratingCount: (count, formatted) => {
        received = [count, formatted];
        return `${formatted} voti`;
    } });
    render(col('rating', { locale: 'de-DE', max: 5, count: context => {
        assert.strictEqual(context.item, item);
        assert.equal(context.value, 4.5);
        return 1234;
    }, ratingDisplayProps: { compact: true, size: 'small' } }), item, 7, localized);
    assert.equal(host.querySelector('.fui-RatingDisplay').getAttribute('aria-label'), '4,5/5, 1.234 voti');
    assert.equal(host.querySelector('.fui-RatingDisplay').textContent, '');
    assert.deepEqual(received, [1234, '1.234']);
    assert.equal(host.querySelectorAll('.fui-RatingItem').length, 1);
});
test('rating defaults to native marigold and preserves an explicit consumer color', () => {
    const selectedClass = options => render(col('rating', options), { value: 4 }).querySelector('.fui-RatingItem__selectedIcon').className;
    const defaultClass = selectedClass({});
    assert.equal(defaultClass, selectedClass({ ratingDisplayProps: { color: 'marigold' } }));
    assert.notEqual(defaultClass, selectedClass({ ratingDisplayProps: { color: 'neutral' } }));
});
test('rating defaults to filled stars and preserves an explicit outline icon', () => {
    const { Star20Filled, Star20Regular } = require('@fluentui/react-icons');
    const paths = options => Array.from(
        render(col('rating', options), { value: 4.5 }).querySelectorAll('.fui-RatingItem svg path'),
        path => path.getAttribute('d'),
    );
    const defaultPaths = paths({});
    assert.ok(defaultPaths.length > 0);
    assert.deepEqual(defaultPaths, paths({ ratingDisplayProps: { icon: Star20Filled } }));
    assert.notDeepEqual(defaultPaths, paths({ ratingDisplayProps: { icon: Star20Regular } }));
});
test('native scalar and mapped cells preserve valid zero false empty and accessible icon labels', () => {
    assert.equal(render(col('number'), { value: 0 }).textContent, '0');
    assert.equal(render(col('text'), { value: '' }).textContent, '');
    assert.equal(render(col('boolean'), { value: false }).textContent, 'Non');
    const icon = render(col('boolean', { display: 'icon' }), { value: true }).querySelector('[role=img]');
    assert.equal(icon.getAttribute('aria-label'), 'Oui');
    assert.ok(icon.querySelector('svg'));
    assert.equal(render(col('badge', { values: { ok: { label: 'Ready' } } }), { value: 'ok' }).textContent, 'Ready');
    assert.equal(render(col('icon', { values: { ok: { label: 'Ready', icon: React.createElement('svg') } } }), { value: 'ok' }).querySelector('[role=img]').getAttribute('aria-label'), 'Ready');
    assert.match(render(col('date'), { value: '2024-02-29' }).textContent, /Feb 29, 2024/);
    assert.match(render(col('dateTime'), { value: '2024-02-29T23:00:00Z' }).textContent, /11:00/);
    assert.equal(render(col('text', { invalidPlaceholder: 'Bad', placeholder: 'Missing' }), { value: {} }).textContent, 'Bad');
});
test('link inline and menu actions receive original item and IDs and disabled actions do not fire', () => {
    const item = Object.freeze({ value: 'Open' });
    let received;
    render(col('link', { href: (row, value) => { assert.strictEqual(row, item); assert.equal(value, 'Open'); return '/target'; }, onNavigate: (_event, data) => received = data }), item).querySelector('a').click();
    assert.strictEqual(received.item, item);
    assert.equal(received.rowId, 7);
    assert.equal(received.value, 'Open');
    assert.equal(received.href, '/target');
    received = undefined;
    render(col('actions', { actions: [{ id: 'run', label: 'Execute very long localized action', placement: 'inline', onClick: (_event, data) => received = data }, { id: 'disabled', label: 'Disabled', placement: 'inline', disabled: true, onClick: () => assert.fail('disabled called') }, { id: 'menu', label: 'Menu action', placement: 'menu', onClick: (_event, data) => received = data }] }), item, 'row-original');
    const buttons = host.querySelectorAll('button');
    buttons[0].click();
    assert.strictEqual(received.item, item);
    assert.equal(received.rowId, 'row-original');
    assert.equal(received.actionId, 'run');
    buttons[1].click();
    assert.equal(received.actionId, 'run');
    act(() => buttons[2].click());
    const menu = document.querySelector('[role=menuitem]');
    assert.equal(menu.textContent, 'Menu action');
    act(() => menu.click());
    assert.equal(received.actionId, 'menu');
});
test('custom JSX keeps identity and original context without stringifying objects', () => {
    const item = { value: { name: 'Original' } };
    const node = React.createElement('strong', {}, 'Custom');
    let received;
    const c = col('custom', { renderCell: context => { received = context; return node; } });
    const context = { column: c, item, rowId: 'row', value: values.classifyDataGridValue(c, item.value, item), strings };
    assert.strictEqual(cells.renderDataGridCell(context), node);
    assert.strictEqual(received.item, item);
    assert.strictEqual(received.value, item.value);
    assert.equal(received.rowId, 'row');
});
test('native people and tag overflow recover original stable labels', () => {
    const people = Object.freeze([{ id: 'a', name: 'Alpha' }, { id: 'b', name: 'Beta' }, { id: 'c', name: 'Gamma' }]);
    render(col('person', { getName: p => p.name }), { value: people[0] });
    assert.match(host.textContent, /Alpha/);
    render(col('people', { getName: p => p.name, getPersonId: p => p.id, maxVisible: 1 }), { value: people });
    const button = host.querySelector('button');
    assert.ok(button);
    act(() => button.click());
    assert.match(document.querySelector('.fui-PopoverSurface').textContent, /Beta.*Gamma/);
    render(col('tags', { values: { a: { label: 'Alpha' }, b: { label: 'Beta' }, c: { label: 'Gamma' } }, maxVisible: 1 }), { value: ['a', 'b', 'c'] });
    const tagTrigger = host.querySelector('button');
    assert.equal(tagTrigger.getAttribute('aria-label'), '2 more tags');
    act(() => tagTrigger.click());
    assert.match(document.querySelector('.fui-PopoverSurface').textContent, /Beta.*Gamma/);
    assert.equal(render(col('tags', { values: {} }), { value: [] }).textContent, 'No tags');
});
test('progress obeys max with readonly rating and authored cells have no inline styles', () => {
    render(col('progress', { max: 200, showValue: true }), { value: 50 });
    const progress = host.querySelector('[role=progressbar]');
    assert.equal(progress.getAttribute('aria-valuemax'), '200');
    assert.equal(progress.getAttribute('aria-valuenow'), '50');
    assert.match(host.textContent, /50/);
    render(col('rating', { max: 10, count: 20 }), { value: 7 });
    assert.ok(host.querySelector('.fui-RatingDisplay'));
    assert.equal(host.querySelector('input'), null);
    render(col('actions', { actions: [{ id: 'a', label: 'Wide label', placement: 'inline', onClick() { } }] }), { value: undefined });
    assert.equal(host.querySelector('[data-grid-actions]').hasAttribute('style'), false);
});
test('link forwards consumer classes attributes events and style escape hatch', () => {
    let copyEvents = 0;
    render(col('link', { href: '/target', linkProps: { className: 'consumer-link', style: { color: 'rgb(255, 0, 0)' }, 'aria-label': 'Consumer navigation', title: 'Consumer title', onCopy: () => copyEvents++ } }), { value: 'Open' });
    const link = host.querySelector('a');
    assert.ok(link.classList.contains('consumer-link'));
    assert.equal(link.style.color, 'rgb(255, 0, 0)');
    assert.equal(link.getAttribute('aria-label'), 'Consumer navigation');
    assert.equal(link.title, 'Consumer title');
    act(() => { link.dispatchEvent(new Event('copy', { bubbles: true })); });
    assert.equal(copyEvents, 1);
});
test('localized boolean tooltip matches visible/icon labels and invalid values never reach typed tooltip callbacks', () => {
    const item = { value: false };
    const c = col('boolean', { falseLabel: 'Explicit no' });
    const value = values.classifyDataGridValue(c, item.value, item);
    assert.equal(cells.getDataGridCellTooltip({ column: c, item, rowId: 1, value, strings, locale: 'en-US' }), 'Explicit no');
    assert.equal(render(c, item).textContent, 'Explicit no');
    const invalidItem = { value: { unexpected: true } };
    let callbacks = 0;
    const invalidColumn = col('number', { tooltip: () => { callbacks++; return 'Typed'; } });
    const invalidValue = values.classifyDataGridValue(invalidColumn, invalidItem.value, invalidItem);
    assert.equal(cells.getDataGridCellTooltip({ column: invalidColumn, item: invalidItem, rowId: 1, value: invalidValue, strings, locale: 'en-US' }), 'Invalid value');
    assert.equal(callbacks, 0);
    const staticColumn = col('number', { tooltip: 'Static description' });
    assert.equal(cells.getDataGridCellTooltip({ column: staticColumn, item: invalidItem, rowId: 1, value: invalidValue, strings, locale: 'en-US' }), 'Static description');
});
test('missing legacy nullable source uses undefined in tooltip callback and callback errors surface', () => {
    let received = 'unread';
    const item = { value: null };
    const c = col('text', { tooltip: context => { received = context.value; return 'Missing value'; } });
    const context = { column: c, item, rowId: 1, value: values.classifyDataGridValue(c, item.value, item), strings, locale: 'en-US' };
    assert.equal(cells.getDataGridCellTooltip(context), 'Missing value');
    assert.equal(received, undefined);
    assert.equal(context.value.raw, null);
    const throwing = col('text', { tooltip: () => { throw Error('tooltip callback'); } });
    assert.throws(() => cells.getDataGridCellTooltip({ ...context, column: throwing }), /tooltip callback/);
});
test('badge defaults to tint and preserves an explicit appearance', () => {
    const c = col('badge', { values: { ok: { label: 'Ready', color: 'success', appearance: undefined } } });
    const badge = render(c, { value: 'ok' }).querySelector('.fui-Badge').className;
    const tinted = render(col('badge', { values: { ok: { label: 'Ready', color: 'success', appearance: 'tint' } } }), { value: 'ok' }).querySelector('.fui-Badge').className;
    assert.equal(badge, tinted);
    const filled = render(col('badge', { values: { ok: { label: 'Ready', color: 'success', appearance: 'filled' } } }), { value: 'ok' }).querySelector('.fui-Badge').className;
    assert.notEqual(filled, tinted);
});
test('person centers text by default and preserves consumer alignment', () => {
    const item = { value: { name: 'Aligned person' } };
    const persona = render(col('person', { getName: p => p.name, personaProps: { textAlignment: undefined } }), item).querySelector('.fui-Persona__primaryText').className;
    const centered = render(col('person', { getName: p => p.name, personaProps: { textAlignment: 'center' } }), item).querySelector('.fui-Persona__primaryText').className;
    assert.equal(persona, centered);
    const start = render(col('person', { getName: p => p.name, personaProps: { textAlignment: 'start' } }), item).querySelector('.fui-Persona__primaryText').className;
    assert.notEqual(start, centered);
});
test('people overflow contains one avatar per person and retains secondary text and presence', () => {
    const people = [{ id: 'a', name: 'Alpha', role: 'Owner' }, { id: 'b', name: 'Beta', role: 'Reviewer' }];
    render(col('people', { getName: p => p.name, getPersonId: p => p.id, getSecondaryText: p => p.role,
        getPresenceProps: () => ({ status: 'available' }), maxVisible: 1 }), { value: people });
    act(() => host.querySelector('button').click());
    const surface = document.querySelector('.fui-AvatarGroupPopover__popoverSurface');
    const item = surface.querySelector('li');
    assert.equal(item.querySelectorAll('.fui-Avatar').length, 1);
    assert.equal(item.querySelectorAll('.fui-Persona__primaryText').length, 1);
    assert.equal(item.querySelector('.fui-Persona__primaryText').textContent, 'Beta');
    assert.equal(item.querySelector('.fui-Persona__secondaryText').textContent, 'Reviewer');
    assert.ok(item.querySelector('.fui-PresenceBadge'));
});
test('boolean icons use native 20px variants with one accessible label', () => {
    const column = col('boolean', { display: 'icon', trueLabel: 'Enabled yes', falseLabel: 'Enabled no' });
    for (const [value, label] of [[true, 'Enabled yes'], [false, 'Enabled no']]) {
        const root = render(column, { value }).querySelector('[role=img]');
        const svg = root.querySelector('svg');
        assert.equal(root.getAttribute('aria-label'), label);
        assert.equal(svg.getAttribute('aria-hidden'), 'true');
        assert.equal(svg.getAttribute('width'), '20');
        assert.equal(svg.getAttribute('height'), '20');
    }
});


test('missing and invalid presentation slots retain outer text spans and nullish defaults', () => {
    const ref = React.createRef(); let clicked = 0, calls = 0;
    const column = col('number', {
        placeholder: { ref, className: 'inner-placeholder', onClick: () => clicked++, children: (Component, props) => { calls++; assert.equal(Component, 'span'); return React.createElement(Component, props, 'Missing slot'); } },
        invalidPlaceholder: { 'data-invalid': '', children: 'Invalid **slot**', as: 'div', dangerouslySetInnerHTML: { __html: 'UNSAFE' } },
    });
    render(column, { value: undefined });
    assert.equal(ref.current.tagName, 'SPAN'); assert.equal(ref.current.parentElement.tagName, 'SPAN'); assert.ok(ref.current.parentElement.className);
    act(() => ref.current.dispatchEvent(new window.MouseEvent('click', { bubbles: true }))); assert.equal(clicked, 1); assert.ok(calls > 0);
    const invalid = render(column, { value: 'bad' }).querySelector('[data-invalid]');
    assert.equal(invalid.tagName, 'SPAN'); assert.equal(invalid.textContent, 'Invalid **slot**'); assert.doesNotMatch(host.innerHTML, /UNSAFE/);
    const before = calls; render(column, { value: 2 }); assert.equal(calls, before);
    for (const value of [undefined, null]) {
        assert.equal(render(col('number', { placeholder: value }), {}).textContent, '—');
        assert.equal(render(col('number', { invalidPlaceholder: value }), { value: 'bad' }).textContent, strings.invalidValue);
    }
    for (const value of [false, true, '', React.createElement(React.Fragment)]) {
        assert.equal(render(col('number', { placeholder: value }), {}).textContent, '');
        assert.equal(render(col('number', { invalidPlaceholder: value }), { value: 'bad' }).textContent, '');
    }
    assert.equal(render(col('sparkline'), {}).textContent, strings.noObservations);
    assert.equal(render(col('custom', { renderCell: () => React.createElement('strong', {}, 'Unchanged custom'), placeholder: { children: 'Missing custom' } }), { value: 'valid' }).querySelector('strong').parentElement.classList.contains('fui-FluentProvider'), true);
});


test('native UI grammar stays separate from Italian rating formatting and tag callbacks retain counts', async () => {
    const italian = await harness.loadModule('packages/spfx-react-toolkit/src/controls/data-grid/state-validation.internal.ts', { nativeCulture: 'it-IT' });
    const french = await harness.loadModule('packages/spfx-react-toolkit/src/controls/data-grid/state-validation.internal.ts', { nativeCulture: 'fr-FR' });
    const baseline = italian.resolveDataGridStrings();
    render(col('rating', { locale: 'it-IT', max: 5, count: 12345 }), { value: 4.5 }, 7, stateValidation.resolveDataGridStrings());
    assert.equal(host.querySelector('.fui-RatingDisplay').getAttribute('aria-label'), '4,5/5, 12.345 votes', 'English UI dictionary with Italian numbers');
    render(col('rating', { locale: 'it-IT', max: 5, count: 12345 }), { value: 4.5 }, 7, baseline);
    assert.equal(host.querySelector('.fui-RatingDisplay').getAttribute('aria-label'), '4,5/5, 12.345 voti');
    assert.equal(french.resolveDataGridStrings().ratingCount(0, '0'), '0 vote', 'Resource source culture chooses French zero grammar');
    assert.equal(baseline.selectRow('$&{rowId}'), 'Seleziona la riga $&{rowId}', 'Template substitution preserves caller data literally');
    let hiddenCount;
    const custom = italian.resolveDataGridStrings({ tagOverflow: count => { hiddenCount = count; return `Hidden ${count}`; } });
    render(col('tags', { values: { a: { label: 'Alpha' }, b: { label: 'Beta' }, c: { label: 'Gamma' } }, maxVisible: 1 }), { value: ['a', 'b', 'c'] }, 7, custom);
    assert.equal(hiddenCount, 2); assert.equal(host.querySelector('button').getAttribute('aria-label'), 'Hidden 2');
    assert.equal(baseline.tagOverflow(1), 'Mostra 1 altro tag');
    assert.equal(baseline.tagOverflow(2), 'Mostra altri 2 tag');
});

const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, act, api, dom;
test.before(async () => {
  dom = harness.createSxDom();
  for (const key of ['Node', 'NodeFilter', 'Element', 'HTMLElement', 'HTMLInputElement', 'Event', 'CustomEvent', 'MouseEvent', 'KeyboardEvent', 'File', 'FileList', 'Blob']) global[key] = dom.window[key];
  dom.window.requestAnimationFrame = callback => setTimeout(() => callback(Date.now()), 0);
  dom.window.cancelAnimationFrame = clearTimeout;
  dom.window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  React = require('react'); ({ act } = require('react-dom/test-utils'));
  api = await harness.loadModule('packages/spfx-react-toolkit/src/controls/form/index.ts', { reactHookForm: true });
});
test.after(async () => { dom?.window.close(); await harness.cleanup(); });
function mount(element) {
  const ReactDOM = require('react-dom'); const { RendererProvider } = require('@griffel/react'); const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
  const renderer = require('@griffel/core').createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  const host = document.createElement('div'); document.body.appendChild(host);
  act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer }, React.createElement(FluentProvider, { theme: webLightTheme }, element)), host); });
  return { host, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
function change(input, value) { input.value = value; act(() => require('react-dom/test-utils').Simulate.change(input, { target: { value } })); }
async function blur(input) { await act(async () => { require('react-dom/test-utils').Simulate.blur(input); await Promise.resolve(); }); }
const initialDate = new Date(2026, 7, 30, 12);
const initialTime = new Date(1970, 0, 1, 9);
const selectedDate = new Date(2026, 7, 31, 12);
const formatDate = date => date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}` : '';
const parseDate = text => text === '2026-08-31' ? new Date(selectedDate) : text === '2026-09-02' ? new Date(2026, 8, 2, 12) : null;
const formatTime = date => `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
const configs = { DatePicker: { allowTextInput: true, formatDate, parseDateFromString: parseDate }, TimePicker: { freeform: true, hourCycle: 'h23', formatDateToTimeString: formatTime } };
for (const member of ['DatePicker', 'TimePicker']) for (const owner of ['resolver', 'rules']) test(`${member} visible rejected draft blocks Root and raw handleSubmit with old valid Date (${owner})`, async () => {
  const F = api.createForm(); assert.ok(F[member], `public ${member} exists`); let methods; let submitted = 0, invalid = 0, domainCalls = 0;
  const initial = member === 'DatePicker' ? initialDate : initialTime;
  function Owner() {
    methods = F.useForm({ defaultValues: { value: initial }, ...(owner === 'resolver' ? { resolver: values => { domainCalls++; return { values, errors: {} }; } } : {}) });
    return React.createElement(F.Root, { methods, onSubmit: () => { submitted++; }, onInvalid: () => { invalid++; } }, React.createElement(F[member], { name: 'value', label: 'Value', ...configs[member], ...(owner === 'rules' ? { rules: { validate: value => { domainCalls++; return value instanceof Date || 'Date needed'; } } } : {}) }));
  }
  const f = mount(React.createElement(Owner));
  try {
    const input = f.host.querySelector('input'); change(input, 'invalid draft');
    await act(async () => { await methods.handleSubmit(() => { submitted++; }, () => { invalid++; })(); });
    assert.equal(submitted, 0); assert.equal(invalid, 1); assert.ok(domainCalls > 0); assert.equal(input.value, 'invalid draft'); assert.deepEqual(methods.getValues('value'), initial);
    await blur(input); assert.equal(input.value, 'invalid draft'); assert.deepEqual(methods.getValues('value'), initial);
    await act(async () => { require('react-dom/test-utils').Simulate.submit(f.host.querySelector('form')); await new Promise(resolve => setImmediate(resolve)); });
    assert.equal(submitted, 0); assert.equal(invalid, 2); assert.equal(methods.getFieldState('value').error.message, member === 'DatePicker' ? 'Complete or correct the date before submitting.' : 'Complete or correct the time before submitting.');
  } finally { f.close(); }
});
for (const member of ['DatePicker', 'TimePicker']) test(`${member} local parse, native ref, callbacks after commit/touch, watch and nullable clear`, async () => {
  const F = api.createForm(); let methods; const ref = React.createRef(); const seen = []; const initial = member === 'DatePicker' ? initialDate : initialTime;
  function Owner() {
    methods = F.useForm({ defaultValues: { value: initial }, mode: 'onBlur' });
    return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F[member], { name: 'value', id: 'local-choice', label: 'Local value', hint: 'Help', ref, ...configs[member], onBlur: () => seen.push(['blur', methods.getFieldState('value').isTouched, methods.getValues('value')]), ...(member === 'DatePicker' ? { onSelectDate: value => seen.push(['change', methods.getValues('value'), value]) } : { onTimeChange: (event, data) => seen.push(['change', methods.getValues('value'), data.selectedTime]) }) }));
  }
  const f = mount(React.createElement(Owner));
  try {
    assert.equal(ref.current.tagName, 'INPUT'); assert.equal(ref.current.id, 'local-choice'); assert.equal(f.host.querySelector('label').htmlFor, 'local-choice');
    assert.match(document.getElementById(ref.current.getAttribute('aria-describedby')).textContent, /Help/);
    change(ref.current, member === 'DatePicker' ? '2026-08-31' : '14:30'); await blur(ref.current);
    const expected = member === 'DatePicker' ? selectedDate : new Date(1970, 0, 1, 14, 30);
    assert.deepEqual(methods.getValues('value'), expected); assert.deepEqual(methods.watch('value'), expected);
    assert.deepEqual(seen[0], ['change', expected, expected]); assert.deepEqual(seen[1], ['blur', true, expected]);
    change(ref.current, ''); await blur(ref.current); assert.equal(methods.getValues('value'), null); assert.equal(methods.watch('value'), null); assert.equal(ref.current.value, '');
  } finally { f.close(); }
});
for (const member of ['DatePicker', 'TimePicker']) test(`${member} rejects bounds visibly, retains old value, retries and clears domain errors`, async () => {
  const F = api.createForm(); let methods; const results = []; const initial = member === 'DatePicker' ? initialDate : initialTime;
  function Owner() { methods = F.useForm({ defaultValues: { value: initial }, mode: 'onBlur' }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F[member], { name: 'value', ...configs[member], inputValidationMessage: 'Correggi il valore', rules: { validate: value => value !== null || 'Required by domain' }, ...(member === 'DatePicker' ? { minDate: initialDate, maxDate: selectedDate, onValidationResult: data => results.push(data) } : { startHour: 9, endHour: 17, onTimeChange: (event, data) => results.push(data) }) })); }
  const f = mount(React.createElement(Owner));
  try {
    const input = f.host.querySelector('input'); const rejected = member === 'DatePicker' ? '2026-09-02' : '18:00'; change(input, rejected); await blur(input);
    assert.equal(input.value, rejected); assert.deepEqual(methods.getValues('value'), initial); assert.equal(results.at(-1).error, 'out-of-bounds'); assert.equal(methods.getFieldState('value').error.message, 'Correggi il valore');
    change(input, member === 'DatePicker' ? '2026-08-31' : '16:59'); await blur(input); assert.equal(methods.getFieldState('value').invalid, false); assert.equal(results.at(-1).error, undefined);
    change(input, ''); await blur(input); assert.equal(methods.getValues('value'), null); assert.equal(methods.getFieldState('value').error.message, 'Required by domain');
  } finally { f.close(); }
});
for (const member of ['DatePicker', 'TimePicker']) test(`${member} same-value reset/resetField and keepDirtyValues preserve only actual dirty drafts`, async () => {
  const F = api.createForm(); let methods; const initial = member === 'DatePicker' ? initialDate : initialTime;
  function Owner() { methods = F.useForm({ defaultValues: { value: initial } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F[member], { name: 'value', ...configs[member] })); }
  const f = mount(React.createElement(Owner));
  try {
    let input = f.host.querySelector('input'); change(input, 'bad draft'); await blur(input);
    act(() => { methods.reset({ value: initial }, { keepDirtyValues: true }); }); input = f.host.querySelector('input'); assert.equal(input.value, member === 'DatePicker' ? formatDate(initial) : formatTime(initial), 'uncommitted draft with unchanged live value is clean');
    change(input, member === 'DatePicker' ? '2026-08-31' : '14:30'); await blur(input); const live = methods.getValues('value'); assert.equal(methods.getFieldState('value').isDirty, true);
    change(input, 'dirty bad draft'); await blur(input); act(() => methods.reset({ value: initial }, { keepDirtyValues: true })); input = f.host.querySelector('input'); assert.equal(input.value, 'dirty bad draft'); assert.deepEqual(methods.getValues('value'), live);
    await act(async () => { await methods.handleSubmit(() => assert.fail('draft blocked'))(); }); assert.ok(methods.getFieldState('value').error);
    act(() => methods.resetField('value', { defaultValue: live })); input = f.host.querySelector('input'); assert.equal(input.value, member === 'DatePicker' ? formatDate(live) : formatTime(live));
    change(input, 'same reset draft'); act(() => methods.reset({ value: live })); input = f.host.querySelector('input'); assert.equal(input.value, member === 'DatePicker' ? formatDate(live) : formatTime(live));
    await act(async () => { await methods.handleSubmit(() => {})(); }); assert.equal(methods.getFieldState('value').invalid, false);
  } finally { f.close(); }
});
test('DateTime parts preserve current scoped sibling values, clear independently and focus native parts', async () => {
  const F = api.createForm(); let methods; const ref = React.createRef(); const seen = [];
  function Owner() { methods = F.useForm({ defaultValues: { record: { appointment: { date: initialDate, time: initialTime } } } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.Object, { name: 'record' }, ({ Fields }) => React.createElement(Fields.DateTime, { name: 'appointment', ref, label: 'Appointment', dateLabel: 'Date', timeLabel: 'Time', rules: { validate: () => 'Parent error' }, datePickerProps: { ...configs.DatePicker, onSelectDate: () => seen.push(methods.getValues('record.appointment')) }, timePickerProps: { ...configs.TimePicker, onTimeChange: () => seen.push(methods.getValues('record.appointment')) } }))); }
  const f = mount(React.createElement(Owner));
  try {
    assert.equal(ref.current.tagName, 'DIV'); assert.equal(ref.current.getAttribute('role'), 'group');
    let [date, time] = ref.current.querySelectorAll('input'); assert.equal(date.name, 'record.appointment.date'); assert.equal(time.name, 'record.appointment.time');
    const externalTime = new Date(1988, 3, 4, 21, 12, 13); act(() => methods.setValue('record.appointment.time', externalTime)); change(date, '2026-08-31'); await blur(date);
    assert.deepEqual(methods.getValues('record.appointment'), { date: selectedDate, time: externalTime }); assert.deepEqual(seen.at(-1), { date: selectedDate, time: externalTime });
    change(time, '14:30'); await blur(time); assert.deepEqual(methods.getValues('record.appointment.date'), selectedDate); assert.equal(methods.getValues('record.appointment.time').getFullYear(), 1988);
    change(date, ''); await blur(date); assert.equal(methods.getValues('record.appointment.date'), null); assert.equal(methods.getValues('record.appointment.time').getHours(), 14);
    await act(async () => { await methods.trigger('record.appointment', { shouldFocus: true }); }); assert.equal(document.activeElement, date);
    assert.equal(ref.current.getAttribute('aria-invalid'), 'true'); assert.match(f.host.textContent, /Parent error/);
  } finally { f.close(); }
});
for (const member of ['DatePicker', 'TimePicker']) test(`${member} array moves draft ownership, removes blockers, preserves dirty draft through actual reset remount`, async () => {
  const F = api.createForm(); let methods, array; const initial = member === 'DatePicker' ? initialDate : initialTime;
  function Owner() { methods = F.useForm({ defaultValues: { rows: [{ value: initial }, { value: initial }] }, resolver: values => ({ values, errors: {} }) }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.Array, { name: 'rows' }, value => { array = value; return value.items.map(({ Fields, key }, index) => React.createElement(Fields[member], { key, name: 'value', label: `Row ${index}`, ...configs[member] })); })); }
  const f = mount(React.createElement(Owner));
  try {
    const dirty = f.host.querySelector('input'); change(dirty, member === 'DatePicker' ? '2026-08-31' : '14:30'); await blur(dirty); const live = methods.getValues('rows.0.value'); change(dirty, 'owned draft'); await blur(dirty);
    act(() => array.move(0, 1)); assert.equal(f.host.querySelectorAll('input')[1], dirty);
    await act(async () => { await methods.handleSubmit(() => assert.fail('moved draft blocks'))(); }); assert.equal(methods.getFieldState('rows.1.value').error.message.includes('before submitting'), true); assert.equal(methods.getFieldState('rows.0.value').error, undefined);
    act(() => methods.reset({ rows: [{ value: initial }, { value: initial }] }, { keepDirtyValues: true }));
    const replaced = f.host.querySelectorAll('input')[1]; assert.notEqual(replaced, dirty); assert.equal(replaced.value, 'owned draft'); assert.deepEqual(methods.getValues('rows.1.value'), live);
    await act(async () => { await methods.handleSubmit(() => assert.fail('remounted draft blocks'))(); }); assert.ok(methods.getFieldState('rows.1.value').error);
    change(replaced, member === 'DatePicker' ? '2026-08-31' : '14:30'); await blur(replaced); assert.equal(replaced.value, member === 'DatePicker' ? '2026-08-31' : '14:30');
    change(replaced, 'removed draft'); await blur(replaced); act(() => array.remove(1));
    let submitted; await act(async () => { await methods.handleSubmit(values => { submitted = values; })(); }); assert.equal(submitted.rows.length, 1); assert.equal(f.host.querySelectorAll('input').length, 1);
  } finally { f.close(); }
});
for (const member of ['DatePicker', 'TimePicker']) test(`${member} targeted reset supersedes pending dirty array handoff without discarding sibling`, async () => {
  const F = api.createForm(); let methods; const initial = member === 'DatePicker' ? initialDate : initialTime;
  function Owner() { methods = F.useForm({ defaultValues: { rows: [{ value: initial }, { value: initial }] } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.Array, { name: 'rows' }, ({ items }) => items.map(({ Fields, key }) => React.createElement(Fields[member], { key, name: 'value', ...configs[member] })))); }
  const f = mount(React.createElement(Owner));
  try {
    const inputs = [...f.host.querySelectorAll('input')];
    for (const [index, input] of inputs.entries()) { change(input, member === 'DatePicker' ? '2026-08-31' : '14:30'); await blur(input); change(input, `dirty ${index}`); await blur(input); }
    const retained = methods.getValues('rows.1.value');
    act(() => { methods.reset({ rows: [{ value: initial }, { value: initial }] }, { keepDirtyValues: true }); methods.resetField('rows.0.value', { defaultValue: initial }); });
    const [target, sibling] = f.host.querySelectorAll('input'); assert.notEqual(target, inputs[0]); assert.equal(target.value, member === 'DatePicker' ? formatDate(initial) : formatTime(initial)); assert.equal(sibling.value, 'dirty 1'); assert.deepEqual(methods.getValues('rows.1.value'), retained);
  } finally { f.close(); }
});
test('TimePicker empty parsing uses deterministic local anchor and preserves explicit/current anchors', async () => {
  const F = api.createForm(); let methods;
  function Owner() { methods = F.useForm({ defaultValues: { empty: null, anchored: null, current: new Date(1988, 3, 4, 9) } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, ...['empty', 'anchored', 'current'].map(name => React.createElement(F.TimePicker, { key: name, name, ...configs.TimePicker, ...(name === 'anchored' ? { dateAnchor: new Date(2001, 5, 6) } : {}) }))); }
  const f = mount(React.createElement(Owner));
  try {
    for (const input of f.host.querySelectorAll('input')) { change(input, '14:30'); await blur(input); }
    for (const [name, expected] of [['empty', [1970, 0, 1]], ['anchored', [2001, 5, 6]], ['current', [1988, 3, 4]]]) { const value = methods.getValues(name); assert.deepEqual([value.getFullYear(), value.getMonth(), value.getDate()], expected); assert.equal(value.getHours(), 14); assert.equal(value.getMinutes(), 30); }
  } finally { f.close(); }
});
test('DatePicker uses copied readonly localization strings and formatting with the real calendar', () => {
  const F = api.createForm(); const strings = Object.freeze({ months: Object.freeze(Array.from({ length: 12 }, (_, index) => `Month ${index + 1}`)), shortMonths: Object.freeze(Array.from({ length: 12 }, (_, index) => `M${index + 1}`)), days: Object.freeze(['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']), shortDays: Object.freeze(['S', 'M', 'T', 'W', 'T', 'F', 'S']), goToToday: 'Vai oggi' }); const formatted = [];
  const f = mount(React.createElement(F.Root, { defaultValues: { date: initialDate }, onSubmit: () => {} }, React.createElement(F.DatePicker, { name: 'date', formatDate, open: true, inlinePopup: true, strings, dateTimeFormatter: { formatDay: date => String(date.getDate()), formatYear: date => String(date.getFullYear()), formatMonth: (date, localized) => { formatted.push(localized); return localized.months[date.getMonth()]; }, formatMonthYear: (date, localized) => `${localized.months[date.getMonth()]} ${date.getFullYear()}`, formatMonthDayYear: (date, localized) => `${date.getDate()} ${localized.months[date.getMonth()]} ${date.getFullYear()}` } })));
  try { assert.match(f.host.textContent, /Vai oggi/); assert.match(f.host.textContent, /Month 8/); assert.ok(formatted.length > 0); assert.equal(formatted[0].goToToday, 'Vai oggi'); assert.equal(strings.months[7], 'Month 8'); }
  finally { f.close(); }
});
test('public date/time types preserve nullable models, refs, callbacks and scoped DateTime parts', () => {
  const ts = require('typescript'); const path = require('node:path');
  const program = ts.createProgram([path.resolve(__dirname, 'fixtures/form-date-time-types.tsx')], { noEmit: true, strict: true, skipLibCheck: true, target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ESNext, moduleResolution: ts.ModuleResolutionKind.NodeJs, jsx: ts.JsxEmit.React, lib: ['lib.es2020.d.ts', 'lib.dom.d.ts'] });
  const diagnostics = ts.getPreEmitDiagnostics(program);
  assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, { getCanonicalFileName: file => file, getCurrentDirectory: () => process.cwd(), getNewLine: () => '\n' }));
});
for (const member of ['DatePicker', 'TimePicker']) test(`${member} disabled scope prevents commits and same-value reset retains focused native target`, async () => {
  const F = api.createForm(); let methods; let enable; const ref = React.createRef(); const initial = member === 'DatePicker' ? initialDate : initialTime; let changes = 0;
  function Owner() { const [disabled, setDisabled] = React.useState(true); enable = () => setDisabled(false); methods = F.useForm({ defaultValues: { value: initial } }); return React.createElement(F.Root, { methods, disabled, onSubmit: () => {} }, React.createElement(F[member], { name: 'value', disabled: false, ref, ...configs[member], ...(member === 'DatePicker' ? { onSelectDate: () => changes++ } : { onTimeChange: () => changes++ }) })); }
  const f = mount(React.createElement(Owner));
  try {
    change(ref.current, member === 'DatePicker' ? '2026-08-31' : '14:30'); await blur(ref.current); assert.deepEqual(methods.getValues('value'), initial); assert.equal(changes, 0);
    act(() => enable()); ref.current.focus(); act(() => methods.reset({ value: initial })); assert.equal(document.activeElement, ref.current); assert.equal(ref.current.tagName, 'INPUT');
  } finally { f.close(); }
});
test('DatePicker Enter commits visible local text before consumer callback and retains rejected text', async () => {
  const F = api.createForm(); let methods; const seen = [];
  function Owner() { methods = F.useForm({ defaultValues: { date: initialDate } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.DatePicker, { name: 'date', ...configs.DatePicker, open: false, onKeyDown: () => seen.push(methods.getValues('date')) })); }
  const f = mount(React.createElement(Owner));
  try { const input = f.host.querySelector('input'); change(input, '2026-08-31'); act(() => require('react-dom/test-utils').Simulate.keyDown(input, { key: 'Enter' })); assert.deepEqual(methods.getValues('date'), selectedDate); assert.deepEqual(seen[0], selectedDate); change(input, 'bad Enter draft'); act(() => require('react-dom/test-utils').Simulate.keyDown(input, { key: 'Enter' })); assert.equal(input.value, 'bad Enter draft'); assert.deepEqual(methods.getValues('date'), selectedDate); }
  finally { f.close(); }
});
test('DatePicker calendar dismissal cannot turn rejected text into the old committed Date', async () => {
  const F = api.createForm(); let methods; let changes = 0;
  function Owner() { methods = F.useForm({ defaultValues: { date: initialDate }, resolver: values => ({ values, errors: {} }) }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.DatePicker, { name: 'date', ...configs.DatePicker, open: true, inlinePopup: true, onSelectDate: () => changes++ })); }
  const f = mount(React.createElement(Owner));
  try {
    const input = f.host.querySelector('input'); change(input, 'popup invalid draft');
    act(() => require('react-dom/test-utils').Simulate.keyDown(input, { key: 'Escape' }));
    assert.equal(input.value, 'popup invalid draft'); assert.deepEqual(methods.getValues('date'), initialDate); assert.equal(changes, 0);
    await act(async () => { await methods.handleSubmit(() => assert.fail('popup draft still blocks'))(); }); assert.ok(methods.getFieldState('date').error);
  } finally { f.close(); }
});
for (const allowTextInput of [false, true]) test(`DatePicker real calendar selection commits with allowTextInput=${allowTextInput}`, async () => {
  const F = api.createForm(); let methods; const seen = [];
  function Owner() { methods = F.useForm({ defaultValues: { date: initialDate } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.DatePicker, { name: 'date', ...configs.DatePicker, allowTextInput, open: true, inlinePopup: true, onSelectDate: date => seen.push([methods.getValues('date'), date]) })); }
  const f = mount(React.createElement(Owner));
  try {
    const button = f.host.querySelector('button[aria-label="31, August, 2026"]'); assert.ok(button); act(() => require('react-dom/test-utils').Simulate.click(button));
    const expected = new Date(2026, 7, 31); assert.deepEqual(methods.getValues('date'), expected); assert.deepEqual(seen.at(-1), [expected, expected]); assert.equal(f.host.querySelector('input').value, '2026-08-31');
  } finally { f.close(); }
});
test('TimePicker real generated option selection uses local deterministic anchor and reports callback after RHF', async () => {
  const F = api.createForm(); let methods; const seen = [];
  function Owner() { methods = F.useForm({ defaultValues: { time: null } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.TimePicker, { name: 'time', formatDateToTimeString: formatTime, startHour: 9, endHour: 17, increment: 30, open: true, inlinePopup: true, onTimeChange: (event, data) => seen.push([methods.getValues('time'), data.selectedTime]) })); }
  const f = mount(React.createElement(Owner));
  try { const option = [...f.host.querySelectorAll('[role="option"]')].find(node => node.textContent === '14:30'); assert.ok(option); act(() => require('react-dom/test-utils').Simulate.click(option)); const expected = new Date(1970, 0, 1, 14, 30); assert.deepEqual(methods.getValues('time'), expected); assert.deepEqual(seen[0], [expected, expected]); }
  finally { f.close(); }
});
for (const timezone of ['America/Los_Angeles', 'UTC', 'Pacific/Auckland']) test(`DatePicker default local calendar parsing preserves date-only ISO and inclusive bounds in ${timezone}`, async () => {
  const previousTimezone = process.env.TZ;
  process.env.TZ = timezone;
  const offset = new Date(2026, 7, 31).getTimezoneOffset();
  assert.ok(timezone === 'UTC' ? offset === 0 : timezone === 'America/Los_Angeles' ? offset > 0 : offset < 0, 'process really uses the requested timezone offset');
  const F = api.createForm(); let methods;
  function Owner() { methods = F.useForm({ defaultValues: { date: null }, mode: 'onBlur' }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.DatePicker, { name: 'date', allowTextInput: true, open: false, minDate: new Date(2026, 7, 31, 12), maxDate: new Date(2026, 7, 31, 12) })); }
  const f = mount(React.createElement(Owner));
  try {
    const input = f.host.querySelector('input'); change(input, '2026-08-31'); await blur(input);
    const value = methods.getValues('date'); assert.ok(value instanceof Date, 'inclusive local boundary is accepted');
    assert.deepEqual([value.getFullYear(), value.getMonth(), value.getDate()], [2026, 7, 31], 'raw RHF local parts preserve the entered day'); assert.equal(value.getHours(), 0); assert.equal(input.value, '2026-08-31'); assert.equal(methods.getFieldState('date').invalid, false);
    change(input, '2026-09-01'); await blur(input); assert.equal(input.value, '2026-09-01'); assert.deepEqual(methods.getValues('date'), value); assert.equal(methods.getFieldState('date').invalid, true);
    change(input, ''); await blur(input); assert.equal(methods.getValues('date'), null);
    change(input, '2026-08-31'); act(() => require('react-dom/test-utils').Simulate.keyDown(input, { key: 'Enter' })); assert.deepEqual(methods.getValues('date'), value); assert.equal(input.value, '2026-08-31');
  } finally { f.close(); if (previousTimezone === undefined) delete process.env.TZ; else process.env.TZ = previousTimezone; }
});
for (const timezone of ['America/Los_Angeles', 'UTC', 'Pacific/Auckland']) test(`DatePicker default local calendar parsing rejects rollover/timestamps and retains leap/year components in ${timezone}`, async () => {
  const previousTimezone = process.env.TZ;
  process.env.TZ = timezone;
  const F = api.createForm(); let methods; const results = [];
  function Owner() { methods = F.useForm({ defaultValues: { date: null }, mode: 'onBlur' }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.DatePicker, { name: 'date', allowTextInput: true, open: false, onValidationResult: data => results.push(data) })); }
  const f = mount(React.createElement(Owner));
  try {
    const input = f.host.querySelector('input');
    for (const [text, expected] of [['2024-02-29', [2024, 1, 29]], ['2000-02-29', [2000, 1, 29]], ['0099-01-02', [99, 0, 2]], ['0000-02-29', [0, 1, 29]]]) {
      change(input, text); await blur(input); const value = methods.getValues('date'); assert.ok(value instanceof Date); assert.deepEqual([value.getFullYear(), value.getMonth(), value.getDate()], expected); assert.equal(input.value, text);
    }
    const previousValue = methods.getValues('date');
    for (const text of ['2023-02-29', '1900-02-29', '2026-04-31', '2026-00-31', '2026-13-01', '2026-01-00', '2026-01-32', '2026-8-31', '08/31/2026', '2026-08-31T00:00:00Z', '2026-08-31T00:00:00-07:00', '2026-08-31T12:30:00']) {
      change(input, text); await blur(input); assert.equal(input.value, text); assert.deepEqual(methods.getValues('date'), previousValue); assert.equal(results.at(-1).error, 'invalid-input');
      await act(async () => { await methods.handleSubmit(() => assert.fail('rejected default text blocks raw submission'))(); }); assert.equal(methods.getFieldState('date').invalid, true);
    }
  } finally { f.close(); if (previousTimezone === undefined) delete process.env.TZ; else process.env.TZ = previousTimezone; }
});
test('DatePicker explicit parser and formatter preserve consumer Date outputs without local normalization', async () => {
  const F = api.createForm(); let methods; const result = new Date('2026-08-31T03:12:13.456Z'); const seen = [];
  function Owner() { methods = F.useForm({ defaultValues: { date: null } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.DatePicker, { name: 'date', allowTextInput: true, open: false, parseDateFromString: text => text === 'consumer timestamp' ? result : null, formatDate: date => date ? `Custom ${date.getTime()}` : '', onSelectDate: date => seen.push(date) })); }
  const f = mount(React.createElement(Owner));
  try { const input = f.host.querySelector('input'); change(input, 'consumer timestamp'); await blur(input); assert.equal(methods.getValues('date').getTime(), result.getTime()); assert.equal(methods.getValues('date').getMilliseconds(), 456); assert.equal(seen[0], result); assert.equal(input.value, `Custom ${result.getTime()}`); }
  finally { f.close(); }
});
test('DatePicker default local calendar parsing stores August 31 rather than August 30 in a negative timezone', async () => {
  const previousTimezone = process.env.TZ; process.env.TZ = 'America/Los_Angeles'; const F = api.createForm(); let methods;
  function Owner() { methods = F.useForm({ defaultValues: { date: null } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.DatePicker, { name: 'date', allowTextInput: true, open: false })); }
  const f = mount(React.createElement(Owner));
  try { const input = f.host.querySelector('input'); change(input, '2026-08-31'); await blur(input); const value = methods.getValues('date'); assert.deepEqual([value.getFullYear(), value.getMonth(), value.getDate()], [2026, 7, 31]); }
  finally { f.close(); if (previousTimezone === undefined) delete process.env.TZ; else process.env.TZ = previousTimezone; }
});

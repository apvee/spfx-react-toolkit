const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, act, api, dom, fluent;
test.before(async () => {
  dom = harness.createSxDom();
  for (const key of ['Node', 'Element', 'HTMLElement', 'HTMLInputElement', 'File', 'FileList', 'Blob']) global[key] = dom.window[key];
  React = require('react'); ({ act } = require('react-dom/test-utils')); fluent = require('@fluentui/react-components');
  api = await harness.loadModule('packages/spfx-react-toolkit/src/controls/form/index.ts', { reactHookForm: true });
});
test.after(async () => { dom?.window.close(); await harness.cleanup(); });
function mount(element) {
  const ReactDOM = require('react-dom'); const { RendererProvider } = require('@griffel/react');
  const renderer = require('@griffel/core').createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  const host = document.createElement('div'); document.body.appendChild(host);
  act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer }, React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme }, element)), host); });
  return { host, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
function children(member) {
  if (member === 'ColorPicker') return React.createElement(fluent.ColorSlider, { channel: 'hue', 'aria-label': 'Hue' });
  if (member === 'SwatchPicker') return ['blue', 'red'].map(value => React.createElement(fluent.ColorSwatch, { key: value, value, color: value, 'aria-label': `${value} display` }));
}
for (const member of ['Rating', 'ColorPicker', 'SwatchPicker']) test(`${member} live domain, root ref, blur order and Field relationships`, async () => {
  const F = api.createForm(); let methods; const ref = React.createRef(); const seen = [];
  const initial = member === 'Rating' ? 2 : member === 'ColorPicker' ? { h: 20, s: .4, v: .8, a: .3 } : 'blue';
  const callback = () => seen.push(methods.getValues('value'));
  function Owner() {
    methods = F.useForm({ defaultValues: { value: initial } });
    return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F[member], {
      name: 'value', ref, id: 'choice', label: 'Choice', hint: 'Help', rules: { validate: () => 'Choose again' },
      onBlur: () => seen.push(methods.getFieldState('value').isTouched),
      ...(member === 'Rating' ? { onChange: callback } : member === 'ColorPicker' ? { onColorChange: callback } : { onSelectionChange: callback })
    }, ...(member === 'Rating' ? [] : [children(member)])));
  }
  const f = mount(React.createElement(Owner));
  try {
    assert.equal(ref.current.tagName, 'DIV'); assert.equal(ref.current.id, 'choice');
    const node = ref.current.querySelector(member === 'SwatchPicker' ? 'button' : 'input');
    assert.ok(node);
    if (member === 'Rating') act(() => require('react-dom/test-utils').Simulate.change(ref.current.querySelectorAll('input')[3]));
    else if (member === 'SwatchPicker') act(() => require('react-dom/test-utils').Simulate.click(ref.current.querySelectorAll('button')[1]));
    else act(() => require('react-dom/test-utils').Simulate.change(node, { target: { value: '90' } }));
    const value = methods.getValues('value');
    assert.equal(member === 'Rating' ? typeof value : member === 'SwatchPicker' ? value : value.h, member === 'Rating' ? 'number' : member === 'SwatchPicker' ? 'red' : 90);
    if (member === 'ColorPicker') { assert.deepEqual(value, { h: 90, s: .4, v: .8, a: .3 }); assert.notEqual(value, initial); assert.equal(initial.h, 20); }
    assert.deepEqual(seen[0], value);
    act(() => require('react-dom/test-utils').Simulate.blur(node)); assert.equal(seen.at(-1), true);
    await act(async () => { await methods.trigger('value', { shouldFocus: true }); });
    assert.match(f.host.textContent, /Choose again/); assert.ok(ref.current.contains(document.activeElement));
    const described = ref.current.getAttribute('aria-describedby'); assert.ok(described); assert.ok(described.split(' ').some(id => document.getElementById(id)?.textContent === 'Help'));
    assert.equal(ref.current.getAttribute('aria-invalid'), 'true'); assert.ok(ref.current.getAttribute('aria-labelledby'));
  } finally { f.close(); }
});
for (const member of ['Rating', 'ColorPicker', 'SwatchPicker']) test(`${member} disabled scope blocks mouse and keyboard writes`, () => {
  const F = api.createForm(); let methods;
  const initial = member === 'Rating' ? 2 : member === 'ColorPicker' ? { h: 20, s: .4, v: .8 } : 'blue';
  function Owner() { methods = F.useForm({ defaultValues: { value: initial } }); return React.createElement(F.Root, { methods, disabled: true, onSubmit: () => {} }, React.createElement(F[member], { name: 'value', disabled: false }, ...(member === 'Rating' ? [] : [children(member)]))); }
  const f = mount(React.createElement(Owner));
  try {
    assert.ok(f.host.querySelector('fieldset').disabled);
    const node = f.host.querySelector(member === 'SwatchPicker' ? 'button' : 'input');
    act(() => { const simulate = require('react-dom/test-utils').Simulate; simulate.click(node); simulate.keyDown(node, { key: 'ArrowRight' }); simulate.change(node, { target: { value: '90' } }); });
    assert.deepEqual(methods.getValues('value'), initial);
  } finally { f.close(); }
});

test('composite public types preserve two-way domains and root refs', () => {
  const ts = require('typescript'); const path = require('node:path');
  const program = ts.createProgram([path.resolve(__dirname, 'fixtures/form-composite-types.tsx')], { noEmit: true, strict: true, skipLibCheck: true, target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ESNext, moduleResolution: ts.ModuleResolutionKind.NodeJs, jsx: ts.JsxEmit.React, lib: ['lib.es2020.d.ts', 'lib.dom.d.ts'] });
  const diagnostics = ts.getPreEmitDiagnostics(program);
  assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, { getCanonicalFileName: file => file, getCurrentDirectory: () => process.cwd(), getNewLine: () => '\n' }));
});

for (const alpha of [undefined, .6]) test(`ColorPicker copies HSV before consumer mutation and preserves emitted alpha with ${alpha === undefined ? 'absent' : 'present'} initial alpha`, () => {
  const F = api.createForm(); let methods;
  const initial = { h: 30, s: .2, v: .9, ...(alpha === undefined ? {} : { a: alpha }) };
  function Owner() { methods = F.useForm({ defaultValues: { color: initial } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.ColorPicker, { name: 'color', onColorChange: (event, data) => { data.color.h = 270; } }, children('ColorPicker'))); }
  const f = mount(React.createElement(Owner));
  try {
    act(() => require('react-dom/test-utils').Simulate.change(f.host.querySelector('input'), { target: { value: '120' } }));
    assert.deepEqual(methods.getValues('color'), { ...initial, h: 120, a: alpha ?? 1 }); assert.equal(initial.h, 30);
    assert.equal(methods.getValues('color').a, alpha ?? 1);
  } finally { f.close(); }
});

test('SwatchPicker error focus skips disabled swatches and unrelated native buttons; root ref stays real', async () => {
  const F = api.createForm(); let methods; const ref = React.createRef();
  function Owner() { methods = F.useForm({ defaultValues: { value: 'red' } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.SwatchPicker, { name: 'value', ref, label: 'Palette', rules: { validate: () => 'Pick again' } }, React.createElement('button', { type: 'button' }, 'Unrelated'), React.createElement(fluent.ColorSwatch, { disabled: true, value: 'blue', color: 'blue' }), React.createElement(fluent.ColorSwatch, { value: 'red', color: 'red', 'aria-label': 'Eligible' }))); }
  const f = mount(React.createElement(Owner));
  try { await act(async () => { await methods.trigger('value', { shouldFocus: true }); }); assert.equal(ref.current.tagName, 'DIV'); assert.equal(document.activeElement.getAttribute('aria-label'), 'Eligible'); }
  finally { f.close(); }
});

test('ColorPicker error focus skips unrelated ranges and disabled primitive inputs', async () => {
  const F = api.createForm(); let methods; const ref = React.createRef();
  function Owner() { methods = F.useForm({ defaultValues: { value: { h: 30, s: .2, v: .9 } } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.ColorPicker, { name: 'value', ref, rules: { validate: () => 'Pick again' } }, React.createElement('input', { type: 'range', 'aria-label': 'Unrelated' }), React.createElement(fluent.ColorSlider, { channel: 'hue', disabled: true }), React.createElement(fluent.AlphaSlider, { 'aria-label': 'Alpha' }))); }
  const f = mount(React.createElement(Owner));
  try { await act(async () => { await methods.trigger('value', { shouldFocus: true }); }); assert.equal(ref.current.tagName, 'DIV'); assert.equal(document.activeElement.getAttribute('aria-label'), 'Alpha'); }
  finally { f.close(); }
});

test('disabled ColorArea capture blocks its composite pointer and keyboard handlers', () => {
  const F = api.createForm(); let methods; let called = 0;
  const initial = { h: 30, s: .2, v: .9 };
  function Owner() { methods = F.useForm({ defaultValues: { value: initial } }); return React.createElement(F.Root, { methods, disabled: true, onSubmit: () => {} }, React.createElement(F.ColorPicker, { name: 'value', onColorChange: () => { called++; } }, React.createElement(fluent.ColorArea, { 'data-area': '', onPointerDown: () => { called++; }, onKeyDown: () => { called++; } }))); }
  const f = mount(React.createElement(Owner));
  try { const area = f.host.querySelector('[data-area]'); act(() => { const simulate = require('react-dom/test-utils').Simulate; simulate.pointerDown(area, { pointerId: 1, clientX: 10, clientY: 10 }); simulate.keyDown(area, { key: 'ArrowUp' }); }); assert.equal(called, 0); assert.deepEqual(methods.getValues('value'), initial); assert.ok([...area.querySelectorAll('input')].every(input => input.matches(':disabled'))); }
  finally { f.close(); }
});

test('ColorPicker defaults to a semantic group associated with its visible label, hint and error', async () => {
  const F = api.createForm(); let methods; const ref = React.createRef();
  function Owner() {
    methods = F.useForm({ defaultValues: { accent: { h: 30, s: .2, v: .9 } } });
    return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.ColorPicker, {
      name: 'accent', ref, label: 'Accent color', hint: 'Choose the brand accent', rules: { validate: () => 'Choose another accent' }
    }, React.createElement(fluent.ColorSlider, { channel: 'hue', 'aria-label': 'Hue' })));
  }
  const f = mount(React.createElement(Owner));
  try {
    const group = ref.current;
    assert.equal(group.tagName, 'DIV');
    assert.equal(group.getAttribute('role'), 'group');
    const labelIds = group.getAttribute('aria-labelledby').split(/\s+/);
    assert.deepEqual(labelIds.map(id => document.getElementById(id)?.textContent), ['Accent color']);
    assert.ok(labelIds.every(id => f.host.contains(document.getElementById(id))));
    assert.equal(group.querySelector('input').getAttribute('aria-label'), 'Hue');
    await act(async () => { await methods.trigger('accent'); });
    const descriptions = group.getAttribute('aria-describedby').split(/\s+/).map(id => document.getElementById(id)?.textContent);
    assert.ok(descriptions.includes('Choose the brand accent'));
    assert.ok(descriptions.includes('Choose another accent'));
    assert.equal(group.getAttribute('aria-invalid'), 'true');
  } finally { f.close(); }
});

test('ColorPicker preserves an explicit named semantic consumer role', () => {
  const F = api.createForm(); const ref = React.createRef();
  const f = mount(React.createElement(F.Root, { defaultValues: { accent: { h: 30, s: .2, v: .9 } }, onSubmit: () => {} }, React.createElement(F.ColorPicker, {
    name: 'accent', ref, role: 'region', label: 'Accent editor'
  }, React.createElement(fluent.ColorSlider, { channel: 'hue', 'aria-label': 'Hue' }))));
  try { assert.equal(ref.current.getAttribute('role'), 'region'); assert.equal(document.getElementById(ref.current.getAttribute('aria-labelledby')).textContent, 'Accent editor'); }
  finally { f.close(); }
});

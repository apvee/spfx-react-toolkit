const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, act, api, dom;
test.before(async () => {
  dom = harness.createSxDom();
  for (const key of ['HTMLElement', 'HTMLInputElement', 'File', 'FileList', 'Blob']) global[key] = dom.window[key];
  React = require('react'); ({ act } = require('react-dom/test-utils'));
  api = await harness.loadModule('tests/fixtures/form-core-entry.tsx', { reactHookForm: true });
});
test.after(async () => { dom?.window.close(); await harness.cleanup(); });
function mount(element) {
  const ReactDOM = require('react-dom'); const { RendererProvider } = require('@griffel/react'); const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
  const renderer = require('@griffel/core').createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  const host = document.createElement('div'); document.body.appendChild(host);
  function render(next = element) { element = next; act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer }, React.createElement(FluentProvider, { theme: webLightTheme }, element)), host); }); }
  render(); return { host, render, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
async function submit(host) { await act(async () => { host.querySelector('form').dispatchEvent(new window.Event('submit', { bubbles: true, cancelable: true })); }); }
function input(host, value) { const node = host.querySelector('input'); act(() => { require('react-dom/test-utils').Simulate.change(node, { target: { value } }); }); }

test('reset(newRecord) updates baseline and resetField updates only its target', () => {
  const F = api.createForm(); let methods;
  function Owner() { methods = F.useForm({ defaultValues: { title: 'First', other: 'Other' } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.Input, { name: 'title' }), React.createElement(F.Input, { name: 'other' })); }
  const f = mount(React.createElement(Owner)); try { act(() => { methods.reset({ title: 'Record', other: 'Baseline' }); }); assert.deepEqual([...f.host.querySelectorAll('input')].map(node => node.value), ['Record', 'Baseline']); act(() => { methods.setValue('title', 'Edit'); methods.setValue('other', 'Keep'); methods.resetField('title'); }); assert.deepEqual(methods.getValues(), { title: 'Record', other: 'Keep' }); act(() => { methods.reset(); }); assert.deepEqual(methods.getValues(), { title: 'Record', other: 'Baseline' }); } finally { f.close(); }
});

test('refresh keepDirtyValues is explicit and retains dirty fields while refreshing clean values', () => {
  const F = api.createForm(); let methods;
  function Owner() { methods = F.useForm({ defaultValues: { title: 'First', other: 'Other' } }); const dirty = methods.formState.dirtyFields; return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.Input, { name: 'title' }), React.createElement(F.Input, { name: 'other' }), React.createElement('output', {}, Object.keys(dirty).join(','))); }
  const f = mount(React.createElement(Owner)); try { act(() => { methods.setValue('title', 'Dirty', { shouldDirty: true }); methods.reset({ title: 'Fresh', other: 'Updated' }, { keepDirtyValues: true }); }); assert.deepEqual(methods.getValues(), { title: 'Dirty', other: 'Updated' }); } finally { f.close(); }
});

test('native Reset calls consumer first, honors cancellation and isolates sibling forms', () => {
  const F = api.createForm(); let a, b; let cancel = true; let calls = 0;
  function Owner({ sibling }) { const methods = F.useForm({ defaultValues: { title: sibling ? 'B' : 'A' } }); if (sibling) b = methods; else a = methods; return React.createElement(F.Root, { methods, onSubmit: () => {}, onReset: event => { calls++; if (cancel) event.preventDefault(); } }, React.createElement(F.Input, { name: 'title' }), React.createElement(F.Reset, {}, 'Reset')); }
  const f = mount(React.createElement(React.Fragment, {}, React.createElement(Owner, {}), React.createElement(Owner, { sibling: true })));
  try { act(() => { a.setValue('title', 'Edited A'); b.setValue('title', 'Edited B'); f.host.querySelector('button').click(); }); assert.equal(calls, 1); assert.equal(a.getValues('title'), 'Edited A'); cancel = false; act(() => { f.host.querySelector('button').click(); }); assert.equal(a.getValues('title'), 'A'); assert.equal(b.getValues('title'), 'Edited B'); } finally { f.close(); }
});

test('useField reset generations target only their field and subscriptions clean up on unmount', () => {
  const F = api.createForm(); let methods; let left, right;
  function Fields() { left = F.useField({ name: 'title' }); right = F.useField({ name: 'other' }); return React.createElement('output', {}, `${left.resetVersion}:${right.resetVersion}`); }
  function Owner() { methods = F.useForm({ defaultValues: { title: '', other: '' } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(Fields)); }
  const f = mount(React.createElement(Owner)); try { act(() => { methods.resetField('title'); }); assert.equal(left.resetVersion, 1); assert.equal(right.resetVersion, 0); act(() => { methods.reset(); }); assert.equal(left.resetVersion, 2); assert.equal(right.resetVersion, 1); } finally { f.close(); } assert.doesNotThrow(() => methods.reset());
});

test('reactive values notify same-valued field drafts only after an actual RHF record reset', () => {
  const F = api.createForm(); let methods;
  function Owner({ values, resetOptions }) { methods = F.useForm({ values, resetOptions }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(api.ResetDraftControl)); }
  const initial = { title: 'Same', other: 'A' }; const f = mount(React.createElement(Owner, { values: initial }));
  try { input(f.host, 'Uncommitted'); f.render(React.createElement(Owner, { values: { ...initial } })); assert.equal(f.host.querySelector('input').value, 'Uncommitted'); f.render(React.createElement(Owner, { values: { title: 'Same', other: 'B' } })); assert.deepEqual(methods.getValues(), { title: 'Same', other: 'B' }); assert.equal(f.host.querySelector('input').value, ''); assert.match(f.host.querySelector('output').textContent, /reactive/); }
  finally { f.close(); }
});

test('async default completion notifies drafts after RHF initialization and respects resetOptions', async () => {
  for (const keepValues of [false, true]) {
    const F = api.createForm(); let methods; let release; const defaults = new Promise(resolve => { release = resolve; });
    function Owner() { methods = F.useForm({ defaultValues: () => defaults, resetOptions: { keepValues } }); const loading = methods.formState.isLoading; return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(api.ResetDraftControl), React.createElement('span', {}, String(loading))); }
    const f = mount(React.createElement(Owner));
    try { input(f.host, 'Uncommitted'); await act(async () => { release({ title: 'Loaded' }); await defaults; }); assert.equal(f.host.querySelector('span').textContent, 'false'); assert.equal(f.host.querySelector('input').value, keepValues ? 'Uncommitted' : ''); assert.match(f.host.querySelector('output').textContent, /async/); }
    finally { f.close(); }
  }
});

test('RHF native-register reset echo does not invoke consumer twice or recurse', () => {
  const F = api.createForm(); let methods; let resets = 0;
  function Owner() { methods = F.useForm({ defaultValues: { title: 'Initial' } }); return React.createElement(F.Root, { methods, onSubmit: () => {}, onReset: () => { resets++; } }, React.createElement('input', { ...methods.register('title') }), React.createElement(F.Reset, {}, 'Reset')); }
  const f = mount(React.createElement(Owner));
  try { act(() => { methods.setValue('title', 'Edit'); f.host.querySelector('button').click(); }); assert.equal(resets, 1); assert.equal(methods.getValues('title'), 'Initial'); act(() => methods.reset()); assert.equal(resets, 1); }
  finally { f.close(); }
});

test('structural move targets the stable current owner before mounted paths update', () => {
  const F = api.createForm(); let methods, left, right;
  function Field({ name, id }) { const binding = F.useField({ name }); if (id === 'a') left = binding; else right = binding; return React.createElement('input', { ref: binding.ref, value: binding.value, onChange: event => binding.onChange(event.target.value) }); }
  function Owner({ moved = false }) { methods = F.useForm({ defaultValues: { rows: [{ title: 'A' }, { title: 'B' }] } }); React.useLayoutEffect(() => api.commitFormResetArrayStructure(methods, 'rows')); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(Field, { key: 'a', id: 'a', name: moved ? 'rows.1.title' : 'rows.0.title' }), React.createElement(Field, { key: 'b', id: 'b', name: moved ? 'rows.0.title' : 'rows.1.title' })); }
  const f = mount(React.createElement(Owner)); let cleanup;
  try {
    cleanup = api.registerFormResetArrayPath(methods, 'rows');
    act(() => { api.synchronizeFormResetArrayOperation(methods, 'rows', { type: 'move', from: 0, to: 1 }); methods.setValue('rows', [{ title: 'B' }, { title: 'A' }]); methods.resetField('rows.1.title'); f.render(React.createElement(Owner, { moved: true })); api.commitFormResetArrayStructure(methods, 'rows'); });
    assert.equal(left.resetVersion, 1, 'moved A receives reset at its current path');
    assert.equal(right.resetVersion, 0, 'B retains its draft generation');
    act(() => methods.resetField('rows.0.title'));
    assert.equal(left.resetVersion, 1); assert.equal(right.resetVersion, 1);
  } finally { cleanup?.(); f.close(); }
});

test('removed owner stops receiving reset notifications before it unmounts', () => {
  const F = api.createForm(); let methods, removed, retained;
  function Field({ name, id }) { const binding = F.useField({ name }); if (id === 'a') removed = binding; else retained = binding; return React.createElement('input', { ref: binding.ref, value: binding.value, onChange: event => binding.onChange(event.target.value) }); }
  function Owner({ removedFirst = false }) { methods = F.useForm({ defaultValues: { rows: [{ title: 'A' }, { title: 'B' }] } }); React.useLayoutEffect(() => api.commitFormResetArrayStructure(methods, 'rows')); return React.createElement(F.Root, { methods, onSubmit: () => {} }, removedFirst ? undefined : React.createElement(Field, { key: 'a', id: 'a', name: 'rows.0.title' }), React.createElement(Field, { key: 'b', id: 'b', name: removedFirst ? 'rows.0.title' : 'rows.1.title' })); }
  const f = mount(React.createElement(Owner)); let cleanup;
  try {
    cleanup = api.registerFormResetArrayPath(methods, 'rows');
    act(() => { api.synchronizeFormResetArrayOperation(methods, 'rows', { type: 'remove', indices: new Set([0]) }); methods.resetField('rows.0.title'); });
    assert.equal(removed.resetVersion, 0, 'removed A is no longer a target'); assert.equal(retained.resetVersion, 0, 'B notification waits until its current path mounts');
    f.render(React.createElement(Owner, { removedFirst: true }));
    assert.equal(retained.resetVersion, 1, 'retained B now owns index zero');
  } finally { cleanup?.(); f.close(); }
  assert.doesNotThrow(() => methods.reset());
});

test('private draft validation follows moved paths and removal before old controls unmount', async () => {
  const F = api.createForm({ resolver: values => ({ values, errors: {} }) }); let methods;
  function Owner() { methods = F.useForm({ defaultValues: { rows: [{ title: 'A' }, { title: 'B' }], sibling: 'C' } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(api.DraftControl, { name: 'rows.0.title' }), React.createElement(api.DraftControl, { name: 'sibling' })); }
  const f = mount(React.createElement(Owner)); let cleanup;
  try {
    cleanup = api.registerFormResetArrayPath(methods, 'rows'); input(f.host, 'incomplete');
    await act(async () => { assert.equal(await methods.trigger(), false); });
    assert.equal(methods.getFieldState('rows.0.title').error.message, 'Draft incomplete');
    await act(async () => { api.synchronizeFormResetArrayOperation(methods, 'rows', { type: 'move', from: 0, to: 1 }); assert.equal(await methods.trigger(), false); });
    assert.equal(methods.getFieldState('rows.0.title').error, undefined); assert.equal(methods.getFieldState('rows.1.title').error.message, 'Draft incomplete');
    await act(async () => { api.synchronizeFormResetArrayOperation(methods, 'rows', { type: 'remove', indices: [1] }); assert.equal(await methods.trigger(), true); });
    assert.equal(methods.getFieldState('rows.1.title').error, undefined);
  } finally { cleanup?.(); f.close(); }
});

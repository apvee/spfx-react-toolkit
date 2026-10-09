const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, act, api, dom;
test.before(async () => {
  dom = harness.createSxDom();
  for (const key of ['HTMLElement', 'HTMLInputElement', 'File', 'FileList', 'Blob']) global[key] = dom.window[key];
  React = require('react'); ({ act } = require('react-dom/test-utils'));
  api = await harness.loadModule('packages/spfx-react-toolkit/src/controls/form/index.ts', { reactHookForm: true });
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

test('owned Root binds Input native refs, events, accessibility and validated values', async () => {
  const F = api.createForm(); const formRef = React.createRef(); const inputRef = React.createRef(); const saved = []; let changed = 0;
  const f = mount(React.createElement(F.Root, { ref: formRef, defaultValues: { title: 'Initial' }, onSubmit: values => saved.push(values), 'data-native': 'yes' },
    React.createElement(F.Input, { ref: inputRef, name: 'title', label: 'Title', hint: 'Helpful', required: true, onChange: () => changed++ }),
    React.createElement(F.Submit, {}, 'Save')));
  try {
    assert.equal(formRef.current.tagName, 'FORM'); assert.equal(formRef.current.noValidate, true);
    assert.equal(formRef.current.dataset.native, 'yes'); assert.equal(inputRef.current.tagName, 'INPUT');
    assert.equal(inputRef.current.value, 'Initial'); assert.equal(f.host.querySelector('label').htmlFor, inputRef.current.id);
    assert.ok(inputRef.current.getAttribute('aria-describedby'));
    input(f.host, 'Edited'); assert.equal(changed, 1); await submit(f.host); assert.deepEqual(saved, [{ title: 'Edited' }]);
  } finally { f.close(); }
});

test('provided Root uses same-suite methods and preserves native noValidate false', async () => {
  const F = api.createForm(); let methods; const saved = [];
  function Owner() { methods = F.useForm({ defaultValues: { title: 'Owned' }, mode: undefined }); return React.createElement(F.Root, { methods, noValidate: false, onSubmit: value => saved.push(value) }, React.createElement(F.Input, { name: 'title' })); }
  const f = mount(React.createElement(Owner));
  try { assert.equal(f.host.querySelector('form').noValidate, false); await submit(f.host); assert.deepEqual(saved, [{ title: 'Owned' }]); assert.equal(methods.formState.isSubmitting, false); }
  finally { f.close(); }
});

test('Root rejects cross-suite methods, initialization alongside methods and owned raw HTML', () => {
  const A = api.createForm(); const B = api.createForm();
  for (const mode of ['cross', 'init', 'raw']) {
    function Bad() { const methods = A.useForm({ defaultValues: { title: '' } }); return React.createElement(mode === 'cross' ? B.Root : A.Root, { methods, ...(mode === 'init' ? { defaultValues: {} } : {}), ...(mode === 'raw' ? { dangerouslySetInnerHTML: { __html: 'unsafe' } } : {}), onSubmit: () => {} }); }
    assert.throws(() => mount(React.createElement(Bad)), /suite|initialization|owned|HTML/i);
  }
});

test('one Root gate spans deferred validation and save; editable fields do not mutate save snapshot', async () => {
  const F = api.createForm(); let releaseValidation, releaseSave; let calls = 0; let submitted; let methods;
  const validation = new Promise(resolve => { releaseValidation = resolve; }); const saving = new Promise(resolve => { releaseSave = resolve; });
  function Owner() { methods = F.useForm({ defaultValues: { title: 'Before' }, resolver: async value => { await validation; return { values: value, errors: {} }; } }); const pending = methods.formState.isSubmitting; return React.createElement(F.Root, { methods, onSubmit: async value => { calls++; submitted = value; await saving; } }, React.createElement(F.Input, { name: 'title' }), React.createElement(F.Submit, {}, 'Save'), React.createElement(F.Reset, {}, 'Reset'), React.createElement('output', {}, String(pending))); }
  const f = mount(React.createElement(Owner));
  const fire = () => f.host.querySelector('form').dispatchEvent(new window.Event('submit', { bubbles: true, cancelable: true }));
  try {
    await act(async () => { fire(); fire(); }); assert.equal(f.host.querySelector('output').textContent, 'true');
    assert.ok([...f.host.querySelectorAll('button')].every(node => node.disabled)); assert.equal(f.host.querySelector('input').disabled, false);
    await act(async () => { releaseValidation(); }); assert.equal(calls, 1); input(f.host, 'During save'); await act(async () => { fire(); }); assert.equal(calls, 1); assert.equal(submitted.title, 'Before');
    await act(async () => { releaseSave(); }); assert.equal(methods.formState.isSubmitting, false);
  } finally { f.close(); }
});

test('disabled Root wins over control false while readOnly retains submitted data', async () => {
  const F = api.createForm(); const saved = [];
  const f = mount(React.createElement(F.Root, { defaultValues: { title: 'Read' }, onSubmit: value => saved.push(value) }, React.createElement(F.Input, { name: 'title', readOnly: true })));
  try { await submit(f.host); assert.deepEqual(saved, [{ title: 'Read' }]); f.render(React.createElement(F.Root, { disabled: true, defaultValues: { title: 'Read' }, onSubmit: value => saved.push(value) }, React.createElement(F.Input, { name: 'title', disabled: false }))); assert.equal(f.host.querySelector('input').disabled, true); await submit(f.host); assert.equal(saved[1].title, undefined); }
  finally { f.close(); }
});


test('direct methods.handleSubmit retains RHF semantics and bypasses Root DOM gate', async () => {
  const F = api.createForm(); let methods; let calls = 0; let release; const validation = new Promise(resolve => { release = resolve; });
  function Owner() { methods = F.useForm({ defaultValues: {}, resolver: async value => { await validation; return { values: value, errors: {} }; } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }); }
  const f = mount(React.createElement(Owner));
  try { const handler = methods.handleSubmit(() => { calls++; }); let first, second; await act(async () => { first = handler(); second = handler(); release(); await Promise.all([first, second]); }); assert.equal(calls, 2); }
  finally { f.close(); }
});

test('Input honors native id consistently with label/hint while generated ids stay stable and unique', () => {
  const F = api.createForm();
  const tree = id => React.createElement(F.Root, { defaultValues: { title: '', other: '', last: '' }, onSubmit: () => {} },
    React.createElement(F.Input, { name: 'title', id, label: 'Explicit', hint: 'Explicit hint' }),
    React.createElement(F.Input, { name: 'other', label: 'Other' }), React.createElement(F.Input, { name: 'last', label: 'Last' }));
  const f = mount(tree('customer-reference'));
  try {
    const inputs = [...f.host.querySelectorAll('input')]; const ids = inputs.map(node => node.id);
    assert.equal(ids[0], 'customer-reference'); assert.equal(new Set(ids).size, 3); assert.ok(ids.every(Boolean));
    assert.equal(f.host.querySelector('label').htmlFor, 'customer-reference');
    const hints = inputs[0].getAttribute('aria-describedby').split(/\s+/).map(id => document.getElementById(id));
    assert.ok(hints.some(node => node?.textContent === 'Explicit hint'));
    f.render(tree('updated-reference'));
    const updated = [...f.host.querySelectorAll('input')]; assert.equal(updated[0].id, 'updated-reference');
    assert.equal(f.host.querySelector('label').htmlFor, 'updated-reference'); assert.deepEqual(updated.slice(1).map(node => node.id), ids.slice(1));
    const updatedHints = updated[0].getAttribute('aria-describedby').split(/\s+/).map(id => document.getElementById(id));
    assert.ok(updatedHints.some(node => node?.textContent === 'Explicit hint'));
  } finally { f.close(); }
});

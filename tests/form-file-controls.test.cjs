const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/form-selection-harness.cjs');
let React, act, api, environment;
test.before(async () => { await harness.setup(); environment = harness; ({ React, act, api } = harness); });
test.after(async () => { await environment?.cleanup(); });
function files(names) {
 const list = document.createElement('input'); list.type = 'file';
 const utils = require('jsdom/lib/jsdom/living/generated/utils');
 for (const name of names) utils.implForWrapper(list.files).push(utils.implForWrapper(new File(['content'], name)));
 return list.files;
}
function choose(node, list) {
 let selection = list;
 Object.defineProperty(node, 'files', { configurable: true, get: () => selection });
 const nativeValue = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value');
 Object.defineProperty(node, 'value', { configurable: true, get: () => nativeValue.get.call(node), set: value => { nativeValue.set.call(node, value); if (value === '') selection = files([]); } });
 act(() => require('react-dom/test-utils').Simulate.change(node));
}
test('FileInput native selection and empty clear write before consumer events with input ref and accessible Field', async () => {
 const F = api.createForm(); const ref = React.createRef(); let methods; const seen = [];
 function Owner() { methods = F.useForm({ defaultValues: { attachments: null } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.FileInput, { name: 'attachments', ref, id: 'file-id', label: 'Attachments', hint: 'PDF files', rules: { validate: files => files !== null || 'Choose a file' }, accept: 'application/pdf', multiple: true, onChange: event => seen.push([methods.getValues('attachments'), event.currentTarget]), onBlur: () => seen.push(methods.getFieldState('attachments').isTouched) })); }
 const f = environment.mount(React.createElement(Owner)); try {
 const node = ref.current; assert.equal(node.type, 'file'); assert.equal(node.id, 'file-id'); assert.equal(node.accept, 'application/pdf'); assert.equal(node.multiple, true); assert.equal(f.host.querySelector('label').htmlFor, node.id); assert.ok(node.getAttribute('aria-describedby'));
 let opened = 0; node.click = () => opened++; act(() => require('react-dom/test-utils').Simulate.click([...f.host.querySelectorAll('button')].find(button => button.textContent === 'Browse'))); assert.equal(opened, 1);
 const selected = files(['report.pdf', 'photo.pdf']); choose(node, selected); assert.equal(methods.getValues('attachments'), selected); assert.deepEqual(seen[0], [selected, node]); assert.equal(f.host.querySelector('[aria-live="polite"]').textContent, 'report.pdf, photo.pdf');
 act(() => require('react-dom/test-utils').Simulate.blur(node)); assert.equal(seen[1], true);
 choose(node, files([])); assert.equal(methods.getValues('attachments'), null); assert.equal(seen[2][0], null);
 choose(node, selected); act(() => require('react-dom/test-utils').Simulate.click([...f.host.querySelectorAll('button')].find(button => button.textContent === 'Clear'))); assert.equal(methods.getValues('attachments'), null); assert.equal(node.value, ''); assert.equal(f.host.querySelector('[aria-live="polite"]').textContent, ''); await act(async () => { await methods.trigger('attachments'); }); assert.equal(node.getAttribute('aria-invalid'), 'true'); assert.ok(node.getAttribute('aria-describedby').split(' ').some(id => document.getElementById(id)?.textContent === 'Choose a file'));
 } finally { f.close(); }
});
test('FileInput disabled ancestry guards Browse, Clear and native events', () => {
 const F = api.createForm(); let methods; let calls = 0;
 function Owner({ disabled }) { methods = F.useForm({ defaultValues: { attachments: null } }); return React.createElement(F.Root, { methods, disabled, onSubmit: () => {} }, React.createElement(F.FileInput, { name: 'attachments', disabled: false, onChange: () => calls++ })); }
 const f = environment.mount(React.createElement(Owner, { disabled: false })); try {
 const node = f.host.querySelector('input'); const selected = files(['a.txt']); choose(node, selected); f.render(React.createElement(Owner, { disabled: true })); assert.equal(node.disabled, true); assert.ok([...f.host.querySelectorAll('button')].every(button => button.disabled)); choose(node, files(['b.txt'])); assert.equal(methods.getValues('attachments'), selected); assert.equal(calls, 1);
 } finally { f.close(); }
});
test('FileInput types enforce full nullable domain, native callbacks and actual ref', () => {
 const ts = require('typescript'); const path = require('node:path'); const program = ts.createProgram([path.resolve(__dirname, 'fixtures/form-file-types.tsx')], { noEmit: true, strict: true, skipLibCheck: true, target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ESNext, moduleResolution: ts.ModuleResolutionKind.NodeJs, jsx: ts.JsxEmit.React, lib: ['lib.es2020.d.ts', 'lib.dom.d.ts'] });
 const diagnostics = ts.getPreEmitDiagnostics(program); assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, { getCanonicalFileName: f => f, getCurrentDirectory: () => process.cwd(), getNewLine: () => '\n' }));
});

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

async function submit(host) { await act(async () => { host.querySelector('form').dispatchEvent(new window.Event('submit', { bubbles: true, cancelable: true })); }); }
function owner(defaultValues, extraProps = {}) {
 const F = api.createForm(); let methods; let submitted;
 function Owner(props) { methods = F.useForm({ defaultValues, ...props }); const dirty = methods.formState.dirtyFields; return React.createElement(F.Root, { methods, onSubmit: value => { submitted = value; } }, React.createElement(F.FileInput, { name: 'attachments' }), React.createElement('output', {}, Object.keys(dirty).join(','))); }
 const f = environment.mount(React.createElement(Owner, extraProps)); return { ...f, F, Owner, get methods() { return methods; }, get submitted() { return submitted; } };
}
test('FileInput normalizes default/external FileLists only for package-owned file fields', async () => {
 const external = files(['existing.txt']); const f = owner({ attachments: external, custom: external }); try {
 assert.equal(f.methods.getValues('attachments'), null); assert.equal(f.methods.getValues('custom'), external); assert.equal(f.host.querySelector('input').files.length, 0);
 act(() => f.methods.reset({ attachments: external, custom: external })); assert.equal(f.methods.getValues('attachments'), null); assert.equal(f.methods.getValues('custom'), external);
 await submit(f.host); assert.equal(f.submitted.attachments, null); assert.equal(f.submitted.custom, external);
 } finally { f.close(); }
});
for (const targeted of [false, true]) test(`FileInput ${targeted ? 'resetField' : 'reset'} clears model, native boundary and submission`, async () => {
 const f = owner({ attachments: null }); try {
 const node = f.host.querySelector('input'); const selected = files(['selected.txt']); choose(node, selected); assert.equal(f.methods.getValues('attachments'), selected);
 // Observe the legal native boundary: product only assigns an empty string, never files.
 let clears = 0; Object.defineProperty(node, 'value', { configurable: true, get: () => '', set: value => { assert.equal(value, ''); clears++; Object.defineProperty(node, 'files', { configurable: true, get: () => files([]) }); } });
 act(() => { if (targeted) f.methods.resetField('attachments', { defaultValue: files(['external.txt']) }); else f.methods.reset({ attachments: files(['external.txt']) }); });
 assert.ok(clears > 0); assert.equal(node.files.length, 0); assert.equal(f.methods.getValues('attachments'), null); assert.equal(f.host.querySelector('[aria-live]').textContent, ''); await submit(f.host); assert.equal(f.submitted.attachments, null);
 } finally { f.close(); }
});
test('FileInput explicit keepDirtyValues preserves a surviving actual native selection', async () => {
 const f = owner({ attachments: null }); try {
 const node = f.host.querySelector('input'); const selected = files(['dirty.txt']); choose(node, selected); act(() => f.methods.reset({ attachments: files(['external.txt']) }, { keepDirtyValues: true }));
 assert.equal(f.host.querySelector('input'), node); assert.equal(node.files, selected); assert.equal(f.methods.getValues('attachments'), selected); assert.equal(f.host.querySelector('[aria-live]').textContent, 'dirty.txt'); await submit(f.host); assert.equal(f.submitted.attachments, selected);
 } finally { f.close(); }
});
test('FileInput reactive values cannot reconstruct native selection', () => {
 const external = files(['external.txt']); const f = owner(undefined, { values: { attachments: external, custom: external } }); try {
 assert.equal(f.methods.getValues('attachments'), null, 'initial reactive model'); f.render(React.createElement(f.Owner, { values: { attachments: external, custom: external, other: 'refresh' } })); assert.equal(f.methods.getValues('attachments'), null, 'refreshed reactive model'); assert.equal(f.methods.getValues('custom'), external);
 } finally { f.close(); }
});
test('FileInput async defaults cannot reconstruct native selection', async () => {
 const external = files(['external.txt']);
 let resolve; const completion = new Promise(done => { resolve = done; }); const a = owner(() => completion); try { await act(async () => { resolve({ attachments: external, custom: external }); await completion; }); assert.equal(a.methods.getValues('attachments'), null, 'async fulfilled model'); assert.equal(a.methods.getValues('custom'), external); } finally { a.close(); }
});
function arrayOwner() {
 const F = api.createForm(); let methods; let array;
 function Owner() { methods = F.useForm({ defaultValues: { rows: [{ attachments: null }, { attachments: null }] } }); const dirty = methods.formState.dirtyFields; return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.Array, { name: 'rows' }, current => { array = current; return current.items.map(({ key, Fields }) => React.createElement(Fields.FileInput, { key, name: 'attachments', 'aria-label': key })); }), React.createElement('output', {}, Object.keys(dirty).join(','))); }
 const f = environment.mount(React.createElement(Owner)); return { ...f, get methods() { return methods; }, get array() { return array; } };
}
test('FileInput move then resetField in one handler clears only the current owner', () => {
 const f = arrayOwner(); try {
 const [first, second] = [...f.host.querySelectorAll('input')]; const selected = files(['first.txt']); const sibling = files(['second.txt']); choose(first, selected); choose(second, sibling);
 act(() => { f.array.move(0, 1); f.methods.resetField('rows.1.attachments'); });
 assert.equal(f.methods.getValues('rows.1.attachments'), null); assert.equal(f.methods.getValues('rows.0.attachments'), sibling); assert.equal(f.host.querySelectorAll('input')[0], second);
 } finally { f.close(); }
});
for (const operation of ['move', 'remove']) test(`FileInput ${operation} then reset in one handler cannot replay lost native selections`, () => {
 const f = arrayOwner(); try {
 for (const input of f.host.querySelectorAll('input')) choose(input, files(['selected.txt']));
 act(() => { if (operation === 'move') f.array.move(0, 1); else f.array.remove(0); f.methods.reset({ rows: [{ attachments: files(['external.txt']) }] }); });
 assert.deepEqual(f.methods.getValues('rows'), [{ attachments: null }]); assert.equal(f.host.querySelectorAll('input').length, 1); assert.equal(f.host.querySelector('input').files.length, 0); assert.equal(f.host.querySelector('[aria-live]').textContent, '');
 } finally { f.close(); }
});
test('FileInput keepDirtyValues clears selections lost during an actual array remount', () => {
 const f = arrayOwner(); try {
 const old = f.host.querySelector('input'); choose(old, files(['selected.txt']));
 act(() => f.methods.reset({ rows: [{ attachments: null }, { attachments: null }] }, { keepDirtyValues: true }));
 assert.notEqual(f.host.querySelector('input'), old); assert.equal(f.host.querySelector('input').files.length, 0); assert.equal(f.methods.getValues('rows.0.attachments'), null);
 } finally { f.close(); }
});
test('FileInput independent forms and unowned custom values remain isolated', () => {
 const first = owner({ attachments: null }); const other = owner({ attachments: null }); try {
 const node = other.host.querySelector('input'); const selected = files(['other.txt']); choose(node, selected); act(() => first.methods.reset());
 assert.equal(other.methods.getValues('attachments'), selected); assert.equal(node.files, selected); assert.equal(other.host.querySelector('[aria-live]').textContent, 'other.txt');
 } finally { first.close(); other.close(); }
});
test('FileInput cannot display externally assigned files as a native selection', () => {
 const f = owner({ attachments: null }); try {
 const node = f.host.querySelector('input'); choose(node, files(['native.txt'])); act(() => f.methods.setValue('attachments', files(['external.txt'])));
 assert.equal(f.methods.getValues('attachments'), null); assert.equal(f.host.querySelector('[aria-live]').textContent, ''); assert.equal(node.value, '');
 } finally { f.close(); }
});
test('FileInput committed reactive keepDirtyValues preserves only a surviving current native selection', () => {
 const initial = { attachments: null, revision: 1 }; const f = owner(undefined, { values: initial, resetOptions: { keepDirtyValues: true } }); try {
 const node = f.host.querySelector('input'); const selected = files(['dirty.txt']); choose(node, selected);
 f.render(React.createElement(f.Owner, { values: { attachments: files(['external.txt']), revision: 2 }, resetOptions: { keepDirtyValues: true } }));
 assert.equal(f.host.querySelector('input'), node); assert.equal(node.files, selected); assert.equal(f.methods.getValues('attachments'), selected); assert.equal(f.methods.getValues('revision'), 2);
 } finally { f.close(); }
});

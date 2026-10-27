const test = require('node:test');
const assert = require('node:assert/strict');
const h = require('./fixtures/form-selection-harness.cjs');
test.before(h.setup); test.after(h.cleanup);
const options = [{ value: 'a', label: 'Alpha' }, { value: 'b', label: 'Beta' }];
function baseline(member, revision = 'first') { return { revision, rows: member === 'Combobox' ? [{ value: 'a' }, { value: '' }] : [{ value: ['a'] }, { value: ['a'] }] }; }
for (const member of ['Combobox', 'ComboboxMulti', 'TagPicker']) for (const source of ['reactive', 'async', 'async-batched']) test(`${member} ${source} array refresh preserves dirty query and clears clean query with committed options`, async () => {
  const F = h.suite(member); let methods, array, release;
  const initial = baseline(member), fresh = baseline(member, 'fresh');
  const loaded = new Promise(resolve => { release = resolve; });
  const loader = () => loaded;
  function Owner({ values = initial, keepDirtyValues = true }) {
    methods = F.useForm(source === 'reactive' ? { values, resetOptions: { keepDirtyValues } } : { defaultValues: loader, resetOptions: { keepDirtyValues } });
    return h.React.createElement(F.Root, { methods, onSubmit: () => {} }, h.React.createElement(F.Array, { name: 'rows' }, context => { array = context; return context.items.map(({ key, Fields }) => h.React.createElement(Fields[member], { key, name: 'value', options, 'aria-label': key })); }));
  }
  const f = h.mount(h.React.createElement(Owner));
  const nodes = () => array.items.map(({ key }) => f.host.querySelector(`input[aria-label="${key}"]`));
  try {
    if (source !== 'reactive') h.act(() => { methods.setValue('rows', initial.rows); methods.setValue('revision', initial.revision); });
    if (member !== 'Combobox') h.act(() => methods.setValue('rows.0.value', ['a', 'b'], { shouldDirty: true }));
    const original = nodes(); h.change(original[0], 'dirty query'); h.change(original[1], 'clean query');
    assert.equal(methods.getFieldState('rows.0.value').isDirty, true); assert.equal(methods.getFieldState('rows.1.value').isDirty, false);
    if (source === 'reactive') f.render(h.React.createElement(Owner, { values: fresh }));
    else if (source === 'async-batched') await h.act(async () => { release(fresh); await loaded; });
    else { release(fresh); await loaded; await new Promise(resolve => setImmediate(resolve)); }
    assert.equal(methods.getValues('revision'), 'fresh'); if (source === 'reactive') assert.notEqual(nodes()[0], original[0]);
    assert.equal(nodes()[0].value, 'dirty query'); if (member === 'ComboboxMulti') h.assertMultiQueryCleared(nodes()[1], 'Alpha'); else assert.equal(nodes()[1].value, '');
    assert.deepEqual(methods.getValues('rows.0.value'), member === 'Combobox' ? '' : ['a', 'b']);
    h.act(() => methods.reset(fresh, { keepDirtyValues: false }));
    if (member === 'ComboboxMulti') h.assertMultiQueryCleared(nodes()[0], 'Alpha'); else assert.equal(nodes()[0].value, member === 'Combobox' ? 'Alpha' : '');
  } finally { f.close(); }
});

test('equal reactive records leave current query and no stale handoff for a replacement row', () => {
  const F = h.suite('ComboboxMulti'); let methods, array;
  const initial = baseline('ComboboxMulti');
  function Owner({ values = initial }) { methods = F.useForm({ values, resetOptions: { keepDirtyValues: true } }); return h.React.createElement(F.Root, { methods, onSubmit: () => {} }, h.React.createElement(F.Array, { name: 'rows' }, context => { array = context; return context.items.map(({ key, Fields }) => h.React.createElement(Fields.ComboboxMulti, { key, name: 'value', options, 'aria-label': key })); })); }
  const f = h.mount(h.React.createElement(Owner)); const node = () => f.host.querySelector(`input[aria-label="${array.items[0].key}"]`);
  try {
    h.act(() => methods.setValue('rows.0.value', ['a', 'b'], { shouldDirty: true })); h.change(node(), 'old query'); const original = node();
    f.render(h.React.createElement(Owner, { values: { ...initial, rows: initial.rows.map(row => ({ ...row })) } }));
    assert.equal(node(), original); assert.equal(node().value, 'old query');
    h.act(() => array.update(0, { value: ['a', 'b'] })); assert.notEqual(node(), original); h.assertMultiQueryCleared(node(), methods.getValues(node().name).map(v => options.find(o => o.value === v).label).join(', '));
  } finally { f.close(); }
});

test('async completion reads committed resetOptions after a suspended render attempts to replace them', async () => {
  const F = h.suite('ComboboxMulti'); let methods, array, release;
  const loaded = new Promise(resolve => { release = resolve; }); const suspended = new Promise(() => {}); const loader = () => loaded;
  function Owner({ blocked = false, keepDirtyValues = true }) {
    methods = F.useForm({ defaultValues: loader, resetOptions: { keepDirtyValues } });
    if (blocked) throw suspended;
    return h.React.createElement(F.Root, { methods, onSubmit: () => {} }, h.React.createElement(F.Array, { name: 'rows' }, context => { array = context; return context.items.map(({ key, Fields }) => h.React.createElement(Fields.ComboboxMulti, { key, name: 'value', options, 'aria-label': key })); }));
  }
  const view = props => h.React.createElement(h.React.Suspense, { fallback: h.React.createElement('span', {}, 'Waiting') }, h.React.createElement(Owner, props));
  const f = h.mount(view({}));
  try {
    h.act(() => { methods.setValue('rows', [{ value: ['a'] }]); methods.setValue('rows.0.value', ['a', 'b'], { shouldDirty: true }); });
    h.change(f.host.querySelector(`input[aria-label="${array.items[0].key}"]`), 'committed query');
    f.render(view({ blocked: true, keepDirtyValues: false }));
    release({ rows: [{ value: ['a'] }] }); await loaded; await new Promise(resolve => setImmediate(resolve));
    assert.deepEqual(methods.getValues('rows.0.value'), ['a', 'b'], 'uncommitted options must not replace committed dirty preservation');
    f.render(view({})); assert.equal(f.host.querySelector('input').value, 'committed query');
  } finally { f.close(); }
});

test('committed reactive resetOptions changes and per-call overrides replace the previous preservation policy', () => {
  const F = h.suite('ComboboxMulti'); let methods, array;
  const initial = baseline('ComboboxMulti');
  function Owner({ values = initial, keepDirtyValues = true }) { methods = F.useForm({ values, resetOptions: { keepDirtyValues } }); return h.React.createElement(F.Root, { methods, onSubmit: () => {} }, h.React.createElement(F.Array, { name: 'rows' }, context => { array = context; return context.items.map(({ key, Fields }) => h.React.createElement(Fields.ComboboxMulti, { key, name: 'value', options, 'aria-label': key })); })); }
  const f = h.mount(h.React.createElement(Owner)); const node = () => f.host.querySelector(`input[aria-label="${array.items[0].key}"]`);
  try {
    h.act(() => methods.setValue('rows.0.value', ['a', 'b'], { shouldDirty: true })); h.change(node(), 'first query');
    f.render(h.React.createElement(Owner, { keepDirtyValues: false })); assert.equal(node().value, 'first query');
    const fresh = baseline('ComboboxMulti', 'fresh'); f.render(h.React.createElement(Owner, { values: fresh, keepDirtyValues: false }));
    assert.deepEqual(methods.getValues('rows.0.value'), ['a']); h.assertMultiQueryCleared(node(), methods.getValues(node().name).map(v => options.find(o => o.value === v).label).join(', '));
    f.render(h.React.createElement(Owner, { values: fresh, keepDirtyValues: true }));
    h.act(() => methods.setValue('rows.0.value', ['a', 'b'], { shouldDirty: true })); h.change(node(), 'second query');
    f.render(h.React.createElement(Owner, { values: baseline('ComboboxMulti', 'last'), keepDirtyValues: true })); assert.equal(node().value, 'second query');
    h.act(() => methods.reset(fresh, { keepDirtyValues: false })); h.assertMultiQueryCleared(node(), methods.getValues(node().name).map(v => options.find(o => o.value === v).label).join(', ')); assert.deepEqual(methods.getValues('rows.0.value'), ['a']);
  } finally { f.close(); }
});

function nativeFiles(name) {
  const node = document.createElement('input'); node.type = 'file';
  const utils = require('jsdom/lib/jsdom/living/generated/utils');
  utils.implForWrapper(node.files).push(utils.implForWrapper(new File(['native'], name)));
  return node.files;
}
const nonResetWrites = {
  'setValues unrelated': (methods, name, value) => methods.setValues({ other: 'second' }),
  'setValues same-name': (methods, name, value) => methods.setValues({ [name]: value, other: 'second' }),
  'setValue unrelated': (methods, name, value) => methods.setValue('other', 'second'),
  'setValue same-name': (methods, name, value) => methods.setValue(name, value),
  'resetDefaultValues': (methods, name, value) => methods.resetDefaultValues({ [name]: value, other: 'second' }),
  'unregister unrelated': methods => methods.unregister('other'),
};
for (const member of ['Combobox', 'ComboboxMulti', 'TagPicker']) for (const [writeName, write] of Object.entries(nonResetWrites)) test(`${member} equal reactive values plus ${writeName} does not acknowledge a reset`, () => {
  const F = h.suite(member); let methods;
  const initial = { choice: member === 'Combobox' ? 'a' : ['a'], other: 'first' };
  function Owner({ values = initial, update = false }) {
    methods = F.useForm({ values });
    h.React.useLayoutEffect(() => { if (update) write(methods, 'choice', methods.getValues('choice')); }, [update]);
    return h.React.createElement(F.Root, { methods, onSubmit: () => {} }, h.React.createElement(F[member], { name: 'choice', options, 'aria-label': 'choice' }));
  }
  const f = h.mount(h.React.createElement(Owner));
  try {
    const node = f.host.querySelector('input'); h.change(node, 'query before non-reset update');
    const live = methods.getValues('choice');
    f.render(h.React.createElement(Owner, { values: { ...initial }, update: true }));
    assert.equal(f.host.querySelector('input'), node, 'RHF skips resetting equal reactive values');
    assert.equal(node.value, 'query before non-reset update'); assert.deepEqual(methods.getValues('choice'), live);
    h.act(() => methods.reset(initial)); assert.equal(node.value, member === 'Combobox' ? 'Alpha' : '', 'a subsequent real reset still clears the draft');
  } finally { f.close(); }
});
for (const [writeName, write] of Object.entries(nonResetWrites)) test(`native selection survives equal reactive values plus ${writeName}`, () => {
  const F = h.suite('FileInput'); let methods;
  const initial = { attachment: null, other: 'first' };
  function Owner({ values = initial, update = false }) {
    methods = F.useForm({ values });
    h.React.useLayoutEffect(() => { if (update) write(methods, 'attachment', methods.getValues('attachment')); }, [update]);
    return h.React.createElement(F.Root, { methods, onSubmit: () => {} }, h.React.createElement(F.FileInput, { name: 'attachment' }));
  }
  const f = h.mount(h.React.createElement(Owner));
  try {
    const node = f.host.querySelector('input[type="file"]'), chosen = nativeFiles('selected.txt'); let selected = chosen;
    Object.defineProperty(node, 'files', { configurable: true, get: () => selected });
    Object.defineProperty(node, 'value', { configurable: true, get: () => '', set: value => { assert.equal(value, ''); const empty = document.createElement('input'); empty.type = 'file'; selected = empty.files; } });
    h.act(() => require('react-dom/test-utils').Simulate.change(node)); assert.equal(methods.getValues('attachment'), chosen);
    f.render(h.React.createElement(Owner, { values: { ...initial }, update: true }));
    assert.equal(f.host.querySelector('input[type="file"]'), node); assert.equal(node.files, chosen); assert.equal(methods.getValues('attachment'), chosen);
    h.act(() => methods.reset(initial)); assert.equal(node.files.length, 0); assert.equal(methods.getValues('attachment'), null);
  } finally { f.close(); }
});
for (const operation of ['append', 'prepend', 'insert', 'remove', 'update', 'replace', 'move', 'swap']) test(`owned native files ${operation} followed by reset keep DOM/model truth without restoring external files`, () => {
  const F = h.suite('FileInput'); let methods, array;
  const external = nativeFiles('external.txt');
  function Owner() { methods = F.useForm({ defaultValues: { rows: [{ attachments: null }, { attachments: null }], custom: external } }); return h.React.createElement(F.Root, { methods, onSubmit: () => {} }, h.React.createElement(F.Array, { name: 'rows' }, context => { array = context; return context.items.map(({ key, Fields }) => h.React.createElement(Fields.FileInput, { key, name: 'attachments', 'aria-label': key })); })); }
  const f = h.mount(h.React.createElement(Owner));
  try {
    for (const [index, item] of array.items.entries()) {
      const node = f.host.querySelector(`input[aria-label="${item.key}"]`); let selected = nativeFiles(`selected${index}.txt`);
      Object.defineProperty(node, 'files', { configurable: true, get: () => selected });
      Object.defineProperty(node, 'value', { configurable: true, get: () => '', set: value => { assert.equal(value, ''); selected = document.createElement('input'); selected.type = 'file'; selected = selected.files; } });
      h.act(() => require('react-dom/test-utils').Simulate.change(node));
    }
    h.act(() => {
      switch (operation) {
        case 'append': array.append({ attachments: null }); break;
        case 'prepend': array.prepend({ attachments: null }); break;
        case 'insert': array.insert(1, { attachments: null }); break;
        case 'remove': array.remove(0); break;
        case 'update': array.update(0, { attachments: null }); break;
        case 'replace': array.replace([{ attachments: null }]); break;
        case 'move': array.move(0, 1); break;
        case 'swap': array.swap(0, 1); break;
      }
      methods.reset({ rows: [{ attachments: external }], custom: external });
    });
    assert.deepEqual(methods.getValues('rows'), [{ attachments: null }]); assert.equal(array.items.length, 1);
    assert.equal(f.host.querySelector('input').files.length, 0); assert.equal(methods.getValues('custom'), external);
  } finally { f.close(); }
});

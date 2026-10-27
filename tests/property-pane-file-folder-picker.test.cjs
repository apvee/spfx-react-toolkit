const test = require('node:test');
const assert = require('node:assert/strict');
const { createPropertyPaneHarness } = require('./fixtures/property-pane-picker-harness.cjs');
const h = createPropertyPaneHarness();
const value = { kind: 'file', tenantId: '11111111-1111-4111-8111-111111111111', endpoint: 'https://tenant.sharepoint.com/_api/v2.0', driveId: 'b!drive', itemId: 'ONE', name: 'Report.pdf', sharePoint: { siteCollectionId: '22222222-2222-4222-8222-222222222222', webId: '33333333-3333-4333-8333-333333333333', listId: '44444444-4444-4444-8444-444444444444' } };
let observed;
// Exercise the actual native provider/host and direct-value adapter while observing its control contract.
h.load('controls/file-folder-picker/file-folder-picker.tsx').FileFolderPicker = props => {
  observed = { props, epoch: h.React.useContext(h.load('controls/pickers/picker-runtime-context.internal.tsx').PickerAuthoritativeEpochContext) };
  return h.React.createElement('button', { disabled: props.disabled }, props.value.map(item => item.name).join(', '));
};
function setup() {
  const factory = h.load('property-pane/file-folder-picker/property-pane-file-folder-picker.tsx').PropertyPaneFileFolderPicker;
  const owner = h.host(), dom = document.createElement('div'), changes = [];
  const base = { instance: owner.instance, key: 'files', value: [value] };
  let latest;
  const render = (overrides = {}, callback = (...args) => changes.push(args)) => {
    latest = factory('files', { ...base, ...overrides });
    h.act(() => latest.properties.onRender(dom, undefined, callback === null ? undefined : callback));
    return latest;
  };
  render();
  return { factory, base, owner, dom, changes, render, close: () => h.act(() => latest.properties.onDispose(dom)) };
}
function emit(values, reason = 'pick') {
  h.act(() => observed.props.onChange({ reason, values, items: [], unresolvedValues: values }));
}
test('native file folder field validates direct target names and stable keys', () => {
  const f = setup();
  try {
    for (const key of ['', 42, {}, 'files.child', 'files[0]', '__proto__', 'constructor', 'prototype']) assert.throws(() => f.factory(key, f.base), /safe direct/);
    assert.throws(() => f.factory('files', { ...f.base, key: ' ' }), /stable/);
    assert.throws(() => f.factory('files', { ...f.base, key: 42 }), /stable/);
  } finally { f.close(); }
});
test('direct-array bridge retains nonreactive draft, adopts reactive values/revisions, and emits independent fresh JSON', () => {
  const f = setup();
  try {
    assert.deepEqual(f.changes, [], 'mount and hydration never write');
    const selected = { ...value, itemId: 'TWO', name: 'Other.pdf' };
    emit([selected]);
    assert.equal(f.changes.length, 1);
    assert.equal(f.changes[0][0], 'files'); assert.equal(f.changes[0][2], true);
    assert.notEqual(f.changes[0][1], observed.props.value);
    assert.notEqual(f.changes[0][1][0], selected);
    assert.notEqual(f.changes[0][1][0].sharePoint, selected.sharePoint);
    f.changes[0][1][0].sharePoint.listId = 'corrupted';
    f.render({ label: 'Cosmetic', value: [{ ...value }] });
    assert.equal(observed.props.value[0].itemId, 'TWO');
    assert.equal(observed.props.value[0].sharePoint.listId, value.sharePoint.listId);
    const draftEpoch = observed.epoch;
    f.render({ valueRevision: 1 });
    assert.equal(observed.props.value[0].itemId, 'ONE'); assert.notEqual(observed.epoch, draftEpoch);
    f.render({ value: [], valueRevision: 2 }); assert.deepEqual(observed.props.value, []);
    assert.equal(f.changes.length, 1);
    assert.equal(f.owner.refreshes, 0);
  } finally { f.close(); }
});
test('malformed native input or actions retain prior confirmed JSON and missing callback disables mutation', () => {
  const f = setup();
  try {
    emit([{ ...value, itemId: 'TWO' }]);
    emit([{ ...value, driveId: 12 }]);
    assert.equal(observed.props.value[0].itemId, 'TWO'); assert.equal(f.changes.length, 1);
    assert.match(observed.props.validationMessage, /Invalid/);
    f.render({ value: [{ kind: 'file', itemId: 2 }], valueRevision: 1 });
    assert.equal(observed.props.value[0].itemId, 'TWO'); assert.equal(observed.props.disabled, true);
    emit([]); assert.equal(f.changes.length, 1);
    f.render({ valueRevision: 2 }, null);
    assert.equal(observed.props.disabled, true); assert.match(observed.props.validationMessage, /callback/i);
    emit([]); assert.equal(f.changes.length, 1);
  } finally { f.close(); }
});
test('latest callback owns changes and disposed controls cannot emit', () => {
  const f = setup();
  let latest;
  try {
    f.render({}, (...args) => { latest = args; });
    const retainedHandler = observed.props.onChange;
    emit([], 'clear'); assert.deepEqual(latest, ['files', [], true]); assert.deepEqual(f.changes, []);
    f.close(); latest = undefined;
    h.act(() => retainedHandler({ reason: 'pick', values: [value], items: [], unresolvedValues: [value] }));
    assert.equal(latest, undefined);
  } finally { f.close(); }
});

test('native field forwards flat capabilities and hint without resetting its draft', () => {
 const f=setup();
 try {
   emit([{...value,itemId:'DRAFT'}]);
   const onReady=()=>{};
   const extra={title:'Native documents',layout:'tiles',hint:'Consumer description',graphEnabled:true,uploadEnabled:true,createFolderEnabled:true,initialSort:{fieldName:'FileLeafRef'},visibleLocations:['site'],onReady};
   f.render(extra);
   for(const [key,expected] of Object.entries(extra)) assert.deepEqual(observed.props[key],expected);
   assert.equal(observed.props.value[0].itemId,'DRAFT');
   assert.equal(f.changes.length,1);
 } finally {f.close();}
});

test('native preset changes preserve nonreactive draft, authoritative epoch, and display seeds without writes', () => {
  const f = setup();
  try {
    emit([{ ...value, itemId: 'DRAFT' }]);
    const models = [{ value, text: 'Known report' }];
    f.render({ preset: 'focused', initialOptions: models });
    const epoch = observed.epoch;
    f.render({ preset: 'extended', initialOptions: models, value: [{ ...value }] });
    assert.equal(observed.props.preset, 'extended');
    assert.equal(observed.props.initialOptions, models);
    assert.equal(observed.props.value[0].itemId, 'DRAFT');
    assert.equal(observed.epoch, epoch);
    assert.equal(f.changes.length, 1);
  } finally { f.close(); }
});

test('file property pane forwards instance strings and keeps saved references unchanged', () => {
  const f = setup();
  const strings = { browse: { files: 'Choose attachments' }, errors: { timeout: 'Try again' } };
  try {
    f.render({ strings });
    assert.equal(observed.props.strings, strings);
    assert.deepEqual(observed.props.value, [value]);
    assert.equal(f.changes.length, 0);
  } finally { f.close(); }
});

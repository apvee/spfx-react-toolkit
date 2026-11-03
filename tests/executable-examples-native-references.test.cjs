const test = require('node:test');
const assert = require('node:assert/strict');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
// Independent expected owners and prerequisites, not a production availability registry.
const cases = [
  ['List', 'list', 'list', 'pickerLists', 'picker-lists', 'Lists'],
  ['Field', 'field', 'field', 'pickerFields', 'picker-fields', 'Fields'],
  ['View', 'view', 'view', 'pickerViews', 'picker-views', 'Views'],
  ['ContentType', 'content-type', 'content type', 'pickerContentTypes', 'picker-content-types', 'Content types'],
  ['ListItem', 'list-item', 'item', 'pickerItems', 'picker-items', 'Items'],
  ['Principal', 'principal', 'principal', 'pickerPrincipals', 'picker-principals', 'Principals'],
];
for (const [name, stem, noun] of cases) test(`native ${name} exact executable availability`, () => {
  const selected = scenario(`Select native ${noun} references`, `packages/spfx-react-toolkit/src/property-pane/pickers/property-pane-${stem}-picker.ts#PropertyPane${name}Picker`);
  assert.deepEqual(selected.entry.publicNames, [`PropertyPane${name}Picker`]);
  assert.deepEqual(Object.fromEntries(Object.entries(selected.entry.metadata).filter(([key]) => key !== 'observe')), { host: 'webPart', kind: 'read', contentPlacement: 'configuration', hostAdapter: 'propertyPane' });
  assert.match(selected.entry.metadata.observe, /Prepare.*Edit.*Close.*saved JSON/);
});
const { createNativeReferenceFixture } = require('./fixtures/executable-examples/native-site-scenario.cjs');
const listId = '22222222-2222-4222-8222-000000000001', web = 'https://tenant.test/sites/root/sub';
const expected = {
  List: ['List 00', { kind: 'list', webUrl: web, listId: '22222222-2222-4222-8222-000000000010' }],
  Field: ['Column 000', { kind: 'field', webUrl: web, listId, fieldId: '22222222-2222-4222-8222-000000000100', internalName: 'Column0' }],
  View: ['View 000', { kind: 'view', webUrl: web, listId, viewId: '22222222-2222-4222-8222-000000000400' }],
  ContentType: ['Content 000', { kind: 'contentType', scope: 'site', webUrl: 'https://tenant.test/sites/root', contentTypeId: '0x01' }],
  ListItem: ['Alpha document', { kind: 'file', webUrl: web, listId, itemId: 2, uniqueId: '22222222-2222-4222-8222-000000000701', serverRelativeUrl: '/sites/root/sub/Docs/alpha.pdf', extension: 'pdf' }],
  Principal: ['Alice', { kind: 'user', loginName: 'alice@tenant.test' }],
};
function prerequisites(f, name) { if (['Field', 'View', 'ListItem'].includes(name)) f.change('Exact list title', 'Tasks'); }
for (const compiled of [false, true]) for (const [name, , noun, target, key, label] of cases) {
  const mode = compiled ? 'COMPILED' : 'SOURCE';
  test(`${mode} ${name} actual public factory pinned parent provider native change and live saved JSON`, async t => {
    const f = createNativeReferenceFixture(t, { subject: name, compiled }); await f.flush();
    assert.equal(f.fields.length, 0); assert.equal(f.calls.length, 0);
    const first = f.container.querySelector('button'), original = f.action('Prepare native scenario');
    if (['Field', 'View', 'ListItem'].includes(name)) { assert.equal(first.disabled, true); f.h.act(() => original()); assert.equal(f.adapter.candidate, undefined); }
    prerequisites(f, name); f.click('Prepare native scenario'); assert.equal(f.calls.length, 0); assert.equal(f.fields.length, 0);
    if (['Field', 'View', 'ListItem'].includes(name)) f.change('Exact list title', 'Docs');
    const config = f.getter(), group = config.pages[0].groups[0], field = group.groupFields[0];
    assert.equal(config.pages.length, 1); assert.equal(group.isGroupNameHidden, true); assert.equal(Object.hasOwn(group, 'groupName'), false); assert.equal(Object.hasOwn(group, 'isCollapsed'), false); assert.equal(group.groupFields.length, 1);
    assert.equal(f.fields.length, 1); assert.equal(f.fields[0].args[0], target);
    const options = f.fields[0].args[1]; assert.equal(options.instance, f.instance); assert.equal(options.key, key); assert.equal(options.label, label); assert.deepEqual(options.value, []);
    if (['Field', 'View', 'ListItem'].includes(name)) assert.equal(options.list, 'Tasks', 'prepared real title is frozen independently of draft edits');
    if (name === 'ListItem') assert.equal(options.textField, 'Title');
    if (name === 'ContentType') { assert.equal(options.scope, 'site'); assert.equal(Object.hasOwn(options, 'list'), false); }
    if (name === 'Principal') { assert.deepEqual(options.principalTypes, ['user']); assert.deepEqual(options.userIdentity, { mode: 'available' }); }
    assert.equal(f.calls.length, 0); f.mountPane(config); await f.flush();
    assert.equal(f.nativeDom.querySelectorAll('.ms-CustomFieldHost').length, 1); assert.equal(f.handlers.size, 1);
    f.open(); if (name === 'Principal' || name === 'ListItem') { f.search(name === 'Principal' ? 'Alice' : 'Alpha'); await f.wait(); } else await f.flush();
    f.choose(expected[name][0]); await f.h.act(async () => { await f.saved; }); await f.flush();
    assert.deepEqual(JSON.parse(JSON.stringify(f.bag[target])), [expected[name][1]], 'independent native saved reference');
    assert.equal(f.changes.length, 1); assert.equal(f.changes[0].name, target); assert.equal(f.changes[0].old, undefined);
    assert.ok(f.calls.length); assert.ok(f.calls.every(request => request.phase.startsWith('native-')));
    assert.equal(f.calls.some(request => /ensureuser/i.test(request.url)), false);
    if (name === 'Principal') assert.ok(f.calls.some(request => request.method === 'POST' && request.body && request.url.includes('clientPeoplePickerSearchUser')));
    else assert.ok(f.calls.every(request => request.method === 'GET'));
    f.click(`Inspect saved ${noun} references`); assert.deepEqual(JSON.parse(f.container.querySelector('pre').textContent), [expected[name][1]]);
    if (name === 'List') {
      f.open(); await f.flush(); f.choose('List 01'); await f.flush();
      assert.deepEqual(f.changes.at(-1).old, [expected.List[1]], 'actual prior SDK array survives pinned parent pre-hook mutation');
      assert.deepEqual(f.bag.pickerLists, [{ kind: 'list', webUrl: web, listId: '22222222-2222-4222-8222-000000000011' }]);
    }
    const replacement = [{ ...expected[name][1], text: 'SDK replacement' }]; f.bag = { [target]: replacement }; f.click(`Inspect saved ${noun} references`); assert.deepEqual(JSON.parse(f.container.querySelector('pre').textContent), replacement);
    const trigger = f.nativeDom.querySelector('[role="combobox"]'), calls = f.calls.length;
    f.setRevision(3); f.h.act(() => original()); await f.flush(); assert.equal(f.nativeDom.querySelector('[role="combobox"]'), trigger); assert.equal(f.calls.length, calls);
    f.setRevision(2); f.click(`Inspect saved ${noun} references`); assert.match(f.container.textContent, /Open the prepared native scenario before inspecting/);
    f.closePane(); assert.equal(f.handlers.size, 0); assert.equal(f.nativeDom.textContent, '');
    f.h.act(() => field.properties.onRender(f.nativeDom, undefined, () => assert.fail('late field cannot save'))); assert.equal(f.nativeDom.textContent, '');
    const implementationFiles = f.origins(); assert.deepEqual(f.dataRoles.map(row => row.id), ['ApveeSpfxReactToolkitPickersStrings']);
    t.diagnostic(JSON.stringify({ subject: name, mode, actualAdapter: f.adapterFile, implementationFiles, nativeData: f.dataRoles, attempts: f.calls.map(({ kind, method, url, query, body, init, phase, identity }) => ({ kind, method, url, query, body, headers: init?.headers, phase, identity })) }));
  });
  test(`${mode} ${name} absent port staging and repeated Reset cancel only own unproduced preparation`, async t => {
    const f = createNativeReferenceFixture(t, { subject: name, compiled, connected: false }); await f.flush();
    assert.match(f.container.textContent, /initialized native scenario host is not connected/); assert.match(f.container.textContent, /existing full integrated configuration independently/);
    assert.equal(f.container.querySelectorAll('input,button').length, 0); f.setRevision(3); f.setRevision(2); assert.equal(f.fields.length + f.calls.length, 0);
  });
  test(`${mode} ${name} repeated Reset invalidates original callbacks without remounting Demo`, async t => {
    const f = createNativeReferenceFixture(t, { subject: name, compiled }); prerequisites(f, name);
    const button = f.container.querySelector('button'), prepare = f.action('Prepare native scenario');
    f.click('Prepare native scenario'); const candidate = f.adapter.candidate; assert.ok(candidate);
    f.setRevision(3); assert.equal(f.adapter.candidate, undefined); f.setRevision(2); f.h.act(() => prepare()); assert.equal(f.adapter.candidate, undefined);
    prerequisites(f, name); f.click('Prepare native scenario'); assert.ok(f.adapter.candidate); assert.notEqual(f.adapter.candidate, candidate);
    assert.equal(f.container.querySelector('button'), button); assert.equal(f.fields.length + f.calls.length, 0);
  });
}

for (const compiled of [false, true]) test(`${compiled ? 'COMPILED' : 'SOURCE'} Principal real pending confirmation cannot save after native cancellation and disposal`, async t => {
  const f = createNativeReferenceFixture(t, { subject: 'Principal', compiled, hold: request => request.body && JSON.parse(request.body).queryParams?.QueryString === 'alice@tenant.test' });
  f.click('Prepare native scenario'); const config = f.getter(); f.mountPane(config); await f.flush();
  f.open(); f.search('Alice'); await f.wait(); f.choose('Alice'); await f.flush();
  assert.equal(f.pending.length, 1); assert.equal(f.changes.length, 0); assert.deepEqual(f.bag, {});
  f.setRevision(3); f.closePane(); await f.settle(f.pending[0]); await f.flush();
  assert.deepEqual(f.bag, {}); assert.equal(f.changes.length, 0); assert.equal(f.handlers.size, 0);
  assert.equal(f.calls.some(request => /ensureuser/i.test(request.url)), false); f.origins();
});

for (const compiled of [false, true]) for (const [name, , noun, target] of cases) test(`${compiled ? 'COMPILED' : 'SOURCE'} ${name} reused revision permanently retires original Prepare and Inspect`, async t => {
  const f = createNativeReferenceFixture(t, { subject: name, compiled, revision: 7 }); await f.flush(); prerequisites(f, name);
  const button = f.container.querySelector('button'), originalPrepare = f.action('Prepare native scenario'), originalInspect = f.action(`Inspect saved ${noun} references`);
  assert.match(f.container.querySelector('[role="status"]').textContent, /Prepare before opening/);
  f.click('Prepare native scenario'); f.mountPane(f.getter()); await f.flush();
  const native = f.nativeDom.querySelector('[role="combobox"]'); assert.ok(native); assert.equal(f.fields.length, 1);
  f.setRevision(3); f.setRevision(7); await f.flush();
  const survives = f.nativeDom.querySelector('[role="combobox"]') === native && f.handlers.size === 1;
  const resetStatus = f.container.querySelector('[role="status"]').textContent;
  f.h.act(() => { originalInspect(); }); const retiredInspectStatus = f.container.querySelector('[role="status"]').textContent;
  f.closePane(); f.h.act(() => { originalPrepare(); });
  const retiredConfiguration = f.getter(), retiredFieldCount = f.fields.length; f.closePane();
  f.bag = {};
  if (['Field', 'View', 'ListItem'].includes(name)) f.change('Exact list title', 'Docs');
  if (name === 'ListItem') f.change('Verified internal display field', 'FileLeafRef');
  f.click('Prepare native scenario'); const current = f.getter(); f.mountPane(current); await f.flush();
  const reference = { ...expected[name][1], text: 'Fresh saved reference' };
  f.bag = { [target]: [reference] }; f.click(`Inspect saved ${noun} references`);
  const freshSaved = f.container.querySelector('pre').textContent;
  f.bag = { [target]: [{ ...reference, text: 'Later SDK reference' }] }; f.h.act(() => { originalInspect(); });
  const retiredSaved = f.container.querySelector('pre').textContent;
  f.click(`Inspect saved ${noun} references`); const currentSaved = JSON.parse(f.container.querySelector('pre').textContent);
  t.diagnostic(JSON.stringify({ reusedRevision: name, compiled, survives, retiredInspectStatus, resetStatus, retiredPages: retiredConfiguration.pages.length, retiredFieldCount, retiredList: f.fields[1]?.args[1].list, freshList: f.fields.at(-1).args[1].list, freshFields: f.fields.length }));
  assert.equal(survives, true, 'Reset cannot dispose the already-produced native session');
  assert.equal(retiredInspectStatus, resetStatus, 'original Inspect stays retired when revision 7 returns');
  assert.equal(retiredConfiguration.pages.length, 0, 'original Prepare cannot create native configuration after ABA');
  assert.equal(retiredFieldCount, 1, 'original Prepare cannot construct a field with retired prerequisites');
  assert.equal(f.fields.length, 2, 'current fresh Prepare remains functional');
  if (['Field', 'View', 'ListItem'].includes(name)) assert.equal(f.fields.at(-1).args[1].list, 'Docs');
  if (name === 'ListItem') assert.equal(f.fields.at(-1).args[1].textField, 'FileLeafRef');
  assert.equal(retiredSaved, freshSaved, 'original Inspect cannot observe through the new live reader');
  assert.deepEqual(JSON.parse(freshSaved), [reference]); assert.deepEqual(currentSaved, [{ ...reference, text: 'Later SDK reference' }]);
  assert.equal(f.container.querySelector('button'), button, 'Demo identity survives both resets');
  assert.equal(f.calls.length, 0, 'prerequisite collection, preparation and reset issue no business query'); f.closePane(); f.origins();
});

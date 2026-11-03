const test = require('node:test');
const assert = require('node:assert/strict');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const owner = 'packages/spfx-react-toolkit/src/property-pane/file-folder-picker/property-pane-file-folder-picker.tsx#PropertyPaneFileFolderPicker';
const title = 'Select native file and folder references';
test('native FileFolder exact independent owning executable availability', () => {
  const { entry } = scenario(title, owner);
  assert.deepEqual(entry.publicNames, ['PropertyPaneFileFolderPicker']);
  assert.deepEqual(Object.fromEntries(Object.entries(entry.metadata).filter(([key]) => key !== 'observe')), { host: 'webPart', kind: 'read', contentPlacement: 'configuration', hostAdapter: 'propertyPane' });
  assert.match(entry.metadata.observe, /Prepare.*Edit.*Browse.*confirm.*cancel.*saved JSON/);
});
const { createNativeReferenceFixture } = require('./fixtures/executable-examples/native-site-scenario.cjs');
const value = (folder = false) => ({ kind: folder ? 'folder' : 'file', tenantId: '11111111-1111-4111-8111-111111111111', endpoint: 'https://tenant.sharepoint.com/sites/HR/_api/v2.0', driveId: 'b!Drive', itemId: folder ? 'Folder' : 'FileA', name: folder ? 'Policies' : 'FileA.pdf', webUrl: 'https://tenant.sharepoint.com/sites/HR/Shared%20Documents/' + (folder ? 'Policies' : 'FileA.pdf'), sharePoint: { siteCollectionId: '22222222-2222-4222-8222-222222222222', webId: '33333333-3333-4333-8333-333333333333', listId: '44444444-4444-4444-8444-444444444444', listItemId: 7, uniqueId: '66666666-6666-4666-8666-666666666666' } });
for (const compiled of [false, true]) {
  const mode = compiled ? 'COMPILED' : 'SOURCE';
  test(`${mode} FileFolder actual pinned native parent hosted confirm and live SDK saved JSON`, async t => {
    const f = createNativeReferenceFixture(t, { subject: 'FileFolder', compiled }); await f.flush();
    assert.equal(f.calls.length + f.fields.length, 0);
    f.click('Prepare native scenario'); assert.equal(f.calls.length + f.fields.length, 0);
    const configuration = f.getter(), group = configuration.pages[0].groups[0];
    assert.equal(configuration.pages.length, 1); assert.equal(group.isGroupNameHidden, true);
    assert.equal(Object.hasOwn(group, 'groupName'), false); assert.equal(Object.hasOwn(group, 'isCollapsed'), false); assert.equal(group.groupFields.length, 1);
    assert.equal(f.fields.length, 1); assert.equal(f.fields[0].args[0], 'fileFolderValues');
    const options = f.fields[0].args[1];
    assert.equal(options.instance, f.instance);
    assert.deepEqual({ ...options, instance: undefined }, { instance: undefined, key: 'file-folder-values', value: [], label: 'Files and folders', preset: 'focused', scope: { kind: 'siteCollection' }, itemTypes: 'all' });
    assert.equal(f.calls.length, 0); f.mountPane(configuration); await f.flush();
    assert.equal(f.nativeDom.querySelectorAll('.ms-CustomFieldHost').length, 1); assert.equal(f.handlers.size, 1);
    assert.equal(f.calls.length, 0, 'mount initializes providers without opening or fetching');
    for (const folder of [false, true]) {
      f.browse(); await f.flush(); assert.equal(f.hosted.posts.length, folder ? 2 : 1, JSON.stringify({ calls: f.calls, text: document.body.textContent }));
      const x = f.hosted.opening(); x.initialize(); assert.deepEqual(x.port.messages, [{ type: 'activate' }]);
      const authenticated = x.result('auth'); x.command('auth', { command: 'authenticate', resource: 'https://tenant.sharepoint.com', type: 'SharePoint' });
      await f.h.act(async () => { assert.equal((await authenticated).data.result, 'token'); });
      const saved = f.nextSaved();
      const confirmed = x.result('pick'); x.command('pick', { command: 'pick', items: [x.raw(folder ? 'Folder' : 'FileA')] });
      await f.h.act(async () => { assert.equal((await confirmed).data.result, 'success'); await saved; }); await f.flush();
      assert.deepEqual(JSON.parse(JSON.stringify(f.bag.fileFolderValues)), [value(folder)]);
      assert.equal(f.changes.length, folder ? 2 : 1); assert.equal(f.changes.at(-1).name, 'fileFolderValues');
      f.click('Inspect saved file and folder references'); assert.deepEqual(JSON.parse(f.container.querySelector('pre').textContent), [value(folder)]);
      assert.equal(x.port.closed, true);
    }
    assert.equal(f.calls.some(r => r.url?.includes('/Subweb/')), false, 'scope uses actual current collection root');
    assert.ok(f.calls.filter(r => r.kind === 'sp').every(r => r.method === 'GET' && r.phase === 'native-hosted-browse'));
    const replacement = [{ ...value(), name: 'SDK replaced bag' }]; f.bag = { fileFolderValues: replacement }; f.click('Inspect saved file and folder references');
    assert.deepEqual(JSON.parse(f.container.querySelector('pre').textContent), replacement);
    f.closePane(); assert.equal(f.handlers.size, 0);
    const implementationFiles = f.origins();
    assert.deepEqual(f.dataRoles.map(row => row.id).sort(), ['ApveeSpfxReactToolkitFileFolderPickerStrings', 'ApveeSpfxReactToolkitStageStrings']);
    t.diagnostic(JSON.stringify({ subject: 'FileFolder', mode, actualAdapter: f.adapterFile, implementationFiles, nativeData: f.dataRoles, attempts: f.calls, messages: f.hosted.messages }));
  });
}
for (const compiled of [false, true]) {
  const mode = compiled ? 'COMPILED' : 'SOURCE';
  test(`${mode} FileFolder missing port stays inactive and unproduced preparation resets without I/O`, async t => {
    const f = createNativeReferenceFixture(t, { subject: 'FileFolder', compiled, connected: false });
    assert.match(f.container.textContent, /selected native scenario is not active/); assert.match(f.container.textContent, /existing full integrated configuration independently/);
    assert.equal(f.container.querySelectorAll('button,input').length, 0); f.setRevision(3); f.setRevision(7); await f.flush();
    assert.equal(f.fields.length + f.calls.length, 0);
  });
  test(`${mode} FileFolder initial7 and ABA retire captured Prepare and Inspect while actual native session survives`, async t => {
    const f = createNativeReferenceFixture(t, { subject: 'FileFolder', compiled });
    const button = f.container.querySelector('button'), prepare = f.action('Prepare native scenario'), inspect = f.action('Inspect saved file and folder references');
    assert.match(f.container.querySelector('[role="status"]').textContent, /Prepare before opening/);
    f.click('Prepare native scenario'); assert.ok(f.adapter.candidate); f.setRevision(3); assert.equal(f.adapter.candidate, undefined);
    f.setRevision(7); f.h.act(() => prepare()); assert.equal(f.adapter.candidate, undefined);
    f.click('Prepare native scenario'); f.mountPane(f.getter()); await f.flush();
    const trigger = f.nativeDom.querySelector('button[aria-haspopup="dialog"]'); assert.ok(trigger);
    f.browse(); await f.flush(); const x = f.hosted.opening(); x.initialize();
    const originalSessionFrame = x.frame;
    f.setRevision(3); f.setRevision(7);
    const status = f.container.querySelector('[role="status"]').textContent;
    f.h.act(() => { prepare(); inspect(); });
    assert.equal(f.container.querySelector('[role="status"]').textContent, status);
    assert.equal(f.nativeDom.querySelector('button[aria-haspopup="dialog"]'), trigger); assert.equal(f.handlers.size, 1);
    assert.equal(x.port.closed, undefined); assert.equal(originalSessionFrame.isConnected, true, 'Reset does not own already-produced hosted lifetime');
    const saved = f.nextSaved(), confirmed = x.result('pick'); x.command('pick', { command: 'pick', items: [x.raw('FileA')] });
    await f.h.act(async () => { await confirmed; await saved; }); await f.flush(); assert.deepEqual(f.bag.fileFolderValues, [value()]);
    f.closePane(); f.h.act(() => prepare()); const fallback = f.getter(); assert.equal(fallback.pages.length, 0); assert.equal(f.fields.length, 1); f.closePane();
    f.click('Prepare native scenario'); f.mountPane(f.getter()); await f.flush();
    f.bag = { fileFolderValues: [value(true)] }; f.click('Inspect saved file and folder references');
    const fresh = f.container.querySelector('pre').textContent; assert.deepEqual(JSON.parse(fresh), [value(true)]);
    f.bag = { fileFolderValues: [value()] }; f.h.act(() => inspect()); assert.equal(f.container.querySelector('pre').textContent, fresh);
    f.click('Inspect saved file and folder references'); assert.deepEqual(JSON.parse(f.container.querySelector('pre').textContent), [value()]);
    assert.equal(f.container.querySelector('button'), button); assert.equal(f.fields.length, 2); f.closePane(); f.origins();
  });
  test(`${mode} FileFolder actual hosted cancel preserves SDK bag and native Close cancels pending validation`, async t => {
    const f = createNativeReferenceFixture(t, { subject: 'FileFolder', compiled, hold: r => r.kind === 'sp' && r.url.includes('/items/Folder') });
    f.click('Prepare native scenario'); const configuration = f.getter(), field = configuration.pages[0].groups[0].groupFields[0]; f.mountPane(configuration); await f.flush();
    f.browse(); await f.flush(); const initial = f.hosted.opening(); initial.initialize();
    const saved = f.nextSaved(), confirmation = initial.result('file'); initial.command('file', { command: 'pick', items: [initial.raw('FileA')] });
    await f.h.act(async () => { await confirmation; await saved; }); await f.flush();
    const bag = f.bag;
    f.browse(); await f.flush(); const cancel = f.hosted.opening(); cancel.initialize();
    const cancelled = cancel.result('cancel'); cancel.command('cancel', { command: 'close' }); await f.h.act(async () => { assert.equal((await cancelled).data.result, 'success'); }); await f.flush();
    assert.equal(f.bag, bag); assert.equal(f.changes.length, 1); assert.equal(cancel.port.closed, true);
    f.browse(); await f.flush(); const late = f.hosted.opening(); late.initialize();
    late.command('pending', { command: 'pick', items: [late.raw('Folder')] }); await f.flush();
    assert.equal(f.pending.length, 1); assert.equal(f.changes.length, 1); const callback = late.port.onmessage;
    f.closePane(); assert.equal(late.port.closed, true); assert.equal(late.port.onmessage, null); assert.equal(f.handlers.size, 0);
    await f.settle(f.pending[0]); f.h.act(() => callback({ data: { type: 'command', id: 'late', data: { command: 'pick', items: [late.raw('FileA')] } } })); await f.flush();
    assert.equal(f.bag, bag); assert.equal(f.changes.length, 1); assert.deepEqual(f.bag.fileFolderValues, [value()]);
    f.h.act(() => field.properties.onRender(f.nativeDom, undefined, () => assert.fail('retired field cannot write'))); assert.equal(f.nativeDom.textContent, '');
    f.origins();
  });
}

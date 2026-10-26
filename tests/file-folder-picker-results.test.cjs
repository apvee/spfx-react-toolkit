const test = require('node:test');
const assert = require('node:assert/strict');
const { ids, web, endpoint, metadata, setup } = require('./fixtures/file-folder-picker/metadata.cjs');
async function state(kind = 'library', extra = {}) { const s = setup(extra); const scope = await s.load('scope').resolveFileFolderPickerScope({ kind, webUrl: web, list: 'Documents', ...(kind === 'folder' ? { folder: 'Policies' } : {}) }, undefined, s.context, s.cache); return { ...s, scope }; }
const pick = (s, ids, policy) => s.load('results').resolveFileFolderPickerItems(ids.map(id => typeof id === 'string' ? s.raw(id) : id), s.scope, s.context, s.cache, policy);
test('minimum raw results hydrate canonical labels, opaque IDs and optional enrichment; root has no numeric ID', async () => {
  const s = await state(); s.items.FileA.name = 'Renamed.PDF';
  const items = await pick(s, ['FileA', 'Root'], { itemTypes: 'all', multiselect: true });
  assert.equal(items[0].text, 'Renamed.PDF'); assert.equal(items[0].extension, '.pdf'); assert.equal(items[0].value.sharePoint.listItemId, 7);
  assert.equal(items[1].value.kind, 'folder'); assert.equal(items[1].value.sharePoint.listItemId, undefined);
});
test('confirmation re-reads cached hydration and rejects a deleted or renamed-ineligible item', async () => {
  const s = await state(); const value = { kind: 'file', tenantId: ids.tenant, endpoint, driveId: 'b!Drive', itemId: 'FileA', name: 'snapshot.pdf' };
  assert.equal((await s.load('results').resolveFileFolderPickerValue(value, s.scope, s.context, s.cache)).text, 'FileA.pdf');
  const count = s.calls.length; s.items.FileA = metadata('FileA', { name: 'renamed.exe' });
  await assert.rejects(pick(s, ['FileA'], { extensions: ['.pdf'] }), /extension/i); assert.ok(s.calls.length > count);
  s.items.FileA.deleted = {}; await assert.rejects(pick(s, ['FileA']), /deleted|unavailable/i);
  assert.equal(await s.load('results').resolveFileFolderPickerValue(value, s.scope, s.context, s.cache), undefined);
});
test('batch acceptance deduplicates order atomically and includes retained unresolved references in count', async () => {
  const s = await state(); const result = await pick(s, ['FileA', 'Child', 'FileA'], { multiselect: true, selectionLimit: 2 }); assert.deepEqual(result.map(o => o.value.itemId), ['FileA', 'Child']);
  assert.equal(s.calls.filter(uri => uri.includes('/items/FileA')).length, 1);
  const retainedValues = [{ kind: 'file', tenantId: ids.tenant, endpoint, driveId: 'b!Drive', itemId: 'unresolved' }];
  await assert.rejects(pick(s, ['FileA', 'Child'], { multiselect: true, selectionLimit: 2, retainedValues }), /limit/i);
  s.items.Child.sharepointIds.siteId = ids.tenant; await assert.rejects(pick(s, ['FileA', 'Child'], { multiselect: true }), /collection/i);
});
test('canonical metadata rejects hostile endpoints, wrong tenant/collection/library, partial and remote results', async () => {
  for (const patch of [{ sharepointIds: { tenantId: ids.site, siteId: ids.site, webId: ids.web, listId: ids.list } }, { sharepointIds: { tenantId: ids.tenant, siteId: ids.tenant, webId: ids.web, listId: ids.list } }, { sharepointIds: { tenantId: ids.tenant, siteId: ids.site, webId: ids.web, listId: ids.tenant } }, { parentReference: { driveId: 'wrong' } }, { remoteItem: {} }, { file: undefined, folder: undefined }, { name: undefined }]) {
    const s = await state(); s.items.FileA = metadata('FileA', patch); await assert.rejects(pick(s, ['FileA']), /tenant|collection|library|drive|remote|metadata|name|facet/i);
  }
  for (const malicious of ['https://other.sharepoint.com/_api/v2.0', 'https://tenant.sharepoint.com/_api/v2.0/drives/x', 'https://graph.microsoft.com/v1.0', 'https://tenant.sharepoint.com/x/../_api/v2.0']) { const s = await state(); const before = s.calls.length; await assert.rejects(pick(s, [{ ...s.raw('FileA'), '@sharePoint.endpoint': malicious }]), /endpoint|reference/i); assert.equal(s.calls.length, before); }
});
test('folder ancestry uses complete parent IDs, handles descendants, sibling rejection and cycles', async () => {
  const s = await state('folder'); assert.equal((await pick(s, ['Child']))[0].value.itemId, 'Child');
  await assert.rejects(pick(s, ['Outside']), /folder|scope/i);
  s.items.Child.parentReference.id = 'Child'; await assert.rejects(pick(s, ['Child']), /cycle|ancestry/i);
  const changed = { ...s.context, userId: 'different' }; await assert.rejects(s.load('results').resolveFileFolderPickerItems([s.raw('Child')], s.scope, changed, s.cache), /source|context/i);
});
test('collection permits authoritative subweb provenance but rejects outside collection; stale callbacks never emit', async () => {
  const s = setup(); s.scope = await s.load('scope').resolveFileFolderPickerScope(undefined, undefined, s.context, s.cache); s.items.FileA.sharepointIds.webId = ids.folder;
  assert.equal((await pick(s, ['FileA']))[0].value.sharePoint.webId, ids.folder);
  await assert.rejects(pick(s, ['FileA'], { isCurrent: () => false }), /stale/i);
  s.cache.dispose(); await assert.rejects(pick(s, ['FileA']), /disposed/i);
});
test('root provenance can be recovered live without fabricating optional item IDs', async () => {
  const s = await state(); delete s.items.Root.sharepointIds;
  const previous = s.context.spClient.get;
  s.context.spClient.get = async uri => new URL(uri).pathname.endsWith('/drives/b!Drive') ? { ok: true, json: async () => ({ id: 'b!Drive', ...(new URL(uri).searchParams.get('$select')?.split(',').includes('sharepointIds') ? {sharepointIds: { tenantId: ids.tenant, siteId: ids.site, webId: ids.web, listId: ids.list }} : {}) }) } : previous(uri);
  const [root] = await pick(s, ['Root'], { itemTypes: 'folders' }); assert.equal(root.value.kind, 'folder'); assert.equal(root.value.sharePoint.listItemId, undefined);
});
test('unreadable/partial results reject the whole batch and unreadable persisted references stay unresolved', async () => {
  const s = await state(); const original = s.context.spClient.get;
  s.context.spClient.get = async uri => uri.includes('/items/Child') ? { ok: false, status: 403, statusText: 'Forbidden' } : original(uri);
  await assert.rejects(pick(s, ['FileA', 'Child'], { multiselect: true }), /403/);
  const value = { kind: 'file', tenantId: ids.tenant, endpoint, driveId: 'b!Drive', itemId: 'Child', name: 'retained' };
  assert.equal(await s.load('results').resolveFileFolderPickerValue(value, s.scope, s.context, s.cache), undefined); assert.equal(value.name, 'retained');
});
test('pending confirmation checks staleness after authoritative reads', async () => {
  const s = await state(); let active = true; const original = s.context.spClient.get;
  s.context.spClient.get = async uri => { const response = await original(uri); if (uri.includes('/items/FileA')) active = false; return response; };
  await assert.rejects(pick(s, ['FileA'], { isCurrent: () => active }), /stale/i);
});
test('item confirmation explicitly selects provenance instead of relying on default metadata or download URLs', async()=>{
 const s=await state();const original=s.context.spClient.get;const reads=[];
 s.context.spClient.get=async uri=>{reads.push(uri);const response=await original(uri);const body=await response.json();if(new URL(uri).pathname.includes('/items/')&&!new URL(uri).searchParams.get('$select')?.split(',').includes('sharepointIds')){const copy={...body};delete copy.sharepointIds;return {...response,json:async()=>copy};}return {...response,json:async()=>body};};
 const [picked]=await pick(s,['FileA']);assert.equal(picked.value.itemId,'FileA');assert.ok(reads.some(uri=>new URL(uri).searchParams.get('$select')?.split(',').includes('sharepointIds')));assert.ok(reads.every(uri=>!uri.includes('downloadUrl')));
});

test('native root alias resolves to a verified canonical root ID; ordinary mismatches still fail', async () => {
  const s = await state('library', { response: (uri, items) => new URL(uri).pathname.endsWith('/items/root') ? items.Root : undefined });
  const [root] = await pick(s, ['root'], { itemTypes: 'folders' });
  assert.equal(root.value.itemId, 'Root');
  assert.equal(root.value.kind, 'folder');
  assert.ok(s.calls.some(uri => new URL(uri).pathname.endsWith('/root')));
  s.items.FileA.id = 'Another';
  await assert.rejects(pick(s, ['FileA']), /identity/i);
  const spoof = await state('library', { response: (uri, items) => new URL(uri).pathname.endsWith('/items/root') ? items.Folder : undefined });
  await assert.rejects(pick(spoof, ['root'], { itemTypes: 'folders' }), /identity|root/i);
});

test('root alias cannot bypass a conflicting authoritative root relationship', async () => {
  const s = await state('library', { response: (uri, items) => new URL(uri).pathname.endsWith('/items/root') ? { ...items.Root, id: 'FalseRoot' } : undefined });
  await assert.rejects(pick(s, ['root'], { itemTypes: 'folders' }), /identity/i);
});

test('successful verified personal authentication cannot authorize personal or Graph result metadata for SharePoint scope',async t=>{
 const s=await state();t.after(()=>{s.cache.dispose();s.h.close();});const listeners=new Set(),sent=[],tokens=[];
 const ownerWindow={addEventListener:(_,fn)=>listeners.add(fn),removeEventListener:(_,fn)=>listeners.delete(fn)};
 const contentWindow={closed:false},port={start(){},close(){},postMessage:m=>sent.push(m)};
 const session=s.load('session').createFileFolderPickerSession({ownerWindow,contentWindow,closeWindow:false,trustedOrigin:'https://tenant.sharepoint.com',channelId:'boundary',allowedResourceOrigins:s.scope.allowedResourceOrigins,graphEnabled:true,
  resolveAuxiliaryAuthenticationOrigins:async()=>[(await s.load('scope').resolveFileFolderPickerScope({kind:'oneDrive'},undefined,s.context,s.cache)).resource],
  requestToken:async request=>{tokens.push(request.resource);return 'verified-personal';},validatePick:items=>s.load('results').resolveFileFolderPickerItems(items,s.scope,s.context,s.cache)});
 t.after(()=>session.cancel());const result=session.run();for(const listener of listeners)listener({source:contentWindow,origin:'https://tenant.sharepoint.com',data:{type:'initialize',channelId:'boundary'},ports:[port]});
 port.onmessage({data:{type:'command',id:'auth',data:{command:'authenticate',resource:'https://tenant-my.sharepoint.com',type:'SharePoint_MSAL'}}});await new Promise(resolve=>setImmediate(resolve));
 assert.deepEqual(tokens,['https://tenant-my.sharepoint.com']);assert.equal(sent.find(m=>m.id==='auth'&&m.type==='result').data.result,'token');assert.deepEqual(s.scope.allowedResourceOrigins,['https://tenant.sharepoint.com']);
 for(const [id,endpoint] of [['personal','https://tenant-my.sharepoint.com/personal/user/_api/v2.0'],['graph','https://graph.microsoft.com/v1.0']]){
  const reads=s.calls.length;port.onmessage({data:{type:'command',id,data:{command:'pick',items:[{...s.raw('FileA'),'@sharePoint.endpoint':endpoint}]}}});await new Promise(resolve=>setImmediate(resolve));
  assert.equal(sent.find(m=>m.id===id&&m.type==='result').data.result,'error');assert.equal(s.calls.length,reads);
 }
 session.cancel();assert.deepEqual(await result,{status:'cancelled'});
});

test('category acceptance uses authoritative case-insensitive filename, keeps folders, rejects mismatches atomically', async () => {
  for (const [category, name] of Object.entries({ images: 'Photo.JPG', video: 'Movie.MP4', audio: 'Track.MP3', media: 'Track.WAV', documents: 'Report.DOCX' })) {
    const s = await state(); s.items.FileA.name = name;
    const before = s.calls.length;
    const options = await pick(s, ['FileA', 'Folder'], { fileCategory: category, itemTypes: 'all', multiselect: true });
    assert.equal(options[0].text, name); assert.equal(options[1].value.kind, 'folder');
    assert.equal(s.calls.slice(before).filter(uri => uri.includes('/items/FileA')).length, 1);
    assert.equal(s.load('results').fileFolderPickerOptionMatchesPolicy(options[0], { fileCategory: category }), true);
    assert.equal(s.load('results').fileFolderPickerOptionMatchesPolicy(options[0], { fileCategory: category === 'images' ? 'audio' : 'images' }), false);
    s.items.FileA.name = 'renamed.exe';
    await assert.rejects(pick(s, ['Folder', 'FileA'], { fileCategory: category, itemTypes: 'all', multiselect: true }), /category/i);
  }
  const s = await state(); s.items.FileA.name = 'unknown'; s.items.FileA.file = { mimeType: 'image/png' };
  await assert.rejects(pick(s, ['FileA'], { fileCategory: 'images' }), /category/i);
  assert.equal((await pick(s, ['FileA']))[0].text, 'unknown');
});

test('category mismatches map to the shared sanitized selection policy error', () => {
  const s = setup();
  const errors = s.h.load('controls/file-folder-picker/file-folder-picker-errors.internal.ts');
  assert.equal(errors.fileFolderPickerFailure(new Error('Selected file does not match the configured category.'), 'selection-invalid').code, 'selection-policy');
});

test('category hydration retains authoritative labels for incompatible saved values but excludes native preselection', async () => {
  const s = await state(); s.items.FileA.name = 'Renamed.exe';
  const value = { kind: 'file', tenantId: ids.tenant, endpoint, driveId: 'b!Drive', itemId: 'FileA', name: 'old.png' };
  const option = await s.load('results').resolveFileFolderPickerValue(value, s.scope, s.context, s.cache, { fileCategory: 'images' });
  assert.equal(option.text, 'Renamed.exe');
  assert.equal(s.load('results').fileFolderPickerOptionMatchesPolicy(option, { fileCategory: 'images' }), false);
  const config = s.load('config').compileFileFolderPickerConfig(s.scope, { fileCategory: 'images', sourceItems: [option] }, { origin: 'https://tenant.sharepoint.com', channelId: 'category' });
  assert.equal(config.selection.sourceItems, undefined);
  await assert.rejects(pick(s, ['FileA'], { fileCategory: 'images' }), /category/i);
});

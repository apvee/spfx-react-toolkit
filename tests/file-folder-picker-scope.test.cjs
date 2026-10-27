const test = require('node:test');
const assert = require('node:assert/strict');
const { ids, web, rootPath, setup } = require('./fixtures/file-folder-picker/metadata.cjs');
const resolve = (s, scope, initial) => s.load('scope').resolveFileFolderPickerScope(scope, initial, s.context, s.cache);
test('default collection resolves actual root independently of current subweb; remote collection starts at root', async () => {
  const s = setup(); const result = await resolve(s);
  assert.equal(result.siteCollectionId, ids.site); assert.equal(result.webUrl, web); assert.equal(result.initial.webUrl, web);
  assert.equal(s.calls.some(uri => uri.includes('/Subweb/_api')), false);
  assert.equal((await resolve(s, { kind: 'siteCollection', siteUrl: web + '/remote' })).webUrl, web);
});
test('library selectors resolve actual document library root and opaque drive/root IDs', async () => {
  for (const list of ['Documents', { kind: 'title', title: "Docs' Exact" }, { kind: 'id', id: ids.list }, { kind: 'url', serverRelativeUrl: rootPath }, { kind: 'path', webRelativePath: 'Shared Documents' }]) {
    const s = setup(); const result = await resolve(s, { kind: 'library', webUrl: web, list });
    assert.equal(result.library.rootFolderPath, rootPath); assert.equal(result.library.driveId, 'b!Drive'); assert.equal(result.library.rootItemId, 'Root'); assert.equal(result.library.rootFolderUniqueId, ids.root);
  }
  const s = setup({ response: uri => uri.includes('RootFolder') ? { BaseType: 0, RootFolder: {} } : undefined });
  await assert.rejects(resolve(s, { kind: 'library', webUrl: web, list: 'Ordinary' }), /document library/i);
});
test('decoded folder names survive REST boundaries; traversal, prefix lookalikes and missing folders fail closed', async () => {
  const s = setup(); const path = "Plans' # % À"; const result = await resolve(s, { kind: 'folder', webUrl: web, list: 'Documents', folder: path });
  assert.equal(result.folder.serverRelativePath, rootPath + '/' + path); assert.equal(result.folder.itemId, 'Folder');
  assert.ok(s.calls.some(uri => decodeURIComponent(uri).includes("Plans'' # % À")));
  for (const folder of ['../secret', 'a/../b', 'https://foreign.test/x', 'a\\b', rootPath + '2/Policy', '%2e%2e/secret']) await assert.rejects(resolve(setup(), { kind: 'folder', webUrl: web, list: 'Documents', folder }), /folder|path|scope/i);
  await assert.rejects(resolve(setup({ response: uri => uri.includes('GetFolderBy') ? new Error('missing folder') : undefined }), { kind: 'folder', webUrl: web, list: 'Documents', folder: 'Missing' }), /missing folder/);
});
test('source URLs reject foreign tenants and traversal before any request; initial locations cannot broaden scope', async () => {
  for (const siteUrl of ['http://tenant.sharepoint.com', 'https://other.sharepoint.com/sites/x', web + '/%2e%2e/x', web + '?x=1', web + '#x', 'https://tenant.sharepoint.com:443/sites/x']) {
    const s = setup(); await assert.rejects(resolve(s, { kind: 'siteCollection', siteUrl }), /URL|tenant|source/i); assert.equal(s.calls.length, 0);
  }
  await assert.rejects(resolve(setup(), undefined, { folder: 'A' }), /list/i);
  const s = setup({ response: uri => uri.includes('/other/_api/site/rootWeb') || uri.includes('/other/_api/web?') ? { Id: ids.web, Url: web + '/other' } : uri.includes('/other/_api/site?') ? { Id: ids.tenant, Url: web + '/other' } : undefined });
  await assert.rejects(resolve(s, { kind: 'library', webUrl: web, list: 'Documents' }, { webUrl: web + '/other', list: 'Documents' }), /collection/i);
  await assert.rejects(resolve(setup(), { kind: 'folder', webUrl: web, list: 'Documents', folder: 'Policies' }, { list: 'Documents', folder: 'Sibling' }), /folder|scope/i);
});
test('OneDrive discovers only the current personal site through a coalesced read and never provisions', async () => {
  const s = setup({ response: uri => uri.includes('GetMyProperties') ? { PersonalUrl: 'https://tenant-my.sharepoint.com/personal/user/' } : undefined });
  const [scope, again] = await Promise.all([resolve(s, { kind: 'oneDrive' }), resolve(s, { kind: 'oneDrive' })]);
  assert.equal(scope.kind, 'oneDrive'); assert.equal(scope.library.driveId, 'b!Drive');
  assert.deepEqual(scope.allowedResourceOrigins, ['https://tenant-my.sharepoint.com']);
  assert.equal(again.baseUrl, scope.baseUrl);
  assert.equal(s.calls.filter(uri => uri.includes('GetMyProperties')).length, 1);
  assert.equal(s.calls.some(uri => /graph|\/me\/|CreatePersonal/i.test(uri)), false);
  for (const PersonalUrl of ['', 'https://other-my.sharepoint.com/personal/user', web]) {
    const bad = setup({ response: uri => uri.includes('GetMyProperties') ? { PersonalUrl } : undefined });
    await assert.rejects(resolve(bad, { kind: 'oneDrive' }), /OneDrive|tenant|source/i);
    assert.equal(bad.calls.length, 1);
  }
});
test('root GUID mismatch and OneDrive ownership/path contradictions fail resolution', async () => {
  const brokenRoot = setup(); brokenRoot.items.Root.sharepointIds.listItemUniqueId = ids.folder;
  await assert.rejects(resolve(brokenRoot, { kind: 'library', webUrl: web, list: 'Documents' }), /root metadata/i);
  const owner = setup({ response: uri => new URL(uri).pathname.endsWith('/drive') ? { id: 'b!Drive', driveType: 'business', owner: { user: { email: 'other@tenant.test' } } } : undefined });
  await assert.rejects(resolve(owner, { kind: 'oneDrive' }), /owner/i);
  const escaped = setup({ response: uri => new URL(uri).pathname.endsWith('/root') ? { id: 'Root', folder: {}, root: {}, parentReference: { driveId: 'b!Drive' }, sharepointIds: { tenantId: ids.tenant, siteId: ids.site, webId: ids.web, listId: ids.list, listItemUniqueId: ids.root }, webUrl: 'https://tenant-my.sharepoint.com/personal/other/Documents' } : undefined });
  await assert.rejects(resolve(escaped, { kind: 'oneDrive' }), /root|path/i);
});

test('personal-domain metadata reads use only its resource token and share equivalent requests', async () => {
  const s = setup(); const resources = [], headers = [];
  const original = s.context.spClient.get;
  s.context.tokenProvider = { getToken: async resource => { resources.push(resource); return 'personal-resource-token'; } };
  s.context.spClient.get = async (uri, config, options) => {
    if (new URL(uri).hostname.endsWith('-my.sharepoint.com')) {
      headers.push(options.headers.Authorization);
      assert.equal(options.headers.Authorization, 'Bearer personal-resource-token');
    } else assert.equal(options.headers.Authorization, undefined);
    return original(uri, config, options);
  };
  await Promise.all([resolve(s, { kind: 'oneDrive' }), resolve(s, { kind: 'oneDrive' })]);
  assert.ok(headers.length > 0);
  assert.ok(resources.every(resource => resource === 'https://tenant-my.sharepoint.com'));
  assert.equal(s.calls.filter(uri => uri.includes('/_api/site?')).length, 1);
});

test('current-user profile email may differ from the login UPN while a different owner remains rejected', async () => {
  const response = uri => uri.includes('GetMyProperties') ? { PersonalUrl: 'https://tenant-my.sharepoint.com/personal/user', Email: 'mail.alias@tenant.test' } :
    new URL(uri).pathname.endsWith('/drive') ? { id: 'b!Drive', driveType: 'business', owner: { user: { email: 'mail.alias@tenant.test' } } } : undefined;
  const result = await resolve(setup({ response }), { kind: 'oneDrive' });
  assert.equal(result.library.driveId, 'b!Drive');
  const other = setup({ response: uri => uri.includes('GetMyProperties') ? { PersonalUrl: 'https://tenant-my.sharepoint.com/personal/user', Email: 'mail.alias@tenant.test' } :
    new URL(uri).pathname.endsWith('/drive') ? { id: 'b!Drive', driveType: 'business', owner: { user: { email: 'unrelated@tenant.test' } } } : undefined });
  await assert.rejects(resolve(other, { kind: 'oneDrive' }), /owner/i);
});

test('missing initial folder may fall back only to verified scope root and never on permissions failure', async () => {
 for (const status of [404,403]) {
  const s=setup({response:uri=>decodeURIComponent(uri).includes('/Missing')?Object.assign(new Error('Missing initial folder'),{status}):undefined});
  const resolver=s.load('scope').resolveFileFolderPickerScope;
  const source={kind:'library',webUrl:web,list:'Documents'};
  if(status===404) {
   await assert.rejects(resolver(source,{list:'Documents',folder:'Missing'},s.context,s.cache));
   const scope=await resolver(source,{list:'Documents',folder:'Missing'},s.context,s.cache,'scopeRoot');
   assert.equal(scope.initial.folderPath,undefined);
   assert.equal(scope.initial.listPath,rootPath);
  } else await assert.rejects(resolver(source,{list:'Documents',folder:'Missing'},s.context,s.cache,'scopeRoot'));
 }
});

test('actual personal-domain HTTP status distinguishes missing initial folder from denied access', async () => {
 for(const status of [404,403]) {
  const s=setup();const original=s.context.spClient.get;
  s.context.spClient.get=async(uri,config,options)=>new URL(uri).hostname==='tenant-my.sharepoint.com'&&decodeURIComponent(uri).includes('/Missing')?{ok:false,status,json:async()=>({})}:original(uri,config,options);
  const resolver=s.load('scope').resolveFileFolderPickerScope;
  if(status===404) {
   const result=await resolver({kind:'oneDrive'},{folder:'Missing'},s.context,s.cache,'scopeRoot');
   assert.equal(result.kind,'oneDrive');assert.equal(result.initial.folderPath,undefined);
   assert.equal(result.library.driveId,'b!Drive');
  }else await assert.rejects(resolver({kind:'oneDrive'},{folder:'Missing'},s.context,s.cache,'scopeRoot'),{code:'onedrive-unavailable'});
 }
});

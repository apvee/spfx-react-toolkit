const { createHarness } = require('../../services-test-harness.cjs');
const ids = { tenant: '11111111-1111-4111-8111-111111111111', site: '22222222-2222-4222-8222-222222222222', web: '33333333-3333-4333-8333-333333333333', list: '44444444-4444-4444-8444-444444444444', root: '55555555-5555-4555-8555-555555555555', folder: '66666666-6666-4666-8666-666666666666' };
const web = 'https://tenant.sharepoint.com/sites/HR';
const endpoint = web + '/_api/v2.0';
const rootPath = '/sites/HR/Shared Documents';
function metadata(id = 'FileA', patch = {}) {
  return { id, name: id + '.pdf', file: {}, parentReference: { driveId: 'b!Drive', id: 'Root' }, sharepointIds: { tenantId: ids.tenant, siteId: ids.site, webId: ids.web, listId: ids.list, listItemId: '7', listItemUniqueId: ids.folder }, webUrl: web + '/Shared%20Documents/' + id + '.pdf', ...patch };
}
function setup(options = {}) {
  const calls = [];
  const h = createHarness({ '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } } });
  const items = { Root: metadata('Root', { name: 'Shared Documents', file: undefined, folder: {}, root: {}, webUrl: web + '/Shared%20Documents', parentReference: { driveId: 'b!Drive' }, sharepointIds: { tenantId: ids.tenant, siteId: ids.site, webId: ids.web, listId: ids.list, listItemUniqueId: ids.root } }), Folder: metadata('Folder', { name: 'Policies', file: undefined, folder: {} }), FileA: metadata(), Child: metadata('Child', { parentReference: { driveId: 'b!Drive', id: 'Folder' } }), Outside: metadata('Outside') };
  const context = { tokenProvider: { getToken: async () => 'fixture-resource-token' }, owner: {}, tenantId: ids.tenant, userId: 'user@tenant.test', webUrl: web + '/Subweb', siteCollectionUrl: web, spClient: { get: async uri => {
    calls.push(uri);
    let result;
    if (options.response) result = options.response(uri, items);
    if (result === undefined) {
      const url = new URL(uri), route = decodeURIComponent(url.pathname);
      const responseWeb = url.hostname === 'tenant-my.sharepoint.com' ? url.origin + '/personal/user' : web;
      if (route.includes('GetMyProperties')) result = { PersonalUrl: 'https://tenant-my.sharepoint.com/personal/user' };
      else if (route.endsWith('/_api/site')) result = { Id: ids.site, Url: responseWeb };
      else if (route.endsWith('/_api/site/rootWeb')) result = { Id: ids.web, Url: responseWeb, ServerRelativeUrl: '/sites/HR' };
      else if (route.endsWith('/_api/web')) result = { Id: ids.web, Url: responseWeb, ServerRelativeUrl: '/sites/HR' };
      else if (route.includes('/lists/') || /\/lists\(/.test(route) || route.includes('GetListUsingPath')) {
        result = route.endsWith('/drive') ? { id: 'b!Drive', driveType: 'documentLibrary' } : { Id: ids.list, Title: 'Documents', BaseTemplate: 850, BaseType: 1, RootFolder: { ServerRelativeUrl: rootPath, UniqueId: ids.root } };
      } else if (route.endsWith('/drive')) result = { id: 'b!Drive', driveType: 'business', owner: { user: { email: context.userId } }, sharepointIds: { tenantId: ids.tenant, siteId: ids.site, webId: ids.web, listId: ids.list } };
      else if (route.includes('GetFolderByServerRelativePath')) { const match = /decodedUrl='(.*)'\)/i.exec(route); result = { ServerRelativeUrl: match?.[1].replace(/''/g, "'"), UniqueId: ids.folder }; }
      else if (route.endsWith('/root')) result = responseWeb === web ? items.Root : { ...items.Root, webUrl: responseWeb + '/Documents' };
      else if (route.includes('/root:/')) result = items.Folder;
      else if (route.includes('/items/')) result = items[route.split('/items/')[1]];
      else throw new Error('Unexpected fixture URI: ' + uri);
    }
    if (result instanceof Error) throw result;
    return { ok: true, status: 200, json: async () => result };
  } } };
  const cache = h.load('controls/pickers/picker-request-cache.internal.ts').createPickerRequestCache(context.owner);
  const load = name => h.load(`controls/file-folder-picker/file-folder-picker-${name}.internal.ts`);
  return { h, load, calls, context, cache, items, raw: id => ({ id, parentReference: { driveId: 'b!Drive' }, '@sharePoint.endpoint': endpoint }) };
}
module.exports = { ids, web, endpoint, rootPath, metadata, setup };

// Finite unavailable hosted form/message/token/REST boundaries. The native fixture
// owns the single loader/DOM/provider; parsing, validation and save stay production.
const assert = require('node:assert/strict');
const { ids, web, endpoint, metadata } = require('../file-folder-picker/metadata.cjs');
function createNativeFileFolderBoundary(t, { h, sdk, calls, unexpected, identity, phase, expectedItemTypes = 'all' }) {
  assert.ok(['files', 'folders', 'all'].includes(expectedItemTypes), 'explicit independent expected item policy');
  const posts = [], tokens = [], messages = [], ports = [];
  const items = { FileA: metadata(), Folder: metadata('Folder', { name: 'Policies', file: undefined, folder: {}, webUrl: web + '/Shared%20Documents/Policies' }) };
  function attempt(request, admit) {
    request.identity = identity; request.phase = phase(); calls.push(request);
    try { assert.equal(request.phase, 'native-hosted-browse'); return admit(); } catch (error) { unexpected.push({ request, message: error.message }); throw error; }
  }
  const provider = { getToken(resource, config) {
    const request = { kind: 'token', resource, config };
    return attempt(request, () => {
      assert.ok(['https://tenant.sharepoint.com', 'https://graph.microsoft.com'].includes(resource));
      assert.equal(config, undefined); tokens.push(request); return Promise.resolve('fixture-resource-token');
    });
  } };
  t.mock.method(window.HTMLFormElement.prototype, 'submit', function () {
    const request = { kind: 'hosted-post', method: this.method, url: this.action, target: this.target,
      body: Object.fromEntries([...this.elements].map(input => [input.name, input.value])) };
    attempt(request, () => {
      const url = new URL(request.url);
      assert.equal(url.origin, 'https://tenant.sharepoint.com'); assert.equal(url.pathname, '/sites/HR/_layouts/15/FilePicker.aspx');
      assert.equal(this.method, 'post'); assert.deepEqual(Object.keys(request.body), ['access_token']);
      assert.equal(request.body.access_token, 'fixture-resource-token');
      assert.deepEqual([...url.searchParams.keys()].sort(), ['filePicker', 'locale']); assert.equal(url.searchParams.get('locale'), 'en-US');
      const config = JSON.parse(url.searchParams.get('filePicker'));
      const frame = [...document.querySelectorAll('iframe')].find(frame => frame.name === this.target); assert.ok(frame);
      const { sourceItems, ...selection } = config.selection;
      assert.deepEqual(selection, { mode: 'single', maximumCount: 1, enablePersistence: false });
      if (sourceItems !== undefined) {
        assert.equal(sourceItems.length, 1);
        assert.ok(['FileA', 'Folder'].includes(sourceItems[0].id));
        assert.deepEqual(sourceItems[0], { id: sourceItems[0].id, parentReference: { driveId: 'b!Drive' }, '@sharePoint.endpoint': endpoint });
      }
      assert.deepEqual({ ...config, selection }, {
        sdk: '8.0', messaging: { origin: 'https://tenant.test', channelId: frame.name },
        entry: { sharePoint: { byPath: { web, fallbackToRoot: false } } }, theme: 'default', list: {},
        leftNav: { enabled: false, preset: 'current-site' }, tray: { prompt: 'selection-summary' }, search: { enabled: true },
        authentication: { enabled: true, tokens: { sharePoint: true, graph: true, substrate: false }, claimsChallenge: { enabled: true } },
        typesAndSources: { mode: expectedItemTypes, locations: { sharePoint: { byPath: { web } } }, access: { mode: 'read' },
          pivots: { oneDrive: false, recent: false, shared: false, sharedLibraries: false, myOrganization: false, site: false } },
        selection, commands: { pick: { action: 'select', select: { urls: { download: false } } },
          upload: { enabled: false }, createFolder: { enabled: false }, browseThisDevice: { enabled: false }, fromALink: { enabled: false }, switchAccount: { mode: 'none' }, manageAccounts: { mode: 'none' } },
      });
      posts.push({ frame, config, url: request.url });
    });
  });
  function route(request) {
    assert.ok(['native-hosted-browse', 'native-render'].includes(request.phase));
    assert.equal(request.identity, identity);
    assert.equal(request.kind, 'sp'); assert.equal(request.method, 'GET'); assert.equal(request.configuration, sdk.SPHttpClient.configurations.v1);
    const url = new URL(request.url); assert.equal(url.origin, 'https://tenant.sharepoint.com');
    assert.equal(request.body, undefined);
    assert.deepEqual(request.init.headers, { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'OData-Version': '' });
    const pathname = decodeURIComponent(url.pathname).replace(/^\/sites\/hr(?=\/)/, '/sites/HR');
    let value;
    if (pathname === '/sites/HR/_api/site') {
      assert.deepEqual([...url.searchParams], [['$select', 'Id,Url']]); value = { Id: ids.site, Url: web };
    } else if (pathname === '/sites/HR/_api/site/rootWeb') {
      assert.deepEqual([...url.searchParams], [['$select', 'Id,Url,ServerRelativeUrl']]); value = { Id: ids.web, Url: web, ServerRelativeUrl: '/sites/HR' };
    } else {
      assert.ok(Object.keys(items).some(id => pathname === '/sites/HR/_api/v2.0/drives/b!Drive/items/' + id), 'finite exact item route ' + pathname);
      assert.deepEqual([...url.searchParams], [['$select', 'id,name,file,folder,parentReference,sharepointIds,size,webUrl,remoteItem,root,deleted']]);
      value = items[pathname.slice(pathname.lastIndexOf('/') + 1)];
    }
    return new Response(JSON.stringify(value), { headers: { 'Content-Type': 'application/json' } });
  }
  function opening(frame = posts.at(-1)?.frame) {
    assert.ok(frame, 'actual native Stage iframe');
    const source = frame.contentWindow, channelId = frame.name, waits = new Map();
    const port = { messages: [], start() { this.started = true; }, close() { this.closed = true; }, postMessage(message) {
      this.messages.push(message); messages.push({ direction: 'out', channelId, message });
      if (message.type === 'result') { waits.get(message.id)?.(message); waits.delete(message.id); }
    } };
    ports.push(port);
    return { frame, port, channelId,
      initialize() {
        const data = { type: 'initialize', channelId }, request = { kind: 'hosted-initialize', origin: 'https://tenant.sharepoint.com', data, channelId };
        attempt(request, () => { messages.push({ direction: 'in', ...request }); h.act(() => { window.dispatchEvent(new window.MessageEvent('message', { origin: request.origin, source, data, ports: [port] })); }); });
      },
      command(id, data) {
        const message = { type: 'command', id, data }, request = { kind: 'hosted-command', channelId, message };
        attempt(request, () => {
          if (data.command === 'authenticate') assert.deepEqual(data, { command: 'authenticate', resource: 'https://tenant.sharepoint.com', type: 'SharePoint' });
          else if (data.command === 'pick') {
            assert.deepEqual(Object.keys(data), ['command', 'items']); assert.equal(data.items.length, 1);
            assert.ok(['FileA', 'Folder'].includes(data.items[0].id));
            assert.deepEqual(data.items[0], { id: data.items[0].id, parentReference: { driveId: 'b!Drive' }, '@sharePoint.endpoint': endpoint });
          } else assert.deepEqual(data, { command: 'close' });
          messages.push({ direction: 'in', ...request }); h.act(() => port.onmessage?.({ data: message }));
        });
      },
      result(id) { const existing = port.messages.find(message => message.type === 'result' && message.id === id); return existing ? Promise.resolve(existing) : new Promise(resolve => waits.set(id, resolve)); },
      raw: id => ({ id, parentReference: { driveId: 'b!Drive' }, '@sharePoint.endpoint': endpoint }),
    };
  }
  return { route, provider, posts, tokens, messages, ports, opening, items, ids, web, endpoint };
}
module.exports = { createNativeFileFolderBoundary };

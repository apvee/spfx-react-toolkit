// Finite native permission, named-AND and direct SPGroup oracle in the existing one-cache DOM harness.
const { createPickerHarness } = require('../picker-source-harness.cjs');
const { scenario } = require('./cards-description-list.cjs');
const fs = require('node:fs'), path = require('node:path'), assert = require('node:assert/strict');
const ReactDOM = require('react-dom');
const { Simulate } = require('react-dom/test-utils');
const { createDOMRenderer } = require('@griffel/core');
const { RendererProvider } = require('@griffel/react');
const fluent = require('@fluentui/react-components');
const owner = 'packages/spfx-react-toolkit/src/controls/sharepoint-permission-trim/sharepoint-permission-trim.tsx#SharePointPermissionTrim';
const single = 'Stage resource permissions and observe protected lifetimes', nested = 'Compose permission and Entra checks with nested AND';
const spGroupOwner = 'packages/spfx-react-toolkit/src/controls/sharepoint-group-trim/sharepoint-group-trim.tsx#SharePointGroupTrim';
const spGroup = 'Stage direct SharePoint group membership and observe protected lifetimes';
const listId = 'aaaaaaaa-1111-4111-8111-111111111111', groupId = 'bbbbbbbb-1111-4111-8111-111111111111';
const web = 'https://tenant.test/sites/team';
function deferred(record) { let resolve, reject; const promise = new Promise((a,b) => { resolve = a; reject = b; }); return Object.assign(record, { promise, resolve, reject, settled: false }); }
function createTrimsFixture(t, title = single, options = {}) {
  const compiled = options.compiled ?? process.env.TRIMS_IMPLEMENTATION === 'compiled';
  const toolkit = {}, context = {}, pageModule = {}, pageNative = { PageContext: { serviceKey: {} } };
  const keys = { page: pageNative.PageContext.serviceKey, sp: {}, graph: {}, theme: {} };
  const native = { SPHttpClient: { serviceKey: keys.sp, configurations: { v1: {} } }, MSGraphClientFactory: { serviceKey: keys.graph } };
  const seen = [], read = fs.readFileSync;
  t.mock.method(fs, 'readFileSync', (file, ...args) => { seen.push(String(file)); return read(file, ...args); });
  const h = createPickerHarness({ '@apvee/spfx-react-toolkit': toolkit, '@microsoft/sp-http': native,
    '@microsoft/sp-page-context': pageNative,
    '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 }, _SPKillSwitch: { isActivated: () => false } },
    '@microsoft/sp-component-base': { ThemeProvider: { serviceKey: keys.theme } },
    './useSPFxContext': context, './useSPFxPageContext': pageModule });
  pageNative.SPPermission = h.load(require.resolve('@microsoft/sp-page-context/lib-commonjs/SPPermission')).default;
  const root = path.resolve(__dirname, '../../../packages/spfx-react-toolkit', compiled ? 'lib' : 'src');
  const library = (stem, tsx = false) => h.load(path.join(root, stem + (compiled ? '.js' : tsx ? '.tsx' : '.ts')));
  Object.assign(context, library('hooks/useSPFxContext')); Object.assign(pageModule, library('hooks/useSPFxPageContext'));
  Object.assign(toolkit, context, pageModule, library('hooks/useSPFxSPHttpClient'), library('core/provider-webpart', true), library('controls/web-part-shell/web-part-shell', true));
  for (const slug of ['sharepoint-permission-trim', 'entra-group-trim', ...(title === spGroup ? ['sharepoint-group-trim'] : [])]) Object.assign(toolkit, library(`controls/${slug}/${slug}`, true));
  const calls = [], initializations = [], captures = [], nativeAttempts = [], protocolErrors = [];
  const selected = scenario(title, title === spGroup ? spGroupOwner : owner);
  // No live-network fallback is admitted by this fixture.
  t.mock.method(global, 'fetch', url => { protocolErrors.push('Unexpected network: ' + url); throw new Error('Unexpected network: ' + url); });
  function dispatch(record) {
    // Record the native boundary before admission: rejected routes are still real attempted I/O.
    nativeAttempts.push(Object.freeze({ ...record, login: record.page.user.loginName, web: record.page.web.absoluteUrl, headers: Object.freeze({ ...(record.init?.headers ?? {}) }), body: record.body ? Object.freeze({ groupIds: Object.freeze([...record.body.groupIds]) }) : undefined }));
    try { return admit(record); }
    catch (error) { protocolErrors.push(error.message); return Promise.reject(error); }
  }
  function admit(record) {
    const url = new URL(record.url, 'https://graph.microsoft.com');
    let phase;
    if (record.kind === 'sp') {
      assert.equal(url.origin, 'https://tenant.test');
      assert.ok(['/sites/team', '/sites/remote', '/sites/new'].some(site => url.pathname.startsWith(site + '/_api/web/')), 'finite selected web');
      assert.equal(record.method, 'GET'); assert.equal(record.configuration, native.SPHttpClient.configurations.v1);
      assert.deepEqual(record.init, { headers: { Accept: 'application/json;odata.metadata=none' } });
      const suffix = decodeURIComponent(url.pathname.slice(url.pathname.indexOf('/_api/web/') + 10));
      const list = "(?:lists/getByTitle\\('(?:Tasks|O''Brien)'\\)|lists\\(guid'aaaaaaaa-1111-4111-8111-111111111111'\\)|GetListUsingPath\\(DecodedUrl='/sites/(?:team|remote|new)/Lists/Tasks'\\))";
      if (title === spGroup) {
        const page = [...url.searchParams].filter(([key]) => key !== '$skiptoken');
        // URLSearchParams.get uses the native null sentinel for an absent query key.
        const continuation = url.searchParams.get('$skiptoken');
        assert.ok(continuation === null || ['names2', 'members2'].includes(continuation), 'finite group continuation');
        if (suffix === 'sitegroups') {
          assert.equal(page[0]?.[0], '$select'); assert.equal(page[0]?.[1], 'Id,Title');
          assert.equal(page[1]?.[0], '$filter');
          assert.ok(["Title eq 'Project Readers'", "Title eq 'Other Readers'", "Title eq 'O''Brien & Readers'"].includes(page[1]?.[1]), 'finite exact group name');
          assert.equal(page.length, 2); phase = 'group-name';
        } else if (/^sitegroups\((?:2|3)\)$/.test(suffix)) { assert.deepEqual(page, [['$select', 'Id,Title']]); assert.equal(continuation, null); phase = 'group-id'; }
        else if (suffix === 'currentuser') { assert.deepEqual(page, [['$select', 'Id']]); assert.equal(continuation, null); phase = 'group-currentuser'; }
        else { assert.match(suffix, /^sitegroups\((?:2|3)\)\/users$/); assert.deepEqual(page, [['$select', 'Id']]); phase = 'group-users'; }
      } else if (suffix === 'effectivebasepermissions' || new RegExp('^' + list + '(?:/items\\((?:1|7)\\))?/effectivebasepermissions$').test(suffix)) { assert.equal(url.search, ''); phase = 'mask'; }
      else if (new RegExp('^' + list + '$').test(suffix)) { assert.deepEqual([...url.searchParams], [['$select', 'Id']]); phase = 'list'; }
      else { assert.equal(suffix, "GetFolderByServerRelativePath(DecodedUrl='/sites/team/Shared Documents/Reports')/ListItemAllFields"); assert.deepEqual([...url.searchParams], [['$select', 'EffectiveBasePermissions,ParentList/Id'], ['$expand', 'ParentList']]); phase = 'folder'; }
    } else {
      assert.equal(url.origin, 'https://graph.microsoft.com');
      if (url.pathname === '/groups') { assert.equal(record.method, 'GET'); assert.deepEqual([...url.searchParams], [['$select', 'id,displayName'], ['$filter', "displayName eq 'Project Readers'"]]); assert.equal(record.body, undefined); phase = 'name'; }
      else { assert.equal(url.pathname, '/me/checkMemberGroups'); assert.equal(url.search, ''); assert.equal(record.method, 'POST'); assert.deepEqual(record.body, { groupIds: [groupId] }); phase = 'member'; }
    }
    const request = deferred({ ...record, capture: Object.freeze({ kind: record.kind, phase, method: record.method, url: record.url, query: Object.freeze([...url.searchParams].map(pair => Object.freeze(pair))), configuration: record.configuration, init: record.init ? Object.freeze({ headers: Object.freeze({ ...record.init.headers }) }) : undefined, body: record.body ? Object.freeze({ groupIds: Object.freeze([...record.body.groupIds]) }) : undefined, login: record.page.user.loginName, web: record.page.web.absoluteUrl, client: record.client, page: record.page, id: record.id }), phase, sequence: calls.length, headers: Object.freeze({ ...(record.headers ?? {}) }), body: record.body === undefined ? undefined : Object.freeze({ groupIds: Object.freeze([...record.body.groupIds]) }) });
    calls.push(request); return request.promise;
  }
  function host(id = 'trims-one', mask = options.mask) {
    const page = { web: { absoluteUrl: web, title: 'Team', permissions: mask === undefined ? undefined : new pageNative.SPPermission({ High: 0, Low: mask }) }, site: { absoluteUrl: web }, user: { id: 17, loginName: id, email: id + '@tenant.test' }, cultureInfo: {} };
    const sp = { get(url, configuration, init) { return dispatch({ kind: 'sp', id, client: sp, page, url, configuration, init, method: 'GET' }); } };
    const graph = { api(url) { return { get: () => dispatch({ kind: 'graph', id, client: graph, page, url, method: 'GET' }), post: body => dispatch({ kind: 'graph', id, client: graph, page, url, method: 'POST', body }) }; } };
    const factory = { getClient(version) { const request = deferred({ kind: 'initialization', id, page, factory, version, value: graph, capture: Object.freeze({ phase: 'graph-initialization', id, page, factory, version, login: page.user.loginName, web: page.web.absoluteUrl }), sequence: initializations.length }); initializations.push(request); try { assert.equal(version, '3'); } catch (error) { protocolErrors.push(error.message); request.reject(error); } return request.promise; } };
    const theme = { tryGetTheme: () => undefined, themeChangedEvent: { add() {}, remove() {} } };
    const services = new Map([[keys.page, page], [keys.sp, sp], [keys.graph, factory], [keys.theme, theme]]);
    const scope = { whenFinished: callback => callback(), consume(key) { if (!services.has(key)) { protocolErrors.push('Unexpected native service key'); throw new Error('Unexpected native service key'); } return services.get(key); } };
    const domElement = document.createElement('div');
    return { instanceId: id, context: { instanceId: id, domElement, serviceScope: scope, pageContext: page, sdks: {}, propertyPane: { refresh() {} } }, properties: { title: 'Trims', showTitle: false }, displayMode: 1, domElement, render() {}, native: { page, sp, graph, factory, services } };
  }
  let instance = host(), revision = options.revision ?? 7, regions;
  const Demo = h.load(selected.file).Demo;
  Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  const renderer = createDOMRenderer(document);
  const container = document.createElement('div'); document.body.appendChild(container);
  function Frame() { return h.React.createElement(Demo, { resetRevision: revision, render: options.fallback ? undefined : next => { regions = next; captures.push(next); return h.React.createElement('section', undefined, next.controls, next.commands, next.status, next.children, next.diagnostics); } }); }
  function tree(child, owner = instance) { return h.React.createElement(RendererProvider, { renderer }, h.React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme }, h.React.createElement(options.directProvider ? toolkit.SPFxWebPartProvider : toolkit.WebPartShell, { instance: owner, surface: 'section' }, child))); }
  function render() { h.act(() => { ReactDOM.render(tree(h.React.createElement(Frame)), container); }); }
  let closed = false; const peers = new Set();
  function unmount() { if (closed) return; h.act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove(); closed = true; }
  t.after(() => { for (const peer of peers) peer.unmount(); unmount(); h.close(); assert.equal(protocolErrors.length, options.expectedProtocolErrors ?? 0, 'native protocol failures cannot be hidden by a product Error branch: ' + protocolErrors.join(' | ')); }); render();
  return { h, toolkit, native, pageNative, keys, root, compiled, seen, calls, nativeAttempts, protocolErrors, initializations, captures, container, host, render, unmount,
    callbacks() {
      const div = container.querySelector('[data-trim-demo="protected"]'), button = container.querySelector('[data-trim-demo="increment"]'); assert.ok(div && button);
      const fiber = node => node[Object.keys(node).find(key => key.startsWith('__reactFiber$'))];
      const host = fiber(div); let owner = host;
      while (owner && owner.type?.name !== 'ProtectedContent') owner = owner.return;
      assert.ok(owner, 'actual protected component fiber');
      let hook = owner.memoizedState, lifetime; while (hook) { const value = hook.memoizedState?.current; if (value && typeof value === 'object' && 'retired' in value && 'active' in value) lifetime = value; hook = hook.next; } assert.ok(lifetime);
      return { lifetime, ref: host.ref, increment: fiber(button).memoizedProps.onClick, current: owner.memoizedProps.current, observe: owner.memoizedProps.observe, node: div };
    },
    get instance() { return instance; }, get regions() { return regions; }, get diagnostics() { return JSON.parse(container.querySelector('pre').textContent); },
    action() { return regions.commands.props.children.props.onClick; },
    click(label = 'Run') { const node = [...container.querySelectorAll('button')].find(n => n.textContent === label); assert.ok(node, label); h.act(() => Simulate.click(node)); },
    change(label, value) { const node = container.querySelector(`[aria-label="${label}"]`); assert.ok(node, label); h.act(() => { node.value = value; Simulate.change(node, { target: { value } }); }); },
    setRevision(value) { revision = value; render(); }, reset() { revision++; render(); }, replace(next) { instance = next; render(); },
    async flush(fn = () => {}) { await h.act(async () => { fn(); for (let i = 0; i < 35; i++) await Promise.resolve(); }); },
    async settle(request, payload, status = 200) { assert.equal(request.settled, false); request.settled = true; await this.flush(() => request.resolve(request.kind === 'sp' ? title === spGroup ? new Response(JSON.stringify(payload), { status, headers: { 'Content-Type': 'application/json' } }) : { ok: status < 400, status, json: async () => payload } : payload)); },
    async reject(request, error = new Error('native request failed')) { request.settled = true; await this.flush(() => request.reject(error)); },
    async initialize(request = initializations.at(-1), error) { request.settled = true; await this.flush(() => error ? request.reject(error) : request.resolve(request.value)); },
    mountPeer(owner = host('trims-peer')) { const view = h.mountComponent(() => tree(h.React.createElement(Demo), owner)); const peer = { ...view, instance: owner, click() { h.act(() => Simulate.click([...view.container.querySelectorAll('button')].find(n => n.textContent === 'Run'))); }, unmount() { view.unmount(); peers.delete(peer); } }; peers.add(peer); return peer; },
    origins() { const files = [...new Set(seen.filter(f => f.startsWith(path.dirname(root) + path.sep) && /\/(src|lib)\//.test(f)))];
      const stems = ['core/provider-webpart','core/context.internal','hooks/useSPFxContext','hooks/useSPFxPageContext','hooks/useSPFxSPHttpClient','hooks/useSPFxMSGraphClient','hooks/useSPFxServiceScope','controls/sharepoint-permission-trim/sharepoint-permission-trim','controls/entra-group-trim/entra-group-trim','controls/access-trim-children.internal','hooks/useSPFxCrossSitePermissions','hooks/useSPFxPermissions','hooks/useAccessRequest.internal','services/spfx-scoped-permissions.internal','services/spfx-entra-groups.internal','helpers/spfx-permissions.helpers', ...(title === spGroup ? ['controls/sharepoint-group-trim/sharepoint-group-trim', 'services/spfx-sharepoint-groups.internal'] : [])];
      for (const stem of stems) assert.ok(files.includes(path.join(root, stem + (compiled ? '.js' : /core\/(provider-webpart|context.internal)|controls\//.test(stem) ? '.tsx' : '.ts'))), 'observed matching implementation ' + stem);
      const ts = require('typescript');
      const inventory = JSON.parse(read(path.join(path.dirname(root), 'localization/resources.json'), 'utf8'));
      const assets = new Set(), dataRoles = [];
      for (const id of ['ApveeSpfxReactToolkitEditableTextStrings', 'ApveeSpfxReactToolkitWebPartShellStrings']) {
        assert.ok(h.nativeResources.requestedIds.includes(id), 'actually requested native dictionary');
        const resource = inventory.resources.find(value => value.id === id);
        const relative = id === 'ApveeSpfxReactToolkitEditableTextStrings' ? 'lib/controls/editable-text/loc/{locale}.js' : 'lib/controls/web-part-shell/loc/{locale}.js';
        assert.equal(resource.path, relative); const asset = path.join(path.dirname(root), relative.replace('{locale}', 'en-us'));
        const ast = ts.createSourceFile(asset, read(asset, 'utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
        assert.equal(ast.statements.length, 1); const expression = ast.statements[0].expression;
        assert.ok(ts.isCallExpression(expression) && expression.expression.text === 'define');
        assert.equal(expression.arguments.length, 2); assert.ok(ts.isArrayLiteralExpression(expression.arguments[0])); assert.equal(expression.arguments[0].elements.length, 0);
        const factory = expression.arguments[1]; assert.ok(ts.isFunctionExpression(factory)); assert.equal(factory.parameters.length, 0); assert.equal(factory.body.statements.length, 1);
        const returned = factory.body.statements[0]; assert.ok(ts.isReturnStatement(returned));
        const literal = node => ts.isStringLiteral(node) || ts.isObjectLiteralExpression(node) && node.properties.every(property => ts.isPropertyAssignment(property) && ts.isStringLiteral(property.name) && literal(property.initializer));
        assert.ok(literal(returned.expression), 'requested native asset is data only'); let leaves = 0; const count = node => { if (ts.isStringLiteral(node)) leaves++; else node.properties.forEach(property => count(property.initializer)); }; count(returned.expression); const bytes = read(asset); dataRoles.push({ id, asset, bytes: bytes.length, sha256: require('node:crypto').createHash('sha256').update(bytes).digest('hex'), literalLeaves: leaves, implementation: 0 }); assets.add(asset);
      }
      assert.deepEqual(files.filter(f => !assets.has(f) && f.includes(compiled ? '/src/' : '/lib/')), [], 'observed opposite implementation is forbidden'); assert.equal(toolkit.useSPFxContext, library('core/context.internal', true).useSPFxContext); return { files, dataRoles }; },
  };
}
module.exports = { createTrimsFixture, single, nested, owner, spGroupOwner, spGroup, listId, groupId, web };

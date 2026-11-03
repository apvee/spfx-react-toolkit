const { createPickerHarness } = require('../picker-source-harness.cjs');
const { scenario } = require('./cards-description-list.cjs');
const nativePicker = require('./react-pickers-native.cjs');
const path = require('node:path'), fs = require('node:fs'), assert = require('node:assert/strict');
const ReactDOM = require('react-dom');
const fluent = require('@fluentui/react-components');
// A finite extension of the accepted Site fixture; each row selects a real public factory.
const subjects = {
  Site: ['site', 'site'], List: ['list', 'list'], Field: ['field', 'field'], View: ['view', 'view'],
  ContentType: ['content-type', 'content type'], ListItem: ['list-item', 'item'], Principal: ['principal', 'principal'], Integrated: ['site', 'metadata and file cascade'], FileFolder: ['file-folder', 'file and folder'],
};
function createNativeReferenceFixture(t, { subject = 'Site', compiled = false, connected = true, revision = 7, hold } = {}) {
  assert.ok(Object.hasOwn(subjects, subject));
  const isIntegrated = subject === 'Integrated';
  if (isIntegrated) {
    const descriptor = Object.getOwnPropertyDescriptor(document, 'createRange');
    // JSDOM 15 has no text Range geometry; geometry itself has independent conformance tests.
    document.createRange = () => ({ selectNodeContents() {}, getClientRects() { return []; } });
    t.after(() => { if (descriptor) Object.defineProperty(document, 'createRange', descriptor); else delete document.createRange; });
  }
  const [stem, noun] = subjects[subject], title = isIntegrated ? 'Configure native metadata and file cascade' : `Select native ${noun} references`;
  const nativePicker = isIntegrated ? require('./react-pickers-native.cjs').createNativePickerRoutes({ origin: 'https://tenant.sharepoint.com', collectionPath: '/sites/hr', webPath: '/sites/hr/subweb', hubPath: '/sites/hub', collectionSiteId: require('../file-folder-picker/metadata.cjs').ids.site }) : require('./react-pickers-native.cjs');
  const isFileFolder = subject === 'FileFolder';
  const owner = isFileFolder ? 'packages/spfx-react-toolkit/src/property-pane/file-folder-picker/property-pane-file-folder-picker.tsx#PropertyPaneFileFolderPicker' : `packages/spfx-react-toolkit/src/property-pane/pickers/property-pane-${stem}-picker.ts#PropertyPane${isIntegrated ? 'Site' : subject}Picker`;
  Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  let dataRoles = [];
  const nativeCallbacks = [], fields = [], calls = [], unexpected = [], pending = [], changes = [], seen = [], keys = { page: {}, sp: {}, graph: {}, theme: {}, token: {} }, handlers = new Set();
  const context = {}, pageHook = {}, factories = {};
  const read = fs.readFileSync;
  t.mock.method(fs, 'readFileSync', (file, ...args) => { seen.push(String(file)); return read(file, ...args); });
  const native = { SPHttpClient: { serviceKey: keys.sp, configurations: { v1: {} } }, MSGraphClientFactory: { serviceKey: keys.graph }, AadTokenProviderFactory: { serviceKey: keys.token } };
  const pinned = require('../native-property-pane/pinned-components.cjs').pinnedComponents();
  const h = createPickerHarness({
    '@apvee/spfx-react-toolkit/property-pane/pickers': factories,
    '@apvee/spfx-react-toolkit/property-pane/file-folder-picker': factories,
    '@microsoft/sp-property-pane': pinned.fields,
    '@apvee/spfx-react-toolkit': { getWebPartShellPropertiesMetadata: () => ({}) },
    '@microsoft/sp-webpart-base': { BaseClientSideWebPart: class {} },
    SpFxReactToolkitTestWebPartStrings: { PropertyPaneDescription: 'Localized native host', BasicGroupName: 'Generic settings', DescriptionFieldLabel: 'Description' },
    './components/WebPartShellDemoHost': { WebPartShellDemoHost: () => undefined }, './components/SpFxReactToolkitTest': { default: () => undefined },
    '@microsoft/sp-core-library': { Version: { parse: value => value }, DisplayMode: { Read: 1, Edit: 2 }, _SPKillSwitch: { isActivated: () => false } },
    '@microsoft/sp-page-context': { PageContext: { serviceKey: keys.page } },
    '@microsoft/sp-component-base': { ThemeProvider: { serviceKey: keys.theme } },
    '@microsoft/sp-http': native,
    './useSPFxContext': context, './useSPFxPageContext': pageHook,
  });
  const root = path.resolve(__dirname, '../../../packages/spfx-react-toolkit', compiled ? 'lib' : 'src');
  const library = (file, tsx = false) => h.load(path.join(root, file + (compiled ? '.js' : tsx ? '.tsx' : '.ts')));
  Object.assign(context, library('hooks/useSPFxContext'));
  Object.assign(pageHook, library('hooks/useSPFxPageContext'));
  for (const name of isIntegrated ? ['Site', 'List', 'Field', 'View', 'ContentType', 'ListItem', 'Principal', 'FileFolder'] : [subject]) {
    const file = name === 'FileFolder' ? 'property-pane/file-folder-picker/property-pane-file-folder-picker' : 'property-pane/pickers/property-pane-' + subjects[name][0] + '-picker';
    const actualFactory = library(file, name === 'FileFolder')['PropertyPane' + name + 'Picker'];
    factories['PropertyPane' + name + 'Picker'] = (...args) => {
      const actual = actualFactory(...args);
      if (isIntegrated) {
        const render = actual.properties.onRender;
        actual.properties.onRender = (element, context, callback) => render(element, context, callback === undefined ? undefined : (...values) => {
          nativeCallbacks.push({ name, values }); callback(...values);
        });
      }
      fields.push({ name, args, actual }); return actual;
    };
  }
  const page = { web: { absoluteUrl: 'https://tenant.test/sites/root/sub' }, site: { absoluteUrl: isIntegrated ? 'https://tenant.sharepoint.com/sites/hr' : 'https://tenant.test/sites/root' },
    user: { loginName: 'current@tenant.test', displayName: 'Current', email: 'current@tenant.test' }, cultureInfo: { isRightToLeft: false }, aadInfo: { tenantId: { toString: () => '11111111-1111-4111-8111-111111111111' } } };
  if (isFileFolder || isIntegrated) {
    const { web, ids } = require('../file-folder-picker/metadata.cjs');
    page.web.absoluteUrl = web + '/Subweb'; page.site.absoluteUrl = web;
    page.aadInfo.tenantId.toString = () => ids.tenant; page.cultureInfo.currentUICultureName = 'en-US';
  }
  let phase = 'selection';
  const hosted = isFileFolder || isIntegrated ? require('./native-file-folder-boundary.cjs').createNativeFileFolderBoundary(t, { h, sdk: native, calls, unexpected, identity: 'native-' + subject + '-' + compiled, phase: () => phase, expectedItemTypes: isIntegrated ? 'files' : 'all' }) : undefined;
  function route(request) {
    if (hosted && (!isIntegrated || new URL(request.url).pathname.toLowerCase().includes('/_api/v2.0') || new URL(request.url).searchParams.get('$select') === 'Id,Url,ServerRelativeUrl' || new URL(request.url).pathname.toLowerCase().includes('/getfilebyserverrelativepath') || new URL(request.url).pathname.toLowerCase().includes('/getfolderbyserverrelativepath'))) return hosted.route(request);
    const url = new URL(request.url), endpoint = decodeURIComponent(url.pathname).toLowerCase();
    assert.equal(url.origin, isIntegrated ? 'https://tenant.sharepoint.com' : 'https://tenant.test');
    if (isIntegrated) {
      const prefix = ['/sites/hr/_api/', '/sites/hr/subweb/_api/', '/sites/hub/_api/'];
      assert.ok(prefix.some(path => endpoint.startsWith(path)), 'finite integrated context paths');
      if (request.kind === 'sp') {
        assert.deepEqual(request.init.headers, { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'OData-Version': '' });
        if (request.method === 'GET') assert.equal(request.body, undefined);
      }
    }
    const result = nativePicker.route(request);
    if (result !== undefined) return result;
    if ((subject === 'Site' || isIntegrated) && request.kind === 'sp' && (endpoint === (isIntegrated ? '/sites/hr/_api/site' : '/sites/root/_api/site') || isIntegrated && endpoint === '/sites/hr/subweb/_api/site') && url.searchParams.get('$select') === 'Id,Url') {
      assert.equal(request.method, 'GET'); assert.deepEqual([...url.searchParams], [['$select', 'Id,Url']]);
      return new Response(JSON.stringify({ Id: nativePicker.siteId, Url: isIntegrated ? 'https://tenant.sharepoint.com/sites/hr' : 'https://tenant.test/sites/root' }), { headers: { 'Content-Type': 'application/json' } });
    }
    if ((subject !== 'Principal' && !isIntegrated) || request.kind !== 'sp') return undefined;
    assert.deepEqual(request.init.headers, { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'OData-Version': '' });
    if (endpoint === (isIntegrated ? '/sites/hr/subweb/_api/sp.ui.applicationpages.clientpeoplepickerwebserviceinterface.clientpeoplepickersearchuser' : '/sites/root/sub/_api/sp.ui.applicationpages.clientpeoplepickerwebserviceinterface.clientpeoplepickersearchuser')) {
      assert.equal(request.method, 'POST'); assert.equal(url.search, '');
      const body = JSON.parse(request.init.body); assert.deepEqual(Object.keys(body), ['queryParams']);
      const q = body.queryParams;
      assert.deepEqual(q, { AllowEmailAddresses: false, AllowMultipleEntities: false, AllUrlZones: false, MaximumEntitySuggestions: 50, PrincipalSource: 15, PrincipalType: isIntegrated ? 15 : 1, QueryString: q.QueryString, SharePointGroupID: 0, UseSubstrateSearch: false });
      assert.ok(['Alice', 'alice@tenant.test'].includes(q.QueryString));
      return new Response(JSON.stringify({ ClientPeoplePickerSearchUser: JSON.stringify([{ Key: 'alice@tenant.test', DisplayText: 'Alice', IsResolved: true, EntityType: 'User', EntityData: { Email: 'alice@tenant.test', UserPrincipalName: 'alice@tenant.test' } }]) }), { headers: { 'Content-Type': 'application/json' } });
    }
    if (endpoint === (isIntegrated ? '/sites/hr/subweb/_api/site' : '/sites/root/sub/_api/site')) {
      assert.equal(request.method, 'GET'); assert.deepEqual([...url.searchParams], [['$select', 'Url']]);
      return new Response(JSON.stringify({ Url: isIntegrated ? 'https://tenant.sharepoint.com/sites/hr' : 'https://tenant.test/sites/root' }), { headers: { 'Content-Type': 'application/json' } });
    }
    if (endpoint === (isIntegrated ? '/sites/hr/subweb/_api/web/siteusers/getbyloginname(@login)' : '/sites/root/sub/_api/web/siteusers/getbyloginname(@login)')) {
      assert.equal(request.method, 'GET'); assert.deepEqual(Object.fromEntries(url.searchParams), { '$select': 'Id,LoginName,Title,Email,PrincipalType', '@login': "'alice@tenant.test'" });
      return new Response(JSON.stringify({ Id: 42, LoginName: 'alice@tenant.test', Title: 'Alice', Email: 'alice@tenant.test', PrincipalType: 1 }), { headers: { 'Content-Type': 'application/json' } });
    }
    return undefined;
  }
  function dispatch(request) {
    // Record every attempted route before admission; a UI error cannot mask rejection.
    request.phase = phase; request.identity = 'native-' + subject + '-' + compiled;
    if (!isFileFolder && !isIntegrated) request.query = Object.fromEntries(new URL(request.url).searchParams);
    request.body = request.init?.body; calls.push(request);
    let response;
    try {
      if (isFileFolder || isIntegrated) request.query = Object.fromEntries(new URL(request.url).searchParams);
      if (request.kind === 'sp') assert.equal(request.configuration, native.SPHttpClient.configurations.v1);
      response = route(request); assert.ok(response, 'bounded native transport ' + request.method + ' ' + request.url);
    } catch (error) { unexpected.push({ request, message: error.message }); throw error; }
    if (hold?.(request)) return new Promise((resolve, reject) => pending.push({ request, response, resolve, reject }));
    return Promise.resolve(response);
  }
  const sp = {
    get: (url, configuration, init) => dispatch({ kind: 'sp', method: 'GET', url, configuration, init }),
    post: (url, configuration, init) => dispatch({ kind: 'sp', method: 'POST', url, configuration, init }),
  };
  t.mock.method(global, 'fetch', (url, init) => dispatch({ kind: 'pnp', url: String(url), init, method: init?.method ?? 'GET' }));
  const theme = { tryGetTheme: () => undefined, themeChangedEvent: { add(_owner, callback) { handlers.add(callback); }, remove(_owner, callback) { handlers.delete(callback); } } };
  const scope = { whenFinished: callback => callback(), consume(key) { if (key === keys.page) return page; if (key === keys.sp) return sp; if (key === keys.theme) return theme; if (key === keys.token && hosted) return { getTokenProvider: async () => hosted.provider }; if (key === keys.graph) return { getClient: async () => { throw new Error('Default native Site fixture needs no Graph'); } }; throw new Error('Unexpected Site service'); } };
  const savedWaiters = [];
  const notifySaved = value => { resolveSaved(value); savedWaiters.splice(0).forEach(resolve => resolve(value)); };
  let resolveSaved;
  const saved = new Promise(resolve => { resolveSaved = resolve; });
  let open = false, bag = {}, lastSdkValues = {}, regions, portConnected = connected, refreshes = 0, automaticRefreshes = 0, contentRefreshes = 0;
  const instanceShape = { context: { instanceId: 'native-' + subject + '-' + compiled, manifest: { id: 'fixture' }, pageContext: page, serviceScope: scope, spHttpClient: sp, propertyPane: { isPropertyPaneOpen: () => open, refresh: () => { refreshes++; } } },
    get properties() { return bag; }, displayMode: 2, domElement: document.createElement('div'), render() {} };
  const appBase = path.resolve(__dirname, '../../../apps/spfx-react-toolkit-test', compiled ? 'lib' : 'src', 'webparts/spFxReactToolkitTest');
  const instance = isIntegrated ? Object.assign(new (h.load(path.join(appBase, 'SpFxReactToolkitTestWebPart.' + (compiled ? 'js' : 'ts'))).default)(), instanceShape) : instanceShape;
  if (isIntegrated) delete instance.render;
  Object.defineProperty(instance, 'properties', { configurable: true, get: () => bag });
  Object.defineProperty(instance, '_properties', { configurable: true, get: () => bag });
  const adapterFile = path.resolve(__dirname, '../../../apps/spfx-react-toolkit-test', compiled ? 'lib' : 'src', 'webparts/spFxReactToolkitTest/components/shared/NativePropertyPaneDemoHost.' + (compiled ? 'js' : 'ts'));
  const Host = h.load(adapterFile).NativePropertyPaneDemoHost;
  const adapter = isIntegrated ? instance.nativeHost : new Host({ instance, readProperties: () => bag, isOpen: () => open, refresh: () => { refreshes++; } });
  const nativeFieldChanged = isIntegrated ? require('../native-property-pane/sdk-schedule.cjs').pinnedFieldDispatch(instance, { onAutomaticRefresh: () => automaticRefreshes++, onContentRefresh: () => contentRefreshes++ }) : undefined;
  const selected = scenario(title, owner), Demo = h.load(selected.file).Demo;
  const container = document.createElement('div'); document.body.appendChild(container);
  const nativeDom = document.createElement('div'); document.body.appendChild(nativeDom);
  function render() { h.act(() => { ReactDOM.render(h.React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme }, h.React.createElement(Demo, {
    resetRevision: revision, ...(portConnected ? { propertyPane: adapter.port } : {}), render: next => { regions = next; return h.React.createElement('section', undefined, next.controls, next.commands, next.status, next.children, next.diagnostics); }
  })), container); }); }
  render();
  const diagnosticState = () => ({
    open, mounted: adapter.mounted.size, witnessed: adapter.witnesses.size,
    produced: adapter.produced !== undefined, nativeView: adapter.nativeView !== undefined,
    fields: fields.length, requests: calls.length, unexpected: unexpected.length,
    pending: pending.length, savedWaiters: savedWaiters.length, handlers: handlers.size,
    hostedPorts: hosted?.ports.length ?? 0, closedPorts: hosted?.ports.filter(port => port.closed).length ?? 0,
  });
  t.after(async () => {
    h.act(() => { adapter.dispose(); ReactDOM.unmountComponentAtNode(nativeDom); ReactDOM.unmountComponentAtNode(container); });
    // Drain retired native continuations while this fixture still owns its transport.
    await h.act(async () => { for (let i = 0; i < 60; i++) await Promise.resolve(); await new Promise(resolve => setImmediate(resolve)); });
    nativeDom.remove(); container.remove(); h.close();
    if (hosted) assert.ok(hosted.ports.every(port => port.closed), 'all owned hosted ports close on actual native disposal');
    assert.deepEqual(unexpected, [], 'all attempted transport routes were finitely admitted');
  });
  return { h, root, library, adapter, saved, hosted, pinned, adapterFile, get dataRoles() { return dataRoles; }, instance, container, nativeDom, nativeCallbacks, diagnosticState, fields, calls, changes, pending, unexpected, handlers, seen, selected,
    get bag() { return bag; }, set bag(value) { bag = value; lastSdkValues = { ...value }; notifySaved(value); }, nextSaved() { return new Promise(resolve => savedWaiters.push(resolve)); }, get regions() { return regions; }, get refreshes() { return refreshes; }, get automaticRefreshes() { return automaticRefreshes; }, get contentRefreshes() { return contentRefreshes; },
    setRevision(value) { revision = value; render(); }, disconnect() { portConnected = false; render(); },
    action(label) { const children = regions.commands?.props.children; const node = (Array.isArray(children) ? children : [children]).find(node => node?.props.children === label); assert.ok(node, label); return node.props.onClick; },
    click(label) { const button = [...container.querySelectorAll('button')].find(node => node.textContent === label); assert.ok(button, label); h.act(() => button.click()); },
    change(label, value) { const input = container.querySelector('[aria-label="' + label + '"]'); assert.ok(input, label); h.act(() => { require('react-dom/test-utils').Simulate.change(input, { target: { value } }); }); },
    builtin(label, value) { const node = nativeDom.querySelector('[aria-label="' + label + '"]'); assert.ok(node, label); phase = 'native-builtin'; h.act(() => require('react-dom/test-utils').Simulate.change(node, { target: { value, checked: value } })); },
    browse() { phase = 'native-hosted-browse'; const button = nativeDom.querySelector('button[aria-haspopup="dialog"]'); assert.ok(button, 'actual native Browse'); assert.equal(button.disabled, false); h.act(() => button.click()); },
    open() { phase = 'native-menu'; const trigger = nativeDom.querySelector('[role="combobox"]'); assert.ok(trigger); h.act(() => { trigger.focus(); trigger.click(); }); },
    search(value) { phase = 'native-search'; const input = document.querySelector('.fui-SearchBox input') ?? nativeDom.querySelector('input[role="combobox"]'); assert.ok(input); h.act(() => { input.value = value; require('react-dom/test-utils').Simulate.change(input, { target: { value } }); }); },
    choose(label) { phase = 'native-confirm'; const option = [...document.querySelectorAll('[role="option"],[role="menuitemcheckbox"]')].find(node => node.textContent.includes(label)); assert.ok(option, label + ': ' + nativeDom.textContent + ' attempts=' + JSON.stringify(calls.map(({method,url})=>({method,url})))); h.act(() => { option.click(); }); },
    async wait(ms = 360) { await h.act(async () => { await new Promise(resolve => setTimeout(resolve, ms)); }); },
    async settle(item) { await h.act(async () => { item.resolve(item.response); for (let i = 0; i < 30; i++) await Promise.resolve(); }); },
    getter() { phase = 'getter'; open = true; return isIntegrated ? instance.getPropertyPaneConfiguration() : adapter.getConfiguration(() => ({ pages: [] })); },
    mount(field, callback = (name, values, reactive) => { const old = lastSdkValues[name]; changes.push({ name, old, values, reactive }); bag = { ...bag, [name]: values }; lastSdkValues = { ...bag }; adapter.fieldChanged(name, old, values); notifySaved(bag); }) { phase = 'native-render'; h.act(() => field.properties.onRender(nativeDom, undefined, callback)); adapter.rendered(); },
    mountPane(configuration) {
      phase = 'native-render'; const { PropertyPane } = pinned;
      h.act(() => { ReactDOM.render(h.React.createElement(PropertyPane, { data: {
        webPartId: instance.context.instanceId, title: 'Native references', configuration: { ...configuration, hideTitleBar: true }, properties: isIntegrated ? require('lodash').cloneDeep(bag) : bag, isReactive: true,
        onRendered: () => { assert.ok(adapter.witnesses.size, 'real native child precedes pinned parent'); if (isIntegrated) instance.onPropertyPaneRendered(); else adapter.rendered(); },
        onPropertyPaneFieldChanged: (name, values, fieldType) => { const old = lastSdkValues[name]; changes.push({ name, old, values, fieldType }); if (nativeFieldChanged) nativeFieldChanged(name, values, fieldType); else { bag = { ...bag, [name]: values }; adapter.fieldChanged(name, old, values); } lastSdkValues = { ...bag }; notifySaved(bag); },
        onConfigurationEvent: () => {},
      } }), nativeDom); });
    },
    closePane() { h.act(() => { ReactDOM.unmountComponentAtNode(nativeDom); }); open = false; if (isIntegrated) instance.onPropertyPaneConfigurationComplete(); else adapter.configurationComplete(); },
    close(field) { h.act(() => field.properties.onDispose(nativeDom)); open = false; if (isIntegrated) instance.onPropertyPaneConfigurationComplete(); else adapter.configurationComplete(); },
    async flush() { await h.act(async () => { for (let i = 0; i < 30; i++) await Promise.resolve(); await new Promise(resolve => setImmediate(resolve)); }); },
    origins() {
      const files = [...new Set(seen.filter(file => file.startsWith(path.dirname(root) + '/') && /\/(src|lib)\//.test(file)))];
      dataRoles = [];
      const data = new Set([...new Set(h.nativeResources.requestedIds)].flatMap(id => {
        const item = JSON.parse(read(path.join(path.dirname(root), 'localization/resources.json'), 'utf8')).resources.find(item => item.id === id);
        assert.ok(item);
        if (isFileFolder || isIntegrated) {
          const paths = { ApveeSpfxReactToolkitFileFolderPickerStrings: 'lib/controls/file-folder-picker/loc/{locale}.js', ApveeSpfxReactToolkitStageStrings: 'lib/controls/stage/loc/{locale}.js', ApveeSpfxReactToolkitPickersStrings: 'lib/controls/pickers/loc/{locale}.js' };
          assert.equal(item.path, paths[id], 'exact actually requested declared FileFolder DATA role');
          const config = JSON.parse(read(path.resolve(__dirname, '../../../apps/spfx-react-toolkit-test/config/config.json'), 'utf8'));
          assert.equal(config.localizedResources[id], '../../node_modules/@apvee/spfx-react-toolkit/' + item.path);
        }
        const asset = path.join(path.dirname(root), item.path.replace('{locale}', 'en-us'));
        const bytes = read(asset), ts = require('typescript');
        const source = ts.createSourceFile(asset, bytes.toString(), ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
        assert.equal(source.statements.length, 1); assert.ok(ts.isExpressionStatement(source.statements[0]));
        const call = source.statements[0].expression; assert.ok(ts.isCallExpression(call) && ts.isIdentifier(call.expression) && call.expression.text === 'define');
        assert.equal(call.arguments.length, 2); assert.ok(ts.isArrayLiteralExpression(call.arguments[0]) && call.arguments[0].elements.length === 0);
        const factory = call.arguments[1]; assert.ok(ts.isFunctionExpression(factory) && factory.parameters.length === 0 && factory.body.statements.length === 1);
        const returned = factory.body.statements[0]; assert.ok(ts.isReturnStatement(returned));
        let literalLeaves = 0;
        function literal(node) { if (ts.isStringLiteral(node)) { literalLeaves++; return true; } return ts.isObjectLiteralExpression(node) && node.properties.every(property => ts.isPropertyAssignment(property) && ts.isStringLiteral(property.name) && literal(property.initializer)); }
        assert.ok(literal(returned.expression), 'only the actually requested declared native string DATA dictionary');
        dataRoles.push({ id, path: asset, bytes: bytes.length, sha256: require('node:crypto').createHash('sha256').update(bytes).digest('hex'), literalLeaves });
        return asset;
      }));
      assert.equal(files.filter(file => !data.has(file)).some(file => file.includes(compiled ? '/src/' : '/lib/')), false);
      const requiredStems = isFileFolder ? ['property-pane/file-folder-picker/property-pane-file-folder-picker', 'property-pane/file-folder-picker/property-pane-file-folder-picker-value.internal', 'property-pane/pickers/property-pane-picker-host.internal', 'controls/file-folder-picker/file-folder-picker', 'controls/file-folder-picker/useFileFolderPicker.internal', 'controls/file-folder-picker/file-folder-picker-scope.internal', 'controls/file-folder-picker/file-folder-picker-results.internal', 'controls/file-folder-picker/file-folder-picker-session.internal', 'controls/file-folder-picker/file-folder-picker-protocol.internal', 'controls/file-folder-picker/file-folder-picker-navigation.internal', 'controls/pickers/picker-request-cache.internal', 'controls/stage/stage', 'core/provider-webpart', 'core/context.internal'] : ['property-pane/pickers/property-pane-' + stem + '-picker', 'property-pane/pickers/property-pane-picker-factory.internal', 'property-pane/pickers/property-pane-picker.internal', 'controls/pickers/' + stem + '-picker', 'core/provider-webpart', 'core/context.internal'];
      for (const requiredStem of requiredStems)
        assert.ok(files.some(file => file === path.join(root, requiredStem + '.js') || file === path.join(root, requiredStem + '.ts') || file === path.join(root, requiredStem + '.tsx')), 'actual matching origin ' + requiredStem);
      assert.ok(seen.includes(adapterFile), 'actual matching private adapter implementation');
      assert.equal(seen.some(file => file.includes('/apps/spfx-react-toolkit-test/' + (compiled ? 'src' : 'lib') + '/') && file.includes('NativePropertyPaneDemoHost.')), false, 'no opposite private adapter');
      return files;
    }
  };
}
const createNativeSiteFixture = (t, options) => createNativeReferenceFixture(t, options);
module.exports = { createNativeSiteFixture, createNativeReferenceFixture };

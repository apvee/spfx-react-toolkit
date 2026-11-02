const { createPickerHarness } = require('../picker-source-harness.cjs');
const { scenario } = require('./cards-description-list.cjs');
const nativePicker = require('./react-pickers-native.cjs');
const path = require('node:path'), fs = require('node:fs'), assert = require('node:assert/strict');
const ReactDOM = require('react-dom');
const fluent = require('@fluentui/react-components');
const title = 'Select native site references';
const owner = 'packages/spfx-react-toolkit/src/property-pane/pickers/property-pane-site-picker.ts#PropertyPaneSitePicker';
function createNativeSiteFixture(t, { compiled = false, connected = true, revision = 7 } = {}) {
  Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  let dataRoles = [];
  const fields = [], calls = [], seen = [], keys = { page: {}, sp: {}, graph: {}, theme: {} }, handlers = new Set();
  const context = {}, pageHook = {}, factories = {};
  const read = fs.readFileSync;
  t.mock.method(fs, 'readFileSync', (file, ...args) => { seen.push(String(file)); return read(file, ...args); });
  const native = { SPHttpClient: { serviceKey: keys.sp, configurations: { v1: {} } }, MSGraphClientFactory: { serviceKey: keys.graph } };
  const h = createPickerHarness({
    '@apvee/spfx-react-toolkit/property-pane/pickers': factories,
    '@microsoft/sp-property-pane': { PropertyPaneFieldType: { Custom: 1 } },
    '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 }, _SPKillSwitch: { isActivated: () => false } },
    '@microsoft/sp-page-context': { PageContext: { serviceKey: keys.page } },
    '@microsoft/sp-component-base': { ThemeProvider: { serviceKey: keys.theme } },
    '@microsoft/sp-http': native,
    './useSPFxContext': context, './useSPFxPageContext': pageHook,
  });
  const root = path.resolve(__dirname, '../../../packages/spfx-react-toolkit', compiled ? 'lib' : 'src');
  const library = (file, tsx = false) => h.load(path.join(root, file + (compiled ? '.js' : tsx ? '.tsx' : '.ts')));
  Object.assign(context, library('hooks/useSPFxContext'));
  Object.assign(pageHook, library('hooks/useSPFxPageContext'));
  const actualFactory = library('property-pane/pickers/property-pane-site-picker').PropertyPaneSitePicker;
  factories.PropertyPaneSitePicker = (...args) => { const actual = actualFactory(...args); fields.push({ args, actual }); return actual; };
  const page = { web: { absoluteUrl: 'https://tenant.test/sites/root/sub' }, site: { absoluteUrl: 'https://tenant.test/sites/root' },
    user: { loginName: 'current@tenant.test', displayName: 'Current', email: 'current@tenant.test' }, cultureInfo: { isRightToLeft: false }, aadInfo: { tenantId: { toString: () => '11111111-1111-4111-8111-111111111111' } } };
  function dispatch(url, configuration, init) {
    const request = { kind: 'sp', method: 'GET', url, configuration, init }; calls.push(request);
    assert.equal(configuration, native.SPHttpClient.configurations.v1);
    const response = nativePicker.route(request); assert.ok(response, 'bounded Site native transport ' + url); return Promise.resolve(response);
  }
  const sp = { get: dispatch, post() { throw new Error('Site fixture forbids tenant writes'); } };
  t.mock.method(global, 'fetch', (url, init) => {
    const request = { kind: 'pnp', url: String(url), init, method: init?.method ?? 'GET' };
    calls.push(request);
    const response = nativePicker.route(request);
    assert.ok(response, 'bounded native Site Search verification only');
    return Promise.resolve(response);
  });
  const theme = { tryGetTheme: () => undefined, themeChangedEvent: { add(_owner, callback) { handlers.add(callback); }, remove(_owner, callback) { handlers.delete(callback); } } };
  const scope = { whenFinished: callback => callback(), consume(key) { if (key === keys.page) return page; if (key === keys.sp) return sp; if (key === keys.theme) return theme; if (key === keys.graph) return { getClient: async () => { throw new Error('Default native Site fixture needs no Graph'); } }; throw new Error('Unexpected Site service'); } };
  let resolveSaved;
  const saved = new Promise(resolve => { resolveSaved = resolve; });
  let open = false, bag = {}, regions, portConnected = connected, refreshes = 0;
  const instance = { context: { instanceId: 'native-site-' + compiled, pageContext: page, serviceScope: scope, spHttpClient: sp, propertyPane: { isPropertyPaneOpen: () => open, refresh: () => { refreshes++; } } },
    get properties() { return bag; }, displayMode: 2, domElement: document.createElement('div'), render() {} };
  const adapterFile = path.resolve(__dirname, '../../../apps/spfx-react-toolkit-test', compiled ? 'lib' : 'src', 'webparts/spFxReactToolkitTest/components/shared/NativePropertyPaneDemoHost.' + (compiled ? 'js' : 'ts'));
  const Host = h.load(adapterFile).NativePropertyPaneDemoHost;
  const adapter = new Host({ instance, readProperties: () => bag, isOpen: () => open, refresh: () => { refreshes++; } });
  const selected = scenario(title, owner), Demo = h.load(selected.file).Demo;
  const container = document.createElement('div'); document.body.appendChild(container);
  const nativeDom = document.createElement('div'); document.body.appendChild(nativeDom);
  function render() { h.act(() => { ReactDOM.render(h.React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme }, h.React.createElement(Demo, {
    resetRevision: revision, ...(portConnected ? { propertyPane: adapter.port } : {}), render: next => { regions = next; return h.React.createElement('section', undefined, next.controls, next.commands, next.status, next.children, next.diagnostics); }
  })), container); }); }
  render();
  t.after(() => { h.act(() => { adapter.dispose(); ReactDOM.unmountComponentAtNode(container); }); nativeDom.remove(); container.remove(); h.close(); });
  return { h, root, library, adapter, saved, adapterFile, get dataRoles() { return dataRoles; }, instance, container, nativeDom, fields, calls, handlers, seen, selected,
    get bag() { return bag; }, set bag(value) { bag = value; resolveSaved(value); }, get regions() { return regions; }, get refreshes() { return refreshes; },
    setRevision(value) { revision = value; render(); }, disconnect() { portConnected = false; render(); },
    action(label) { const children = regions.commands?.props.children; const node = (Array.isArray(children) ? children : [children]).find(node => node?.props.children === label); assert.ok(node, label); return node.props.onClick; },
    click(label) { const button = [...container.querySelectorAll('button')].find(node => node.textContent === label); assert.ok(button, label); h.act(() => button.click()); },
    getter() { open = true; return adapter.getConfiguration(() => ({ pages: [] })); },
    mount(field, callback = (name, values) => { bag = { ...bag, [name]: values }; adapter.fieldChanged(name, undefined, values); }) { h.act(() => field.properties.onRender(nativeDom, undefined, callback)); adapter.rendered(); },
    close(field) { h.act(() => field.properties.onDispose(nativeDom)); open = false; adapter.configurationComplete(); },
    async flush() { await h.act(async () => { for (let i = 0; i < 30; i++) await Promise.resolve(); await new Promise(resolve => setImmediate(resolve)); }); },
    origins() {
      const files = [...new Set(seen.filter(file => file.startsWith(path.dirname(root) + '/') && /\/(src|lib)\//.test(file)))];
      dataRoles = [];
      const data = new Set([...new Set(h.nativeResources.requestedIds)].flatMap(id => {
        const item = JSON.parse(read(path.join(path.dirname(root), 'localization/resources.json'), 'utf8')).resources.find(item => item.id === id);
        assert.ok(item);
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
      for (const stem of ['property-pane/pickers/property-pane-site-picker', 'property-pane/pickers/property-pane-picker-factory.internal', 'property-pane/pickers/property-pane-picker.internal', 'controls/pickers/site-picker', 'core/provider-webpart', 'core/context.internal'])
        assert.ok(files.some(file => file === path.join(root, stem + '.js') || file === path.join(root, stem + '.ts') || file === path.join(root, stem + '.tsx')), 'actual matching origin ' + stem);
      assert.ok(seen.includes(adapterFile), 'actual matching private adapter implementation');
      assert.equal(seen.some(file => file.includes('/apps/spfx-react-toolkit-test/' + (compiled ? 'src' : 'lib') + '/') && file.includes('NativePropertyPaneDemoHost.')), false, 'no opposite private adapter');
      return files;
    }
  };
}
module.exports = { createNativeSiteFixture };

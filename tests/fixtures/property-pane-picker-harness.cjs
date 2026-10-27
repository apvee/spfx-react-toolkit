const { createPickerHarness } = require('./picker-source-harness.cjs');
function createPropertyPaneHarness() {
  // The shared loader assumes TSX; this existing .ts hook contains an async single-generic arrow.
  // Compile the real hook with its filename rather than replacing invocation/lifecycle behavior.
  const hookFile = require('node:path').resolve(__dirname, '../../packages/spfx-react-toolkit/src/hooks/useAsyncInvoke.internal.ts');
  const hookModule = { exports: {} };
  const hookCode = require('typescript').transpileModule(require('node:fs').readFileSync(hookFile, 'utf8'), { fileName: hookFile, compilerOptions: { module: require('typescript').ModuleKind.CommonJS, target: require('typescript').ScriptTarget.ES2020 } }).outputText;
  new Function('require', 'module', 'exports', hookCode)(require, hookModule, hookModule.exports);
  const keys = { page: {}, sp: {}, graph: {}, theme: {} };
  let readContext, readPage;
  const h = createPickerHarness({
    '@microsoft/sp-http': { SPHttpClient: { serviceKey: keys.sp, configurations: { v1: {} } }, MSGraphClientFactory: { serviceKey: keys.graph } },
    '@microsoft/sp-page-context': { PageContext: { serviceKey: keys.page } },
    '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 } },
    '@microsoft/sp-component-base': { ThemeProvider: { serviceKey: keys.theme } },
    '@microsoft/sp-property-pane': { PropertyPaneFieldType: { Custom: 1 } },
    './useSPFxContext': { useSPFxContext: () => readContext() },
    './useSPFxPageContext': { useSPFxPageContext: () => readPage() },
    './useAsyncInvoke.internal': hookModule.exports,
  });
  Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  readContext = h.load('core/context.internal.tsx').useSPFxContext;
  readPage = h.load('hooks/useSPFxPageContext.ts').useSPFxPageContext;
  function host(id = 'owner') {
    let graphInitializations = 0, requests = 0, refreshes = 0;
    const handlers = new Set();
    const page = { web: { absoluteUrl: 'https://tenant.test/sites/hr' }, site: { absoluteUrl: 'https://tenant.test/sites/hr' }, user: { loginName: 'current@tenant.test', displayName: 'Current', email: 'current@tenant.test' },
      cultureInfo: { isRightToLeft: false }, aadInfo: { tenantId: { toString: () => '11111111-1111-4111-8111-111111111111' } } };
    const graph = { api() { const api = { version() { return api; }, header() { return api; }, get: async () => { requests++; return { value: [] }; } }; return api; } };
    const sp = { get: async () => { requests++; return { ok: true, status: 200, json: async () => ({ value: [] }) }; }, post: async () => { requests++; return { ok: true, status: 200, json: async () => ({ ClientPeoplePickerSearchUser: JSON.stringify([{ Key: 'alice@tenant.test', DisplayText: 'Alice', IsResolved: true, EntityType: 'User', EntityData: { Email: 'alice@tenant.test' } }]) }) }; } };
    const theme = { tryGetTheme: () => undefined, themeChangedEvent: { add(_observer, handler) { handlers.add(handler); }, remove(_observer, handler) { handlers.delete(handler); } } };
    const scope = { whenFinished: callback => callback(), consume(key) { if (key === keys.page) return page; if (key === keys.sp) return sp; if (key === keys.graph) return { getClient: async () => { graphInitializations++; return graph; } }; if (key === keys.theme) return theme; throw new Error('Unexpected native service'); } };
    const instance = { context: { instanceId: id, pageContext: page, serviceScope: scope, spHttpClient: sp, propertyPane: { refresh() { refreshes++; } } }, properties: {}, displayMode: 1, domElement: document.createElement('div'), render() {} };
    return { instance, handlers, get graphInitializations() { return graphInitializations; }, get requests() { return requests; }, get refreshes() { return refreshes; } };
  }
  return { ...h, host };
}
module.exports = { createPropertyPaneHarness };

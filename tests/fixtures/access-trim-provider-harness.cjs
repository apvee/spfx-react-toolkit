const { createHarness } = require('../services-test-harness.cjs');
const ReactDOM = require('react-dom');
const { Permission } = require('./access-trim-harness.cjs');
const { loadHook } = require('../async-test-harness.cjs');
function createProviderHarness() {
  const keys = { page: {}, sp: {}, graph: {}, theme: {} };
  let useContext; let usePage;
  const h = createHarness({
    '@microsoft/sp-http': { SPHttpClient: { serviceKey: keys.sp, configurations: { v1: {} } }, MSGraphClientFactory: { serviceKey: keys.graph } },
    '@microsoft/sp-page-context': { SPPermission: Permission, PageContext: { serviceKey: keys.page } },
    '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 } },
    '@microsoft/sp-component-base': { ThemeProvider: { serviceKey: keys.theme } },
    './useSPFxContext': { useSPFxContext: () => useContext() },
    './useSPFxPageContext': { useSPFxPageContext: () => usePage() },
    './useAsyncInvoke.internal': { useAsyncInvoke: loadHook('useAsyncInvoke.internal', 'useAsyncInvoke') },
  });
  useContext = h.load('core/context.internal.tsx').useSPFxContext;
  usePage = h.load('hooks/useSPFxPageContext.ts').useSPFxPageContext;
  const Provider = h.load('core/provider-base.internal.tsx').SPFxProviderBase;
  const theme = { tryGetTheme: () => ({}), themeChangedEvent: { add() {}, remove() {} } };
  function host(id, page, sp, graph) {
    const serviceScope = { whenFinished: fn => fn(), consume: key => {
      if (key === keys.page) return page;
      if (key === keys.sp) return sp;
      if (key === keys.graph) return { getClient: async () => graph };
      if (key === keys.theme) return theme;
      throw new Error('Unexpected SDK boundary');
    } };
    return { context: { instanceId: id, serviceScope }, properties: {}, displayMode: 1, domElement: document.createElement('div'), render() {} };
  }
  function mount(hosts, Control, props, children) {
    const container = document.createElement('div'); document.body.appendChild(container); const observed = [];
    function App() {
      h.React.useLayoutEffect(() => { observed.push(container.textContent); });
      return h.React.createElement(h.React.Fragment, undefined, ...hosts.map((instance, index) => h.React.createElement(Provider, { instance, key: index }, h.React.createElement(Control, props, children))));
    }
    const fixture = { container, observed, render(nextHosts = hosts, nextProps = props) { hosts = nextHosts; props = nextProps; h.act(() => { ReactDOM.render(h.React.createElement(App), container); }); }, unmount() { h.act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove(); } };
    fixture.render(); return fixture;
  }
  return { h, host, mount };
}
module.exports = { createProviderHarness };

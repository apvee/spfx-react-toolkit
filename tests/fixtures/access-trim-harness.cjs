const { createHarness } = require('../services-test-harness.cjs');
const ReactDOM = require('react-dom');
const { deferred } = require('../async-test-harness.cjs');
class Permission {
  constructor(mask) { this.mask = mask; }
  get value() { return { High: this.mask.High ?? 0, Low: this.mask.Low }; }
  hasPermission(permission) { return (this.mask.Low & permission.mask.Low) === permission.mask.Low; }
}
function setup() {
  const state = {
    provider: {},
    page: { web: { absoluteUrl: 'https://tenant.test/sites/current' }, user: { loginName: 'first', email: 'first@test' } },
    graphReady: true,
  };
  const spCalls = []; const graphCalls = [];
  state.spResponse = async url => ({ ok: true, json: async () => url.includes('currentuser') ? { Id: 17 } : url.includes('/users') ? { value: [{ Id: 17 }] } : url.includes('sitegroups(') ? { Id: 2, Title: 'Readers' } : { High: '0', Low: '1' } });
  state.graphResponse = async (url, body) => body ? { value: body.groupIds } : url === '/me?$select=id' ? { id: 'aaaaaaaa-0000-0000-0000-000000000001' } : { value: [] };
  const spClient = () => ({ get: (url, config, options) => { spCalls.push({ url, config, options }); return state.spResponse(url); } });
  const graphClient = () => ({ api: url => ({ get: () => { graphCalls.push({ url }); return state.graphResponse(url); }, post: body => { graphCalls.push({ url, body }); return state.graphResponse(url, body); } }) });
  state.sp = spClient(); state.graph = graphClient();
  const h = createHarness({
    '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } },
    '@microsoft/sp-page-context': { SPPermission: Permission },
    '../hooks/useSPFxContext': { useSPFxContext: () => state.provider },
    './useSPFxContext': { useSPFxContext: () => state.provider },
    '../../hooks/useSPFxContext': { useSPFxContext: () => state.provider },
    './useSPFxPageContext': { useSPFxPageContext: () => state.page },
    '../../hooks/useSPFxPageContext': { useSPFxPageContext: () => state.page },
    './useSPFxSPHttpClient': { useSPFxSPHttpClient: () => ({ client: state.sp }) },
    '../../hooks/useSPFxSPHttpClient': { useSPFxSPHttpClient: () => ({ client: state.sp }) },
    '../../hooks/useSPFxMSGraphClient': { useSPFxMSGraphClient: () => ({ client: state.graphReady ? state.graph : undefined, isInitializing: !state.graphReady, initError: state.graphError }) },
  });
  function mount(Control, props, children) {
    const container = document.createElement('div'); document.body.appendChild(container);
    const observed = [];
    function Observer() {
      h.React.useLayoutEffect(() => { observed.push(container.textContent); });
      return h.React.createElement(Control, props, children);
    }
    const fixture = {
      container, observed,
      render(nextProps = props, nextChildren = children) {
        props = nextProps; children = nextChildren;
        h.act(() => { ReactDOM.render(h.React.createElement(Observer), container); });
      },
      unmount() { h.act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove(); },
    };
    fixture.render(); return fixture;
  }
  function branches(Control, granted = h.React.createElement('b', { key: 'grant' }, 'grant')) {
    return [granted,
      h.React.createElement(Control.Loading, { key: 'loading' }, h.React.createElement('i', null, 'loading')),
      h.React.createElement(Control.Fallback, { key: 'fallback' }, h.React.createElement('i', null, 'denied')),
      h.React.createElement(Control.ErrorFallback, { key: 'error' }, h.React.createElement('i', null, 'error'))];
  }
  return { h, state, spCalls, graphCalls, mount, branches, replaceSP() { state.sp = spClient(); }, replaceGraph() { state.graph = graphClient(); } };
}
async function flush(h, fn = () => {}) { await h.act(async () => { fn(); for (let i = 0; i < 40; i++) await Promise.resolve(); }); }
module.exports = { setup, flush, deferred, Permission };

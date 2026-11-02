const path = require('node:path');
const fs = require('node:fs');
const assert = require('node:assert/strict');
const { loadHook } = require('../../async-test-harness.cjs');
const { createHarness } = require('../../services-test-harness.cjs');
const { scenario } = require('./cards-description-list.cjs');
const ReactDOM = require('react-dom');
const { Simulate } = require('react-dom/test-utils');
const { createDOMRenderer } = require('@griffel/core');
const { RendererProvider } = require('@griffel/react');
const fluent = require('@fluentui/react-components');
const nativeGlobals = ['NodeFilter', 'Element', 'HTMLElement'];
const oldGlobals = nativeGlobals.map(name => Object.getOwnPropertyDescriptor(global, name));
for (const name of nativeGlobals) Object.defineProperty(global, name, { configurable: true, writable: true, value: window[name] });
require('node:test').after(() => nativeGlobals.forEach((name, index) => { if (oldGlobals[index]) Object.defineProperty(global, name, oldGlobals[index]); else delete global[name]; }));
class Permission {
  constructor(value) { this.value = value; }
  hasPermission(permission) { return this === Permission.fullMask || ((this.value.Low & permission.value.Low) === permission.value.Low && (this.value.High & permission.value.High) === permission.value.High); }
}
for (const [name, Low] of Object.entries({ viewListItems: 1, addListItems: 2, manageLists: 0x800, manageWeb: 0x40000000, emptyMask: 0 })) Permission[name] = new Permission({ High: 0, Low });
Permission.fullMask = new Permission({ High: 0x7fffffff, Low: 0xffffffff });
const owners = ['Context','PageContext','InstanceInfo','DisplayMode','EnvironmentInfo','UserInfo','SiteInfo','ListInfo','LocaleInfo','HubSiteInfo','Teams','PageType','CorrelationInfo','Permissions','CrossSitePermissions','ContainerInfo','ContainerSize','ThemeInfo','Fluent9ThemeInfo','ServiceScope'];
function page() {
  return { site: { id: '11111111-0000-0000-0000-000000000001', absoluteUrl: 'https://tenant.test/sites/current', serverRelativeUrl: '/sites/current' },
    web: { id: '22222222-0000-0000-0000-000000000002', title: 'Project intranet', absoluteUrl: 'https://tenant.test/sites/current', serverRelativeUrl: '/sites/current', timeZoneInfo: { id: 4, description: 'Regional group', offset: -60, standardOffset: 0, daylightOffset: -60 } },
    user: { displayName: 'Giulia Rossi', loginName: 'giulia', email: 'giulia@contoso.com', isExternalGuestUser: false },
    cultureInfo: { currentCultureName: 'it-IT', currentUICultureName: 'en-US', isRightToLeft: false },
    legacyPageContext: { hubSiteId: 'aaaaaaaa-0000-0000-0000-000000000001', isHubSite: false, siteClassification: 'Internal' }, page: { type: 'SitePage' } };
}
function createContextFixture(t, title, ownerKey, options = {}) {
  const toolkit = {}, actualContext = {}, actualPage = {};
  const keys = { page: {}, sp: {}, theme: {} };
  const h = createHarness({ '@apvee/spfx-react-toolkit': toolkit,
    '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 } },
    '@microsoft/sp-component-base': { ThemeProvider: { serviceKey: keys.theme } },
    '@microsoft/sp-page-context': { PageContext: { serviceKey: keys.page }, SPPermission: Permission },
    '@microsoft/sp-http': { SPHttpClient: { serviceKey: keys.sp, configurations: { v1: {} } } },
    './useSPFxContext': actualContext, './useSPFxPageContext': actualPage,
    './useAsyncInvoke.internal': { useAsyncInvoke: loadHook('useAsyncInvoke.internal', 'useAsyncInvoke') } });
  Object.assign(actualContext, h.load('hooks/useSPFxContext.ts'));
  Object.assign(actualPage, h.load('hooks/useSPFxPageContext.ts'));
  for (const suffix of owners) Object.assign(toolkit, h.load(`hooks/useSPFx${suffix}.ts`));
  Object.assign(toolkit, h.load('core/provider-webpart.tsx'), h.load('controls/web-part-shell/web-part-shell.tsx'), h.load('hooks/useSPFxProperties.ts'), h.load('hooks/useSx.ts'), h.load('helpers/styles/index.ts'), h.load('controls/relative-time/relative-time.tsx'));
  Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  const calls = [], handlers = new Set(), observations = [], roots = new Set();
  const themeService = { current: undefined, tryGetTheme() { return this.current; }, themeChangedEvent: { add: (_owner, fn) => handlers.add(fn), remove: (_owner, fn) => handlers.delete(fn) } };
  function client() { return { get(url, configuration, init) { let resolve, reject; const promise = new Promise((a,b) => { resolve=a; reject=b; }); calls.push({ url, configuration, init, resolve: value => resolve({ ok: true, json: async () => value }), reject }); return promise; } }; }
  function host(id = 'context-one', nativePage = page()) {
    const sp = client();
    const scope = { whenFinished: fn => fn(), consume(key) { if (key === keys.page) return nativePage; if (key === keys.sp) return sp; if (key === keys.theme) return themeService; throw new Error('Unknown native service'); } };
    const domElement = document.createElement('div');
    return { instanceId: id, context: { instanceId: id, domElement, pageContext: nativePage, serviceScope: scope, propertyPane: { refresh() {} }, sdks: {} }, properties: { title: 'Native host', showTitle: false }, displayMode: 1, domElement, render() {} };
  }
  let instance = host(), revision = 0, selected = scenario(title, ownerKey), Demo, parentRenders = 0;
  // Negative control changes only the emitted child's scope comparison, in memory.
  // It uses the established loader/compiler and restores the filesystem reader immediately.
  const read = fs.readFileSync;
  if (options.parentOnly) fs.readFileSync = (file, ...args) => {
    const source = read(file, ...args);
    if (String(file) !== selected.file) return source;
    const marker = 'const current = sameScope(run.scope, scope);';
    assert.equal(source.split(marker).length, 2, 'One emitted Hub child comparison');
    return source.replace(marker, 'const current = true;');
  };
  try { Demo = h.load(selected.file).Demo; } finally { fs.readFileSync = read; }
  const base = path.resolve(__dirname, '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
  const Shell = h.load(path.join(base, 'WebPartShellDemoHost.tsx')).WebPartShellDemoHost;
  const Provider = toolkit.SPFxWebPartProvider;
  const { DemoCard } = h.load(path.join(base, 'shared/DemoCard.tsx'));
  const container = document.createElement('div'); document.body.appendChild(container); roots.add(container);
  const renderer = createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  function Frame() { parentRenders++; return h.React.createElement(Demo, { resetRevision: revision, render: regions => h.React.createElement(DemoCard, { title: selected.entry.title, iconName: 'Info', contentPlacement: 'configuration', ...regions }) }); }
  function render() {
    h.act(() => { ReactDOM.render(h.React.createElement(RendererProvider, { renderer }, h.React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme },
      h.React.createElement(options.directProvider ? Provider : Shell, { instance }, h.React.createElement(h.React.Profiler, { id: 'Context', onRender: () => observations.push({ text: container.textContent, hubId: instance.context.pageContext.legacyPageContext?.hubSiteId, parentRenders }) }, h.React.createElement(Frame))))), container); });
  }
  function unmount() { if (!roots.has(container)) return; h.act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove(); roots.delete(container); }
  t.after(() => { unmount(); h.close(); });
  render();
  return { h, toolkit, keys, calls, observations, container, handlers, themeService, page, host, client, render, unmount,
    mountPeer(nativeHost = host('context-peer')) {
      const PeerDemo = Demo;
      const peer = h.mountComponent(() => h.React.createElement(RendererProvider, { renderer }, h.React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme }, h.React.createElement(Provider, { instance: nativeHost }, h.React.createElement(PeerDemo)))));
      return { ...peer, instance: nativeHost, click(name) { const node = [...peer.container.querySelectorAll('button')].find(button => button.textContent === name); assert.ok(node, name); h.act(() => { Simulate.click(node); }); } };
    },
    get entry() { return selected.entry; }, get instance() { return instance; }, get parentRenders() { return parentRenders; },
    replace(next) { instance = next; render(); }, reset() { revision++; render(); },
    select(nextTitle, nextOwner) { selected = scenario(nextTitle, nextOwner); Demo = h.load(selected.file).Demo; render(); },
    click(name) { const button = [...container.querySelectorAll('button')].find(node => node.textContent === name); assert.ok(button, name); h.act(() => Simulate.click(button)); },
    change(label, value) { const node = container.querySelector(`[aria-label="${label}"]`); assert.ok(node, label); h.act(() => { node.value = value; Simulate.change(node, { target: { value } }); }); },
    async flush(fn = () => {}) { await h.act(async () => { fn(); for (let i=0;i<20;i++) await Promise.resolve(); }); },
  };
}
module.exports = { createContextFixture, Permission, page };

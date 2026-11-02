// Import the established DOM/bootstrap in the same order as Context; no second DOM universe.
const { loadHook } = require('../../async-test-harness.cjs');
const { createHarness } = require('../../services-test-harness.cjs');
const { scenario } = require('./cards-description-list.cjs');
const path = require('node:path');
const fs = require('node:fs');
const assert = require('node:assert/strict');
const ReactDOM = require('react-dom');
const { Simulate } = require('react-dom/test-utils');
const { createDOMRenderer } = require('@griffel/core');
const { RendererProvider } = require('@griffel/react');
const fluent = require('@fluentui/react-components');
const globals = ['NodeFilter', 'Element', 'HTMLElement'];
const previous = globals.map(name => Object.getOwnPropertyDescriptor(global, name));
for (const name of globals)
    Object.defineProperty(global, name, { configurable: true, value: window[name] });
require('node:test').after(() => globals.forEach((name, i) => { if (previous[i])
    Object.defineProperty(global, name, previous[i]);
else
    delete global[name]; }));
function deferred(record = {}) { let resolve, reject; const promise = new Promise((a, b) => { resolve = a; reject = b; }); return Object.assign(record, { promise, resolve, reject }); }
function token(scopes = 'User.Read', aud = 'https://graph.microsoft.com') { return 'e30.' + Buffer.from(JSON.stringify({ aud, scp: scopes, exp: 9999999999 })).toString('base64url') + '.signature'; }
function createClientsFixture(t, title, options = {}) {
    const compiled = options.compiled ?? process.env.CLIENTS_IMPLEMENTATION === 'compiled';
    const asyncModule = compiled ? {} : { useAsyncInvoke: loadHook('useAsyncInvoke.internal', 'useAsyncInvoke') };
    const toolkit = {}, context = {}, pageModule = {};
    const keys = Object.fromEntries(['page', 'theme', 'http', 'sp', 'graph', 'aad', 'token'].map(key => [key, {}]));
    const native = Object.fromEntries([['HttpClient', 'http'], ['SPHttpClient', 'sp'], ['MSGraphClientFactory', 'graph'], ['AadHttpClientFactory', 'aad'], ['AadTokenProviderFactory', 'token'], ['AadHttpClient', 'aadClient']].map(([name, key]) => [name, { serviceKey: keys[key], configurations: { v1: {} } }]));
    const h = createHarness({ '@apvee/spfx-react-toolkit': toolkit, '@microsoft/sp-http': native,
        '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 } },
        '@microsoft/sp-page-context': { PageContext: { serviceKey: keys.page } },
        '@microsoft/sp-component-base': { ThemeProvider: { serviceKey: keys.theme } },
        './useSPFxContext': context, './useSPFxPageContext': pageModule,
        // The generic TS arrow is parsed as TSX by services-test-harness; reuse its established TS-only loader.
        // This module imports React only, owns state per call and creates no provider/store/context.
        './useAsyncInvoke.internal': asyncModule,
        ...(options.oneDriveService ? { '../services/spfx-onedrive-app-data.service': { createSPFxOneDriveAppDataService: () => options.oneDriveService } } : {}),
    });
    const libraryRoot = path.resolve(__dirname, '../../../packages/spfx-react-toolkit', compiled ? 'lib' : 'src');
    function library(file) { return h.load(path.join(libraryRoot, file + (compiled ? '.js' : /provider|web-part-shell/.test(file) ? '.tsx' : '.ts'))); }
    if (compiled) Object.assign(asyncModule, library('hooks/useAsyncInvoke.internal'));
    Object.assign(context, library('hooks/useSPFxContext'));
    Object.assign(pageModule, library('hooks/useSPFxPageContext'));
    for (const name of ['Context', 'PageContext', 'ServiceScope', 'InstanceInfo', 'HttpClient', 'SPHttpClient', 'MSGraphClient', 'AadHttpClient', 'AadTokenProvider', 'ApiPermissionPrecheck', 'Properties', 'Fluent9ThemeInfo'])
        Object.assign(toolkit, library('hooks/useSPFx' + name));
    if (options.graphData) for (const name of ['UserInfo', 'UserPhoto', 'OneDriveAppData']) Object.assign(toolkit, library('hooks/useSPFx' + name));
    Object.assign(toolkit, library('hooks/useSx'), library('core/provider-webpart'), library('controls/web-part-shell/web-part-shell'), library('helpers/styles/index'), library('helpers/spfx-api-permission-precheck.helpers'));
    Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
    const calls = [], initializations = [], tokens = [], observers = [], commits = [], downloads = [];
    if (options.graphData) t.mock.method(global, 'fetch', (url, init) => {
        assert.equal(url, 'https://download.fixture.invalid/app-data');
        const request = deferred({ url, init }); downloads.push(request); return request.promise;
    });
    const themeHandlers = new Set();
    const themeService = { tryGetTheme: () => undefined, themeChangedEvent: { add: (_o, fn) => themeHandlers.add(fn), remove: (_o, fn) => themeHandlers.delete(fn) } };
    function http(kind) { const client = { get(url, configuration, init) { const request = deferred({ kind, client, url, configuration, init }); calls.push(request); return request.promise; } }; return client; }
    function host(id = 'clients-one') {
        const page = { web: { absoluteUrl: 'https://tenant.test/sites/' + id, title: id }, site: { absoluteUrl: 'https://tenant.test/sites/' + id }, user: { displayName: 'Fixture User', loginName: 'fixture-user', email: 'fixture@example.invalid' }, cultureInfo: {} };
        const popup = new Map(), redirect = new Map();
        const event = handlers => ({ add(observer, handler) { observers.push(observer); handlers.set(handler, observer); }, remove(_observer, handler) { handlers.delete(handler); } });
        const provider = { popupEvent: event(popup), onBeforeRedirectEvent: event(redirect), getToken(resource, settings) { const request = deferred({ id, resource, settings }); tokens.push(request); return request.promise; } };
        const graph = { api(url) {
            const headers = {};
            function request(method, body) { const pending = deferred({ kind: 'graph', client: graph, url, fields: query.fields, method, headers: { ...headers }, body }); calls.push(pending); return pending.promise; }
            const query = { select(fields) { query.fields = fields; return query; }, header(name, value) { headers[name] = value; return query; }, get() { return request('GET'); }, put(body) { return request('PUT', body); } };
            return query;
        } };
        function init(kind, value, target) { const request = deferred({ id, kind, target, value }); initializations.push(request); return request.promise; }
        const services = new Map([[keys.page, page], [keys.theme, themeService], [keys.http, http('http')], [keys.sp, http('sp')],
            [keys.graph, { getClient: version => init('graph', graph, version) }],
            [keys.aad, { getClient: resource => init('aad', http('aad'), resource) }],
            [keys.token, { getTokenProvider: () => init('token', provider) }]]);
        const scope = { whenFinished: callback => callback(), consume(key) { if (!services.has(key))
                throw new Error('Missing native service'); return services.get(key); } };
        const domElement = document.createElement('div');
        return { instanceId: id, context: { instanceId: id, domElement, pageContext: page, serviceScope: scope, sdks: {}, propertyPane: { refresh() { } } }, properties: { title: 'Clients fixture', showTitle: false }, displayMode: 1, domElement, render() { }, native: { services, provider, popup, redirect } };
    }
    let redrawChild;
    let instance = host(), revision = options.revision ?? 0, selected = scenario(title, options.ownerKey), regions;
    const source = fs.readFileSync;
    if (options.unmasked)
        fs.readFileSync = (file, ...args) => { const value = source(file, ...args); return String(file) === selected.file ? value.replace('observedScope === serviceScope', 'true').replace('observedTarget === info.resourceUrl', 'true') : value; };
    let Demo;
    try {
        Demo = h.load(selected.file).Demo;
    }
    finally {
        fs.readFileSync = source;
    }
    const base = path.resolve(__dirname, '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
    const Shell = h.load(path.join(base, 'WebPartShellDemoHost.tsx')).WebPartShellDemoHost;
    const { DemoCard } = h.load(path.join(base, 'shared/DemoCard.tsx'));
    const renderer = createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
    const container = document.createElement('div');
    document.body.appendChild(container);
    function Frame() { const [, redraw] = h.React.useState(0); redrawChild = () => redraw(value => value + 1); return h.React.createElement(Demo, { resetRevision: revision, render: options.fallback ? undefined : next => { regions = next; return h.React.createElement(DemoCard, { title, iconName: 'Info', contentPlacement: 'configuration', ...next }); } }); }
    function render() { h.act(() => { ReactDOM.render(h.React.createElement(RendererProvider, { renderer }, h.React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme }, h.React.createElement(options.directProvider ? toolkit.SPFxWebPartProvider : Shell, { instance }, h.React.createElement(h.React.Profiler, { id: 'Clients', onRender: () => commits.push({ id: instance.instanceId, text: container.textContent, buttons: [...container.querySelectorAll('button')].map(node => ({ text: node.textContent, disabled: node.disabled })) }) }, h.React.createElement(Frame))))), container); }); }
    let closed = false;
    function unmount() { if (closed)
        return; h.act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove(); closed = true; }
    t.after(() => { unmount(); h.close(); });
    render();
    return { h, toolkit, keys, native, container, calls, downloads, initializations, tokens, commits, observers, themeHandlers, host, render, unmount,
        mountPeer(nativeHost = host('clients-peer')) { const peer = h.mountComponent(() => h.React.createElement(RendererProvider, { renderer }, h.React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme }, h.React.createElement(toolkit.SPFxWebPartProvider, { instance: nativeHost }, h.React.createElement(Demo))))); return { ...peer, instance: nativeHost, click(name) { const node = [...peer.container.querySelectorAll('button')].find(n => n.textContent === name); assert.ok(node, name); h.act(() => { Simulate.click(node); }); } }; },
        get instance() { return instance; }, get regions() { return regions; }, get entry() { return selected.entry; },
        select(nextTitle, ownerKey) { selected = scenario(nextTitle, ownerKey); Demo = h.load(selected.file).Demo; render(); },
        replace(next) { instance = next; render(); },
        changeScope(next) { instance.context.serviceScope = next.context.serviceScope; h.act(() => redrawChild()); }, reset() { revision++; render(); },
        setRevision(value) { revision = value; render(); },
        action(name) { const children = regions?.commands?.props?.children; const nodes = Array.isArray(children) ? children : [children]; const button = nodes.find(node => node?.props?.children === name); assert.ok(button, name); return button.props.onClick; },
        click(name) { const node = [...container.querySelectorAll('button')].find(node => node.textContent === name); assert.ok(node, name); h.act(() => { Simulate.click(node); }); },
        change(label, value) { const node = container.querySelector(`[aria-label="${label}"]`); assert.ok(node, label); h.act(() => { node.value = value; Simulate.change(node, { target: { value } }); }); },
        response(index, payload) { calls[index].resolve({ ok: true, status: 200, json: async () => payload }); },
        async flush(fn = () => { }) { await h.act(async () => { fn(); for (let i = 0; i < 24; i++)
            await Promise.resolve(); }); },
        async ready() { await this.flush(); await this.flush(() => initializations.filter(request => !request.settled).forEach(request => { request.settled = true; request.resolve(request.value); })); },
    };
}
module.exports = { createClientsFixture, token };

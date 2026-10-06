const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { JSDOM } = require('jsdom');
const dom = new JSDOM('<!doctype html><html><body></body></html>', { url: 'https://example.test' });
global.window = dom.window;
global.document = dom.window.document;
global.navigator = dom.window.navigator;
global.requestAnimationFrame = fn => setTimeout(fn, 0);
global.cancelAnimationFrame = clearTimeout;
// Node's MessageChannel keeps React 17's browser scheduler alive after the suite.
global.MessageChannel = undefined;
const React = require('react');
const ReactDOM = require('react-dom');
const { act } = require('react-dom/test-utils');
function loadHook(file, name, stubs = {}) {
    const modules = new Map();
    function load(fileName) {
        if (modules.has(fileName)) return modules.get(fileName).exports;
        const module = { exports: {} };
        modules.set(fileName, module);
        const source = fs.readFileSync(path.join(process.env.ASYNC_HOOK_SOURCE || path.join(__dirname, '../packages/spfx-react-toolkit/src/hooks'), fileName + '.ts'), 'utf8');
        const requireBoundary = key => {
            if (key === 'react') return React;
            // Exercise the actual shared list lifecycle, retaining existing service/SDK boundaries.
            if (key === './useSPFxPnPList.internal') return load('useSPFxPnPList.internal');
            return stubs[key] || (() => { throw new Error('Missing boundary: ' + key); })();
        };
        new Function('require', 'module', 'exports', ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 } }).outputText)(requireBoundary, module, module.exports);
        return module.exports;
    }
    return load(file)[name];
}
function mount(hook, props) {
    const container = document.createElement('div');
    document.body.appendChild(container);
    let result;
    function Probe(p) { result = hook(p); return null; }
    const render = next => { props = next; act(() => { ReactDOM.render(React.createElement(Probe, props), container); }); };
    render(props);
    return { get current() { return result; }, render, unmount() { act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove(); } };
}
function deferred() { let resolve, reject; const promise = new Promise((a, b) => { resolve = a; reject = b; }); return { promise, resolve, reject }; }
function start(fn) { let result; act(() => { result = fn(); }); return result; }
async function settle(fn) { await act(async () => { await fn(); }); }
async function flush() { await settle(async () => { for (let i = 0; i < 8; i++)
    await Promise.resolve(); }); }
function catalogBoundary(service, client = {}) {
    const discoverAppCatalogUrl = async () => 'https://tenant.test/catalog';
    const checkWritePermission = async () => true;
    return { './useAppCatalogUrl.internal': { useAppCatalogUrl() { const isMountedRef = React.useRef(true); React.useEffect(() => () => { isMountedRef.current = false; }, []); return { spHttpClient: client, discoverAppCatalogUrl, checkWritePermission, isMountedRef }; } },
        '../services/spfx-tenant-property.service': { createSPFxTenantPropertyService: () => service },
        '../services/spfx-tenant-key-value-store.service': { createSPFxTenantKeyValueStoreService: () => service } };
}
function driveBoundary(service) { const client = {}; return { './useSPFxMSGraphClient': { useSPFxMSGraphClient: () => ({ client, isReady: true, isInitializing: false, initError: undefined }) }, '../services/spfx-onedrive-app-data.service': { createSPFxOneDriveAppDataService: () => service } }; }
function pnpBoundary(factory, kind) { const context = { sp: {}, isInitialized: true }; return { './useSPFxPnPContext': { useSPFxPnPContext: () => context }, ['../services/spfx-pnp-' + kind + '.service']: { ['createSPFxPnP' + (kind === 'list' ? 'List' : 'Search') + 'Service']: factory } }; }
const listPage = (items, nextSkip = items.length, hasMore = false) => ({ items, nextSkip, hasMore, effectivePageSize: 1 });
const searchPage = (ids, totalResults = ids.length) => ({ results: ids.map(id => ({ id, data: { Title: id } })), totalResults, refiners: [] });
module.exports = { React, loadHook, mount, deferred, start, settle, flush, catalogBoundary, driveBoundary, pnpBoundary, listPage, searchPage };

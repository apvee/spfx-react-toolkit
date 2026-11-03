// Bounded scenario composition on the existing filename-aware picker harness: one DOM/cache/provider universe.
const { createPickerHarness } = require('../picker-source-harness.cjs');
const { scenario } = require('./cards-description-list.cjs');
const fs = require('node:fs'), path = require('node:path'), assert = require('node:assert/strict');
const ReactDOM = require('react-dom');
const { Simulate } = require('react-dom/test-utils');
const { createDOMRenderer } = require('@griffel/core');
const { RendererProvider } = require('@griffel/react');
const fluent = require('@fluentui/react-components');
const nativePicker = require('./react-pickers-native.cjs');
const web = 'https://tenant.test/sites/root/sub', collection = 'https://tenant.test/sites/root';
const tenant = '11111111-1111-4111-8111-111111111111';
const titles = {
    SitePicker: ['site-picker', 'Explore site selection and contextual search'],
    ListPicker: ['list-picker', 'Explore list catalogs and independent consumer choices'],
    FieldPicker: ['field-picker', 'Explore site columns and list fields'],
    ContentTypePicker: ['content-type-picker', 'Explore site and list content type references'],
    ViewPicker: ['view-picker', 'Explore scoped view catalogs'],
    ListItemPicker: ['list-item-picker', 'Explore list item labels and query pages'],
    PrincipalPicker: ['principal-picker', 'Explore principal discovery and explicit identity resolution'],
};
function deferred() { let resolve, reject; const promise = new Promise((a, b) => { resolve = a; reject = b; }); return { promise, resolve, reject }; }
function json(value, status = 200) { return new Response(JSON.stringify(value), { status, headers: { 'Content-Type': 'application/json' } }); }
function createReactPickersFixture(t, name = 'PrincipalPicker', options = {}) {
    const compiled = options.compiled ?? process.env.PICKER_IMPLEMENTATION === 'compiled';
    Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
    const toolkit = {}, context = {}, pageModule = {};
    const keys = { page: {}, sp: {}, graph: {}, theme: {} };
    const native = { SPHttpClient: { serviceKey: keys.sp, configurations: { v1: {} } }, MSGraphClientFactory: { serviceKey: keys.graph } };
    const seen = [], originalRead = fs.readFileSync;
    const selected = scenario(titles[name][1], `packages/spfx-react-toolkit/src/controls/pickers/${titles[name][0]}.tsx#${name}`);
    t.mock.method(fs, 'readFileSync', (file, ...args) => {
        seen.push(String(file));
        let value = originalRead(file, ...args);
        if (String(file) === selected.file && typeof value === 'string' && options.mutate)
            value = options.mutate(value);
        return value;
    });
    const h = createPickerHarness({ '@apvee/spfx-react-toolkit': toolkit,
        '@microsoft/sp-http': native, '@microsoft/sp-page-context': { PageContext: { serviceKey: keys.page } },
        '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 }, _SPKillSwitch: { isActivated: () => false } },
        '@microsoft/sp-component-base': { ThemeProvider: { serviceKey: keys.theme } },
        './useSPFxContext': context, './useSPFxPageContext': pageModule,
    });
    const root = path.resolve(__dirname, '../../../packages/spfx-react-toolkit', compiled ? 'lib' : 'src');
    const library = (file, tsx = false) => h.load(path.join(root, file + (compiled ? '.js' : tsx ? '.tsx' : '.ts')));
    Object.assign(context, library('hooks/useSPFxContext'));
    Object.assign(pageModule, library('hooks/useSPFxPageContext'));
    Object.assign(toolkit, context, pageModule, library('core/provider-webpart', true));
    // Exact known TSX owner path, never the client helper's provider-only extension heuristic.
    for (const [control, [file]] of Object.entries(titles))
        Object.assign(toolkit, library('controls/pickers/' + file, true));
    const calls = [], initializations = [], pending = [];
    const subject = { mounts: 0, unmounts: 0, props: undefined };
    const Actual = toolkit[name];
    toolkit[name] = function ObservedPicker(props) { subject.props = props; h.React.useEffect(() => { subject.mounts++; return () => { subject.unmounts++; }; }, []); return h.React.createElement(Actual, options.subjectProps ? options.subjectProps(props) : props); };
    let hold = options.hold, override = options.route;
    t.mock.method(global, 'fetch', (url, init) => { if (name !== 'SitePicker')
        throw new Error('Unexpected live network: ' + url); return dispatch({ kind: 'pnp', url: String(url), init, method: init?.method ?? 'GET' }); });
    function row(login = 'alice@tenant.test') { const alice = login.includes('alice'); return { Id: alice ? 42 : 43, LoginName: login, Title: alice ? 'Alice' : 'Bob', Email: login, PrincipalType: 1 }; }
    function route(request) {
        if (override) {
            const response = override(request);
            if (response !== undefined)
                return response;
        }
        const url = new URL(request.url), endpoint = url.pathname.toLowerCase();
        assert.equal(url.origin, 'https://tenant.test');
        if (request.kind === 'sp') {
            assert.equal(request.configuration, native.SPHttpClient.configurations.v1, 'actual native v1 object');
            assert.deepEqual(request.init.headers, { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'OData-Version': '' });
            const domainResponse = nativePicker.route(request);
            if (domainResponse !== undefined)
                return domainResponse;
            if (endpoint.endsWith('clientpeoplepickerwebserviceinterface.clientpeoplepickersearchuser')) {
                assert.equal(request.method, 'POST');
                assert.equal(url.search, '');
                const q = JSON.parse(request.init.body).queryParams;
                assert.deepEqual(Object.keys(q).sort(), ['AllUrlZones', 'AllowEmailAddresses', 'AllowMultipleEntities', 'MaximumEntitySuggestions', 'PrincipalSource', 'PrincipalType', 'QueryString', 'SharePointGroupID', 'UseSubstrateSearch'].sort());
                assert.deepEqual([q.AllowEmailAddresses, q.AllowMultipleEntities, q.AllUrlZones, q.PrincipalSource, q.PrincipalType, q.SharePointGroupID, q.UseSubstrateSearch], [false, false, false, 15, 1, 0, false]);
                assert.equal(q.MaximumEntitySuggestions, 50);
                assert.equal(typeof q.QueryString, 'string');
                const rows = ['alice@tenant.test', 'bob@tenant.test'].filter(login => login.includes(q.QueryString.toLowerCase()) || (login.startsWith('alice') ? 'Alice' : 'Bob').toLowerCase().startsWith(q.QueryString.toLowerCase()));
                return json({ ClientPeoplePickerSearchUser: JSON.stringify(rows.map(login => ({ Key: login, DisplayText: login.startsWith('alice') ? 'Alice' : 'Bob', IsResolved: true, EntityType: 'User', EntityData: { Email: login, UserPrincipalName: login } }))) });
            }
            if (endpoint.endsWith('/_api/site')) {
                assert.equal(request.method, 'GET');
                assert.deepEqual([...url.searchParams], [['$select', 'Url']]);
                return json({ Url: collection });
            }
            if (endpoint.endsWith('/_api/web/siteusers/getbyid(43)')) {
                assert.equal(request.method, 'GET');
                assert.equal(url.searchParams.get('$select'), 'Id,LoginName,Title,Email,PrincipalType');
                return json(row('bob@tenant.test'));
            }
            if (endpoint.includes('/_api/web/siteusers/getbyloginname')) {
                assert.equal(request.method, 'GET');
                assert.equal(url.searchParams.get('$select'), 'Id,LoginName,Title,Email,PrincipalType');
                const login = url.searchParams.get('@login').slice(1, -1);
                return login.includes('alice') && options.registered !== true ? json({ error: 'not registered' }, 404) : json(row(login));
            }
            if (endpoint.endsWith('/_api/web/currentuser')) {
                assert.equal(request.method, 'GET');
                assert.equal(url.searchParams.get('$select'), 'Id,Title,LoginName,Email,PrincipalType');
                return json({ Id: 41, Title: 'Current user', LoginName: 'current@tenant.test', Email: 'current@tenant.test', PrincipalType: 1 });
            }
            if (endpoint.endsWith('/_api/web/sitegroups')) {
                assert.equal(request.method, 'GET');
                assert.equal(url.searchParams.get('$select'), 'Id,Title,Description');
                assert.equal(url.searchParams.get('$top'), '50');
                assert.match(url.searchParams.get('$filter'), /^startswith\(Title,'[^']*'\)$/);
                return json({ value: [{ Id: 51, Title: 'Group SharePoint', Description: 'Native group' }] });
            }
            if (endpoint.endsWith('/_api/web/sitegroups/getbyid(51)')) {
                assert.equal(request.method, 'GET');
                assert.equal(url.searchParams.get('$select'), 'Id,Title,Description');
                return json({ Id: 51, Title: 'Group SharePoint', Description: 'Native group' });
            }
            if (endpoint.endsWith('/_api/web/ensureuser')) {
                assert.equal(request.method, 'POST');
                assert.equal(url.search, '');
                const body = JSON.parse(request.init.body);
                assert.deepEqual(Object.keys(body), ['logonName']);
                assert.ok(['alice@tenant.test', 'bob@tenant.test'].includes(body.logonName));
                return json(row(body.logonName));
            }
        }
        if (request.kind === 'pnp') {
            const response = nativePicker.route(request);
            if (response !== undefined)
                return response;
        }
        if (request.kind === 'graph') {
            assert.equal(request.version, 'v1.0');
            if (url.pathname === '/me/people' || url.pathname === '/me/memberOf') {
                assert.equal(url.searchParams.get('$top'), '50');
                return { value: [] };
            }
            if (/^\/groups\/[^/]+$/.test(url.pathname)) {
                const id = url.pathname.split('/').pop(), index = ['44444444-4444-4444-8444-444444444444', '55555555-5555-4555-8555-555555555555', '66666666-6666-4666-8666-666666666666'].indexOf(id);
                assert.ok(index >= 0);
                assert.equal(url.searchParams.get('$select'), 'id,displayName,groupTypes,securityEnabled,mailEnabled,mail,description');
                return { id, displayName: ['Group Microsoft365', 'Group Security', 'Group Distribution'][index], groupTypes: index === 0 ? ['Unified'] : [], securityEnabled: index === 1, mailEnabled: index !== 1 };
            }
            if (url.pathname === '/groups') {
                assert.equal(url.searchParams.get('$select'), 'id,displayName,groupTypes,securityEnabled,mailEnabled,mail,description');
                assert.match(url.searchParams.get('$filter'), /^startswith\(displayName,'[^']*'\)$/);
                assert.equal(url.searchParams.get('$top'), '50');
                return { value: [{ id: '44444444-4444-4444-8444-444444444444', displayName: 'Group Microsoft365', groupTypes: ['Unified'], securityEnabled: false, mailEnabled: true }, { id: '55555555-5555-4555-8555-555555555555', displayName: 'Group Security', groupTypes: [], securityEnabled: true, mailEnabled: false }, { id: '66666666-6666-4666-8666-666666666666', displayName: 'Group Distribution', groupTypes: [], securityEnabled: false, mailEnabled: true }] };
            }
            if (/^\/users\/[^/]+$/.test(url.pathname)) {
                assert.equal(url.searchParams.get('$select'), 'id,displayName,userPrincipalName,mail');
                const login = decodeURIComponent(url.pathname.slice('/users/'.length));
                return { id: login.includes('alice') ? '22222222-2222-4222-8222-222222222222' : '33333333-3333-4333-8333-333333333333', displayName: login.includes('alice') ? 'Alice' : 'Bob', userPrincipalName: login, mail: login };
            }
        }
        throw new Error('Unsupported native picker fixture route: ' + request.method + ' ' + request.url);
    }
    function dispatch(request) {
        calls.push(request);
        // Validate every route even while its response is held. No universal JSON success fallback.
        const response = route(request);
        if (hold?.(request)) {
            const deferredRequest = Object.assign(request, deferred(), { response });
            pending.push(deferredRequest);
            return deferredRequest.promise;
        }
        return Promise.resolve(response);
    }
    function host(id = 'pickers-one') {
        const page = { web: { absoluteUrl: web, title: 'Subweb' }, site: { absoluteUrl: collection }, user: { loginName: 'current@tenant.test', displayName: 'Current', email: 'current@tenant.test' }, aadInfo: { tenantId: { toString: () => tenant } }, cultureInfo: { isRightToLeft: false } };
        const sp = { get: (url, configuration, init) => dispatch({ kind: 'sp', id, client: sp, url, configuration, init, method: 'GET' }), post: (url, configuration, init) => dispatch({ kind: 'sp', id, client: sp, url, configuration, init, method: 'POST' }) };
        const graph = { api(uri) { const request = { kind: 'graph', id, client: graph, url: new URL(uri, 'https://tenant.test').href, uri, method: 'GET', headers: {} }; const query = { version(value) { request.version = value; return query; }, header(key, value) { request.headers[key] = value; return query; }, get() { return dispatch(request); } }; return query; } };
        const theme = { tryGetTheme: () => undefined, themeChangedEvent: { add() { }, remove() { } } };
        const services = new Map([[keys.page, page], [keys.sp, sp], [keys.theme, theme], [keys.graph, { getClient(version) { assert.equal(version, '3'); const request = Object.assign({ id, value: graph }, deferred()); initializations.push(request); return request.promise; } }]]);
        const scope = { whenFinished: callback => callback(), consume(key) { assert.ok(services.has(key), 'finite native service key'); return services.get(key); } };
        const domElement = document.createElement('div');
        return { instanceId: id, context: { instanceId: id, domElement, pageContext: page, serviceScope: scope, spHttpClient: sp, sdks: {}, propertyPane: { refresh() { } } }, properties: { title: 'Picker fixture', showTitle: false }, displayMode: 1, domElement, render() { }, native: { sp, graph, services } };
    }
    let instance = host(), revision = options.revision ?? 0, regions;
    const Demo = h.load(selected.file).Demo;
    let Shell, Explorer, groups;
    if (options.explorer) {
        Object.assign(toolkit, library('controls/web-part-shell/web-part-shell', true), library('hooks/useSPFxProperties'));
        const base = path.resolve(__dirname, '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
        Shell = h.load(path.join(base, 'WebPartShellDemoHost.tsx')).WebPartShellDemoHost;
        Explorer = h.load(path.join(base, 'ExampleExplorer.tsx')).ExampleExplorer;
        groups = h.load(path.join(base, 'generated/jsdoc-examples/catalogue.ts')).executableGroups.map(group => ({ ...group, features: group.features.filter(feature => feature.title === name) })).filter(group => group.features.length);
    }
    const renderer = createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
    const container = document.createElement('div');
    document.body.appendChild(container);
    function Frame() { return Explorer ? h.React.createElement(Explorer, { groups }) : h.React.createElement(Demo, { resetRevision: revision, render: options.fallback ? undefined : next => { regions = next; return h.React.createElement('section', undefined, next.controls, next.commands, next.status, next.children, next.diagnostics); } }); }
    function tree(child, owner = instance) { return h.React.createElement(RendererProvider, { renderer }, h.React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme }, h.React.createElement(Shell ?? toolkit.SPFxWebPartProvider, { instance: owner }, child))); }
    function render() { h.act(() => { ReactDOM.render(tree(h.React.createElement(Frame)), container); }); }
    const peers = new Set();
    let closed = false;
    function unmount() { if (closed)
        return; h.act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove(); closed = true; }
    t.after(() => { for (const peer of peers) peer.unmount(); unmount(); h.close(); });
    render();
    function findNode(label) { const node = [...container.querySelectorAll('input,select,button')].find(n => n.getAttribute('aria-label') === label || (n.id && [...container.querySelectorAll('label')].some(l => l.htmlFor === n.id && l.textContent === label))); assert.ok(node, label); return node; }
    return { h, toolkit, library, root, compiled, seen, calls, pending, initializations, container, host, render, unmount, subject, entry: selected.entry,
        mountPeer(owner = host('pickers-peer')) { const node = document.createElement('div'); document.body.appendChild(node); function Peer() { return h.React.createElement(Demo, { render: next => { return h.React.createElement('section', undefined, next.controls, next.commands, next.status, next.children, next.diagnostics); } }); } h.act(() => { ReactDOM.render(tree(h.React.createElement(Peer), owner), node); }); const peer = { container: node, instance: owner, get diagnostics() { return JSON.parse(node.querySelector('pre').textContent); }, click(label) { const button = [...node.querySelectorAll('button')].find(n => n.textContent === label); assert.ok(button); h.act(() => { button.click(); }); }, unmount() { h.act(() => { ReactDOM.unmountComponentAtNode(node); }); node.remove(); peers.delete(peer); } }; peers.add(peer); return peer; },
        get instance() { return instance; }, get regions() { return regions; }, get diagnostics() { const pre = container.querySelector('pre'); return pre ? JSON.parse(pre.textContent) : undefined; },
        hold(predicate) { hold = predicate; }, route(fn) { override = fn; },
        change(label, value) { const node = findNode(label); h.act(() => { node.value = value; Simulate.change(node, { target: { value } }); }); },
        toggle(label) { const node = findNode(label); h.act(() => { node.checked = !node.checked; Simulate.change(node, { target: { checked: node.checked } }); }); },
        action(label) { const nodes = regions.commands.props.children; const element = nodes.find(n => n?.props?.children === label); assert.ok(element, label); return element.props.onClick; },
        click(label) { const node = [...container.querySelectorAll('button')].find(n => n.textContent === label); assert.ok(node, label); h.act(() => { node.click(); }); },
        trigger() { return container.querySelector('[role="combobox"]'); },
        open() { const node = this.trigger(); assert.ok(node, 'actual Fluent picker trigger'); h.act(() => { node.click(); }); },
        search(value) { const node = document.querySelector('.fui-SearchBox input') ?? container.querySelector('input[role="combobox"]'); assert.ok(node, 'native picker search'); h.act(() => { node.value = value; Simulate.change(node, { target: { value } }); }); },
        choose(text) { const node = [...document.querySelectorAll('[role="option"],[role="menuitemcheckbox"]')].find(n => n.textContent.includes(text)); assert.ok(node, 'actual option ' + text); h.act(() => { node.click(); }); },
        reset() { revision++; render(); }, setRevision(value) { revision = value; render(); }, replace(next) { instance = next; render(); },
        async flush(fn = () => { }) { await h.act(async () => { fn(); for (let i = 0; i < 30; i++)
            await Promise.resolve(); await new Promise(resolve => setImmediate(resolve)); for (let i = 0; i < 30; i++)
            await Promise.resolve(); }); },
        async ready(error) { await this.flush(); await this.flush(() => initializations.filter(p => !p.settled).forEach(p => { p.settled = true; if (error)
            p.reject(error);
        else
            p.resolve(p.value); })); },
        async wait(ms = 330) { await h.act(async () => { await new Promise(resolve => setTimeout(resolve, ms)); }); await this.flush(); },
        async settle(request, failure) { await this.flush(() => failure ? request.reject(failure) : request.resolve(request.response)); },
        verifyUniverse() { assert.equal(toolkit.SPFxWebPartProvider, library('core/provider-webpart', true).SPFxWebPartProvider); assert.equal(toolkit.useSPFxContext, library('core/context.internal', true).useSPFxContext); },
        origins() {
            const packageRoot = path.dirname(root) + path.sep;
            const files = [...new Set(seen.filter(file => file.startsWith(packageRoot) && /\/(src|lib)\//.test(file)))];
            const required = [
                'core/provider-webpart', 'core/context.internal', 'hooks/useSPFxContext', 'hooks/useSPFxPageContext',
                'hooks/useSPFxSPHttpClient', 'hooks/useSPFxMSGraphClient', 'hooks/useSPFxServiceScope',
                'controls/pickers/usePickerRuntime.internal', 'controls/pickers/usePickerSourceContext.internal',
                'controls/pickers/picker-request-cache.internal', 'controls/pickers/picker-metadata-source.internal',
                'controls/pickers/picker-renderer.internal', 'controls/pickers/principal-identity.internal',
                ...Object.values(titles).flatMap(([file]) => ['controls/pickers/' + file, 'controls/pickers/' + file + '.source.internal']),
            ];
            for (const file of required) {
                const tsx = /provider-webpart|context.internal|picker-renderer.internal|controls\/pickers\/(site|list|field|content-type|view|list-item|principal)-picker$/.test(file);
                assert.ok(files.includes(path.join(root, file + (compiled ? '.js' : tsx ? '.tsx' : '.ts'))), 'actual origin ' + file);
            }
            // The established source harness intentionally evaluates emitted data-only AMD dictionaries.
            // Classify only the exact versioned inventory path for an actually requested native resource.
            const inventory = JSON.parse(originalRead(path.join(path.dirname(root), 'localization/resources.json'), 'utf8'));
            const resourceId = 'ApveeSpfxReactToolkitPickersStrings';
            const declared = inventory.resources.find(resource => resource.id === resourceId);
            assert.equal(declared.path, 'lib/controls/pickers/loc/{locale}.js');
            const nativeAssets = new Set();
            if (h.nativeResources.requestedIds.includes(resourceId)) {
                const asset = path.join(path.dirname(root), declared.path.replace('{locale}', 'en-us'));
                const ts = require('typescript');
                const source = ts.createSourceFile(asset, originalRead(asset, 'utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
                const statement = source.statements[0];
                assert.equal(source.statements.length, 1);
                assert.equal(ts.isExpressionStatement(statement), true);
                const define = statement.expression;
                assert.equal(ts.isCallExpression(define) && ts.isIdentifier(define.expression) && define.expression.text === 'define', true);
                assert.equal(define.arguments.length, 2);
                assert.equal(ts.isArrayLiteralExpression(define.arguments[0]) && define.arguments[0].elements.length === 0, true);
                const factory = define.arguments[1];
                assert.equal(ts.isFunctionExpression(factory) && factory.parameters.length === 0 && factory.body.statements.length === 1, true);
                const returned = factory.body.statements[0];
                assert.equal(ts.isReturnStatement(returned), true);
                const literal = node => ts.isStringLiteral(node) || ts.isObjectLiteralExpression(node) && node.properties.every(property =>
                    ts.isPropertyAssignment(property) && ts.isStringLiteral(property.name) && literal(property.initializer));
                assert.equal(literal(returned.expression), true, 'only nested literal string data, no implementation statements');
                assert.equal(h.nativeResources.require(resourceId).__locale, 'en-US');
                nativeAssets.add(asset);
            }
            assert.equal(files.filter(file => !nativeAssets.has(file)).some(file => file.includes(compiled ? '/src/' : '/lib/')), false, 'all observed implementation paths use matching mode');
            return files;
        },
    };
}
module.exports = { createReactPickersFixture };

// One existing filename-aware loader/DOM/provider universe; only native SDK/transport/form/Port boundaries are doubled.
const { createPickerHarness } = require('../picker-source-harness.cjs');
const { scenario } = require('./cards-description-list.cjs');
const fs = require('node:fs'), path = require('node:path'), assert = require('node:assert/strict');
const ReactDOM = require('react-dom');
const { Simulate } = require('react-dom/test-utils');
const { createDOMRenderer } = require('@griffel/core');
const { RendererProvider } = require('@griffel/react');
const fluent = require('@fluentui/react-components');
const { ids, web, endpoint, rootPath, metadata } = require('../file-folder-picker/metadata.cjs');
const personal = 'https://tenant-my.sharepoint.com/personal/user';
const title = 'Explore hosted files and folders with explicit source policies';
const owner = 'packages/spfx-react-toolkit/src/controls/file-folder-picker/file-folder-picker.tsx#FileFolderPicker';
function deferred() { let resolve, reject; const promise = new Promise((a, b) => { resolve = a; reject = b; }); return { promise, resolve, reject }; }
function json(value, status = 200) { return new Response(JSON.stringify(value), { status, headers: { 'Content-Type': 'application/json' } }); }
function createFileFolderFixture(t, options = {}) {
    const compiled = options.compiled ?? process.env.FILE_FOLDER_IMPLEMENTATION === 'compiled';
    Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
    const selected = scenario(title, owner), seen = [], originalRead = fs.readFileSync;
    t.mock.method(fs, 'readFileSync', (file, ...args) => {
        seen.push(String(file));
        let value = originalRead(file, ...args);
        if (String(file) === selected.file && typeof value === 'string' && options.mutate)
            value = options.mutate(value);
        if (typeof value === 'string' && options.runtimeMutation && String(file).endsWith('/file-folder-picker-session.internal.' + (compiled ? 'js' : 'ts')))
            value = options.runtimeMutation(value);
        return value;
    });
    const toolkit = {}, context = {}, pageModule = {}, keys = { page: {}, sp: {}, token: {}, theme: {} };
    const sdk = { SPHttpClient: { serviceKey: keys.sp, configurations: { v1: {} } }, AadTokenProviderFactory: { serviceKey: keys.token } };
    const h = createPickerHarness({ '@apvee/spfx-react-toolkit': toolkit, '@microsoft/sp-http': sdk,
        '@microsoft/sp-page-context': { PageContext: { serviceKey: keys.page } },
        '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 }, _SPKillSwitch: { isActivated: () => false } },
        '@microsoft/sp-component-base': { ThemeProvider: { serviceKey: keys.theme } }, './useSPFxContext': context, './useSPFxPageContext': pageModule });
    const root = path.resolve(__dirname, '../../../packages/spfx-react-toolkit', compiled ? 'lib' : 'src');
    const library = (file, tsx = false) => h.load(path.join(root, file + (compiled ? '.js' : tsx ? '.tsx' : '.ts')));
    Object.assign(context, library('hooks/useSPFxContext'));
    Object.assign(pageModule, library('hooks/useSPFxPageContext'));
    Object.assign(toolkit, context, pageModule, library('core/provider-webpart', true), library('hooks/useSPFxServiceScope'), library('helpers/spfx-page-context.helpers'), library('controls/file-folder-picker/file-folder-picker', true), library('controls/accordion-group/accordion-group', true));
    const calls = [], initializations = [], acquisitions = [], tokens = [], pending = [], posts = [], changes = [];
    const subject = { mounts: 0, unmounts: 0, props: undefined };
    const Actual = toolkit.FileFolderPicker;
    toolkit.FileFolderPicker = function ObservedPicker(props) {
        subject.props = props;
        h.React.useEffect(() => { subject.mounts++; return () => { subject.unmounts++; }; }, []);
        return h.React.createElement(Actual, { ...props, onChange(data) { changes.push(data); props.onChange?.(data); } });
    };
    let hold = options.hold, override = options.route;
    t.mock.method(global, 'fetch', url => { throw new Error('Unexpected live network: ' + url); });
    // Capture the real hidden POST at its local form boundary. No production session/config/navigation replacement.
    t.mock.method(window.HTMLFormElement.prototype, 'submit', function () {
        const url = new URL(this.action), frame = [...document.querySelectorAll('iframe')].find(f => f.name === this.target);
        assert.ok(frame, 'real Stage iframe target');
        assert.equal(this.method, 'post');
        const opening = subject.props?.scope.kind === 'oneDrive' ? personal : web;
        const expected = new URL(opening);
        assert.equal(url.origin, expected.origin, 'finite hosted POST opening origin');
        assert.equal(url.pathname, expected.pathname + '/_layouts/15/FilePicker.aspx', 'finite hosted POST opening path');
        assert.deepEqual([...url.searchParams.keys()].sort(), ['filePicker', 'locale']);
        assert.equal(url.searchParams.get('locale'), 'en-US');
        assert.equal(this.elements.access_token.value, 'fixture-resource-token');
        posts.push({ config: JSON.parse(url.searchParams.get('filePicker')), locale: url.searchParams.get('locale'), frame, action: url.origin + url.pathname });
    });
    const items = {
        Root: metadata('Root', { name: 'Shared Documents', file: undefined, folder: {}, root: {}, webUrl: web + '/Shared%20Documents', parentReference: { driveId: 'b!Drive' }, sharepointIds: { tenantId: ids.tenant, siteId: ids.site, webId: ids.web, listId: ids.list, listItemUniqueId: ids.root } }),
        Folder: metadata('Folder', { name: 'Policies', file: undefined, folder: {}, webUrl: web + '/Shared%20Documents/Policies' }),
        FileA: metadata(), FileB: metadata('FileB'), Child: metadata('Child', { parentReference: { driveId: 'b!Drive', id: 'Folder' } }),
    };
    function route(request) {
        if (override) {
            const value = override(request);
            if (value !== undefined)
                return value;
        }
        const url = new URL(request.url), route = decodeURIComponent(url.pathname).replace(/^\/sites\/hr/i, '/sites/HR'), isPersonal = url.origin === 'https://tenant-my.sharepoint.com';
        assert.ok(['https://tenant.sharepoint.com', 'https://tenant-my.sharepoint.com'].includes(url.origin), 'finite metadata origin');
        assert.equal(request.method, 'GET');
        assert.equal(request.configuration, sdk.SPHttpClient.configurations.v1);
        assert.deepEqual(request.init.headers, isPersonal ? { Accept: 'application/json;odata=nometadata', 'OData-Version': '', Authorization: 'Bearer fixture-resource-token' } : { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'OData-Version': '' });
        const responseWeb = isPersonal ? personal : web, responsePath = isPersonal ? '/personal/user' : '/sites/HR';
        const base = isPersonal ? '/personal/user' : '/sites/HR';
        const allowedSelect = (value, expand) => { assert.equal(url.searchParams.get('$select'), value); assert.deepEqual([...url.searchParams.keys()].sort(), expand ? ['$expand', '$select'] : ['$select']); };
        if (route === web.slice(new URL(web).origin.length) + '/_api/SP.UserProfiles.PeopleManager/GetMyProperties') {
            allowedSelect('PersonalUrl,Email');
            return json({ PersonalUrl: personal, Email: 'user@tenant.test' });
        }
        if (route === base + '/_api/site') {
            allowedSelect('Id,Url');
            return json({ Id: ids.site, Url: responseWeb });
        }
        if (route === base + '/_api/site/rootWeb' || route === base + '/_api/web') {
            allowedSelect('Id,Url,ServerRelativeUrl');
            return json({ Id: ids.web, Url: responseWeb, ServerRelativeUrl: responsePath });
        }
        if (route === base + "/_api/web/lists/getByTitle('Documents')") {
            allowedSelect('Id,Title,BaseTemplate');
            return json({ Id: ids.list, Title: 'Documents', BaseTemplate: 850 });
        }
        if (route === base + `/_api/web/lists(guid'${ids.list}')`) {
            allowedSelect('BaseType,RootFolder/ServerRelativeUrl,RootFolder/UniqueId', 'RootFolder');
            assert.equal(url.searchParams.get('$expand'), 'RootFolder');
            return json({ BaseType: 1, RootFolder: { ServerRelativeUrl: rootPath, UniqueId: ids.root } });
        }
        if (route === base + `/_api/v2.0/sites/tenant.sharepoint.com,${ids.site},${ids.web}/lists/${ids.list}/drive`) {
            assert.equal(url.search, '');
            return json({ id: 'b!Drive', driveType: 'documentLibrary' });
        }
        if (route === base + '/_api/v2.0/drive') {
            allowedSelect('id,driveType,owner,sharepointIds');
            return json({ id: 'b!Drive', driveType: 'business', owner: { user: { email: 'user@tenant.test' } } });
        }
        if (route === base + '/_api/v2.0/drives/b!Drive/root') {
            assert.deepEqual([...url.searchParams.keys()], ['$select']);
            assert.ok(['id,name,folder,root,parentReference,sharepointIds,webUrl,remoteItem,deleted', 'id,folder,root,parentReference,deleted,remoteItem', 'id,folder,root,parentReference,sharepointIds,deleted,remoteItem'].includes(url.searchParams.get('$select')));
            return json(isPersonal ? { ...items.Root, webUrl: personal + '/Documents' } : items.Root);
        }
        if (route === base + '/_api/v2.0/drives/b!Drive/root:/Policies') {
            allowedSelect('id,name,folder,root,parentReference,sharepointIds,webUrl,remoteItem,deleted');
            return json(items.Folder);
        }
        if (route === base + "/_api/web/GetFolderByServerRelativePath(decodedUrl='/sites/HR/Shared Documents/Policies')") {
            allowedSelect('ServerRelativeUrl,UniqueId');
            return json({ ServerRelativeUrl: rootPath + '/Policies', UniqueId: ids.folder });
        }
        const itemPrefix = base + '/_api/v2.0/drives/b!Drive/items/';
        if (route.startsWith(itemPrefix)) {
            allowedSelect('id,name,file,folder,parentReference,sharepointIds,size,webUrl,remoteItem,root,deleted');
            const id = route.slice(itemPrefix.length);
            assert.ok(Object.hasOwn(items, id) || id === 'root', 'finite item ID');
            return json(id === 'root' ? items.Root : items[id]);
        }
        throw new Error('Unsupported FileFolder fixture route: ' + request.method + ' ' + request.url);
    }
    function dispatch(request) {
        calls.push(request);
        const response = route(request);
        if (hold?.(request)) {
            const entry = Object.assign(request, deferred(), { response });
            pending.push(entry);
            return entry.promise;
        }
        return Promise.resolve(response);
    }
    function host(id = 'file-folder-one') {
        const page = { web: { absoluteUrl: web, title: 'HR' }, site: { absoluteUrl: web }, user: { loginName: 'user@tenant.test', displayName: 'User', email: 'user@tenant.test' }, aadInfo: { tenantId: { toString: () => ids.tenant } }, cultureInfo: { isRightToLeft: false, currentUICultureName: 'en-US' } };
        const sp = { get(url, configuration, init) { return dispatch({ id, client: sp, url, configuration, init, method: 'GET' }); }, post() { throw new Error('No native content write route'); } };
        const provider = { getToken(resource, config) {
                assert.ok(['https://tenant.sharepoint.com', 'https://tenant-my.sharepoint.com', 'https://graph.microsoft.com'].includes(resource));
                const request = Object.assign({ id, resource, config, response: 'fixture-resource-token' }, deferred());
                tokens.push(request);
                if (!options.holdToken?.(request))
                    request.resolve(request.response);
                return request.promise;
            } };
        const theme = { tryGetTheme: () => undefined, themeChangedEvent: { add() { }, remove() { } } };
        const services = new Map([[keys.page, page], [keys.sp, sp], [keys.theme, theme], [keys.token, { getTokenProvider() { const request = Object.assign({ id, value: provider }, deferred()); initializations.push(request); return request.promise; } }]]);
        const scope = { whenFinished: callback => callback(), consume(key) { assert.ok(services.has(key), 'finite native SDK key'); acquisitions.push({ id, key }); return services.get(key); } };
        const domElement = document.createElement('div');
        return { instanceId: id, context: { instanceId: id, domElement, pageContext: page, serviceScope: scope, spHttpClient: sp, sdks: {}, propertyPane: { refresh() { } } }, properties: { title: 'File fixture', showTitle: false }, displayMode: 1, domElement, render() { }, native: { sp, provider, services, page } };
    }
    let instance = host(), revision = options.revision ?? 0, regions;
    const Demo = h.load(selected.file).Demo;
    let Shell, Explorer, groups;
    if (options.explorer) {
        Object.assign(toolkit, library('controls/web-part-shell/web-part-shell', true), library('hooks/useSPFxProperties'));
        const base = path.resolve(__dirname, '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
        Shell = h.load(path.join(base, 'WebPartShellDemoHost.tsx')).WebPartShellDemoHost;
        Explorer = h.load(path.join(base, 'ExampleExplorer.tsx')).ExampleExplorer;
        groups = h.load(path.join(base, 'generated/jsdoc-examples/catalogue.ts')).executableGroups.map(group => ({ ...group, features: group.features.filter(feature => feature.title === 'FileFolderPicker') })).filter(group => group.features.length);
    }
    const renderer = createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
    const container = document.createElement('div');
    document.body.appendChild(container);
    function Frame() { return Explorer ? h.React.createElement(Explorer, { groups }) : h.React.createElement(Demo, { resetRevision: revision, render: options.fallback ? undefined : next => { regions = next; return h.React.createElement('section', undefined, next.controls, next.commands, next.status, next.children, next.diagnostics); } }); }
    function tree(child, owner = instance) { return h.React.createElement(RendererProvider, { renderer }, h.React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme }, h.React.createElement(Shell ?? toolkit.SPFxWebPartProvider, { instance: owner }, child))); }
    function render() { h.act(() => { ReactDOM.render(tree(h.React.createElement(Frame)), container); }); }
    let closed = false;
    const peers = new Set();
    function unmount() {
        if (closed)
            return;
        h.act(() => { ReactDOM.unmountComponentAtNode(container); });
        container.remove();
        closed = true;
    }
    t.after(() => {
        for (const peer of peers)
            peer.unmount();
        unmount();
        h.close();
    });
    render();
    // Operate the real native disclosure, rather than assuming collapsed Fluent panels keep DOM mounted.
    function readDiagnostics(root = container) {
        const header = [...root.querySelectorAll('button')].find(node => node.textContent === 'Diagnostics and selection JSON');
        assert.ok(header, 'native diagnostics disclosure');
        const collapsed = header.getAttribute('aria-expanded') === 'false';
        if (collapsed)
            h.act(() => { header.click(); });
        try {
            const pre = root.querySelector('pre');
            assert.ok(pre, 'rendered public diagnostics');
            return JSON.parse(pre.textContent);
        }
        finally {
            if (collapsed)
                h.act(() => { header.click(); });
        }
    }
    function findNode(label) {
        const find = () => [...container.querySelectorAll('input,select,button')].find(n => n.getAttribute('aria-label') === label || (n.id && [...container.querySelectorAll('label')].some(l => l.htmlFor === n.id && l.textContent === label)));
        let node = find();
        if (!node || node.closest('.fui-AccordionPanel')?.hasAttribute('inert')) {
            const options = [...container.querySelectorAll('button')].find(n => n.textContent === 'Options');
            if (options?.getAttribute('aria-expanded') === 'false')
                h.act(() => { options.click(); });
            node = find();
        }
        assert.ok(node, label);
        return node;
    }
    function opening(frame = document.querySelector('iframe')) {
        assert.ok(frame, 'actual Stage iframe');
        const source = frame.contentWindow, channelId = frame.name;
        const port = { messages: [], start() { this.started = true; }, close() { this.closed = true; }, postMessage(message) { this.messages.push(message); } };
        return { frame, port, source, channelId,
            initialize(extra = {}) { h.act(() => { window.dispatchEvent(new window.MessageEvent('message', { origin: 'https://tenant.sharepoint.com', source, data: { type: 'initialize', channelId }, ports: [port], ...extra })); }); },
            notify() { h.act(() => { port.onmessage?.({ data: { type: 'notification', data: { notification: 'page-loaded' } } }); }); },
            command(id, data) { h.act(() => { port.onmessage?.({ data: { type: 'command', id, data } }); }); },
        };
    }
    return { h, toolkit, library, root, compiled, seen, calls, pending, initializations, acquisitions, tokens, posts, changes, items, subject, container, host, render, unmount, opening,
        get instance() { return instance; }, get regions() { return regions; }, get diagnostics() { return readDiagnostics(); },
        hold(predicate) { hold = predicate; }, route(fn) { override = fn; }, node: findNode,
        change(label, value) { const node = findNode(label); h.act(() => { node.value = value; Simulate.change(node, { target: { value } }); }); },
        toggle(label) { const node = findNode(label); h.act(() => { node.checked = !node.checked; Simulate.change(node, { target: { checked: node.checked } }); }); },
        action(label) { const element = regions.commands.props.children.find(n => n?.props?.children === label); assert.ok(element, label); return element.props.onClick; },
        click(label) { const node = [...container.querySelectorAll('button')].find(n => n.textContent === label); assert.ok(node, label); h.act(() => node.click()); },
        trigger() { return container.querySelector('button[aria-haspopup="dialog"]'); },
        browse() { const node = this.trigger(); assert.ok(node, 'actual native Browse'); assert.equal(node.disabled, false); h.act(() => node.click()); },
        reset() { revision++; render(); }, replace(next) { instance = next; render(); },
        async flush(fn = () => { }) {
            await h.act(async () => {
                fn();
                for (let i = 0; i < 30; i++)
                    await Promise.resolve();
                await new Promise(resolve => setImmediate(resolve));
                for (let i = 0; i < 30; i++)
                    await Promise.resolve();
            });
        },
        async ready() { await this.flush(); await this.flush(() => initializations.filter(p => !p.settled).forEach(p => { p.settled = true; p.resolve(p.value); })); },
        async settle(request, error) { await this.flush(() => error ? request.reject(error) : request.resolve(request.response)); },
        mountPeer(owner = host('file-folder-peer')) {
            const node = document.createElement('div');
            document.body.appendChild(node);
            let peerRevision = 0;
            const peer = { container: node, instance: owner, get diagnostics() { return readDiagnostics(node); },
                render() { h.act(() => { ReactDOM.render(tree(h.React.createElement(Demo, { resetRevision: peerRevision }), owner), node); }); },
                click(label) { const button = [...node.querySelectorAll('button')].find(n => n.textContent === label); assert.ok(button); h.act(() => button.click()); },
                browse() { const button = node.querySelector('button[aria-haspopup="dialog"]'); assert.ok(button); h.act(() => button.click()); }, reset() { peerRevision++; this.render(); },
                unmount() { h.act(() => { ReactDOM.unmountComponentAtNode(node); }); node.remove(); peers.delete(peer); } };
            peers.add(peer);
            peer.render();
            return peer;
        },
        sourceContext(owner = instance) { return { owner: owner.context, spClient: owner.native.sp, tokenProvider: owner.native.provider, tenantId: ids.tenant, userId: owner.native.page.user.loginName, webUrl: web, siteCollectionUrl: web }; },
        origins() {
            assert.equal(toolkit.SPFxWebPartProvider, library('core/provider-webpart', true).SPFxWebPartProvider);
            assert.equal(toolkit.useSPFxContext, library('core/context.internal', true).useSPFxContext);
            const packageRoot = path.dirname(root) + path.sep;
            const files = [...new Set(seen.filter(file => file.startsWith(packageRoot) && /\/(src|lib)\//.test(file)))];
            const required = { 'controls/accordion-group/accordion-group': true, 'core/provider-webpart': true, 'core/provider-base.internal': true, 'core/context.internal': true, 'controls/file-folder-picker/file-folder-picker': true, 'controls/stage/stage': true };
            for (const stem of ['useFileFolderPicker.internal', 'file-folder-picker-scope.internal', 'file-folder-picker-results.internal', 'file-folder-picker-config.internal', 'file-folder-picker-presets.internal', 'file-folder-picker-session.internal', 'file-folder-picker-protocol.internal', 'file-folder-picker-navigation.internal', 'useFileFolderPickerTheme.internal'])
                required['controls/file-folder-picker/' + stem] = false;
            for (const stem of ['useSPFxContext', 'useSPFxPageContext', 'useSPFxServiceScope', 'useSPFxSPHttpClient', 'useSPFxAadTokenProvider'])
                required['hooks/' + stem] = false;
            for (const stem of ['usePickerSourceContext.internal', 'picker-request-cache.internal'])
                required['controls/pickers/' + stem] = false;
            required['helpers/spfx-page-context.helpers'] = false;
            for (const [stem, tsx] of Object.entries(required))
                assert.ok(files.includes(path.join(root, stem + (compiled ? '.js' : tsx ? '.tsx' : '.ts'))), 'observed actual origin ' + stem);
            const inventory = JSON.parse(originalRead(path.join(path.dirname(root), 'localization/resources.json'), 'utf8'));
            const nativeAssets = new Set(), dataRoles = [];
            for (const [id, expected] of [['ApveeSpfxReactToolkitFileFolderPickerStrings', 'lib/controls/file-folder-picker/loc/{locale}.js'], ['ApveeSpfxReactToolkitPickersStrings', 'lib/controls/pickers/loc/{locale}.js'], ['ApveeSpfxReactToolkitStageStrings', 'lib/controls/stage/loc/{locale}.js']]) {
                if (!h.nativeResources.requestedIds.includes(id))
                    continue;
                const declared = inventory.resources.find(r => r.id === id);
                assert.equal(declared.path, expected);
                const asset = path.join(path.dirname(root), declared.path.replace('{locale}', 'en-us')), ts = require('typescript'), bytes = originalRead(asset), source = ts.createSourceFile(asset, bytes.toString(), ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
                assert.equal(source.statements.length, 1);
                const statement = source.statements[0];
                assert.equal(ts.isExpressionStatement(statement), true);
                const define = statement.expression;
                assert.equal(ts.isCallExpression(define) && ts.isIdentifier(define.expression) && define.expression.text === 'define', true);
                assert.equal(define.arguments.length, 2);
                assert.equal(ts.isArrayLiteralExpression(define.arguments[0]) && define.arguments[0].elements.length === 0, true);
                const factory = define.arguments[1];
                assert.equal(ts.isFunctionExpression(factory) && factory.parameters.length === 0 && factory.body.statements.length === 1, true);
                const returned = factory.body.statements[0];
                assert.equal(ts.isReturnStatement(returned), true);
                let leaves = 0;
                const literal = node => {
                    if (ts.isStringLiteral(node)) {
                        leaves++;
                        return true;
                    }
                    return ts.isObjectLiteralExpression(node) && node.properties.every(p => ts.isPropertyAssignment(p) && ts.isStringLiteral(p.name) && literal(p.initializer));
                };
                assert.equal(literal(returned.expression), true, 'only literal string data; zero implementation');
                dataRoles.push({ id, asset, bytes: bytes.length, sha: require('crypto').createHash('sha256').update(bytes).digest('hex'), leaves, implementation: 0 });
                nativeAssets.add(asset);
            }
            assert.equal(dataRoles.some(r => r.id === 'ApveeSpfxReactToolkitFileFolderPickerStrings'), true);
            assert.equal(files.filter(file => !nativeAssets.has(file)).some(file => file.includes(compiled ? '/src/' : '/lib/')), false, 'all observed implementations use matching mode: ' + files.filter(file => !nativeAssets.has(file) && file.includes(compiled ? '/src/' : '/lib/')).join(', '));
            return { files, dataRoles };
        },
    };
}
module.exports = { createFileFolderFixture, title, owner, ids, web, endpoint, rootPath };

const test = require('node:test');
const assert = require('node:assert/strict');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const titles = {
    useSPFxAadTokenProvider: 'Inspect client readiness and token-provider initialization',
    useSPFxHttpClient: 'Load an external todo with tracked invocation state',
    useSPFxSPHttpClient: 'Read the current SharePoint web title and URL',
    useSPFxMSGraphClient: 'Read the current Graph profile after initialization',
    useSPFxAadHttpClient: 'Initialize a custom API client and call a path',
    useSPFxApiPermissionPrecheck: 'Manually check Graph and custom API delegated scopes',
};
test('Clients initialization checkpoint has the two actual source-owned subjects', () => {
    for (const owner of ['useSPFxAadTokenProvider', 'useSPFxAadHttpClient'])
        assert.ok(scenario(titles[owner]).file);
});
const { createClientsFixture } = require('./fixtures/executable-examples/clients-scenarios.cjs');
for (const owner of ['useSPFxAadTokenProvider', 'useSPFxAadHttpClient'])
    test(`${owner} masks ready A on the first B commit before effects`, async (t) => {
        const f = createClientsFixture(t, titles[owner]);
        await f.ready();
        if (owner === 'useSPFxAadHttpClient') {
            f.change('Resource URL or App ID URI', 'https://api.test');
            f.click('Initialize AAD client');
            await f.ready();
        }
        assert.match(f.container.textContent, owner === 'useSPFxAadTokenProvider' ? /Token provider: Ready/ : /Ready: Yes/);
        const start = f.commits.length;
        const input = f.container.querySelector('input');
        f.changeScope(f.host('clients-two'));
        if (input)
            assert.equal(f.container.querySelector('input'), input, 'same mounted native input');
        const first = f.commits[start];
        assert.ok(first);
        assert.doesNotMatch(first.text, owner === 'useSPFxAadTokenProvider' ? /Token provider: Ready/ : /Ready: Yes/);
        await f.ready();
        assert.match(f.container.textContent, owner === 'useSPFxAadTokenProvider' ? /Token provider: Ready/ : /Ready: Yes/);
    });
for (const owner of ['useSPFxAadTokenProvider', 'useSPFxAadHttpClient'])
    test(`${owner} initialization checkpoint detects omitted source masking`, async (t) => {
        const f = createClientsFixture(t, titles[owner], { unmasked: true });
        await f.ready();
        if (owner === 'useSPFxAadHttpClient') {
            f.change('Resource URL or App ID URI', 'https://api.test');
            f.click('Initialize AAD client');
            await f.ready();
        }
        const start = f.commits.length;
        f.changeScope(f.host('clients-two'));
        assert.match(f.commits[start].text, owner === 'useSPFxAadTokenProvider' ? /Token provider: Ready/ : /Ready: Yes/, 'negative control exposes the previous native ready state before effects');
    });
test('AAD resource A to B retains controls and masks previous readiness before effects', async (t) => {
    const f = createClientsFixture(t, titles.useSPFxAadHttpClient);
    await f.ready();
    f.change('Resource URL or App ID URI', 'https://a.test');
    f.click('Initialize AAD client');
    await f.ready();
    assert.match(f.container.textContent, /Ready: Yes/);
    const input = f.container.querySelector('input');
    input.focus();
    f.change('Resource URL or App ID URI', 'https://b.test');
    const start = f.commits.length;
    f.click('Initialize AAD client');
    assert.equal(f.container.querySelector('input'), input);
    assert.equal(document.activeElement, input);
    assert.doesNotMatch(f.commits[start].text, /Ready: Yes/);
    assert.equal(f.calls.length, 0);
    await f.ready();
    assert.match(f.container.textContent, /Ready: Yes/);
});
test('exact six Clients owners expose explicit source-owned cases', () => {
    for (const [owner, title] of Object.entries(titles)) {
        const entry = scenario(title).entry;
        assert.equal(entry.featureTitle, owner);
        assert.equal(entry.metadata.host, 'webPart');
        assert.equal(entry.metadata.contentPlacement, 'configuration');
        assert.equal(entry.metadata.kind, owner === 'useSPFxAadTokenProvider' ? 'snapshot' : 'read');
    }
});
for (const [owner, command] of [['useSPFxHttpClient', 'Load external todo'], ['useSPFxSPHttpClient', 'Load current web'], ['useSPFxMSGraphClient', 'Load /me'], ['useSPFxAadHttpClient', 'Call path']])
    test(`${owner} actual request newest result and pending truth survive reverse completion and Reset`, async (t) => {
        const f = createClientsFixture(t, titles[owner]);
        await f.ready();
        if (owner === 'useSPFxAadHttpClient') {
            f.change('Resource URL or App ID URI', 'https://api.test');
            f.click('Initialize AAD client');
            await f.ready();
        }
        assert.equal(f.calls.length, 0);
        const action = f.action(command);
        let a, b;
        f.h.act(() => { a = action(); b = action(); });
        assert.equal(f.calls.length, 2);
        assert.match(f.container.textContent, /Loading/);
        await f.flush(() => { if (owner === 'useSPFxMSGraphClient')
            f.calls[1].resolve({ title: 'Newest', completed: false, userId: 0, empty: '' });
        else
            f.response(1, { title: 'Newest', completed: false, userId: 0, empty: '' }); });
        assert.match(f.container.textContent, /Newest/);
        assert.match(f.container.textContent, /Loading/);
        await f.flush(() => f.calls[0].reject(new Error('Obsolete failure')));
        await Promise.all([a, b]);
        assert.match(f.container.textContent, /Newest/);
        assert.doesNotMatch(f.container.textContent, /Obsolete failure|Loading: underlying/);
        f.reset();
        await f.flush(() => action());
        assert.equal(f.calls.length, 2, 'original action is invalidated before native I/O');
        assert.doesNotMatch(f.container.textContent, /Newest/);
    });
test('mounted precheck completes through actual native tokens and blocks reset retry until original promise settles', async (t) => {
    const f = createClientsFixture(t, titles.useSPFxApiPermissionPrecheck);
    await f.ready();
    assert.equal(f.tokens.length, 0);
    f.click('Check permissions');
    assert.equal(f.tokens.length, 1);
    assert.equal(f.instance.native.popup.size, 1);
    f.reset();
    assert.match(f.container.textContent, /Previous check may still be completing/);
    assert.ok([...f.container.querySelectorAll('button')].find(n => n.textContent === 'Check permissions').disabled);
    await f.flush(() => f.tokens[0].resolve(require('./fixtures/executable-examples/clients-scenarios.cjs').token()));
    assert.equal(f.instance.native.popup.size, 0);
    assert.equal(f.instance.native.redirect.size, 0);
    assert.doesNotMatch(f.container.textContent, /available to the current/);
    f.click('Retry without cache');
    assert.equal(f.tokens[1].settings.useCachedToken, false);
    await f.flush(() => f.tokens[1].resolve(require('./fixtures/executable-examples/clients-scenarios.cjs').token()));
    assert.match(f.container.textContent, /ready/);
    assert.match(f.container.textContent, /User.Read/);
});
const commands = { useSPFxHttpClient: 'Load external todo', useSPFxSPHttpClient: 'Load current web', useSPFxMSGraphClient: 'Load /me', useSPFxAadHttpClient: 'Call path' };
async function requestFixture(t, owner, options) { const f = createClientsFixture(t, titles[owner], options); await f.ready(); if (owner === 'useSPFxAadHttpClient') {
    f.change('Resource URL or App ID URI', 'https://api.test');
    f.click('Initialize AAD client');
    await f.ready();
} return f; }
for (const owner of Object.keys(commands))
    test(`${owner} retained commands cannot start old scope work; old rejection after Reset is caught`, async (t) => {
        const f = await requestFixture(t, owner);
        const old = f.action(commands[owner]);
        let pending;
        f.h.act(() => { pending = old(); });
        assert.equal(f.calls.length, 1);
        f.reset();
        await f.flush(() => f.calls[0].reject(new Error('Reset obsolete')));
        await pending;
        assert.doesNotMatch(f.container.textContent, /Reset obsolete/);
        await f.flush(() => old());
        assert.equal(f.calls.length, 1);
        // Capture a current command, then change native scope without asking the provider to remount its child.
        if (owner === 'useSPFxAadHttpClient') {
            f.change('Resource URL or App ID URI', 'https://api.test');
            f.click('Initialize AAD client');
            await f.ready();
        }
        const current = f.action(commands[owner]);
        f.changeScope(f.host('second'));
        await f.ready();
        await f.flush(() => current());
        assert.equal(f.calls.length, 1);
        const last = f.action(commands[owner]);
        f.unmount();
        await f.flush(() => last());
        assert.equal(f.calls.length, 1);
    });
for (const owner of Object.keys(commands))
    test(`${owner} latest rejection remains after older success and all pending settles`, async (t) => {
        const f = await requestFixture(t, owner);
        const action = f.action(commands[owner]);
        let first, last;
        f.h.act(() => { first = action(); last = action(); });
        await f.flush(() => f.calls[1].reject(new Error('Latest failed')));
        assert.match(f.container.textContent, /Latest failed/);
        assert.match(f.container.textContent, /Loading: underlying/);
        await f.flush(() => owner === 'useSPFxMSGraphClient' ? f.calls[0].resolve({ title: 'Old success' }) : f.response(0, { title: 'Old success' }));
        await Promise.all([first, last]);
        assert.match(f.container.textContent, /Latest failed/);
        assert.doesNotMatch(f.container.textContent, /Old success|Loading: underlying/);
    });
test('SharePoint current web changes on the same native client invalidate old URLs without I/O', async (t) => {
    const f = await requestFixture(t, 'useSPFxSPHttpClient');
    const old = f.action('Load current web');
    const native = f.instance.native.services.get(f.keys.sp);
    f.instance.context.pageContext.web.absoluteUrl = 'https://tenant.test/sites/next/';
    await f.flush(() => old());
    assert.equal(f.calls.length, 0, 'live page URL is checked even before child redraw');
    f.render();
    assert.equal(f.instance.native.services.get(f.keys.sp), native);
    assert.equal(f.calls.length, 0);
    f.click('Load current web');
    assert.equal(f.calls[0].url, 'https://tenant.test/sites/next/_api/web?$select=Title,Url');
    await f.flush(() => f.response(0, { Title: 'Next', Url: 'https://tenant.test/sites/next' }));
    assert.match(f.container.textContent, /Next/);
});
test('AAD drafts invalidate retained initialization/path actions and normalize one slash', async (t) => {
    const f = await requestFixture(t, 'useSPFxAadHttpClient');
    const original = f.action('Call path'), initialize = f.action('Initialize AAD client');
    const count = f.initializations.length;
    f.change('Resource URL or App ID URI', 'https://b.test');
    await f.flush(() => initialize());
    await f.flush(() => original());
    assert.equal(f.calls.length, 0);
    assert.equal(f.initializations.length, count);
    f.click('Initialize AAD client');
    await f.ready();
    for (const path of ['api/health', '/api/health']) {
        f.change('Path', path);
        f.click('Call path');
        assert.equal(f.calls.at(-1).url, 'https://b.test/api/health');
        await f.flush(() => f.response(f.calls.length - 1, { status: 'healthy' }));
    }
});
test('actual Graph requests retain native projection and token status selection requests no business data', async (t) => {
    const status = createClientsFixture(t, titles.useSPFxAadTokenProvider, { revision: 7 });
    await status.ready();
    assert.equal(status.calls.length, 0);
    assert.equal(status.tokens.length, 0);
    assert.ok(status.initializations.some(request => request.kind === 'graph' && request.target === '3'));
    status.reset();
    assert.match(status.container.textContent, /Token provider: Ready/);
    status.unmount();
    const graph = await requestFixture(t, 'useSPFxMSGraphClient');
    graph.click('Load /me');
    assert.equal(graph.calls[0].url, '/me');
    assert.equal(graph.calls[0].fields, 'displayName,mail,userPrincipalName');
    await graph.flush(() => graph.calls[0].resolve({ displayName: 'Ada', mail: 'a@test', userPrincipalName: 'ada@test' }));
    assert.match(graph.container.textContent, /Ada/);
});
test('HTTP scalar rows preserve false zero empty and eight fields; private/cycle diagnostics stay collapsed', async (t) => {
    const f = await requestFixture(t, 'useSPFxHttpClient');
    const payload = { title: 'Visible response', completed: false, userId: 0, empty: '', a: 1, b: 2, c: 3, d: 4, ninth: 'omitted', nested: { diagnosticOnly: 'nested payload', _secret: 'hidden', serviceScope: { secret: 'hidden' } } };
    payload.nested.cycle = payload;
    f.click('Load external todo');
    await f.flush(() => f.response(0, payload));
    const result = f.container.querySelector('[aria-label="HttpClient results"]');
    assert.match(result.textContent, /Visible response|false|0/);
    assert.ok([...result.querySelectorAll('dt')].some(n => n.textContent === 'empty'));
    assert.doesNotMatch(result.textContent, /nested payload|omitted/);
    const trigger = [...f.container.querySelectorAll('button')].find(node => node.textContent === 'Advanced diagnostics');
    assert.equal(trigger.getAttribute('aria-expanded'), 'false');
    f.click('Advanced diagnostics');
    f.click('Raw response');
    const raw = f.container.querySelector('[aria-label="Raw response JSON"]');
    assert.match(raw.textContent, /nested payload/);
    assert.match(raw.textContent, /\[Private\]/);
    assert.match(raw.textContent, /\[Circular\]/);
    assert.doesNotMatch(raw.textContent, /hidden/);
});
test('independent actual providers retain separate native clients and pending outcomes', async (t) => {
    const f = await requestFixture(t, 'useSPFxHttpClient');
    const peer = f.mountPeer();
    await f.ready();
    f.click('Load external todo');
    peer.click('Load external todo');
    assert.equal(f.calls.length, 2);
    assert.notEqual(f.calls[0].client, f.calls[1].client);
    await f.flush(() => f.response(1, { title: 'Peer only' }));
    assert.match(peer.container.textContent, /Peer only/);
    assert.doesNotMatch(f.container.textContent, /Peer only/);
    await f.flush(() => f.response(0, { title: 'Main only' }));
    assert.match(f.container.textContent, /Main only/);
    assert.doesNotMatch(peer.container.textContent, /Main only/);
    peer.unmount();
});
for (const owner of Object.keys(titles))
    test(`${owner} portable fallback is mounted with the actual provider`, async (t) => { const f = createClientsFixture(t, titles[owner], { fallback: true, revision: 5 }); await f.ready(); assert.ok(f.container.textContent.length > 20); assert.equal(f.calls.length, 0); assert.equal(f.tokens.length, 0); });
test('precheck uses sequential resources, passive guards, uncached retry and original promise settlement after target edit', async (t) => {
    const { token } = require('./fixtures/executable-examples/clients-scenarios.cjs');
    const f = createClientsFixture(t, titles.useSPFxApiPermissionPrecheck);
    await f.ready();
    f.change('Resource URL or App ID URI', ' https://custom.test ');
    const action = f.action('Check permissions');
    let pending;
    f.h.act(() => { pending = action(); });
    assert.equal(f.tokens.length, 1);
    let popup = 0, redirect = 0;
    for (const handler of f.instance.native.popup.keys())
        handler({ cancel() { popup++; } });
    for (const handler of f.instance.native.redirect.keys())
        handler({ cancel() { redirect++; } });
    assert.equal(popup, 1);
    assert.equal(redirect, 1);
    f.change('Resource URL or App ID URI', 'https://other.test');
    await f.flush(() => action());
    assert.equal(f.tokens.length, 1, 'retained action cannot start another check');
    await f.flush(() => f.tokens[0].resolve(token()));
    assert.equal(f.tokens.length, 2);
    assert.equal(f.tokens[1].resource, 'https://custom.test');
    assert.equal(f.instance.native.popup.size, 1, 'guard spans the complete sequential check');
    await f.flush(() => f.tokens[1].resolve(token('user_impersonation', 'https://custom.test')));
    await pending;
    assert.equal(f.instance.native.popup.size, 0);
    assert.ok(f.observers.every(observer => observer.isDisposed));
    assert.doesNotMatch(f.container.textContent, /available to the current/);
    f.click('Retry without cache');
    assert.equal(f.tokens[2].settings.useCachedToken, false);
    await f.flush(() => f.tokens[2].resolve(token('Wrong.Scope')));
    assert.equal(f.tokens[3].settings.useCachedToken, false);
    await f.flush(() => f.tokens[3].resolve(token('wrong', 'https://other.test')));
    assert.match(f.container.textContent, /actionRequired/);
    assert.match(f.container.textContent, /not present/);
    assert.doesNotMatch(f.container.textContent, /\.signature/);
});
test('precheck timeout detaches guards when public promise settles, while original token acquisition can still settle later', async (t) => {
    const { token } = require('./fixtures/executable-examples/clients-scenarios.cjs');
    const f = createClientsFixture(t, titles.useSPFxApiPermissionPrecheck);
    await f.ready();
    t.mock.timers.enable({ apis: ['setTimeout'] });
    f.click('Check permissions');
    const native = f.tokens[0];
    f.reset();
    await f.flush(() => t.mock.timers.tick(15000));
    assert.equal(f.instance.native.popup.size, 0);
    assert.equal(f.instance.native.redirect.size, 0);
    assert.doesNotMatch(f.container.textContent, /Previous check may still/);
    await f.flush(() => native.resolve(token()));
    assert.doesNotMatch(f.container.textContent, /available to the current/);
});
test('precheck catches event-registration failures and cleans partially registered passive guards', async (t) => {
    const f = createClientsFixture(t, titles.useSPFxApiPermissionPrecheck);
    await f.ready();
    f.instance.native.provider.onBeforeRedirectEvent.add = () => { throw new Error('Registration failed'); };
    f.click('Check permissions');
    await f.flush();
    assert.match(f.container.textContent, /Registration failed/);
    assert.equal(f.tokens.length, 0);
    assert.equal(f.instance.native.popup.size, 0);
    assert.ok(f.observers.every(observer => observer.isDisposed));
});
test('precheck unmount leaves native guard until its promise settles and catches obsolete failure', async (t) => {
    const f = createClientsFixture(t, titles.useSPFxApiPermissionPrecheck);
    await f.ready();
    f.click('Check permissions');
    f.unmount();
    assert.equal(f.instance.native.popup.size, 1);
    await f.flush(() => f.tokens[0].reject(new Error('consent required')));
    assert.equal(f.instance.native.popup.size, 0);
    assert.equal(f.instance.native.redirect.size, 0);
});
for (const [owner, kind] of [['useSPFxAadTokenProvider', 'token'], ['useSPFxAadHttpClient', 'aad'], ['useSPFxMSGraphClient', 'graph']])
    test(`${owner} renders actual factory acquisition rejection separately from business requests`, async (t) => {
        const f = createClientsFixture(t, titles[owner]);
        if (owner === 'useSPFxAadHttpClient') {
            f.change('Resource URL or App ID URI', 'https://api.test');
            f.click('Initialize AAD client');
        }
        await f.flush();
        t.mock.method(console, 'error', () => { });
        const request = f.initializations.find(r => r.kind === kind);
        assert.ok(request);
        await f.flush(() => request.reject(new Error('Acquisition failed')));
        assert.match(f.container.textContent, /Acquisition failed/);
        assert.equal(f.calls.length, 0);
        assert.equal(f.tokens.length, 0);
    });
test('Reset preserves AAD input identity/focus and baseline while retained native request is still pending', async (t) => {
    const f = await requestFixture(t, 'useSPFxAadHttpClient');
    const input = f.container.querySelector('input');
    input.focus();
    const action = f.action('Call path');
    let pending;
    f.h.act(() => { pending = action(); });
    f.reset();
    assert.equal(f.container.querySelector('input'), input);
    assert.equal(document.activeElement, input);
    assert.equal(input.value, '');
    assert.equal(f.container.querySelector('[aria-label="Path"]').value, '/api/health');
    assert.match(f.container.textContent, /Loading: underlying requests are still completing/);
    await f.flush(() => f.calls[0].reject(new Error('Old AAD request')));
    await pending;
    assert.doesNotMatch(f.container.textContent, /Old AAD request/);
});
test('native theme and renderer-only updates retain AAD drafts result and focused control without requests', async t => {
    const f = await requestFixture(t, 'useSPFxAadHttpClient');
    f.click('Call path');
    await f.flush(() => f.response(0, { status: 'Retained healthy result' }));
    const input = f.container.querySelector('input');
    input.focus();
    const initializations = f.initializations.length;
    f.h.act(() => {
        const theme = require('@fluentui/react').createTheme({ palette: { themePrimary: '#ab1234' } });
        for (const handler of f.themeHandlers) handler({ theme });
    });
    f.render();
    assert.equal(f.container.querySelector('input'), input);
    assert.equal(document.activeElement, input);
    assert.equal(input.value, 'https://api.test');
    assert.match(f.container.textContent, /Retained healthy result/);
    assert.equal(f.initializations.length, initializations);
    assert.equal(f.calls.length, 1);
});
test('returning an AAD draft to its old value does not resurrect an old command', async t => {
    const f = await requestFixture(t, 'useSPFxAadHttpClient');
    const original = f.action('Call path');
    f.change('Path', '/different');
    f.change('Path', '/api/health');
    await f.flush(() => { original(); });
    assert.equal(f.calls.length, 0);
});
test('reset revision returning to a previous number still invalidates the original command', async t => {
    const f = await requestFixture(t, 'useSPFxHttpClient');
    const original = f.action('Load external todo');
    f.reset();
    f.setRevision(0);
    await f.flush(() => { original(); });
    assert.equal(f.calls.length, 0);
});

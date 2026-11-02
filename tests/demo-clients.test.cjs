const assert = require('node:assert/strict');
const { test } = require('node:test');
const { createClientsFixture } = require('./fixtures/executable-examples/clients-scenarios.cjs');
const titles = ['Inspect client readiness and token-provider initialization', 'Load an external todo with tracked invocation state', 'Read the current SharePoint web title and URL', 'Read the current Graph profile after initialization', 'Initialize a custom API client and call a path', 'Manually check Graph and custom API delegated scopes'];
test('client response metadata stays visible while raw payload diagnostics start collapsed', async (t) => {
    const f = createClientsFixture(t, titles[1]);
    await f.ready();
    assert.equal(f.calls.length, 0);
    assert.equal(f.container.querySelector('[aria-label$=" preview"]'), null);
    assert.ok(f.container.querySelector('[aria-label="HttpClient results"]'));
    f.click('Load external todo');
    await f.flush(() => f.response(0, { title: 'Visible response', completed: false, userId: 8, nested: { diagnosticOnly: 'raw nested payload' } }));
    const results = f.container.querySelector('[aria-label="HttpClient results"]');
    assert.match(results.textContent, /Visible response/);
    assert.match(results.textContent, /false/);
    assert.match(results.textContent, /8/);
    assert.doesNotMatch(results.textContent, /raw nested payload/);
    const trigger = [...f.container.querySelectorAll('button')].find(n => n.textContent === 'Advanced diagnostics');
    assert.equal(trigger.getAttribute('aria-expanded'), 'false');
    f.click('Advanced diagnostics');
    assert.equal(trigger.getAttribute('aria-expanded'), 'true');
    f.click('Raw response');
    assert.match(f.container.querySelector('[aria-label="Raw response JSON"]').textContent, /raw nested payload/);
});
test('each client selection mounts one current result and performs no business invocations', async (t) => {
    for (const title of titles) {
        const f = createClientsFixture(t, title);
        await f.ready();
        assert.equal(f.container.querySelectorAll('[data-demo-section]').length, 1);
        assert.equal(f.container.querySelectorAll('[data-demo-configuration]').length, 1);
        assert.equal(f.container.querySelectorAll('[data-demo-result]').length, 1);
        assert.equal(f.calls.length, 0);
        assert.equal(f.tokens.length, 0);
        f.unmount();
    }
});
test('client completion from an obsolete native service identity cannot publish into the current result', async (t) => {
    const f = createClientsFixture(t, titles[1]);
    await f.ready();
    f.click('Load external todo');
    f.changeScope(f.host('current'));
    await f.ready();
    await f.flush(() => f.response(0, { title: 'Obsolete identity' }));
    assert.doesNotMatch(f.container.textContent, /Obsolete identity/);
    f.click('Load external todo');
    await f.flush(() => f.response(1, { title: 'Current identity' }));
    assert.match(f.container.textContent, /Current identity/);
});
test('AAD configuration draft stays source-owned without native acquisition or business I/O', async (t) => {
    const f = createClientsFixture(t, titles[4]);
    await f.ready();
    const factories = f.initializations.length;
    f.change('Resource URL or App ID URI', 'api://simulated');
    assert.equal(f.container.querySelector('input').value, 'api://simulated');
    assert.equal(f.initializations.length, factories);
    assert.equal(f.calls.length, 0);
    assert.equal(f.tokens.length, 0);
});

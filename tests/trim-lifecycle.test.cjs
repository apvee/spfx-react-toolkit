const test = require('node:test');
const assert = require('node:assert/strict');
const { setup, flush, deferred, Permission } = require('./fixtures/access-trim-harness.cjs');
const { createProviderHarness } = require('./fixtures/access-trim-provider-harness.cjs');
const definitions = [
  { family: 'permission', path: 'sharepoint-permission-trim/sharepoint-permission-trim', symbol: 'SharePointPermissionTrim', props: () => ({ permissions: [new Permission({ Low: 1 })], scope: 'list', listTitle: 'Tasks' }), next: p => ({ ...p, listTitle: 'Other' }), transport: 'sp', grant: () => ({ ok: true, json: async () => ({ High: 0, Low: 1 }) }), deny: () => ({ ok: true, json: async () => ({ High: 0, Low: 0 }) }) },
  { family: 'sp', path: 'sharepoint-group-trim/sharepoint-group-trim', symbol: 'SharePointGroupTrim', props: () => ({ groups: [{ id: 2 }] }), next: p => ({ ...p, webUrl: 'https://tenant.test/sites/other' }), transport: 'sp', grant: () => ({ ok: true, json: async () => ({ Id: 2, Title: 'Readers' }) }), deny: () => ({ ok: false, status: 403 }) },
  { family: 'entra', path: 'entra-group-trim/entra-group-trim', symbol: 'EntraGroupTrim', props: () => ({ groups: [{ id: 'aaaaaaaa-0000-0000-0000-000000000002' }] }), next: p => ({ ...p, groups: [{ id: 'aaaaaaaa-0000-0000-0000-000000000003' }] }), transport: 'graph', grant: () => ({ value: ['aaaaaaaa-0000-0000-0000-000000000002'] }), deny: () => ({ value: [] }) },
];
function load(f, d) { return f.h.load(`controls/${d.path}.tsx`)[d.symbol]; }
test('unkeyed consumers remount across structural branches while same branch state and marker keys persist', async t => {
  const f = setup(); t.after(() => f.h.close()); const C = load(f, definitions[0]); const pending = deferred();
  f.state.spResponse = () => pending.promise;
  let mounts = 0; let unmounts = 0;
  function Child({ branch }) {
    const [value, set] = f.h.React.useState(0);
    f.h.React.useEffect(() => { mounts++; return () => { unmounts++; }; }, []);
    return f.h.React.createElement('button', { onClick: () => set(value + 1) }, `${branch}:${value}`);
  }
  const child = branch => f.h.React.createElement(Child, { branch });
  const children = errorKey => [child('allowed'),
    f.h.React.createElement(C.Loading, { key: 'loading' }, child('loading')),
    f.h.React.createElement(C.Fallback, { key: 'denied' }, child('denied')),
    f.h.React.createElement(C.ErrorFallback, { key: errorKey }, child('error'))];
  const props = { permissions: [new Permission({ Low: 1 })] };
  const m = f.mount(C, props, children('error-one'));
  assert.equal(m.container.textContent, 'loading:0'); f.h.act(() => { m.container.firstChild.click(); });
  await flush(f.h, () => pending.resolve({ ok: true, json: async () => ({ High: 0, Low: 0 }) }));
  assert.equal(m.container.textContent, 'denied:0'); assert.equal(mounts, 2); assert.equal(unmounts, 1);
  f.h.act(() => { m.container.firstChild.click(); }); m.render({ ...props }, children('error-one'));
  assert.equal(m.container.textContent, 'denied:1'); assert.equal(mounts, 2); assert.equal(unmounts, 1);
  const invalid = { ...props, match: 'invalid' }; m.render(invalid, children('error-one'));
  assert.equal(m.container.textContent, 'error:0'); assert.equal(mounts, 3); assert.equal(unmounts, 2);
  f.h.act(() => { m.container.firstChild.click(); }); m.render(invalid, children('error-one'));
  assert.equal(m.container.textContent, 'error:1'); assert.equal(mounts, 3);
  m.render(invalid, children('error-two')); assert.equal(m.container.textContent, 'error:0'); assert.equal(mounts, 4); assert.equal(unmounts, 3);
  f.state.page.web.permissions = new Permission({ Low: 1 }); m.render(props, children('error-two'));
  assert.equal(m.container.textContent, 'allowed:0'); assert.equal(mounts, 5); assert.equal(unmounts, 4);
  assert.equal(m.container.innerHTML, '<button>allowed:0</button>');
});
for (const d of definitions) {
  test(`${d.family} mounts active consumers only, has no wrappers and retains keyed child state`, async t => {
    const f = setup(); t.after(() => f.h.close()); const C = load(f, d); const counts = { grant: 0, loading: 0, denied: 0, error: 0 };
    const mounts = { ...counts }; const unmounts = { ...counts };
    function Child({ branch }) { counts[branch]++; f.h.React.useEffect(() => { mounts[branch]++; return () => { unmounts[branch]++; }; }, []); const [value, set] = f.h.React.useState(0); return f.h.React.createElement('button', { onClick: () => set(value + 1) }, `${branch}:${value}`); }
    const e = (branch, key = branch) => f.h.React.createElement(Child, { branch, key });
    const children = [e('grant'), f.h.React.createElement(C.Loading, { key: 'loading' }, e('loading')), f.h.React.createElement(C.Fallback, { key: 'fallback' }, e('denied')), f.h.React.createElement(C.ErrorFallback, { key: 'error' }, e('error'))];
    const props = d.props(); const m = f.mount(C, props, children);
    assert.equal(counts.grant, 0); assert.equal(counts.denied, 0); assert.equal(counts.error, 0); await flush(f.h);
    assert.deepEqual(mounts, { grant: 1, loading: 1, denied: 0, error: 0 }); assert.equal(unmounts.loading, 1);
    assert.equal(m.container.children.length, 1); assert.equal(m.container.firstChild.tagName, 'BUTTON');
    f.h.act(() => { m.container.firstChild.click(); }); assert.equal(m.container.textContent, 'grant:1');
    m.render({ ...props }, [f.h.React.createElement(C.ErrorFallback, { key: 'error' }, e('error')), e('grant'), f.h.React.createElement(C.Loading, { key: 'loading' }, e('loading'))]);
    await flush(f.h); assert.equal(m.container.textContent, 'grant:1'); assert.equal(counts.error, 0);
    assert.equal(mounts.grant, 1); assert.equal(unmounts.grant, 0);
  });
  for (const kind of ['duplicate', 'wrong-family']) test(`${d.family} ${kind} markers select error and never request`, async t => {
    const f = setup(); t.after(() => f.h.close()); const C = load(f, d); const other = load(f, definitions.find(x => x !== d));
    const children = f.branches(C); children.push(f.h.React.createElement(kind === 'duplicate' ? C.Loading : other.Loading, { key: 'invalid' }, 'invalid'));
    const m = f.mount(C, d.props(), children); await flush(f.h); assert.equal(m.container.textContent, 'error');
    assert.equal(f.spCalls.length + f.graphCalls.length, 0);
  });
  test(`${d.family} detectable fragment-wrapped branches cannot become protected children`, async t => {
    const f = setup(); t.after(() => f.h.close()); const C = load(f, d);
    const children = [f.h.React.createElement(f.h.React.Fragment, { key: 'fragment' }, f.h.React.createElement(C.Loading, undefined, 'nested loading')), f.h.React.createElement(C.ErrorFallback, { key: 'error' }, 'error')];
    const m = f.mount(C, d.props(), children); await flush(f.h); assert.equal(m.container.textContent, 'error');
    assert.equal(f.spCalls.length + f.graphCalls.length, 0);
  });
  test(`${d.family} equal inline arrays keep settled grant without refetch; validation replacement revokes`, async t => {
    const f = setup(); t.after(() => f.h.close()); const C = load(f, d); const props = d.props(); const m = f.mount(C, props, f.branches(C)); await flush(f.h);
    const count = f.spCalls.length + f.graphCalls.length; assert.equal(m.container.textContent, 'grant');
    m.render(d.props()); await flush(f.h); assert.equal(f.spCalls.length + f.graphCalls.length, count);
    m.render({ ...props, match: 'invalid' }); assert.equal(m.observed.at(-1), 'error'); await flush(f.h); assert.equal(f.spCalls.length + f.graphCalls.length, count);
  });
  for (const identity of ['target', 'client', 'user', 'provider']) test(`${d.family} ${identity} replacement masks previous grant before effects`, async t => {
    const f = setup(); t.after(() => f.h.close()); const C = load(f, d); const props = d.props(); const m = f.mount(C, props, f.branches(C)); await flush(f.h);
    assert.equal(m.container.textContent, 'grant'); const pending = deferred(); f.state[`${d.transport}Response`] = () => pending.promise;
    if (identity === 'client') d.transport === 'sp' ? f.replaceSP() : f.replaceGraph();
    if (identity === 'user') f.state.page.user.loginName = 'second';
    if (identity === 'provider') f.state.provider = {};
    m.render(identity === 'target' ? d.next(props) : props);
    assert.equal(m.observed.at(-1), 'loading'); assert.equal(m.container.textContent, 'loading');
    m.unmount(); await flush(f.h, () => pending.reject(new Error('obsolete'))); assert.equal(m.container.textContent, '');
  });
  for (const old of ['success', 'error']) test(`${d.family} obsolete ${old} cannot replace newer result`, async t => {
    const f = setup(); t.after(() => f.h.close()); const C = load(f, d); const pending = deferred(); let first = true;
    const original = f.state[`${d.transport}Response`];
    f.state[`${d.transport}Response`] = (...args) => { if (first) { first = false; return pending.promise; } return original(...args); };
    const props = d.props(); const m = f.mount(C, props, f.branches(C)); await flush(f.h);
    m.render(d.next(props)); await flush(f.h); assert.equal(m.container.textContent, 'grant');
    await flush(f.h, () => old === 'success' ? pending.resolve(d.grant()) : pending.reject(new Error('obsolete')));
    assert.equal(m.container.textContent, 'grant');
  });
  test(`${d.family} older success cannot replace current failure; new valid target can recover`, async t => {
    const f = setup(); t.after(() => f.h.close()); const C = load(f, d); const pending = deferred(); let first = true; const original = f.state[`${d.transport}Response`];
    f.state[`${d.transport}Response`] = () => { if (first) { first = false; return pending.promise; } return Promise.reject(new Error('current failed')); };
    const props = d.props(); const m = f.mount(C, props, f.branches(C)); await flush(f.h);
    m.render(d.next(props)); await flush(f.h); assert.equal(m.container.textContent, 'error');
    await flush(f.h, () => pending.resolve(d.grant())); assert.equal(m.container.textContent, 'error');
    f.state[`${d.transport}Response`] = original; m.render(props); assert.equal(m.observed.at(-1), 'loading'); await flush(f.h); assert.equal(m.container.textContent, 'grant');
  });
}
for (const d of definitions.slice(1)) test(`${d.family} inflight decisions deduplicate per provider only and never persist`, async t => {
  const f = setup(); t.after(() => f.h.close()); const C = load(f, d); const pending = deferred(); const original = f.state[`${d.transport}Response`]; let n = 0;
  f.state[`${d.transport}Response`] = (...args) => n++ === 0 ? pending.promise : original(...args);
  const a = f.mount(C, d.props(), f.branches(C)); const b = f.mount(C, d.props(), f.branches(C)); await flush(f.h);
  assert.equal(n, 1); f.state.provider = {}; const c = f.mount(C, d.props(), f.branches(C)); await flush(f.h); assert.ok(n > 1);
  assert.equal(c.container.textContent, 'grant'); await flush(f.h, () => pending.resolve(d.grant()));
  assert.equal(a.container.textContent, 'grant'); assert.equal(b.container.textContent, 'grant');
  const before = n; f.mount(C, d.props(), f.branches(C)); await flush(f.h); assert.ok(n > before);
});
for (const d of definitions) test(`${d.family} actual provider instances isolate inflight decisions even with identical native client and user`, async t => {
  const f = createProviderHarness(); t.after(() => f.h.close()); const C = load(f, d);
  const page = { web: { absoluteUrl: 'https://tenant.test/sites/current' }, user: { loginName: 'same', email: 'same@test' } };
  const pending = deferred(); let calls = 0;
  const sp = { get: async url => {
    if (url.includes('effectivebasepermissions') || (url.includes('sitegroups') && !url.includes('/users'))) { calls++; return pending.promise; }
    return { ok: true, json: async () => url.includes('currentuser') ? { Id: 17 } : { value: [{ Id: 17 }] } };
  } };
  const graph = { api: () => ({ post: () => { calls++; return pending.promise; } }) };
  const hosts = [f.host('one', page, sp, graph), f.host('two', page, sp, graph)];
  const m = f.mount(hosts, C, d.props(), [f.h.React.createElement('b', { key: 'granted' }, 'grant'), f.h.React.createElement(C.Loading, { key: 'loading' }, 'loading')]);
  await flush(f.h); assert.equal(calls, 2); assert.equal(m.container.textContent, 'loadingloading');
  await flush(f.h, () => pending.resolve(d.grant())); assert.equal(m.container.innerHTML, '<b>grant</b><b>grant</b>');
  m.unmount();
});

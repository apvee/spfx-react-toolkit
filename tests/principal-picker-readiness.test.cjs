const test = require('node:test');
const assert = require('node:assert/strict');
const { createPropertyPaneHarness } = require('./fixtures/property-pane-picker-harness.cjs');
const tenantId = '11111111-1111-4111-8111-111111111111';
const id = '22222222-2222-4222-8222-222222222222';
const groupId = '33333333-3333-4333-8333-333333333333';
const web = 'https://tenant.test/sites/hr';
const loginName = 'alice@tenant.test';
const flush = async h => h.act(async () => {
  for (let i = 0; i < 15; i++) await Promise.resolve();
  await new Promise(resolve => setTimeout(resolve, 5));
  for (let i = 0; i < 15; i++) await Promise.resolve();
});
function fixture({ group = false, seed = false, mode, membership = false, empty = false } = {}) {
  const h = createPropertyPaneHarness(), owner = h.host(), dom = document.createElement('div');
  const frames = [], errors = [], changes = [], graphCalls = [], spCalls = [];
  let complete, fail, initializations = 0;
  const pending = new Promise((resolve, reject) => { complete = resolve; fail = reject; });
  const factory = { getClient: () => { initializations++; return pending; } };
  const scope = owner.instance.context.serviceScope, consume = scope.consume.bind(scope);
  scope.consume = key => { const service = consume(key); return service?.getClient === undefined ? service : factory; };
  const userRow = { id, displayName: 'Live Alice', userPrincipalName: loginName, mail: loginName };
  const groupRow = { id: groupId, displayName: 'Live group', groupTypes: ['Unified'], securityEnabled: false, mailEnabled: true };
  const graph = { api(uri) {
    const api = { version() { return api; }, header() { return api; },
      get: async () => { graphCalls.push({ uri }); return uri.startsWith('/groups?') ? { value: [groupRow] } : group ? groupRow : userRow; },
      post: async body => { graphCalls.push({ uri, body }); return { value: [groupId] }; } };
    return api;
  } };
  owner.instance.context.spHttpClient.get = async uri => {
    spCalls.push({ uri });
    return { ok: true, status: 200, json: async () => uri.includes('/_api/site?') ? { Url: web } : { Id: 3, Title: 'Alice', LoginName: loginName, Email: loginName, PrincipalType: 1 } };
  };
  owner.instance.context.spHttpClient.post = async (uri, config, options) => {
    spCalls.push({ uri, body: JSON.parse(options.body) });
    return { ok: true, status: 200, json: async () => ({ ClientPeoplePickerSearchUser: JSON.stringify([{ Key: loginName, DisplayText: 'Alice', IsResolved: true, EntityType: 'User', EntityData: { Email: loginName } }]) }) };
  };
  const value = group ? { kind: 'microsoft365Group', identity: { tenantId, id: groupId }, text: 'Saved group' } :
    { kind: 'user', loginName, ...(mode === undefined && !membership ? { entra: { tenantId, id } } : {}), text: 'Saved Alice' };
  const model = group ? h.load('controls/pickers/principal-models.internal.ts').graphGroupOption({ ...groupRow, displayName: 'Saved group' }, { tenantId }) :
    { key: 'seed', kind: 'user', text: 'Saved Alice', value, loginName };
  const renderer = h.load('controls/pickers/picker-renderer.internal.tsx'), Render = renderer.PickerRenderer;
  let runtime;
  renderer.PickerRenderer = props => {
    runtime = props.runtime;
    frames.push({ loading: runtime.loading, selected: runtime.selectedItems.map(o => o.text), validation: runtime.validation, error: runtime.error?.message, phase: runtime.searchFeedback.phase });
    return h.React.createElement(Render, props);
  };
  const factoryFn = h.load('property-pane/pickers/property-pane-principal-picker.ts').PropertyPanePrincipalPicker;
  const props = { instance: owner.instance, key: 'readiness', value: empty ? [] : [value], label: 'Principals', debounceMs: 0, presentation: 'tagPicker',
    principalTypes: group ? ['microsoft365Group'] : ['user'], initialOptions: seed ? [model] : undefined,
    userIdentity: mode === undefined || mode === 'available' ? undefined : { mode }, userScope: membership ? { kind: 'groups', groups: [{ source: 'entra', id: groupId }] } : undefined,
    onError: (error, details) => errors.push({ message: error.message, operation: details.operation, source: details.source }) };
  const field = factoryFn('values', props);
  h.act(() => field.properties.onRender(dom, undefined, (...args) => changes.push(args)));
  return { h, dom, owner, graph, graphCalls, spCalls, errors, changes, frames, props, factoryFn, get runtime() { return runtime; }, get initializations() { return initializations; },
    release: () => h.act(async () => complete(graph)), reject: () => h.act(async () => fail(new Error('Graph initialization denied'))),
    dispose: () => h.act(() => field.properties.onDispose(dom)),
    async search(query) { const input = dom.querySelector('input'); h.act(() => { input.click(); input.value = query; require('react-dom/test-utils').Simulate.change(input, { target: { value: query }, currentTarget: { value: query } }); }); await flush(h); } };
}
for (const [name, options] of [
  ['saved Graph group hydration', { group: true }],
  ['seeded Graph group eligibility', { group: true, seed: true }],
  ['seeded Entra user eligibility', { seed: true }],
  ['seeded user Graph membership', { seed: true, membership: true }],
]) {
  test(`${name} waits for client initialization without transient unresolved feedback`, async () => {
    const f = fixture(options);
    try {
      await flush(f.h);
      assert.equal(f.runtime.loading, true);
      assert.equal(f.runtime.validation, undefined);
      assert.deepEqual(f.errors, []);
      assert.equal(f.graphCalls.length, 0);
      assert.match(f.dom.textContent, /Saved/);
      await f.release(); await flush(f.h);
      assert.equal(f.runtime.loading, false);
      assert.equal(f.runtime.selectedItems.length, 1);
      assert.equal(f.runtime.validation, undefined);
      assert.deepEqual(f.errors, []);
      assert.ok(f.frames.every(frame => frame.validation === undefined && frame.error === undefined));
      assert.equal(f.initializations, 1);
      assert.equal(f.graphCalls.filter(call => call.body === undefined).length, 1);
      if (options.membership) assert.deepEqual(f.graphCalls.find(call => call.body !== undefined).body, { groupIds: [groupId] });
      assert.equal(f.spCalls.some(call => call.uri.includes('ensureuser')), false);
      assert.deepEqual(f.changes, []);
    } finally { f.dispose(); }
  });
}
test('typed Graph group search remains pending until initialization completes', async () => {
  const f = fixture({ group: true, empty: true });
  try {
    await f.search('group');
    assert.equal(f.runtime.searchFeedback.phase, 'pending');
    assert.deepEqual(f.errors, []);
    await f.release(); await flush(f.h);
    assert.equal(f.runtime.searchFeedback.phase, 'complete');
    assert.deepEqual(f.runtime.searchOptions.map(o => o.text), ['Live group']);
    assert.equal(f.graphCalls.length, 1);
    assert.ok(f.frames.every(frame => frame.phase !== 'error'));
    assert.deepEqual(f.changes, []);
  } finally { f.dispose(); }
});
test('explicit both-ID resolution started before Graph readiness commits once after readiness', async () => {
  const f = fixture({ seed: true, mode: 'both' });
  try {
    await flush(f.h);
    assert.equal(f.runtime.selectedItems.length, 1);
    const resolve = [...f.dom.querySelectorAll('button')].find(button => button.textContent === 'Resolve IDs');
    assert.ok(resolve);
    f.h.act(() => require('react-dom/test-utils').Simulate.click(resolve)); await flush(f.h);
    assert.equal(f.runtime.confirming, true);
    assert.equal(f.runtime.error, undefined);
    assert.deepEqual(f.changes, []);
    await f.release(); await flush(f.h);
    assert.equal(f.runtime.confirming, false);
    assert.equal(f.runtime.error, undefined);
    assert.equal(f.runtime.canResolve, false);
    assert.equal(f.changes.length, 1);
    assert.equal(f.changes[0][0], 'values');
    assert.deepEqual(f.changes[0][1][0].entra, { tenantId, id });
    assert.deepEqual(f.changes[0][1][0].sharePoint, { siteCollectionUrl: web, id: 3 });
    assert.deepEqual(f.errors, []);
    assert.equal(f.spCalls.some(call => call.uri.includes('ensureuser')), false);
  } finally { f.dispose(); }
});
test('terminal Graph initialization failure settles Graph-dependent selection and reports its cause', async () => {
  const f = fixture({ group: true, seed: true });
  try {
    await flush(f.h); await f.reject(); await flush(f.h);
    assert.equal(f.runtime.loading, false);
    assert.match(f.runtime.validation, /unresolved/);
    assert.deepEqual(f.errors.map(error => error.message), ['Graph initialization denied']);
    assert.equal(f.errors[0].source, 'graph');
    assert.deepEqual(f.graphCalls, []);
    assert.deepEqual(f.changes, []);
  } finally { f.dispose(); }
});
test('SP-only selection and search remain usable when optional Graph initialization fails', async () => {
  const f = fixture({ seed: true, mode: 'available' });
  try {
    await flush(f.h);
    assert.equal(f.runtime.selectedItems.length, 1);
    assert.equal(f.runtime.loading, false);
    await f.reject(); await f.search('alice');
    assert.equal(f.runtime.selectedItems.length, 1);
    assert.equal(f.runtime.searchFeedback.phase, 'complete');
    assert.equal(f.runtime.error, undefined);
    assert.deepEqual(f.errors, []);
    assert.deepEqual(f.graphCalls, []);
    assert.equal(f.spCalls.some(call => call.body?.queryParams?.QueryString === 'alice'), true);
    assert.deepEqual(f.changes, []);
  } finally { f.dispose(); }
});
test('unmounting while Graph initialization is pending prevents later Graph work and callbacks', async () => {
  const f = fixture({ group: true, seed: true });
  await flush(f.h);
  f.dispose();
  await f.release(); await flush(f.h);
  assert.deepEqual(f.graphCalls, []);
  assert.deepEqual(f.errors, []);
  assert.deepEqual(f.changes, []);
});


test('replacing the native owner cancels the former pending Graph dependency', async () => {
  const f = fixture({ group: true, seed: true });
  let next;
  try {
    await flush(f.h);
    const oldInput = f.dom.querySelector('input');
    const replacement = f.h.host('replacement');
    next = f.factoryFn('values', { ...f.props, instance: replacement.instance, value: [] });
    f.h.act(() => next.properties.onRender(f.dom, undefined, (...args) => f.changes.push(args)));
    await flush(f.h);
    assert.notEqual(f.dom.querySelector('input'), oldInput);
    await f.release(); await flush(f.h);
    assert.deepEqual(f.graphCalls, []);
    assert.deepEqual(f.errors, []);
    assert.deepEqual(f.changes, []);
    assert.deepEqual(f.runtime.selectedOptions, []);
  } finally { if (next) f.h.act(() => next.properties.onDispose(f.dom)); else f.dispose(); }
});

test('private Graph-only source readiness accepts an asynchronous client without SharePoint', async () => {
  const h = createPropertyPaneHarness();
  const graph = { api() { const api = { version() { return api; }, get: async () => ({ id: groupId, displayName: 'Group', groupTypes: ['Unified'], securityEnabled: false, mailEnabled: true }) }; return api; } };
  const context = { owner: {}, tenantId, userId: 'current', webUrl: web, siteCollectionUrl: web, graphClientIdentity: {}, getGraphClient: async () => graph };
  const source = h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource({ principalTypes: ['microsoft365Group'] }, context);
  assert.equal(source.ready, true);
  const cache = h.load('controls/pickers/picker-request-cache.internal.ts').createPickerRequestCache(context.owner);
  const value = { kind: 'microsoft365Group', identity: { tenantId, id: groupId } };
  const result = await source.hydrate([{ kind: 'value', key: 'group', value }], { context, cache, query: '', pageSize: 50, generation: 0 });
  assert.equal(result.resolved[0].option.text, 'Group');
  cache.dispose();
});

test('disposing a pending explicit resolution settles the action without Graph work or writes', { timeout: 5000 }, async () => {
  const f = fixture({ seed: true, mode: 'both' });
  await flush(f.h);
  let pending;
  f.h.act(() => { pending = f.runtime.resolve({ type: 'click' }); });
  await flush(f.h);
  assert.equal(f.runtime.confirming, true);
  f.dispose();
  await pending;
  await f.release(); await flush(f.h);
  assert.deepEqual(f.graphCalls, []);
  assert.deepEqual(f.errors, []);
  assert.deepEqual(f.changes, []);
});


test('unseeded Graph hydration publishes the terminal initialization cause and clears it after removal', async () => {
  const f = fixture({ group: true });
  try {
    await flush(f.h); await f.reject(); await flush(f.h);
    assert.equal(f.runtime.loading, false);
    assert.equal(f.runtime.error?.message, 'Graph initialization denied');
    assert.deepEqual(f.errors, [{ message: 'Graph initialization denied', operation: 'resolve', source: 'graph' }]);
    assert.match(f.dom.textContent, /Saved group/);
    await f.h.act(async () => f.runtime.clear({ type: 'click' }));
    assert.equal(f.runtime.error, undefined, 'the removed selection cannot retain a hydration error');
    assert.deepEqual(f.runtime.selectedOptions, []);
  } finally { f.dispose(); }
});

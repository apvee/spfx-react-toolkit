const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const option = name => ({ key: name, text: name, url: `https://tenant.test/sites/${name}`, value: { kind: 'site', url: `https://tenant.test/sites/${name}` } });
const context = { owner: {}, tenantId: '11111111-1111-4111-8111-111111111111', userId: 'u', webUrl: 'https://tenant.test/sites/a', siteCollectionUrl: 'https://tenant.test/sites/a' };
const flush = async h => h.act(async () => { await new Promise(resolve => setTimeout(resolve, 4)); });
function source(overrides = {}) {
  return { identity: 'sites', family: 'site', context, ready: true, transport: 'sharePoint', accepts: () => true, discover: async () => ({ options: [option('remote')], completeness: 'complete' }), hydrate: async inputs => ({ resolved: [], unresolved: inputs.map(i => ({ key: i.key })) }), resolve: async () => ({ resolved: [], unresolved: [] }), verifyCandidate: async option => option, ...overrides };
}
test('reactive consumer options preserve selections and do not repeat typed queries on cosmetic changes', async () => {
  const h = createHarness();
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let requests = 0;
  const src = source({ discover: async () => { requests++; return { options: [option('remote')], completeness: 'complete' }; } });
  const props = { defaultSelectedOptions: ['a'], initialOptions: [option('a')], availableOptions: [option('b')], debounceMs: 0 };
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), props);
  h.act(() => { root.current.setOpen(true); root.current.setQuery('re'); });
  await flush(h);
  assert.equal(requests, 1);
  root.render({ ...props, availableOptions: [], label: 'Renamed', presentation: 'tagPicker', onOptionSelect: () => {} });
  await flush(h);
  assert.equal(requests, 1);
  assert.deepEqual(root.current.selectedOptions, ['a']);
  assert.equal(root.current.selectedItems[0].text, 'a');
  h.act(() => root.current.setQuery(''));
  assert.deepEqual(root.current.sections, []);
  root.unmount();
});
test('older query responses are hidden and unknown opaque selections remain unresolved', async () => {
  const h = createHarness();
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const completions = new Map();
  const src = source({ discover: request => new Promise(resolve => completions.set(request.query, resolve)) });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), { defaultSelectedOptions: ['opaque'], debounceMs: 0 });
  h.act(() => { root.current.setOpen(true); root.current.setQuery('aa'); });
  await flush(h);
  h.act(() => root.current.setQuery('bb'));
  await flush(h);
  await h.act(async () => completions.get('bb')({ options: [option('b')], completeness: 'complete' }));
  await h.act(async () => completions.get('aa')({ options: [option('a')], completeness: 'complete' }));
  assert.deepEqual(root.current.searchOptions.map(o => o.key), ['b']);
  assert.deepEqual(root.current.unresolvedOptions, ['opaque']);
  assert.deepEqual(root.current.selectedOptions, ['opaque']);
  root.unmount();
});
test('failed confirmation preserves prior selection and emits no incomplete value', async () => {
  const h = createHarness();
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let emitted = 0;
  const src = source({ verifyCandidate: async () => { throw new Error('Rejected identity'); } });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), { defaultSelectedOptions: ['a'], initialOptions: [option('a')], availableOptions: [option('b')], onOptionSelect: () => emitted++ });
  await h.act(async () => root.current.select({ type: 'click' }, 'b'));
  assert.equal(emitted, 0);
  assert.deepEqual(root.current.selectedOptions, ['a']);
  assert.match(root.current.error.message, /Rejected/);
  root.unmount();
});
test('unmount during candidate verification cannot start confirmation writes', async () => {
  const h = createHarness();
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let finish, confirms = 0;
  const src = source({ verifyCandidate: () => new Promise(resolve => { finish = resolve; }), confirm: async values => { confirms++; return values; } });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), { availableOptions: [option('a')] });
  let pending;
  h.act(() => { pending = root.current.select({ type: 'click' }, 'a'); });
  root.unmount();
  await h.act(async () => { finish(option('a')); await pending; });
  assert.equal(confirms, 0);
});
test('blur does not cancel pending confirmation and rejected blur validation is reported', async () => {
  const h = createHarness();
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let finish, errors = 0;
  const src = source({ verifyCandidate: () => new Promise(resolve => { finish = resolve; }) });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), { availableOptions: [option('a')], validationTrigger: 'blur', validateSelection: async () => { throw new Error('validator failed'); }, onError: () => errors++ });
  let pending;
  h.act(() => { pending = root.current.select({ type: 'click' }, 'a'); });
  await h.act(async () => root.current.blur());
  await h.act(async () => { finish(option('a')); await pending; });
  assert.equal(root.current.confirming, false);
  assert.deepEqual(root.current.selectedOptions, ['a']);
  await h.act(async () => root.current.blur());
  assert.equal(errors, 1);
  assert.match(root.current.error.message, /validator failed/);
  root.unmount();
});
test('removing a default principal never rebinds the remaining key to another initial reference', async () => {
  const h = createHarness();
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const inputs = [];
  const src = source({ hydrate: async values => { inputs.push(values); return { resolved: [], unresolved: values.map(i => ({ key: i.key })) }; } });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), { defaultPrincipals: [{ kind: 'login', loginName: 'alice' }, { kind: 'login', loginName: 'bob' }] });
  await flush(h);
  await h.act(async () => { await root.current.remove({ type: 'click' }, 'principal-initial:0'); });
  await flush(h);
  assert.deepEqual(root.current.selectedOptions, ['principal-initial:1']);
  assert.ok(inputs.flat().filter(input => input.key === 'principal-initial:1').every(input => input.reference.loginName === 'bob'));
  root.unmount();
});
test('every visible option remains selectable beyond historical model-cache capacity', async () => {
  const h = createHarness();
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const src = source({ discover: async () => ({ options: Array.from({ length: 200 }, (_, i) => option(String(i))), completeness: 'complete' }) });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), { suggestionsLimit: 200, debounceMs: 0 });
  h.act(() => { root.current.setOpen(true); root.current.setQuery('xx'); });
  await flush(h);
  await h.act(async () => root.current.select({ type: 'click' }, '0'));
  assert.deepEqual(root.current.selectedOptions, ['0']);
  root.unmount();
});
test('retry generation ignores an older pending continuation', async () => {
  const h = createHarness();
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let finish, pages = 0;
  const src = source({ discover: async request => request.continuation ? new Promise(resolve => { finish = resolve; }) : { options: [option(pages++ === 0 ? 'old' : 'new')], continuation: 'cursor', completeness: 'partial' } });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), { debounceMs: 0 });
  h.act(() => { root.current.setOpen(true); root.current.setQuery('xx'); });
  await flush(h);
  let pending;
  h.act(() => { pending = root.current.loadMore(); });
  await flush(h);
  h.act(() => root.current.retry());
  await flush(h);
  await h.act(async () => { finish({ options: [option('obsolete')], completeness: 'complete' }); await pending; });
  assert.deepEqual(root.current.searchOptions.map(o => o.key), ['new']);
  root.unmount();
});
test('typed search adopts consumer identity models and reports retained source conflicts', async () => {
  const h = createHarness();
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const consumer = { ...option('a'), text: 'Consumer' };
  const remote = { ...option('a'), key: 'remote-alias', text: 'Remote' };
  const src = source({ discover: async () => ({ options: [remote], completeness: 'complete' }) });
  const props = { availableOptions: [consumer], debounceMs: 0 };
  const root = h.mount(p => usePickerRuntime(p.props, p.source, { multiselect: false }), { props, source: src });
  h.act(() => { root.current.setOpen(true); root.current.setQuery('xx'); });
  await flush(h);
  assert.equal(root.current.searchOptions[0].text, 'Consumer');
  await h.act(async () => root.current.select({ type: 'click' }, 'a'));
  root.render({ props, source: source({ identity: 'other-scope', accepts: () => false }) });
  assert.deepEqual(root.current.selectedOptions, ['a']);
  assert.match(root.current.validation, /ambito/);
  root.unmount();
});
test('native authoritative revision cancels pending confirmation without replacing source or cache', async () => {
  const h = createHarness();
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const { PickerAuthoritativeEpochContext } = h.load('controls/pickers/picker-runtime-context.internal.tsx');
  let finish, confirms = 0, current;
  const src = source({ verifyCandidate: () => new Promise(resolve => { finish = resolve; }), confirm: async values => { confirms++; return values; } });
  function Probe() { current = usePickerRuntime({ selectedOptions: [], availableOptions: [option('a')] }, src, { multiselect: false }); return null; }
  const container = document.createElement('div'); document.body.appendChild(container);
  const ReactDOM = require('react-dom');
  const render = epoch => h.act(() => { ReactDOM.render(h.React.createElement(PickerAuthoritativeEpochContext.Provider, { value: epoch }, h.React.createElement(Probe)), container); });
  render('0');
  let pending;
  h.act(() => { pending = current.select({ type: 'click' }, 'a'); });
  render('1');
  await h.act(async () => { finish(option('a')); await pending; });
  assert.equal(confirms, 0);
  h.act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove();
});
test('consumer presearch candidates are scope verified, cached by identity and updated without resetting selections', async () => {
  const h = createHarness();
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let calls = 0;
  const src = source({ preflight: true, verifyCandidate: async o => { calls++; return o.key === 'outside' ? undefined : o; } });
  const props = { defaultSelectedOptions: ['selected'], initialOptions: [option('selected')], availableOptions: [option('inside'), option('outside')] };
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), props);
  h.act(() => { root.current.setOpen(true); }); await flush(h);
  assert.deepEqual(root.current.sections.flatMap(s => s.options.map(o => o.key)), ['inside']);
  root.render({ ...props, availableOptions: [{ ...option('inside'), text: 'Renamed' }, option('new')] }); await flush(h);
  assert.equal(calls, 4);
  assert.deepEqual(root.current.sections.flatMap(s => s.options.map(o => o.text)), ['Renamed', 'new']);
  assert.deepEqual(root.current.selectedOptions, ['selected']);
  root.unmount();
});
test('change validation gates removal and clearing just as it gates selection', async () => {
  const h = createHarness();
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let emitted = 0;
  const root = h.mount(p => usePickerRuntime(p, source(), { multiselect: true }), { defaultSelectedOptions: ['a'], initialOptions: [option('a')], validateSelection: async items => items.length === 0 ? 'Required' : undefined, onOptionSelect: () => emitted++ });
  await h.act(async () => { await root.current.remove({ type: 'click' }, 'a'); });
  await h.act(async () => { await root.current.clear({ type: 'click' }); });
  assert.equal(emitted, 0);
  assert.deepEqual(root.current.selectedOptions, ['a']);
  assert.equal(root.current.validation, 'Required');
  root.unmount();
});
test('a rejected alias rebind cannot show one identity and confirm another', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const first = { ...option('a'), key: 'same' }, next = { ...option('b'), key: 'same' };
  const changes = [];
  const root = h.mount(p => usePickerRuntime(p, source(), { multiselect: false }), { availableOptions: [first], onOptionSelect: (_, data) => changes.push(data) });
  root.render({ availableOptions: [next], onOptionSelect: (_, data) => changes.push(data) });
  assert.deepEqual(root.current.sections.flatMap(section => section.options), []);
  assert.match(root.current.error.message, /rebound/);
  root.unmount();
});
test('consumer scope verification does not silently truncate before an eligible twenty-first option', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const options = Array.from({ length: 21 }, (_, i) => option(`candidate${i}`));
  const src = source({ preflight: true, verifyCandidate: async o => o.key === 'candidate20' ? o : undefined });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), { availableOptions: options });
  h.act(() => { root.current.setOpen(true); }); await flush(h);
  assert.deepEqual(root.current.sections.flatMap(s => s.options.map(o => o.key)), ['candidate20']);
  root.unmount();
});
test('continuation failures notify the latest error callback', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const errors = [];
  const src = source({ discover: async req => { if (req.continuation) throw new Error('page failure'); return { options: [option('a')], continuation: 'next', completeness: 'partial' }; } });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), { debounceMs: 0, onError: (error, details) => errors.push([error.message, details.operation]) });
  h.act(() => { root.current.setOpen(true); root.current.setQuery('aa'); }); await flush(h);
  await h.act(async () => { await root.current.loadMore(); });
  assert.deepEqual(errors, [['page failure', 'search']]);
  assert.equal(root.current.searchFeedback.phase, 'error');
  assert.equal(root.current.searchFeedback.error.message, 'page failure');
  assert.deepEqual(root.current.searchOptions.map(item => item.key), ['a'], 'failed pagination retains usable rows');
  root.unmount();
});
test('explicit refresh invalidates nested personal requests after source configuration changes', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let calls = 0, identity = 'one';
  const root = h.mount(props => usePickerRuntime(props, source({ identity, personalSuggestions: async request => {
    const options = await request.cache.request('personal-transport', async () => { calls++; return [option('a')]; }, request.partition);
    return { people: options, completeness: 'complete' };
  } }), { multiselect: false }), { refreshKey: 0 });
  h.act(() => { root.current.setOpen(true); }); await flush(h); assert.equal(calls, 1);
  h.act(() => { root.current.setOpen(false); });
  identity = 'two'; root.render({ refreshKey: 0 }); await flush(h); assert.equal(calls, 1);
  root.render({ refreshKey: 1 });
  h.act(() => { root.current.setOpen(true); }); await flush(h); assert.equal(calls, 2);
  root.unmount();
});
test('seeded selections undergo read-only scope verification even with complete display models', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let calls = 0;
  const src = source({ preflight: true, verifyCandidate: async () => { calls++; return undefined; } });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), { defaultSelectedOptions: ['subweb'], initialOptions: [option('subweb')] });
  await flush(h);
  assert.equal(calls, 1);
  assert.deepEqual(root.current.selectedItems, []);
  assert.deepEqual(root.current.unresolvedOptions, ['subweb']);
  assert.deepEqual(root.current.selectedOptions, ['subweb']);
  assert.match(root.current.validation, /risolti|ambito/);
  root.unmount();
});

test('adding a candidate preserves unresolved scope-rejected keys without promoting their models in events or validation', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const changes = [], validated = [];
  const src = source({ preflight: true, accepts: o => o.key === 'b', verifyCandidate: async o => o.key === 'b' ? o : undefined });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), { defaultSelectedOptions: ['a'], initialOptions: [option('a')], availableOptions: [option('b')], validateSelection: items => { validated.push(items.map(o => o.key)); }, onOptionSelect: (_, data) => changes.push(data) });
  await flush(h);
  await h.act(async () => { await root.current.select({ type: 'click' }, 'b'); });
  assert.deepEqual(changes[0].selectedOptions, ['a', 'b']);
  assert.deepEqual(changes[0].selectedItems.map(o => o.key), ['b']);
  assert.deepEqual(changes[0].selectedValues, [option('b').value]);
  assert.deepEqual(changes[0].unresolvedOptions, ['a']);
  assert.deepEqual(validated, [['b']]);
  root.unmount();
});
test('bounded partial diagnostics stay visible alongside healthy rows without offering retry', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const src = source({ discover: async () => ({ options: [option('a')], completeness: 'partial', errors: [{ source: 'sharePoint', error: new Error('Maximum 1000 results: refine'), retryable: false }] }) });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), { debounceMs: 0 });
  h.act(() => { root.current.setOpen(true); root.current.setQuery('aa'); }); await flush(h);
  assert.equal(root.current.searchOptions.length, 1);
  assert.equal(root.current.warning, 'Maximum 1000 results: refine');
  assert.equal(root.current.error, undefined);
  root.unmount();
});
test('personal failure is displayed with healthy consumer suggestions and cleared after retry', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let fail = true;
  const src = source({ personalSuggestions: async () => ({ people: [], errors: fail ? [{ source: 'people', error: new Error('People denied') }] : [], completeness: fail ? 'partial' : 'complete' }) });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), { availableOptions: [option('a')] });
  h.act(() => { root.current.setOpen(true); }); await flush(h);
  assert.equal(root.current.sections[0].options[0].key, 'a');
  assert.equal(root.current.error?.message, 'People denied');
  fail = false; h.act(() => { root.current.retry(); }); await flush(h);
  assert.equal(root.current.error, undefined);
  root.unmount();
});

test('explicit ID action is offered only when selected users lack IDs required by their mode', () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const user = { kind: 'user', key: 'alice', text: 'Alice', loginName: 'alice@tenant.test', value: { kind: 'user', loginName: 'alice@tenant.test' } };
  const src = source({ confirm: async options => options });
  const props = { defaultSelectedOptions: ['alice'], initialOptions: [user], userIdentity: { mode: 'available' } };
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), props);
  assert.equal(root.current.canResolve, false);
  root.render({ ...props, userIdentity: { mode: 'entra' } });
  assert.equal(root.current.canResolve, true);
  root.render({ ...props, userIdentity: { mode: 'entra' }, initialOptions: [{ ...user, value: { ...user.value, entra: { tenantId: context.tenantId, id: 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa' } } }] });
  assert.equal(root.current.canResolve, false);
  root.unmount();
});

test('a user with an old collection registration can explicitly resolve IDs required in the new destination', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const user = { kind: 'user', key: 'alice', text: 'Alice', loginName: 'alice@tenant.test', value: { kind: 'user', loginName: 'alice@tenant.test', sharePoint: { id: 3, siteCollectionUrl: 'https://tenant.test/sites/old' } } };
  const src = source({ preflight: true, confirm: async options => options, needsIdentityResolution: async () => true });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), { defaultSelectedOptions: ['alice'], initialOptions: [user], userIdentity: { mode: 'sharePoint' } });
  await flush(h);
  assert.equal(root.current.canResolve, true);
  root.unmount();
});

test('metadata initial catalog loads on open, caps after consumer dedup and keeps remote search independent', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const rows = Array.from({ length: 35 }, (_, i) => option(`entry${i}`)), requests = [];
  const src = source({ discover: async req => { requests.push(req.query); return { options: req.query === '' ? rows : [option('outside')], completeness: 'complete' }; } });
  const props = { availableOptions: [{ ...rows[0], text: 'Consumer first' }], debounceMs: 0 };
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false, initialCatalogLimit: 30 }), props);
  try {
    assert.deepEqual(requests, []);
    h.act(() => root.current.setOpen(true)); await flush(h);
    const options = root.current.sections.flatMap(s => s.options);
    assert.equal(options.length, 30); assert.equal(options[0].text, 'Consumer first');
    assert.equal(root.current.catalogHint, 'Sono disponibili altri elementi.');
    root.render({ ...props, label: 'Cosmetic', availableOptions: [{ ...rows[0], text: 'Renamed' }] }); await flush(h);
    assert.equal(requests.filter(q => q === '').length, 1);
    h.act(() => root.current.setQuery('outside')); await flush(h);
    assert.deepEqual(root.current.searchOptions.map(o => o.key), ['outside']);
  } finally { root.unmount(); }
});
test('initial catalog distinguishes exact cap from lookahead and obeys disabled/default suppression', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  for (const size of [0, 29, 30, 31]) {
    const src = source({ discover: async () => ({ options: Array.from({ length: size }, (_, i) => option(`entry${i}`)), completeness: 'complete' }) });
    const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true, initialCatalogLimit: 30 }), {});
    try { h.act(() => root.current.setOpen(true)); await flush(h); assert.equal(root.current.sections.flatMap(s => s.options).length, Math.min(size, 30)); assert.equal(Boolean(root.current.catalogHint), size > 30); }
    finally { root.unmount(); }
  }
  let calls = 0;
  const src = source({ discover: async () => { calls++; return { options: [], completeness: 'complete' }; } });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false, initialCatalogLimit: 10 }), { disabled: true });
  try { h.act(() => root.current.setOpen(true)); await flush(h); root.render({ showDefaultSuggestions: false }); await flush(h); assert.equal(calls, 0); }
  finally { root.unmount(); }
});
test('pending selected scope and hydration do not flash unresolved errors but terminal missing values do', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let finish;
  const src = source({ preflight: true, verifyCandidate: () => new Promise(resolve => { finish = resolve; }) });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), { defaultSelectedOptions: ['a'], initialOptions: [option('a')] });
  try { await flush(h); assert.equal(root.current.loading, true); assert.equal(root.current.validation, undefined); await h.act(async () => finish(option('a'))); assert.equal(root.current.loading, false); assert.equal(root.current.validation, undefined); }
  finally { root.unmount(); }
  let hydrate;
  const missing = source({ hydrate: () => new Promise(resolve => { hydrate = resolve; }) });
  const pending = h.mount(p => usePickerRuntime(p, missing, { multiselect: true }), { defaultSelectedOptions: ['missing'] });
  try { await flush(h); assert.equal(pending.current.validation, undefined); assert.equal(pending.current.loading, true); await h.act(async () => hydrate({ resolved: [], unresolved: [{ key: 'missing' }] })); assert.match(pending.current.validation, /non sono risolti/); assert.equal(pending.current.loading, false); }
  finally { pending.unmount(); }
});
test('a settled rejected selection remains invalid while another identity is pending', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let finish;
  const src = source({ preflight: true, verifyCandidate: async o => o.key === 'a' ? undefined : await new Promise(resolve => { finish = resolve; }) });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), { selectedOptions: ['a'], initialOptions: [option('a'), option('b')] });
  try { await flush(h); assert.match(root.current.validation, /non sono risolti/); root.render({ selectedOptions: ['a', 'b'], initialOptions: [option('a'), option('b')] }); await flush(h); assert.equal(root.current.loading, true); assert.match(root.current.validation, /non sono risolti/); await h.act(async () => finish(option('b'))); }
  finally { root.unmount(); }
});

test('successful asynchronous hydration never exposes an intermediate unresolved validation', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let finish; const observed = [];
  const src = source({ hydrate: () => new Promise(resolve => { finish = resolve; }) });
  const root = h.mount(p => { const result = usePickerRuntime(p, src, { multiselect: true }); observed.push(result.validation); return result; }, { defaultSelectedOptions: ['a'] });
  try { await flush(h); await h.act(async () => finish({ resolved: [{ key: 'a', option: option('a') }], unresolved: [] })); assert.equal(root.current.selectedItems.length, 1); assert.deepEqual(observed.filter(Boolean), []); }
  finally { root.unmount(); }
});
test('case-only edits preserve accumulated search pages and equivalent requests', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const calls = [];
  const src = source({ caseInsensitive: true, discover: async req => { calls.push(req.query); return req.continuation ? { options: [option('second')], completeness: 'complete' } : { options: [option('first')], continuation: 'next', completeness: 'partial' }; } });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), { debounceMs: 0 });
  try {
    h.act(() => { root.current.setOpen(true); root.current.setQuery('DOCUMENTI'); }); await flush(h);
    await h.act(async () => root.current.loadMore());
    h.act(() => root.current.setQuery('Documenti')); await flush(h);
    assert.deepEqual(root.current.searchOptions.map(o => o.key), ['first', 'second']);
    assert.equal(root.current.query, 'Documenti'); assert.deepEqual(calls, ['documenti', 'documenti']);
  } finally { root.unmount(); }
});
test('catalog selection and limit updates reuse pages and retain source partial diagnostics', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const rows = Array.from({ length: 35 }, (_, i) => option(`entry${i}`)); let calls = 0;
  const src = source({ discover: async () => { calls++; return { options: rows, completeness: 'partial' }; } });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true, initialCatalogLimit: p.initialCatalogLimit ?? 30 }), { selectedOptions: [] });
  try {
    h.act(() => root.current.setOpen(true)); await flush(h);
    root.render({ selectedOptions: ['entry0'], initialOptions: [rows[0]], initialCatalogLimit: 40 }); await flush(h);
    assert.equal(calls, 1); assert.equal(root.current.sections.flatMap(s => s.options).length, 34);
    assert.match(root.current.warning, /parziale/);
    assert.deepEqual(root.current.selectedOptions, ['entry0']);
  } finally { root.unmount(); }
});

test('catalog filters once, fills after selected exclusion and remains bounded when the source continues', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let requests = 0;
  const rows = Array.from({ length: 35 }, (_, i) => option(`entry${i}`));
  const src = source({ discover: async req => { requests++; return { options: req.continuation ? rows.slice(20) : rows.slice(0, 20), continuation: req.continuation ? undefined : 'next', completeness: req.continuation ? 'complete' : 'partial' }; } });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true, initialCatalogLimit: 30 }), { defaultSelectedOptions: ['entry0', 'entry1'], initialOptions: rows.slice(0, 2), filterOptions: options => options.slice(1) });
  try { h.act(() => root.current.setOpen(true)); await flush(h); const result = root.current.sections.flatMap(s => s.options); assert.equal(requests, 2); assert.equal(result.length, 30); assert.equal(result[0].key, 'entry3'); assert.match(root.current.catalogHint, /altri elementi/); }
  finally { root.unmount(); }
});
test('stale hydration cannot publish unresolved validation after source binding changes', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let oldFinish, newFinish; const observations = [];
  const first = source({ identity: 'first', hydrate: () => new Promise(resolve => { oldFinish = resolve; }) });
  const second = source({ identity: 'second', hydrate: () => new Promise(resolve => { newFinish = resolve; }) });
  const root = h.mount(p => { const result = usePickerRuntime({ selectedOptions: ['a'] }, p, { multiselect: true }); observations.push(result.validation); return result; }, first);
  try {
    await flush(h); root.render(second); await flush(h);
    await h.act(async () => oldFinish({ resolved: [], unresolved: [{ key: 'a' }] }));
    assert.equal(root.current.loading, true); assert.equal(root.current.validation, undefined);
    await h.act(async () => newFinish({ resolved: [{ key: 'a', option: option('a') }], unresolved: [] }));
    assert.deepEqual(root.current.selectedOptions, ['a']); assert.equal(root.current.loading, false); assert.deepEqual(observations.filter(Boolean), []);
  } finally { root.unmount(); }
});

test('complete catalogs show all available rows without the ListPicker limit or hint', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const rows = Array.from({ length: 245 }, (_, i) => option(`entry${i}`));
  const src = source({ completeCatalog: true, discover: async () => ({ options: rows, completeness: 'complete' }) });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), { defaultSelectedOptions: ['entry0'], initialOptions: [rows[0]] });
  try { h.act(() => root.current.setOpen(true)); await flush(h); assert.equal(root.current.sections.flatMap(s => s.options).length, 244); assert.equal(root.current.catalogHint, undefined); assert.deepEqual(root.current.selectedOptions, ['entry0']); }
  finally { root.unmount(); }
});

test('contextual Site suggestions share composition/cache and report hub failures as SharePoint', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let requests = 0; const errors = [];
  const src = source({ personalSuggestions: async () => { requests++; return { currentSite: [option('current')], hubSite: [option('hub')], errors: [{ source: 'hubSite', error: new Error('Hub unavailable') }], completeness: 'partial' }; } });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), { onError: (_, details) => errors.push(details) });
  try {
    h.act(() => root.current.setOpen(true)); await flush(h);
    assert.deepEqual(root.current.sections.map(s => s.source), ['currentSite', 'hubSite']);
    assert.deepEqual(errors[0], { source: 'sharePoint', operation: 'load', suggestionSource: 'hubSite' });
    assert.match(root.current.error.message, /Hub unavailable/);
    root.render({ label: 'Cosmetic' }); await flush(h); assert.equal(requests, 1);
    h.act(() => root.current.setOpen(false)); h.act(() => root.current.setOpen(true)); await flush(h); assert.equal(requests, 1);
  } finally { root.unmount(); }
});


test('pinned Site context survives typed queries, consumer cap and updates without duplicates or repeated loading', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let loads = 0;
  const src = source({ persistentSuggestions: true, personalSuggestions: async () => { loads++; return { currentSite: [option('current')], hubSite: [option('hub')], completeness: 'complete' }; }, discover: async () => ({ options: [option('current'), option('hub'), option('remote')], completeness: 'complete' }) });
  const props = { availableOptions: Array.from({length: 6}, (_, i) => option(`consumer${i}`)), debounceMs: 0 };
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), props);
  try {
    h.act(() => { root.current.setQuery('unrelated'); root.current.setOpen(true); }); await flush(h);
    assert.deepEqual(root.current.pinnedSections.map(s => s.source), ['currentSite', 'hubSite']);
    assert.deepEqual(root.current.searchOptions.map(o => o.key), ['remote']);
    h.act(() => root.current.setQuery('')); await flush(h);
    assert.equal(root.current.sections.flatMap(s => s.options).length, 5);
    root.render({ ...props, availableOptions: [{...option('current'), key: 'consumer-current', text: 'Consumer title'}] }); await flush(h);
    assert.equal(root.current.pinnedSections[0].options[0].text, 'Consumer title');
    root.render({ ...props, selectedOptions: ['current'], initialOptions: [option('current')] }); await flush(h);
    assert.deepEqual(root.current.pinnedSections.map(s => s.source), ['hubSite']);
    assert.deepEqual(root.current.selectedOptions, ['current']); assert.equal(loads, 1);
  } finally { root.unmount(); }
});

test('search feedback is pending throughout debounce and ignores obsolete query completion', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const completions = new Map();
  const src = source({ discover: req => new Promise(resolve => completions.set(req.query, resolve)) });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), { debounceMs: 15 });
  try {
    h.act(() => { root.current.setOpen(true); root.current.setQuery('old'); });
    assert.equal(root.current.searchFeedback.phase, 'pending');
    await h.act(async () => { await new Promise(resolve => setTimeout(resolve, 25)); });
    h.act(() => root.current.setQuery('new'));
    assert.equal(root.current.searchFeedback.phase, 'pending');
    await h.act(async () => completions.get('old')({ options: [], completeness: 'complete' }));
    assert.equal(root.current.searchFeedback.phase, 'pending');
    await h.act(async () => { await new Promise(resolve => setTimeout(resolve, 25)); });
    await h.act(async () => completions.get('new')({ options: [], completeness: 'complete' }));
    assert.equal(root.current.searchFeedback.phase, 'complete');
    assert.equal(root.current.searchFeedback.matchesVisibleChoices, false);
    h.act(() => root.current.retry());
    assert.equal(root.current.searchFeedback.phase, 'pending', 'retry cannot publish a previous empty result during debounce');
  } finally { root.unmount(); }
});

test('search feedback distinguishes fixed/selected matches and excludes contextual failures from the query outcome', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const src = source({ persistentSuggestions: true, personalSuggestions: async () => ({ currentSite: [option('current')], errors: [{ source: 'hubSite', error: new Error('Hub denied') }], completeness: 'partial' }), discover: async req => ({ options: req.query === 'visible' ? [option('current'), option('selected')] : [], completeness: 'complete' }) });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), { defaultSelectedOptions: ['selected'], initialOptions: [option('selected')], debounceMs: 0 });
  try {
    h.act(() => { root.current.setOpen(true); root.current.setQuery('visible'); }); await flush(h);
    assert.equal(root.current.searchFeedback.phase, 'complete'); assert.equal(root.current.searchFeedback.matchesVisibleChoices, true);
    assert.equal(root.current.searchFeedback.error, undefined); assert.equal(root.current.searchOptions.length, 0);
    assert.equal(root.current.error.message, 'Hub denied');
    h.act(() => root.current.setQuery('absent')); await flush(h);
    assert.equal(root.current.searchFeedback.phase, 'complete'); assert.equal(root.current.searchFeedback.matchesVisibleChoices, false);
    assert.equal(root.current.pinnedSections.length, 1); assert.deepEqual(root.current.selectedOptions, ['selected']);
  } finally { root.unmount(); }
});

test('reopening a settled typed menu preserves accumulated pages without a pending frame', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let requests = 0;
  const observed = [];
  const src = source({ caseInsensitive: true, discover: async req => {
    requests++;
    return req.continuation ? { options: [option('second')], completeness: 'complete' } : { options: [option('first')], continuation: 'next', completeness: 'partial' };
  } });
  const root = h.mount(p => {
    const result = usePickerRuntime(p, src, { multiselect: true });
    observed.push({ phase: result.searchFeedback.phase, keys: result.searchOptions.map(o => o.key) });
    return result;
  }, { debounceMs: 0 });
  try {
    h.act(() => { root.current.setOpen(true); root.current.setQuery('query'); }); await flush(h);
    await h.act(async () => root.current.loadMore());
    h.act(() => root.current.setOpen(false));
    observed.length = 0;
    h.act(() => root.current.setOpen(true)); await flush(h);
    assert.deepEqual(root.current.searchOptions.map(o => o.key), ['first', 'second']);
    assert.equal(requests, 2, 'reopen reuses both completed pages');
    assert.ok(observed.every(frame => frame.phase === 'complete' && frame.keys.join(',') === 'first,second'), JSON.stringify(observed));
    h.act(() => root.current.retry()); await flush(h);
    assert.equal(requests, 3, 'explicit retry still loads a fresh first page');
    assert.deepEqual(root.current.searchOptions.map(o => o.key), ['first']);
  } finally { root.unmount(); }
});

test('consumer eligibility reports pending feedback without empty or cosmetic-update flashes', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let finish, requests = 0;
  const observed = [];
  const src = source({ preflight: true, verifyCandidate: candidate => { requests++; return new Promise(resolve => { finish = () => resolve(candidate); }); } });
  const props = { availableOptions: [option('inside')] };
  const root = h.mount(p => {
    const result = usePickerRuntime(p, src, { multiselect: false });
    if (result.open) observed.push({ loading: result.loading, keys: result.sections.flatMap(s => s.options.map(o => o.key)) });
    return result;
  }, props);
  try {
    h.act(() => root.current.setOpen(true)); await flush(h);
    assert.equal(root.current.loading, true, 'eligibility is pending, not an empty menu');
    assert.deepEqual(root.current.sections, []);
    await h.act(async () => finish());
    assert.equal(root.current.loading, false);
    assert.deepEqual(root.current.sections.flatMap(s => s.options.map(o => o.key)), ['inside']);
    assert.ok(observed.every(frame => frame.loading || frame.keys.includes('inside')), JSON.stringify(observed));
    observed.length = 0;
    root.render({ availableOptions: [{ ...option('inside'), text: 'Renamed' }], label: 'Cosmetic' }); await flush(h);
    h.act(() => root.current.setOpen(false)); h.act(() => root.current.setOpen(true)); await flush(h);
    assert.equal(requests, 1);
    assert.ok(observed.every(frame => !frame.loading && frame.keys.includes('inside')), JSON.stringify(observed));
    assert.equal(root.current.sections[0].options[0].text, 'Renamed');
  } finally { root.unmount(); }
});

test('lazy contextual suggestions stay pending until settled and refresh without stale rows', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let finish, requests = 0;
  const observed = [];
  const src = source({ persistentSuggestions: true, personalSuggestions: () => { requests++; return new Promise(resolve => { finish = resolve; }); } });
  const root = h.mount(p => {
    const result = usePickerRuntime(p, src, { multiselect: false });
    if (result.open) observed.push({ loading: result.loading, keys: result.pinnedSections.flatMap(s => s.options.map(o => o.key)) });
    return result;
  }, { refreshKey: 0 });
  try {
    h.act(() => root.current.setOpen(true)); await flush(h);
    assert.equal(root.current.loading, true, 'context metadata is pending, not an empty menu');
    await h.act(async () => finish({ currentSite: [option('current')], completeness: 'complete' }));
    assert.equal(root.current.loading, false);
    assert.ok(observed.every(frame => frame.loading || frame.keys.includes('current')), JSON.stringify(observed));
    observed.length = 0;
    root.render({ refreshKey: 0, label: 'Cosmetic' });
    h.act(() => root.current.setOpen(false)); h.act(() => root.current.setOpen(true)); await flush(h);
    assert.equal(requests, 1);
    assert.ok(observed.every(frame => !frame.loading && frame.keys.includes('current')), JSON.stringify(observed));
    root.render({ refreshKey: 1 }); await flush(h);
    assert.equal(requests, 2);
    assert.equal(root.current.loading, true);
    assert.deepEqual(root.current.pinnedSections, [], 'refresh cannot retain a contextual row without current verification');
    await h.act(async () => finish({ currentSite: [], completeness: 'complete' }));
    assert.equal(root.current.loading, false);
    assert.deepEqual(root.current.pinnedSections, []);
  } finally { root.unmount(); }
});

test('retained search pages still respect cache expiry and page-size changes', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const { createPickerRequestCache } = h.load('controls/pickers/picker-request-cache.internal.ts');
  const { PickerRuntimeContext } = h.load('controls/pickers/picker-runtime-context.internal.tsx');
  let now = 0;
  const requests = [];
  const cache = createPickerRequestCache(context.owner, { now: () => now });
  const src = source({ discover: async req => { requests.push(req.pageSize); return { options: [option(`size${req.pageSize}-request${requests.length}`)], completeness: 'complete' }; } });
  const Wrapper = props => h.React.createElement(PickerRuntimeContext.Provider, { value: cache }, props.children);
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), { debounceMs: 0, suggestionsLimit: 10 }, Wrapper);
  try {
    h.act(() => { root.current.setOpen(true); root.current.setQuery('query'); }); await flush(h);
    h.act(() => root.current.setOpen(false));
    now = 300001;
    h.act(() => root.current.setOpen(true)); await flush(h);
    assert.deepEqual(requests, [10, 10], 'expired first page reloads on reopen');
    assert.deepEqual(root.current.searchOptions.map(o => o.key), ['size10-request2']);
    root.render({ debounceMs: 0, suggestionsLimit: 20 });
    assert.equal(root.current.searchFeedback.phase, 'pending', 'a different page size cannot show the old result as settled');
    await flush(h);
    assert.deepEqual(root.current.searchOptions.map(o => o.key), ['size20-request3']);
    root.render({ debounceMs: 0, suggestionsLimit: 10 }); await flush(h);
    assert.deepEqual(requests, [10, 10, 20]);
    assert.deepEqual(root.current.searchOptions.map(o => o.key), ['size10-request2'], 'switching back adopts the matching cached page');
  } finally { root.unmount(); cache.dispose(); }
});

test('reopening during continuation preserves pending rows and adopts the completed page once', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let finish, requests = 0;
  const src = source({ discover: async req => {
    requests++;
    return req.continuation ? new Promise(resolve => { finish = resolve; }) : { options: [option('first')], continuation: 'next', completeness: 'partial' };
  } });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), { debounceMs: 0 });
  try {
    h.act(() => { root.current.setOpen(true); root.current.setQuery('query'); }); await flush(h);
    let pending;
    h.act(() => { pending = root.current.loadMore(); }); await flush(h);
    h.act(() => root.current.setOpen(false)); h.act(() => root.current.setOpen(true)); await flush(h);
    assert.equal(root.current.loading, true);
    assert.deepEqual(root.current.searchOptions.map(o => o.key), ['first']);
    await h.act(async () => { finish({ options: [option('second')], completeness: 'complete' }); await pending; });
    assert.equal(requests, 2);
    assert.equal(root.current.searchFeedback.phase, 'complete');
    assert.deepEqual(root.current.searchOptions.map(o => o.key), ['first', 'second']);
  } finally { root.unmount(); }
});

test('an older continuation cannot replace results after the search page size changes', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let finish;
  const src = source({ discover: async req => req.continuation ? new Promise(resolve => { finish = resolve; }) :
    { options: [option(`size${req.pageSize}`)], continuation: 'next', completeness: 'partial' } });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), { debounceMs: 0, suggestionsLimit: 10 });
  try {
    h.act(() => { root.current.setOpen(true); root.current.setQuery('query'); }); await flush(h);
    let pending;
    h.act(() => { pending = root.current.loadMore(); }); await flush(h);
    root.render({ debounceMs: 0, suggestionsLimit: 20 }); await flush(h);
    await h.act(async () => { finish({ options: [option('obsolete')], completeness: 'complete' }); await pending; });
    assert.deepEqual(root.current.searchOptions.map(o => o.key), ['size20']);
    assert.equal(root.current.searchFeedback.phase, 'partial');
    assert.equal(root.current.loading, false);
  } finally { root.unmount(); }
});

for (const replacement of ['expiration', 'explicit sibling refresh']) {
  test(`reopening adopts a sibling's replacement search entry after ${replacement}`, async () => {
    const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
    const { createPickerRequestCache } = h.load('controls/pickers/picker-request-cache.internal.ts');
    const { PickerRuntimeContext } = h.load('controls/pickers/picker-runtime-context.internal.tsx');
    let now = 0, requests = 0;
    const cache = createPickerRequestCache(context.owner, { now: () => now });
    const src = source({ discover: async () => ({ options: [option(`generation${++requests}`)], completeness: 'complete' }) });
    const Wrapper = props => h.React.createElement(PickerRuntimeContext.Provider, { value: cache }, props.children);
    const useRuntime = p => usePickerRuntime(p, src, { multiselect: true });
    const first = h.mount(useRuntime, { debounceMs: 0 }, Wrapper);
    const sibling = h.mount(useRuntime, { debounceMs: 0 }, Wrapper);
    try {
      h.act(() => { first.current.setOpen(true); first.current.setQuery('query'); }); await flush(h);
      assert.deepEqual(first.current.searchOptions.map(o => o.key), ['generation1']);
      h.act(() => first.current.setOpen(false));
      if (replacement === 'expiration') now = 300001;
      else h.act(() => sibling.current.retry());
      h.act(() => { sibling.current.setOpen(true); sibling.current.setQuery('query'); }); await flush(h);
      assert.deepEqual(sibling.current.searchOptions.map(o => o.key), ['generation2']);
      h.act(() => first.current.setOpen(true)); await flush(h);
      assert.deepEqual(first.current.searchOptions.map(o => o.key), ['generation2'], 'a valid replacement cache entry cannot validate older retained results');
      assert.equal(requests, 2, 'reopen adopts the current cache entry without another transport request');
    } finally { first.unmount(); sibling.unmount(); cache.dispose(); }
  });
}

test('a continuation from a replaced shared entry cannot restore older rows after reopen', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const { createPickerRequestCache } = h.load('controls/pickers/picker-request-cache.internal.ts');
  const { PickerRuntimeContext } = h.load('controls/pickers/picker-runtime-context.internal.tsx');
  let finish, requests = 0;
  const cache = createPickerRequestCache(context.owner);
  const src = source({ discover: async req => req.continuation ? new Promise(resolve => { finish = resolve; }) :
    { options: [option(`generation${++requests}`)], continuation: 'next', completeness: 'partial' } });
  const Wrapper = props => h.React.createElement(PickerRuntimeContext.Provider, { value: cache }, props.children);
  const useRuntime = p => usePickerRuntime(p, src, { multiselect: true });
  const first = h.mount(useRuntime, { debounceMs: 0 }, Wrapper);
  const sibling = h.mount(useRuntime, { debounceMs: 0 }, Wrapper);
  try {
    h.act(() => { first.current.setOpen(true); first.current.setQuery('query'); }); await flush(h);
    let pending;
    h.act(() => { pending = first.current.loadMore(); }); await flush(h);
    h.act(() => { first.current.setOpen(false); sibling.current.retry(); sibling.current.setOpen(true); sibling.current.setQuery('query'); }); await flush(h);
    h.act(() => first.current.setOpen(true)); await flush(h);
    assert.deepEqual(first.current.searchOptions.map(o => o.key), ['generation2']);
    await h.act(async () => { finish({ options: [option('obsolete')], completeness: 'complete' }); await pending; });
    assert.deepEqual(first.current.searchOptions.map(o => o.key), ['generation2']);
    assert.equal(requests, 2);
  } finally { first.unmount(); sibling.unmount(); cache.dispose(); }
});
test('reopening a replaced entry immediately invalidates old continuations before a nonzero debounce', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const { createPickerRequestCache } = h.load('controls/pickers/picker-request-cache.internal.ts');
  const { PickerRuntimeContext } = h.load('controls/pickers/picker-runtime-context.internal.tsx');
  let finish, requests = 0;
  const frames = [];
  const cache = createPickerRequestCache(context.owner);
  const src = source({ discover: async req => req.continuation ? new Promise(resolve => { finish = resolve; }) :
    { options: [option(`generation${++requests}`)], continuation: 'next', completeness: 'partial' } });
  const Wrapper = props => h.React.createElement(PickerRuntimeContext.Provider, { value: cache }, props.children);
  const first = h.mount(p => {
    const result = usePickerRuntime(p, src, { multiselect: true });
    if (result.open) frames.push({ phase: result.searchFeedback.phase, keys: result.searchOptions.map(o => o.key) });
    return result;
  }, { debounceMs: 80 }, Wrapper);
  const sibling = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), { debounceMs: 0 }, Wrapper);
  try {
    h.act(() => { first.current.setOpen(true); first.current.setQuery('query'); });
    await h.act(async () => new Promise(resolve => setTimeout(resolve, 100)));
    let pending;
    h.act(() => { pending = first.current.loadMore(); }); await flush(h);
    h.act(() => { first.current.setOpen(false); sibling.current.retry(); sibling.current.setOpen(true); sibling.current.setQuery('query'); }); await flush(h);
    frames.length = 0;
    h.act(() => first.current.setOpen(true));
    await h.act(async () => { finish({ options: [option('obsolete')], completeness: 'complete' }); await pending; });
    assert.deepEqual(first.current.searchOptions.map(o => o.key), ['generation2'], 'the cached replacement is adopted before the debounce delay');
    assert.ok(frames.every(frame => !frame.keys.includes('obsolete')), JSON.stringify(frames));
    assert.equal(requests, 2, 'replacement adoption starts no additional discovery request');
    await h.act(async () => new Promise(resolve => setTimeout(resolve, 100)));
    assert.equal(requests, 2, 'no redundant first-page timer remains after adoption');
  } finally { first.unmount(); sibling.unmount(); cache.dispose(); }
});
test('reopening a replaced settled entry never announces retained old rows as settled', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const { createPickerRequestCache } = h.load('controls/pickers/picker-request-cache.internal.ts');
  const { PickerRuntimeContext } = h.load('controls/pickers/picker-runtime-context.internal.tsx');
  let requests = 0;
  const frames = [], cache = createPickerRequestCache(context.owner);
  const src = source({ discover: async () => ({ options: [option(`generation${++requests}`)], completeness: 'complete' }) });
  const Wrapper = props => h.React.createElement(PickerRuntimeContext.Provider, { value: cache }, props.children);
  const first = h.mount(p => {
    const result = usePickerRuntime(p, src, { multiselect: true });
    if (result.open) frames.push({ phase: result.searchFeedback.phase, keys: result.searchOptions.map(o => o.key) });
    return result;
  }, { debounceMs: 0 }, Wrapper);
  const sibling = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), { debounceMs: 0 }, Wrapper);
  try {
    h.act(() => { first.current.setOpen(true); first.current.setQuery('query'); }); await flush(h);
    h.act(() => { first.current.setOpen(false); sibling.current.retry(); sibling.current.setOpen(true); sibling.current.setQuery('query'); }); await flush(h);
    first.render({ debounceMs: 80 }); frames.length = 0;
    h.act(() => first.current.setOpen(true)); await flush(h);
    assert.ok(frames.every(frame => !frame.keys.includes('generation1') || frame.phase === 'pending'), JSON.stringify(frames));
    assert.deepEqual(first.current.searchOptions.map(o => o.key), ['generation2']);
    assert.equal(requests, 2);
  } finally { first.unmount(); sibling.unmount(); cache.dispose(); }
});
test('first-page cache provenance preserves one debounced request and its pending completion', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let finish, requests = 0;
  const src = source({ caseInsensitive: true, discover: () => { requests++; return new Promise(resolve => { finish = resolve; }); } });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), { debounceMs: 20 });
  try {
    h.act(() => { root.current.setOpen(true); root.current.setQuery('Query'); });
    await h.act(async () => new Promise(resolve => setTimeout(resolve, 40)));
    assert.equal(requests, 1); assert.equal(root.current.loading, true);
    root.render({ debounceMs: 20, label: 'Cosmetic' }); h.act(() => root.current.setQuery('QUERY'));
    await h.act(async () => finish({ options: [option('resolved')], completeness: 'complete' }));
    assert.deepEqual(root.current.searchOptions.map(o => o.key), ['resolved']);
    assert.equal(root.current.searchFeedback.phase, 'complete');
    await h.act(async () => new Promise(resolve => setTimeout(resolve, 40)));
    assert.equal(requests, 1, 'subscribing to the current in-flight entry must not restart its timer or transport');
  } finally { root.unmount(); }
});

test('ContentType consumer models override equivalent catalog rows and update labels without resetting query or selection', async () => {
  const calls = [];
  const h = createHarness({ '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } } });
  const web = context.webUrl;
  const rows = ['0x0101', '0x0101AB', '0x0101AC'].map((StringId, index) => ({ StringId, Name: index === 0 ? 'Selected' : 'Same name', Group: 'Documents', Hidden: false, ReadOnly: false, Sealed: false }));
  const ctContext = { ...context, spClient: { get: async uri => { calls.push(uri); return { ok: true, status: 200, json: async () => uri.includes('/rootWeb?') ? { Url: web } : uri.includes('getById') ? rows.find(row => decodeURIComponent(uri).toLowerCase().includes("'" + row.StringId.toLowerCase() + "'")) : { value: rows } }; } } };
  const { createContentTypePickerSource } = h.load('controls/pickers/content-type-picker.source.internal.ts');
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const ct = (index, key, text = rows[index].Name) => ({ kind: 'contentType', key, text, name: rows[index].Name, group: rows[index].Group, hidden: false, readOnly: false, sealed: false, id: rows[index].StringId, webUrl: web,
    value: { kind: 'contentType', scope: 'site', webUrl: web, contentTypeId: rows[index].StringId } });
  const selected = ct(0, 'retained'), consumer = ct(1, 'consumer', 'Consumer label');
  const props = { defaultSelectedOptions: ['retained'], initialOptions: [selected], availableOptions: [consumer], debounceMs: 0 };
  const root = h.mount(p => usePickerRuntime(p, createContentTypePickerSource(p, ctContext), { multiselect: false }), props);
  try {
    h.act(() => root.current.setOpen(true)); await flush(h);
    assert.deepEqual(root.current.selectedOptions, ['retained']); assert.equal(root.current.selectedItems[0].key, 'retained');
    const options = root.current.sections.flatMap(s => s.options);
    assert.equal(options.length, 2); assert.equal(options[0].key, 'consumer'); assert.equal(options[0].text, 'Consumer label');
    assert.equal(options.filter(o => o.id === '0x0101AB').length, 1);
    h.act(() => root.current.setQuery('sa')); await flush(h);
    const count = calls.length;
    root.render({ ...props, label: 'Cosmetic', presentation: 'tagPicker', availableOptions: [{ ...consumer, text: 'Updated consumer' }] }); await flush(h);
    assert.equal(calls.length, count); assert.equal(root.current.query, 'sa'); assert.deepEqual(root.current.selectedOptions, ['retained']);
    assert.equal(root.current.searchOptions.find(o => o.key === 'consumer').text, 'Updated consumer');
    h.act(() => root.current.setQuery('')); await flush(h);
    root.render({ ...props, availableOptions: [] }); await flush(h);
    assert.equal(root.current.selectedItems[0].text, 'Selected'); assert.equal(calls.length, count);
    await h.act(async () => root.current.remove({ type: 'click' }, 'retained'));
    assert.deepEqual(root.current.selectedOptions, []);
  } finally { root.unmount(); }
});


test('selection acceptance is explicit for controlled consumers even when props are retained or synchronously updated', async () => {
  for (const update of [false, true]) {
    const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
    const changes = [];
    const props = { selectedOptions: [], availableOptions: [option('a')], onOptionSelect: (_event, data) => {
      changes.push(data.selectedOptions);
      if (update) root.render({ ...props, selectedOptions: data.selectedOptions });
    } };
    const root = h.mount(p => usePickerRuntime(p, source(), { multiselect: false }), props);
    try {
      h.act(() => root.current.setQuery('query'));
      let accepted;
      await h.act(async () => { accepted = await root.current.select({ type: 'click' }, 'a'); });
      assert.equal(accepted, true);
      assert.deepEqual(changes, [['a']]);
      assert.deepEqual(root.current.selectedOptions, update ? ['a'] : []);
      assert.equal(root.current.query, '');
    } finally { root.unmount(); }
  }
});

test('refused selection returns false and preserves query and existing selection', async () => {
  for (const scenario of ['missing', 'disabled', 'limit', 'verification', 'confirmation', 'validation']) {
    const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
    const changes = [];
    const src = source({ verifyCandidate: async candidate => scenario === 'verification' ? undefined : candidate,
      confirm: scenario === 'confirmation' ? async () => [] : undefined });
    const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), {
      defaultSelectedOptions: ['a'], initialOptions: [option('a')], availableOptions: [option('b')], disabled: scenario === 'disabled',
      selectionLimit: scenario === 'limit' ? 1 : undefined,
      validateSelection: scenario === 'validation' ? () => 'Refused' : undefined,
      onOptionSelect: (_event, data) => changes.push(data),
    });
    try {
      h.act(() => root.current.setQuery('query'));
      let accepted;
      await h.act(async () => { accepted = await root.current.select({ type: 'click' }, scenario === 'missing' ? 'missing' : 'b'); });
      assert.equal(accepted, false, scenario);
      assert.equal(root.current.query, 'query', scenario);
      assert.deepEqual(root.current.selectedOptions, ['a'], scenario);
      assert.deepEqual(changes, [], scenario);
    } finally { root.unmount(); }
  }
});

test('pending and invalidated selection attempts return false without accepting stale work', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  let finish, emitted = 0;
  const src = source({ verifyCandidate: () => new Promise(resolve => { finish = resolve; }) });
  const props = { availableOptions: [option('a')], onOptionSelect: () => emitted++ };
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: false }), props);
  try {
    let pending, accepted;
    h.act(() => { pending = root.current.select({ type: 'click' }, 'a'); });
    await h.act(async () => { accepted = await root.current.select({ type: 'click' }, 'a'); });
    assert.equal(accepted, false, 'confirming guard');
    root.render({ ...props, selectedOptions: ['replacement'] });
    await h.act(async () => { finish(option('a')); accepted = await pending; });
    assert.equal(accepted, false, 'selection revision invalidates the action');
    assert.equal(emitted, 0);
  } finally { root.unmount(); }
});

test('selected aliases exclude pinned current and hub rows and typed results across canonical URL variants', async () => {
  const h = createHarness(), { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const selected = { ...option('current'), key: 'retained', url: 'https://tenant.test/sites/CURRENT/', value: { kind: 'site', url: 'https://tenant.test/sites/CURRENT/' } };
  const src = source({ persistentSuggestions: true, personalSuggestions: async () => ({ currentSite: [option('current')], hubSite: [{ ...option('current'), key: 'hub-alias' }], completeness: 'complete' }),
    discover: async () => ({ options: [{ ...option('current'), key: 'search-alias' }, option('other')], completeness: 'complete' }) });
  const root = h.mount(p => usePickerRuntime(p, src, { multiselect: true }), { defaultSelectedOptions: ['retained'], initialOptions: [selected], debounceMs: 0 });
  try {
    h.act(() => { root.current.setOpen(true); root.current.setQuery('current'); }); await flush(h);
    assert.deepEqual(root.current.pinnedSections, []);
    assert.deepEqual(root.current.searchOptions.map(row => row.key), ['other']);
    assert.equal(root.current.searchFeedback.matchesVisibleChoices, true);
  } finally { root.unmount(); }
});

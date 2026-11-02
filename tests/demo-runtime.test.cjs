const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./fixtures/web-part-shell-harness.cjs');
const { pilotModules, button } = require('./fixtures/executable-examples/pilot-harness.cjs');
const { Simulate, act } = require('react-dom/test-utils');
const { activeRules } = require('./fixtures/sx-harness.cjs');
const titles = {
  properties: 'Edit the current host description with merge and updater',
  session: 'Save and clear an instance scoped session value',
  local: 'Save and clear an instance scoped local value',
  logger: 'Write native log levels and inspect recent calls',
  performance: 'Measure a cancellable local timed operation',
};
// Independent expectations for this Runtime cohort only, never a production catalogue inventory.
const runtimeCases = [
  { ownerKey: 'packages/spfx-react-toolkit/src/hooks/useSPFxProperties.ts#useSPFxProperties', title: titles.properties, kind: 'write' },
  { ownerKey: 'packages/spfx-react-toolkit/src/hooks/useSPFxStorage.ts#useSPFxSessionStorage', title: titles.session, kind: 'write' },
  { ownerKey: 'packages/spfx-react-toolkit/src/hooks/useSPFxStorage.ts#useSPFxLocalStorage', title: titles.local, kind: 'write' },
  { ownerKey: 'packages/spfx-react-toolkit/src/hooks/useSPFxLogger.ts#useSPFxLogger', title: titles.logger, kind: 'write' },
  { ownerKey: 'packages/spfx-react-toolkit/src/hooks/useSPFxPerformance.ts#useSPFxPerformance', title: titles.performance, kind: 'read' },
];
function assertRuntimeCatalogue(entries) {
  const owned = entries.filter(entry => runtimeCases.some(expected => expected.ownerKey === entry.ownerKey));
  for (const expected of runtimeCases) {
    const matches = owned.filter(entry => entry.ownerKey === expected.ownerKey && entry.title === expected.title);
    assert.equal(matches.length, 1, 'Exactly one Runtime case: ' + expected.ownerKey + '/' + expected.title);
    const entry = matches[0];
    assert.equal(entry.metadata.host, 'webPart');
    assert.equal(entry.metadata.contentPlacement, 'configuration');
    assert.equal(entry.metadata.kind, expected.kind);
  }
  assert.equal(owned.length, 5, 'Exactly five Runtime cases for the expected canonical owners');
}
function fixture(t, kind, options = {}) {
  const h = createHarness(); t.after(() => h.close());
  const modules = pilotModules(h);
  options.boundary?.(modules.toolkit, h.React);
  const scenario = modules.scenario(titles[kind]);
  let reset, regions;
  const host = h.instance(options.id || 'runtime-one');
  host.properties.description = 'Initial description';
  Object.assign(host.properties, options.properties);
  if (options.absentDescription) delete host.properties.description;
  let refreshes = 0;
  host.context.propertyPane.refresh = () => { refreshes++; };
  const Demo = scenario.Demo;
  const { DemoCard } = modules.load('shared/DemoCard.tsx');
  function Frame() {
    const [revision, setRevision] = h.React.useState(0);
    reset = () => setRevision(value => value + 1);
    return h.React.createElement(Demo, { resetRevision: revision, render: next => {
      regions = next;
      return h.React.createElement(DemoCard, { title: scenario.entry.title, iconName: 'Info', contentPlacement: 'configuration', ...next });
    } });
  }
  const view = h.mount(host, h.React.createElement(Frame));
  let currentHost = host;
  const timers = new Map(), retained = new Map(); let sequence = 100000;
  const set = window.setTimeout.bind(window), clear = window.clearTimeout.bind(window);
  t.mock.method(window, 'setTimeout', (callback, delay, ...args) => {
    if (![250, 3000].includes(delay)) return set(callback, delay, ...args);
    const id = ++sequence; timers.set(id, { callback, delay }); retained.set(id, callback); return id;
  });
  t.mock.method(window, 'clearTimeout', id => { if (retained.has(id)) timers.delete(id); else clear(id); });
  const command = name => {
    const element = h.React.Children.toArray(regions.commands.props.children).find(node => node.props.children === name);
    assert.ok(element, name); return element.props.onClick;
  };
  return { h, host, view, entry: scenario.entry, catalogue: modules.catalogue, timers, retained, command,
    get container() { return view.container; },
    get refreshes() { return refreshes; },
    click(name) { h.act(() => Simulate.click(button(view.container, name))); },
    change(value) { const input = view.container.querySelector('input'); h.act(() => Simulate.change(input, { target: { value } })); },
    reset() { h.act(reset); },
    replace(id = host.context.instanceId) { currentHost = h.instance(id, currentHost.context.serviceScope); currentHost.properties.description = 'Replacement description'; view.render(currentHost); return currentHost; },
    expire(id) { timers.delete(id); h.act(() => retained.get(id)()); },
    unmount: view.unmount,
  };
}
async function flush(callback = () => {}) { await act(async () => { callback(); await Promise.resolve(); await Promise.resolve(); }); }

test('Runtime has exactly five source-owned cases with canonical owners and configuration placement', t => {
  for (const [kind, title] of Object.entries(titles)) {
    const f = fixture(t, kind);
    assert.equal(f.entry.title, title);
    assertRuntimeCatalogue(f.catalogue.executableExamples);
    assert.equal(f.entry.metadata.host, 'webPart');
    assert.equal(f.entry.metadata.contentPlacement, 'configuration');
    const symbol = { properties: 'useSPFxProperties', session: 'useSPFxSessionStorage', local: 'useSPFxLocalStorage', logger: 'useSPFxLogger', performance: 'useSPFxPerformance' }[kind];
    const ownerFile = kind === 'local' || kind === 'session' ? 'useSPFxStorage' : symbol;
    assert.equal(f.entry.ownerKey, `packages/spfx-react-toolkit/src/hooks/${ownerFile}.ts#${symbol}`);
    assert.equal(f.entry.metadata.kind, kind === 'performance' ? 'read' : 'write');
    f.unmount();
  }
});

test('Runtime replaces notification ownership; retained old expiry cannot erase the new feedback, reset or departure', t => {
  const f = fixture(t, 'properties');
  assert.equal(f.timers.size, 0);
  f.click('setProperties'); const first = [...f.timers.keys()][0];
  f.click('updateProperties'); const second = [...f.timers.keys()][0];
  assert.equal(f.timers.size, 1); assert.notEqual(first, second);
  f.expire(first);
  assert.match(f.container.textContent, /Properties updated with updateProperties/);
  assert.ok(f.timers.has(second), 'old expiry cannot lose the replacement handle');
  f.reset(); assert.equal(f.timers.size, 0);
  f.expire(second); assert.doesNotMatch(f.container.textContent, /Properties updated with/);
  f.click('setProperties'); f.unmount(); assert.equal(f.timers.size, 0);
});

test('Properties merge and complete replacement preserve unrelated keys, synchronize drafts and reject retained commands', t => {
  const native = { keep: 1 };
  const f = fixture(t, 'properties', { properties: { nativeKey: native } });
  assert.equal(f.refreshes, 0, 'mount does not update the host');
  const input = f.container.querySelector('input'); f.change('User description');
  const old = f.command('setProperties');
  f.click('setProperties'); assert.equal(f.host.properties.description, 'User description');
  f.click('updateProperties'); assert.equal(f.host.properties.description, 'User description updated');
  assert.equal(f.host.properties.nativeKey, native); assert.equal(f.host.properties.title, 'initial'); assert.equal(f.host.properties.extra, 'keep');
  const writes = f.refreshes;
  f.change('uncommitted'); f.reset(); assert.equal(f.container.querySelector('input'), input); assert.equal(input.value, 'User description updated');
  f.h.act(old); assert.equal(f.host.properties.description, 'User description updated');
  assert.equal(f.refreshes, writes, 'reset and retained commands do not write the host');
  const beforeReplacement = f.command('setProperties'); const next = f.replace();
  assert.equal(input.value, 'Replacement description');
  f.h.act(beforeReplacement); assert.equal(next.properties.description, 'Replacement description');
  const afterReplacement = f.command('setProperties'); f.unmount(); f.h.act(afterReplacement);
  assert.equal(next.properties.description, 'Replacement description');
});

for (const kind of ['session', 'local']) {
  const label = kind === 'session' ? 'Session' : 'Local';
  test(`${label} real scoped storage: explicit saves/clear, defaults, reset, two-instance isolation and retained command invalidation`, t => {
    const f = fixture(t, kind); const other = fixture(t, kind, { id: 'runtime-two' });
    const storage = kind === 'session' ? sessionStorage : localStorage;
    const key = `spfx:runtime-one:demo-${kind}-key`, otherKey = `spfx:runtime-two:demo-${kind}-key`;
    const writes = t.mock.method(Object.getPrototypeOf(storage), 'setItem');
    const removes = t.mock.method(Object.getPrototypeOf(storage), 'removeItem');
    assert.equal(storage.getItem(key), null, 'mount does not persist fallback');
    assert.equal(button(f.container, 'Save ' + label).disabled, true);
    f.change('Owned value'); const stale = f.command('Save ' + label); const staleClear = f.command('Clear ' + label); f.click('Save ' + label);
    assert.equal(storage.getItem(key), '"Owned value"'); assert.equal(storage.getItem(otherKey), null);
    assert.equal(f.container.querySelector('input').value, ''); assert.match(f.container.textContent, /persistence attempted/);
    const input = f.container.querySelector('input'); f.change('unsaved'); f.reset();
    assert.equal(f.container.querySelector('input'), input); assert.equal(input.value, ''); assert.match(f.container.textContent, /Owned value/);
    assert.equal(writes.mock.callCount(), 1); assert.equal(removes.mock.callCount(), 0, 'reset performs no persistence operation');
    f.h.act(() => { stale(); staleClear(); }); assert.equal(storage.getItem(key), '"Owned value"');
    storage.setItem(key, '"same-document external"'); f.reset();
    assert.match(f.container.textContent, /Owned value/, 'reset uses current hook state, without forced reread');
    const beforeContext = f.command('Clear ' + label), beforeSave = f.command('Save ' + label); f.replace('runtime-three'); f.h.act(() => { beforeContext(); beforeSave(); });
    assert.equal(storage.getItem(key), '"same-document external"');
    f.change('New scope'); const unmountedSave = f.command('Save ' + label); f.click('Save ' + label); const unmounted = f.command('Clear ' + label);
    f.unmount(); f.h.act(() => { unmounted(); unmountedSave(); }); assert.equal(storage.getItem(`spfx:runtime-three:demo-${kind}-key`), '"New scope"');
    other.change('Other'); other.click('Save ' + label); assert.equal(storage.getItem(otherKey), '"Other"');
    other.reset(); other.click('Clear ' + label); assert.equal(storage.getItem(otherKey), null); assert.doesNotMatch(other.container.textContent, /persistence attempted/);
    const reentry = fixture(t, kind, { id: 'runtime-one' }); assert.match(reentry.container.textContent, /same-document external/);
  });
  test(`${label} storage keeps false and empty JSON visible and reports best-effort blocked persistence`, t => {
    const storage = kind === 'session' ? sessionStorage : localStorage;
    storage.setItem(`spfx:runtime-one:demo-${kind}-key`, 'false');
    const f = fixture(t, kind); assert.match(f.container.textContent, /false/);
    const prototype = Object.getPrototypeOf(storage);
    const write = t.mock.method(prototype, 'setItem', () => { throw new Error('blocked'); });
    const remove = t.mock.method(prototype, 'removeItem', () => { throw new Error('blocked'); });
    f.change('Local only'); f.click('Save ' + label); assert.match(f.container.textContent, /Local only/); assert.equal(write.mock.callCount(), 1);
    f.click('Clear ' + label); assert.equal(remove.mock.callCount(), 1); assert.doesNotMatch(f.container.textContent, /Current value: Local only/);
    write.mock.restore(); remove.mock.restore();
    storage.setItem(`spfx:runtime-one:demo-${kind}-key`, '""');
    f.unmount(); const empty = fixture(t, kind); assert.equal(empty.container.querySelector('output').textContent, 'Current value: ');
  });
}

test('Runtime feedback occupies one configuration status and logger uses native levels with last-five collapsed history', t => {
  const properties = fixture(t, 'properties'); properties.click('setProperties');
  assert.equal(properties.container.querySelectorAll('[role=status]').length, 1);
  properties.unmount();
  const calls = [];
  for (const level of ['info', 'warn', 'error']) t.mock.method(console, level, (...args) => calls.push({ level, args }));
  const f = fixture(t, 'logger'); assert.equal(calls.length, 0);
  const stale = f.command('Info');
  f.click('Info');
  const styles = () => activeRules(f.view.renderer, f.container.querySelector('[data-runtime-result]').className).join('\n');
  assert.match(styles(), /colorNeutralBackground2/); assert.match(styles(), /fontFamilyMonospace/);
  assert.match(styles(), /spacingVerticalS/); assert.match(styles(), /borderRadiusMedium/);
  f.click('Warning'); assert.match(styles(), /colorStatusWarningBackground1/);
  assert.equal(f.container.querySelector('[data-runtime-result]').textContent, 'WARNING: SPFx toolkit demo warning');
  const disclosure = button(f.container, 'Advanced diagnostics'); assert.equal(disclosure.getAttribute('aria-expanded'), 'false');
  f.h.act(() => Simulate.click(disclosure));
  const details = f.container.querySelector('[aria-label="Recent log calls"]');
  assert.match(details.textContent, /SPFx toolkit demo info/); assert.match(details.textContent, /SPFx toolkit demo warning/);
  for (let i = 0; i < 5; i++) f.click('Error');
  assert.equal(f.container.querySelectorAll('[aria-label="Recent log calls"] li').length, 5);
  assert.match(styles(), /colorStatusDangerBackground1/);
  assert.equal(f.container.querySelector('[data-runtime-result]').hasAttribute('style'), false);
  assert.deepEqual(calls.map(call => call.level), ['info', 'warn', 'error', 'error', 'error', 'error', 'error']);
  assert.ok(calls.every(call => /SPFx toolkit demo (info|warning|error)$/.test(call.args[0])));
  f.reset(); f.h.act(stale); assert.equal(calls.length, 7); assert.equal(f.container.querySelector('[aria-label="Recent log calls"]'), null);
  const replaced = f.command('Info'); f.replace(); f.h.act(replaced); assert.equal(calls.length, 7);
  const departed = f.command('Info'); f.unmount(); f.h.act(departed); assert.equal(calls.length, 7);
});

test('Runtime real performance cancels before callback entry and releases held waits, retaining external timeline entries', async t => {
  const f = fixture(t, 'performance');
  const perf = require('node:perf_hooks').performance;
  const descriptor = Object.getOwnPropertyDescriptor(global, 'performance');
  Object.defineProperty(global, 'performance', { configurable: true, value: perf });
  t.after(() => { perf.clearMarks('runtime-external'); perf.clearMeasures('spfx-toolkit-demo'); Object.defineProperty(global, 'performance', descriptor); });
  perf.mark('runtime-external');
  assert.equal(f.timers.size, 0);
  const old = f.command('Run timed operation');
  f.click('Run timed operation'); f.reset();
  await flush();
  assert.equal(f.timers.size, 0, 'cancel before callback entry creates no delay');
  assert.doesNotMatch(f.container.textContent, /Timed operation completed/);
  f.h.act(old); await flush(); assert.equal(f.timers.size, 0, 'retained reset command cannot start');
  f.click('Run timed operation'); await flush();
  assert.equal(f.timers.size, 1); const held = [...f.timers.keys()][0];
  f.reset(); await flush(); assert.equal(f.timers.size, 0);
  assert.equal(perf.getEntriesByType('mark').filter(mark => mark.name.startsWith('spfx-toolkit-demo-start-')).length, 0, 'cancel releases owned wait and hook start mark');
  f.click('Run timed operation'); await flush(); const fresh = [...f.timers.keys()][0];
  f.expire(held); await flush();
  assert.match(f.container.textContent, /Timing operation/); assert.ok(f.timers.has(fresh));
  f.expire(fresh); await flush();
  assert.match(f.container.textContent, /Timed operation completed/);
  assert.match(f.container.textContent, /Duration: [0-9.]+ms/);
  assert.equal(perf.getEntriesByName('runtime-external').length, 1, 'unrelated explicit marks remain');
  assert.ok(perf.getEntriesByName('spfx-toolkit-demo').length >= 1, 'the actual hook retains measurement entries');
  f.click('Run timed operation'); await flush();
  const replaced = f.command('Run timed operation'); f.replace(); await flush();
  assert.equal(f.timers.size, 0); f.h.act(replaced); await flush(); assert.equal(f.timers.size, 0);
  f.click('Run timed operation'); await flush(); f.unmount(); await flush(); assert.equal(f.timers.size, 0);
});

for (const outcome of ['resolve', 'reject']) {
  test(`Performance injected operation boundary: old ${outcome} cannot alter a fresh run, same-ID context or unmounted UI`, async t => {
    const calls = [];
    const f = fixture(t, 'performance', { boundary(toolkit, React) {
      const useReal = toolkit.useSPFxPerformance;
      toolkit.useSPFxPerformance = () => {
        const actual = useReal();
        return React.useMemo(() => ({ ...actual, time: (name, callback) => new Promise((resolve, reject) => calls.push({ name, callback, resolve, reject })) }), [actual]);
      };
    } });
    const original = f.command('Run timed operation');
    f.h.act(() => { original(); original(); });
    assert.equal(calls.length, 1, 'synchronous ownership blocks duplicate runs before React commits');
    assert.equal(calls[0].name, 'spfx-toolkit-demo');
    f.reset(); f.click('Run timed operation'); assert.equal(calls.length, 2);
    await flush(() => calls[0][outcome](outcome === 'resolve' ? { result: 'Obsolete result', durationMs: 999 } : new Error('obsolete failure')));
    assert.match(f.container.textContent, /Timing operation/); assert.doesNotMatch(f.container.textContent, /Obsolete|obsolete failure|999/);
    await flush(() => calls[1].reject(new Error('Independent operation failure')));
    assert.equal(f.container.querySelector('[role=alert]').textContent, 'Independent operation failure');
    assert.equal(button(f.container, 'Run timed operation').disabled, false);
    f.click('Run timed operation'); const oldContext = f.command('Run timed operation');
    f.replace(); await flush(() => calls[2].resolve({ result: 'Old host result', durationMs: 999 }));
    assert.doesNotMatch(f.container.textContent, /Old host result/);
    f.h.act(oldContext); assert.equal(calls.length, 3, 'old same-ID native-context command cannot start');
    f.click('Run timed operation'); assert.equal(calls.length, 4);
    const errors = [];
    const error = t.mock.method(console, 'error', (...args) => errors.push(args));
    f.unmount(); await flush(() => calls[3].reject(new Error('departed failure')));
    assert.deepEqual(errors, [], 'late failure is handled without React unmounted-state warnings');
    error.mock.restore();
  });
}

for (const kind of ['session', 'local']) test(`${kind} feedback replaces its timer and stale expiry cannot erase a later save`, t => {
  const f = fixture(t, kind), label = kind === 'session' ? 'Session' : 'Local';
  f.change('First'); f.click('Save ' + label); const first = [...f.timers.keys()][0];
  f.change('Second'); f.click('Save ' + label); const second = [...f.timers.keys()][0];
  assert.notEqual(first, second); assert.equal(f.timers.size, 1);
  f.expire(first); assert.match(f.container.textContent, /persistence attempted/); assert.ok(f.timers.has(second));
  f.click('Clear ' + label); assert.equal(f.timers.size, 1, 'clear does not install a new notification');
  f.expire(second); assert.equal(f.container.querySelector('[role=status]'), null);
  f.change('Departure'); f.click('Save ' + label); f.unmount(); assert.equal(f.timers.size, 0);
});

test('Logger exposes a native console sink failure without fabricating a successful UI call', t => {
  t.mock.method(console, 'info', () => { throw new Error('Synthetic console failure'); });
  const f = fixture(t, 'logger'); f.click('Info');
  assert.equal(f.container.querySelector('[role=alert]').textContent, 'Synthetic console failure');
  assert.equal(f.container.querySelector('[data-runtime-result]'), null);
  assert.equal(f.container.querySelector('[aria-label="Recent log calls"]'), null);
});

test('Properties missing description displays empty and adopts native host edits without writing on mount or reset', t => {
  const f = fixture(t, 'properties', { absentDescription: true });
  const input = f.container.querySelector('input');
  assert.equal(input.value, ''); assert.equal(f.refreshes, 0);
  assert.equal(Object.hasOwn(f.host.properties, 'description'), false);
  f.host.properties.description = 'Property pane description'; f.view.render(f.host);
  assert.equal(input.value, 'Property pane description');
  f.change('Uncommitted local draft'); f.reset();
  assert.equal(f.container.querySelector('input'), input); assert.equal(input.value, 'Property pane description');
  assert.equal(f.refreshes, 0);
});

test('Runtime catalogue contract accepts unrelated owner growth without mutating the real catalogue', t => {
  const f = fixture(t, 'properties'); const entries = f.catalogue.executableExamples;
  const original = entries.slice();
  assertRuntimeCatalogue(entries);
  const unrelated = entries.find(entry => !runtimeCases.some(expected => expected.ownerKey === entry.ownerKey));
  assert.ok(unrelated, 'real catalogue includes another canonical owner');
  const title = 'Additional unrelated source-owned case';
  const id = require('node:crypto').createHash('sha256').update(JSON.stringify([unrelated.ownerKey, title])).digest('hex');
  assertRuntimeCatalogue([...entries, { ...unrelated, id, title }]);
  assert.deepEqual(entries, original, 'ownership checks leave the real catalogue unchanged');
});
test('Runtime catalogue contract rejects each missing expected canonical owner/title pair', t => {
  const f = fixture(t, 'properties'); const entries = f.catalogue.executableExamples;
  for (const expected of runtimeCases) {
    const remaining = entries.filter(entry => entry.ownerKey !== expected.ownerKey || entry.title !== expected.title);
    assert.throws(() => assertRuntimeCatalogue(remaining), error => error.code === 'ERR_ASSERTION' && error.message.includes('Exactly one Runtime case: ' + expected.ownerKey + '/' + expected.title));
  }
});
test('Runtime catalogue contract rejects each duplicated expected canonical owner/title pair', t => {
  const f = fixture(t, 'properties'); const entries = f.catalogue.executableExamples;
  for (const expected of runtimeCases) {
    const entry = entries.find(entry => entry.ownerKey === expected.ownerKey && entry.title === expected.title);
    assert.ok(entry);
    assert.throws(() => assertRuntimeCatalogue([...entries, { ...entry }]), error => error.code === 'ERR_ASSERTION' && error.message.includes('Exactly one Runtime case: ' + expected.ownerKey + '/' + expected.title));
  }
});

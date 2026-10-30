const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}
function sdkFixture(initial = { app: { theme: 'default', host: { name: 'TeamsModern' } } }) {
  const pending = deferred();
  let listener;
  const calls = { context: 0, registration: 0 };
  const sdk = { app: {
    getContext() { calls.context++; return pending.promise; },
    registerOnThemeChangeHandler(callback) { calls.registration++; listener = callback; },
  } };
  return { sdk, calls, pending, initial, emit(theme) { assert.equal(typeof listener, 'function'); listener(theme); } };
}
function runtimeFixture(t, sdk) {
  const React = require('react');
  const HostContext = React.createContext(undefined);
  const h = createHarness({ './useSPFxContext': { useSPFxContext: () => ({ spfxContext: React.useContext(HostContext) }) } });
  t.after(h.close);
  h.context.current = { sdks: { microsoftTeams: sdk } };
  const state = h.load('core/state.internal.tsx');
  const createStore = h.load('core/runtime-store.internal.ts').createSPFxRuntimeStore;
  const useTeams = h.load('hooks/useSPFxTeams.ts').useSPFxTeams;
  function mount(hook = useTeams, store = createStore()) {
    function Wrapper({ children }) {
      return h.React.createElement(HostContext.Provider, { value: h.context.current },
        h.React.createElement(state.SPFxRuntimeStoreContext.Provider, { value: store }, children));
    }
    const view = h.mount(hook, {}, Wrapper);
    return { ...view, store, get current() { return view.current; } };
  }
  async function flush(action = () => {}) {
    await h.act(async () => { action(); for (let i = 0; i < 10; i++) await Promise.resolve(); });
  }
  return { ...h, mount, useTeams, createStore, flush };
}

test('two Teams hooks in one runtime store acquire context and register once, retaining live subscription', async t => {
  const sdk = sdkFixture();
  const f = runtimeFixture(t, sdk.sdk);
  const view = f.mount(() => [f.useTeams(), f.useTeams()]);
  await f.flush(() => sdk.pending.resolve(sdk.initial));
  assert.equal(sdk.calls.context, 1);
  assert.equal(sdk.calls.registration, 1);
  f.act(() => sdk.emit('DARK'));
  assert.equal(view.store.getState().teams.theme, 'dark');
  assert.deepEqual(view.current.map(info => info.theme), ['dark', 'dark']);
  view.render();
  assert.equal(sdk.calls.context, 1);
});

test('two stores share one SDK dispatcher and keep other runtime properties independent through cleanup', async t => {
  const sdk = sdkFixture(); const f = runtimeFixture(t, sdk.sdk);
  const first = f.mount(); const second = f.mount();
  first.store.setState({ properties: { title: 'one' } });
  second.store.setState({ properties: { title: 'two' } });
  await f.flush(() => sdk.pending.resolve(sdk.initial));
  assert.equal(sdk.calls.registration, 1);
  f.act(() => sdk.emit('contrast'));
  assert.equal(first.store.getState().teams.theme, 'highContrast');
  assert.equal(second.store.getState().teams.theme, 'highContrast');
  assert.equal(first.store.getState().properties.title, 'one');
  assert.equal(second.store.getState().properties.title, 'two');
  first.unmount(); const snapshot = first.store.getState();
  f.act(() => sdk.emit('dark'));
  assert.equal(first.store.getState(), snapshot);
  assert.equal(second.store.getState().teams.theme, 'dark');
  second.unmount(); const last = second.store.getState();
  f.act(() => sdk.emit('default'));
  assert.equal(second.store.getState(), last);
});

test('distinct SDKs remain isolated', async t => {
  const first = sdkFixture(); const second = sdkFixture(); const f = runtimeFixture(t, first.sdk);
  const a = f.mount(); await f.flush(() => first.pending.resolve(first.initial));
  f.context.current = { sdks: { microsoftTeams: second.sdk } };
  const b = f.mount(); await f.flush(() => second.pending.resolve(second.initial));
  f.act(() => first.emit('dark'));
  assert.equal(a.store.getState().teams.theme, 'dark');
  assert.equal(b.store.getState().teams.theme, 'default');
  f.act(() => second.emit('contrast'));
  assert.equal(a.store.getState().teams.theme, 'dark');
  assert.equal(b.store.getState().teams.theme, 'highContrast');
});

test('live theme while getContext is pending wins over stale initial theme', async t => {
  const sdk = sdkFixture(); const f = runtimeFixture(t, sdk.sdk); const view = f.mount();
  f.act(() => sdk.emit('highContrast'));
  await f.flush(() => sdk.pending.resolve({ app: { theme: 'default', host: { name: 'Office' } } }));
  assert.equal(view.store.getState().teams.theme, 'highContrast');
  assert.equal(view.store.getState().teams.context.app.host.name, 'Office');
});

test('SDK replacement and unmount ignore old pending contexts and old live events', async t => {
  const old = sdkFixture(); const next = sdkFixture(); const f = runtimeFixture(t, old.sdk); const view = f.mount();
  f.context.current = { sdks: { microsoftTeams: { teamsJs: next.sdk } } }; view.render();
  await f.flush(() => next.pending.resolve({ app: { theme: 'dark' } }));
  const snapshot = view.store.getState().teams;
  await f.flush(() => old.pending.resolve({ app: { theme: 'default' } }));
  f.act(() => old.emit('contrast'));
  assert.equal(view.store.getState().teams, snapshot);
  view.unmount();
  f.act(() => next.emit('default'));
  assert.equal(view.store.getState().teams, snapshot);
});

test('unmounted pending store stays unchanged when context resolves', async t => {
  const sdk = sdkFixture(); const f = runtimeFixture(t, sdk.sdk); const view = f.mount();
  view.unmount(); const snapshot = view.store.getState();
  await f.flush(() => sdk.pending.resolve(sdk.initial));
  f.act(() => sdk.emit('dark'));
  assert.equal(view.store.getState(), snapshot);
});

test('one hook unmount keeps same-store remaining subscription active and remount does not reregister', async t => {
  const sdk = sdkFixture(); const f = runtimeFixture(t, sdk.sdk); const store = f.createStore();
  const a = f.mount(f.useTeams, store); const b = f.mount(f.useTeams, store);
  await f.flush(() => sdk.pending.resolve(sdk.initial)); a.unmount();
  f.act(() => sdk.emit('dark')); assert.equal(store.getState().teams.theme, 'dark');
  b.unmount(); const c = f.mount(f.useTeams, store); await f.flush();
  assert.equal(sdk.calls.registration, 1);
  f.act(() => sdk.emit('contrast')); assert.equal(store.getState().teams.theme, 'highContrast');
  c.unmount();
});

test('v2 context and registration preferred over v1', async t => {
  let v1Contexts = 0; let v1Registrations = 0; const sdk = sdkFixture();
  sdk.sdk.getContext = () => { v1Contexts++; };
  sdk.sdk.registerOnThemeChangeHandler = () => { v1Registrations++; };
  const f = runtimeFixture(t, sdk.sdk); const view = f.mount();
  await f.flush(() => sdk.pending.resolve(sdk.initial));
  assert.equal(v1Contexts, 0); assert.equal(v1Registrations, 0);
  assert.equal(view.store.getState().teams.supported, true);
});

test('failed v2 acquisition falls back to v1 and reports failure', async t => {
  const errors = t.mock.method(console, 'error', () => {}); let listener; let v1Contexts = 0;
  const sdk = { app: { getContext: async () => { throw new Error('v2 unavailable'); } }, getContext(callback) { v1Contexts++; callback({ theme: 'contrast' }); }, registerOnThemeChangeHandler(callback) { listener = callback; } };
  const f = runtimeFixture(t, sdk); const view = f.mount(); await f.flush();
  assert.equal(v1Contexts, 1); assert.equal(view.store.getState().teams.theme, 'highContrast');
  f.act(() => listener('dark')); assert.equal(view.store.getState().teams.theme, 'dark');
  assert.ok(errors.mock.calls.some(call => call.arguments.some(argument => argument?.message === 'v2 unavailable')));
});

test('theme registration failure reports error and preserves acquired initial context', async t => {
  const errors = t.mock.method(console, 'error', () => {});
  const initial = { app: { theme: 'dark' } };
  const sdk = { app: { getContext: async () => initial, registerOnThemeChangeHandler() { throw new Error('registration failed'); } } };
  const f = runtimeFixture(t, sdk); const view = f.mount(); await f.flush();
  assert.equal(view.store.getState().teams.context, initial);
  assert.equal(view.store.getState().teams.theme, 'dark');
  assert.ok(errors.mock.calls.some(call => call.arguments.some(argument => argument?.message === 'registration failed')));
});

test('failed SDK acquisition is observable and remains unsupported', async t => {
  const errors = t.mock.method(console, 'error', () => {});
  const f = runtimeFixture(t, { app: { getContext: async () => { throw new Error('context failed'); } } });
  const view = f.mount(); await f.flush();
  assert.equal(view.store.getState().teams.supported, false);
  assert.equal(view.store.getState().teams.initialized, true);
  assert.ok(errors.mock.calls.some(call => call.arguments.some(argument => argument?.message === 'context failed')));
});

for (const [name, expected] of [['TeamsModern', 'Teams'], ['Office', 'Office'], ['OutlookWin32', 'Outlook'], ['UnknownHost', 'SharePoint']]) {
  test('environment uses SPFx SDK result for ' + name + ' without SDKs on PageContext', async t => {
    const sdk = sdkFixture(); const f = runtimeFixture(t, sdk.sdk);
    assert.equal(f.pageContext.sdks, undefined);
    const useEnvironment = f.load('hooks/useSPFxEnvironmentInfo.ts').useSPFxEnvironmentInfo;
    let info;
    const view = f.mount(() => { info = useEnvironment(); return info; });
    await f.flush(() => sdk.pending.resolve({ app: { theme: 'dark', host: { name } } }));
    assert.equal(info.type, expected);
    assert.equal(info.isTeams, expected === 'Teams');
    assert.equal(info.isOffice, expected === 'Office');
    assert.equal(info.isOutlook, expected === 'Outlook');
    view.unmount();
  });
}

test('environment updates on SDK replacement and preserves local/on-prem fallback', async t => {
  const first = sdkFixture(); const next = sdkFixture(); const f = runtimeFixture(t, first.sdk);
  const useEnvironment = f.load('hooks/useSPFxEnvironmentInfo.ts').useSPFxEnvironmentInfo;
  let info; const view = f.mount(() => { info = useEnvironment(); return info; });
  await f.flush(() => first.pending.resolve({ app: { host: { name: 'Office' } } }));
  assert.equal(info.type, 'Office');
  f.context.current = { sdks: { microsoftTeams: next.sdk } }; view.render();
  assert.equal(info.type, 'SharePoint');
  await f.flush(() => next.pending.resolve({ app: { host: { name: 'Outlook' } } }));
  assert.equal(info.type, 'Outlook');
  f.context.current = {}; f.pageContext.legacyPageContext = { isOnPremises: true }; view.render(); await f.flush();
  assert.equal(info.type, 'SharePointOnPrem');
  f.pageContext.web.absoluteUrl = 'https://localhost:4321/workbench.aspx'; view.render();
  assert.equal(info.type, 'Local'); assert.equal(info.isWorkbench, true);
});

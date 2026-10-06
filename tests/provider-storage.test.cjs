const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./provider-test-harness.cjs');

function run(body) {
  const harness = createHarness();
  try { body(harness); } finally { harness.close(); }
}
function observeProvider(h) {
  const { useSPFxProperties } = h.load('hooks/useSPFxProperties.ts');
  const { useSPFxDisplayMode } = h.load('hooks/useSPFxDisplayMode.ts');
  const { useSPFxRuntimeSelector } = h.load('core/state.internal.tsx');
  let latest;
  function Child() {
    latest = { ...useSPFxProperties(), ...useSPFxDisplayMode(), theme: useSPFxRuntimeSelector(state => state.theme) };
    return h.React.createElement('span', null, latest.properties?.title);
  }
  return { Child, get latest() { return latest; } };
}

test('provider observes displayMode changes on the same host instance', () => run(h => {
  const host = h.instance();
  const view = observeProvider(h);
  const fixture = h.mount(host, view.Child);
  for (const mode of [2, 1]) {
    host.displayMode = mode;
    fixture.render();
    assert.equal(view.latest.mode, mode);
    assert.equal(view.latest.isEdit, mode === 2);
  }
}));

test('Property Pane in-place mutation after hook setter reaches consumers', () => run(h => {
  const host = h.instance();
  let refreshes = 0;
  host.context.propertyPane.refresh = () => refreshes++;
  const view = observeProvider(h);
  const fixture = h.mount(host, view.Child);
  h.act(() => view.latest.setProperties({ title: 'hook' }));
  assert.equal(host.properties.title, 'hook');
  assert.equal(refreshes, 1);
  host.properties.title = 'pane';
  fixture.render();
  assert.equal(view.latest.properties.title, 'pane');
  assert.equal(fixture.container.textContent, 'pane');
  assert.equal(refreshes, 1, 'host-origin changes must not refresh the pane');
  const currentProperties = view.latest.properties;
  fixture.render();
  assert.equal(view.latest.properties, currentProperties, 'an unchanged parent rerender preserves runtime snapshot');
  assert.equal(refreshes, 1);
}));

test('provider observes in-place property deletion before and after hook updates', () => run(h => {
  const host = h.instance();
  const view = observeProvider(h);
  const fixture = h.mount(host, view.Child);
  delete host.properties.extra;
  fixture.render();
  assert.equal('extra' in view.latest.properties, false);
  h.act(() => view.latest.setProperties({ extra: 'again' }));
  delete host.properties.extra;
  fixture.render();
  assert.equal('extra' in view.latest.properties, false);
  h.act(() => view.latest.updateProperties(() => ({ title: 'only' })));
  assert.deepEqual(host.properties, { title: 'only' });
}));

test('two providers keep properties and subscriptions isolated through cleanup', () => run(h => {
  const a = h.instance('a'); const b = h.instance('b');
  const viewA = observeProvider(h); const viewB = observeProvider(h);
  const fixtureA = h.mount(a, viewA.Child); const fixtureB = h.mount(b, viewB.Child);
  h.act(() => viewA.latest.setProperties({ title: 'changed-a' }));
  assert.equal(viewB.latest.properties.title, 'initial');
  assert.equal(b.properties.title, 'initial');
  const oldSetter = viewA.latest.setProperties;
  fixtureA.unmount();
  assert.equal(a.context.serviceScope.handlers.size, 0);
  h.act(() => oldSetter({ title: 'after-unmount' }));
  assert.equal(a.properties.title, 'changed-a');
  h.act(() => b.context.serviceScope.emit({ name: 'b-theme' }));
  assert.equal(viewB.latest.theme.name, 'b-theme');
  fixtureB.unmount();
  assert.equal(b.context.serviceScope.handlers.size, 0);
}));

test('replacing a ready scope hides children until the new scope finishes', () => run(h => {
  const host = h.instance();
  const oldScope = host.context.serviceScope;
  const view = observeProvider(h);
  const fixture = h.mount(host, view.Child);
  const pending = h.scope(false);
  host.context = { ...host.context, serviceScope: pending };
  fixture.render();
  assert.equal(fixture.container.textContent, '');
  assert.equal(oldScope.handlers.size, 0);
  h.act(() => pending.finish());
  assert.equal(fixture.container.textContent, 'initial');
  assert.equal(pending.handlers.size, 1);
}));

test('old scope completion cannot release a newer pending scope', () => run(h => {
  const first = h.scope(false); const second = h.scope(false);
  const host = h.instance('pending', first);
  const view = observeProvider(h);
  const fixture = h.mount(host, view.Child);
  host.context.serviceScope = second;
  fixture.render();
  h.act(() => first.finish());
  assert.equal(fixture.container.textContent, '');
  h.act(() => second.finish());
  assert.equal(fixture.container.textContent, 'initial');
}));

for (const persisted of [false, true]) test('inline object storage default settles (' + (persisted ? 'persisted' : 'missing') + ')', () => run(h => {
  const { useSPFxLocalStorage } = h.load('hooks/useSPFxStorage.ts');
  if (persisted) localStorage.setItem('spfx:one:prefs', '{"count":7}');
  let renders = 0; let latest;
  function Child() {
    assert.ok(++renders < 15, 'storage must settle without repeated effect updates');
    latest = useSPFxLocalStorage('prefs', { count: 0 });
    return null;
  }
  const fixture = h.mount(h.instance(), Child);
  assert.deepEqual(latest.value, { count: persisted ? 7 : 0 });
  h.act(() => latest.setValue(previous => ({ count: previous.count + 1 })));
  fixture.render();
  assert.deepEqual(latest.value, { count: persisted ? 8 : 1 });
  assert.ok(renders < 10);
}));

test('default changes preserve current value and remove uses latest default', () => run(h => {
  const { useSPFxLocalStorage } = h.load('hooks/useSPFxStorage.ts');
  let latest;
  function Child({ fallback }) { latest = useSPFxLocalStorage('prefs', fallback); return null; }
  const fixture = h.mount(h.instance(), Child, { fallback: { count: 1 } });
  fixture.render(undefined, { fallback: { count: 8 } });
  assert.deepEqual(latest.value, { count: 1 }, 'a new fallback alone does not reset a missing key');
  h.act(() => latest.setValue({ count: 9 }));
  const currentValue = latest.value;
  fixture.render(undefined, { fallback: { count: 2 } });
  assert.equal(latest.value, currentValue, 'a new fallback alone does not reread persisted objects');
  assert.deepEqual(latest.value, { count: 9 });
  h.act(() => latest.remove());
  assert.deepEqual(latest.value, { count: 2 });
  assert.equal(localStorage.getItem('spfx:one:prefs'), null);
}));

test('key and storage kind switches use persisted values or the current fallback', () => run(h => {
  const { useSPFxLocalStorage, useSPFxSessionStorage } = h.load('hooks/useSPFxStorage.ts');
  let latest;
  // Both public wrappers delegate to the same internal hook, preserving hook order.
  function Child({ storage = 'local', name = 'a', fallback = 1 }) {
    latest = (storage === 'local' ? useSPFxLocalStorage : useSPFxSessionStorage)(name, fallback);
    return null;
  }
  localStorage.setItem('spfx:one:b', '20');
  sessionStorage.setItem('spfx:one:b', '30');
  const fixture = h.mount(h.instance(), Child);
  h.act(() => latest.setValue(10));
  fixture.render(undefined, { name: 'b', fallback: 2 });
  assert.equal(latest.value, 20);
  fixture.render(undefined, { name: 'b', storage: 'session', fallback: 3 });
  assert.equal(latest.value, 30);
  fixture.render(undefined, { name: 'missing', storage: 'session', fallback: 4 });
  assert.equal(latest.value, 4);
  fixture.render(h.instance('other'), { name: 'b', fallback: 5 });
  assert.equal(latest.value, 5);
}));

test('storage events affect only the matching provider and listeners detach on unmount', () => run(h => {
  const { useSPFxLocalStorage } = h.load('hooks/useSPFxStorage.ts');
  const storageListeners = new Set();
  const add = h.window.addEventListener.bind(h.window);
  const remove = h.window.removeEventListener.bind(h.window);
  h.window.addEventListener = (type, listener, options) => {
    if (type === 'storage') storageListeners.add(listener);
    add(type, listener, options);
  };
  h.window.removeEventListener = (type, listener, options) => {
    if (type === 'storage') storageListeners.delete(listener);
    remove(type, listener, options);
  };
  let valueA; let valueB;
  function A() { valueA = useSPFxLocalStorage('prefs', 0); return null; }
  function B() { valueB = useSPFxLocalStorage('prefs', 0); return null; }
  const fixtureA = h.mount(h.instance('a'), A);
  const fixtureB = h.mount(h.instance('b'), B);
  localStorage.setItem('spfx:a:prefs', '11');
  h.act(() => h.window.dispatchEvent(new h.window.StorageEvent('storage', { key: 'spfx:a:prefs', storageArea: localStorage })));
  assert.equal(valueA.value, 11); assert.equal(valueB.value, 0);
  h.act(() => valueB.setValue(8));
  assert.equal(localStorage.getItem('spfx:b:prefs'), '8');
  fixtureA.unmount(); fixtureB.unmount();
  assert.equal(storageListeners.size, 0, 'provider unmount removes its browser storage listeners');
  h.window.addEventListener = add;
  h.window.removeEventListener = remove;
  const prior = console.error; const errors = []; console.error = (...args) => errors.push(args);
  try { h.act(() => h.window.dispatchEvent(new h.window.StorageEvent('storage', { key: 'spfx:a:prefs', storageArea: localStorage }))); }
  finally { console.error = prior; }
  assert.deepEqual(errors, []);
}));


test('in-place scope replacement hides children immediately and ignores completion after unmount', () => run(h => {
  const host = h.instance();
  const oldScope = host.context.serviceScope;
  const pending = h.scope(false);
  const view = observeProvider(h);
  const fixture = h.mount(host, view.Child);
  host.context.serviceScope = pending;
  fixture.render();
  assert.equal(fixture.container.textContent, '');
  assert.equal(oldScope.handlers.size, 0);
  fixture.unmount();
  h.act(() => pending.finish());
  assert.equal(pending.handlers.size, 0);
}));

test('replacing host property bags and provider instances preserves host-origin state', () => run(h => {
  const first = h.instance('first');
  let refreshes = 0;
  first.context.propertyPane.refresh = () => refreshes++;
  const view = observeProvider(h);
  const fixture = h.mount(first, view.Child);
  first.properties = { title: 'replacement' };
  fixture.render();
  assert.deepEqual(view.latest.properties, { title: 'replacement' });
  assert.equal(refreshes, 0);
  const second = h.instance('second');
  second.properties = { title: 'second-host' };
  fixture.render(second);
  assert.equal(view.latest.properties.title, 'second-host');
  assert.equal(first.context.serviceScope.handlers.size, 0);
  assert.equal(second.context.serviceScope.handlers.size, 1);
  h.act(() => view.latest.setProperties({ title: 'second-hook' }));
  assert.equal(first.properties.title, 'replacement');
  assert.equal(second.properties.title, 'second-hook');
}));

test('external storage deletion uses latest fallback and ignores a different key or storage area', () => run(h => {
  const { useSPFxLocalStorage } = h.load('hooks/useSPFxStorage.ts');
  let latest;
  function Child({ fallback }) { latest = useSPFxLocalStorage('prefs', fallback); return null; }
  localStorage.setItem('spfx:one:prefs', '{"count":9}');
  const fixture = h.mount(h.instance(), Child, { fallback: { count: 1 } });
  const currentValue = latest.value;
  fixture.render(undefined, { fallback: { count: 2 } });
  assert.equal(latest.value, currentValue);
  localStorage.removeItem('spfx:one:prefs');
  h.act(() => h.window.dispatchEvent(new h.window.StorageEvent('storage', { key: 'spfx:one:other', storageArea: localStorage })));
  assert.equal(latest.value, currentValue);
  h.act(() => h.window.dispatchEvent(new h.window.StorageEvent('storage', { key: 'spfx:one:prefs', storageArea: sessionStorage })));
  assert.equal(latest.value, currentValue);
  h.act(() => h.window.dispatchEvent(new h.window.StorageEvent('storage', { key: 'spfx:one:prefs', storageArea: localStorage })));
  assert.deepEqual(latest.value, { count: 2 });
}));

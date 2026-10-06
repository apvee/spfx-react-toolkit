const assert = require('node:assert/strict');
const test = require('node:test');
const { React, loadHook, mount, deferred, start, settle, flush } = require('./async-test-harness.cjs');
const ROOT = 'https://tenant.test/sites/collection';
function fixture(overrides = {}, onLayout, mountHook = mount) {
  const boundary = { client: {}, pageContext: { site: { absoluteUrl: ROOT }, web: { absoluteUrl: ROOT } }, baseUrl: 'https://unrelated.test' };
  const calls = [];
  const factories = [];
  const implementations = {
    ensureListReady: async () => {}, get: async key => ({ key, value: 1, description: undefined, id: 1 }),
    list: async () => [], save: async () => {}, remove: async () => {}, canCurrentUserWrite: async () => true,
    ...overrides,
  };
  const hook = loadHook('useSPFxSiteKeyValueStore', 'useSPFxSiteKeyValueStore', {
    './useSPFxSPHttpClient': { useSPFxSPHttpClient: () => boundary },
    './useSPFxPageContext': { useSPFxPageContext: () => boundary.pageContext },
    '../services/spfx-site-key-value-store.service': { createSPFxSiteKeyValueStoreService(client) {
      factories.push(client);
      return Object.fromEntries(Object.entries(implementations).map(([name, fn]) => [name, (...args) => {
        calls.push({ name, args, client });
        return fn(...args);
      }]));
    } },
  });
  const view = mountHook(() => {
    const result = hook();
    React.useLayoutEffect(() => { if (onLayout) onLayout(result); }, [boundary.client, boundary.pageContext.site.absoluteUrl]);
    return result;
  });
  return { boundary, calls, factories, view, rerender: () => view.render({}) };
}
function methodCalls(f, name) { return f.calls.filter(call => call.name === name); }

test('mount checks permissions without provisioning; root and two subsites share one target and callbacks', async () => {
  const f = fixture();
  try {
    await flush();
    assert.equal(f.view.current.isReady, true);
    assert.equal(f.view.current.canWrite, true);
    assert.deepEqual(f.calls.map(call => call.name), ['canCurrentUserWrite']);
    const original = f.view.current;
    for (const webUrl of [ROOT, ROOT + '/one', ROOT + '/two']) {
      f.boundary.pageContext.web.absoluteUrl = webUrl;
      f.boundary.baseUrl = webUrl;
      f.rerender();
      assert.equal(f.view.current.get, original.get);
      await settle(async () => {
        assert.equal((await f.view.current.get('a')).value, 1);
        assert.deepEqual(await f.view.current.list(), []);
        await f.view.current.save('a', { enabled: true }, 'description');
        await f.view.current.remove('a');
      });
    }
    assert.equal(f.factories.length, 1);
    assert.equal(methodCalls(f, 'ensureListReady').length, 0);
    for (const call of f.calls) {
      const index = call.name === 'get' || call.name === 'remove' ? 1 : call.name === 'save' ? 2 : 0;
      assert.equal(call.args[index], ROOT);
    }
    assert.deepEqual(methodCalls(f, 'save')[0].args, ['a', { enabled: true }, ROOT, 'description']);
  } finally { f.view.unmount(); }
});

for (const missing of ['client', 'root']) {
  test(`missing ${missing} rejects every operation and becomes ready once restored`, async () => {
    const f = fixture();
    try {
      if (missing === 'client') f.boundary.client = undefined;
      else f.boundary.pageContext.site.absoluteUrl = ' /// ';
      f.rerender();
      await flush();
      assert.equal(f.view.current.isReady, false);
      assert.equal(f.view.current.canWrite, false);
      const before = f.calls.length;
      await settle(async () => {
        for (const fn of [() => f.view.current.get('a'), () => f.view.current.list(), () => f.view.current.save('a', 1), () => f.view.current.remove('a')]) {
          await assert.rejects(fn, /available|required|ready/i);
        }
      });
      assert.equal(f.calls.length, before);
      if (missing === 'client') f.boundary.client = {};
      else f.boundary.pageContext.site.absoluteUrl = ROOT;
      f.rerender();
      await flush();
      assert.equal(f.view.current.isReady, true);
    } finally { f.view.unmount(); }
  });
}

for (const method of ['get', 'list']) {
  test(`${method} publishes read failures and returns fallback`, async () => {
    const f = fixture({ [method]: async () => { throw 'read denied'; } });
    try {
      await settle(async () => assert.deepEqual(await f.view.current[method]('a'), method === 'get' ? undefined : []));
      assert.equal(f.view.current.error.message, 'read denied');
      assert.equal(f.view.current.writeError, undefined);
      assert.equal(f.view.current.isLoading, false);
    } finally { f.view.unmount(); }
  });
}
for (const method of ['save', 'remove']) {
  test(`${method} publishes write failures, rejects, and does not refresh permission`, async () => {
    const failure = new Error('write denied');
    const f = fixture({ [method]: async () => { throw failure; } });
    try {
      await flush();
      await settle(async () => assert.rejects(() => f.view.current[method]('a', 1), error => error === failure));
      assert.equal(f.view.current.writeError, failure);
      assert.equal(f.view.current.error, undefined);
      assert.equal(f.view.current.isWriting, false);
      assert.equal(methodCalls(f, 'canCurrentUserWrite').length, 1);
    } finally { f.view.unmount(); }
  });
}

test('successful writes refresh permission and latest permission result wins', async () => {
  const permissions = [deferred(), deferred(), deferred()];
  let index = 0;
  const f = fixture({ canCurrentUserWrite: () => permissions[index++].promise });
  try {
    await settle(async () => { await f.view.current.save('a', 1); await f.view.current.remove('a'); });
    assert.equal(index, 3);
    await settle(async () => { permissions[2].resolve(false); });
    await settle(async () => { permissions[1].resolve(true); permissions[0].resolve(true); });
    assert.equal(f.view.current.canWrite, false);
  } finally { f.view.unmount(); }
});

test('permission lookup rejection yields false without an operation error', async () => {
  const f = fixture({ canCurrentUserWrite: async () => { throw new Error('permission failed'); } });
  try { await flush(); assert.equal(f.view.current.canWrite, false); assert.equal(f.view.current.error, undefined); }
  finally { f.view.unmount(); }
});

for (const channel of ['read', 'write']) {
  test(`${channel} counts all operations resolving in reverse order and publishes only latest-started errors`, async () => {
    const first = deferred(), second = deferred(), third = deferred();
    let index = 0;
    const operations = [first, second, third];
    const method = channel === 'read' ? 'get' : 'save';
    const flag = channel === 'read' ? 'isLoading' : 'isWriting';
    const errorKey = channel === 'read' ? 'error' : 'writeError';
    const f = fixture({ [method]: () => operations[index++].promise });
    try {
      const invoke = () => f.view.current[method]('a', 1);
      const a = start(invoke), b = start(invoke);
      // Register rejection handlers before settling deferred writes.
      const observedA = a.catch(error => error), observedB = b.catch(error => error);
      assert.equal(f.view.current[flag], true);
      const newest = new Error('newest');
      await settle(async () => { second.reject(newest); await observedB; });
      assert.equal(f.view.current[flag], true);
      assert.equal(f.view.current[errorKey], newest);
      await settle(async () => { first.reject(new Error('older')); await observedA; });
      assert.equal(f.view.current[flag], false);
      assert.equal(f.view.current[errorKey], newest);
      const c = start(invoke);
      assert.equal(f.view.current[errorKey], undefined);
      await settle(async () => { third.resolve(undefined); await c; });
      assert.equal(f.view.current[flag], false);
    } finally { f.view.unmount(); }
  });
}

test('operations clear only their own error channel and read/write flags remain independent', async () => {
  const read = deferred(), write = deferred();
  let reads = 0, writes = 0;
  const f = fixture({ get: () => ++reads === 1 ? Promise.reject(new Error('read')) : read.promise,
    save: () => ++writes === 1 ? Promise.reject(new Error('write')) : write.promise });
  try {
    await settle(async () => { await f.view.current.get('a'); await assert.rejects(() => f.view.current.save('a', 1)); });
    const readCall = start(() => f.view.current.get('a'));
    assert.equal(f.view.current.error, undefined);
    assert.equal(f.view.current.writeError.message, 'write');
    const writeCall = start(() => f.view.current.save('a', 1));
    assert.equal(f.view.current.writeError, undefined);
    assert.equal(f.view.current.isLoading, true);
    assert.equal(f.view.current.isWriting, true);
    await settle(async () => { read.resolve(undefined); await readCall; });
    assert.equal(f.view.current.isLoading, false);
    assert.equal(f.view.current.isWriting, true);
    await settle(async () => { write.resolve(undefined); await writeCall; });
  } finally { f.view.unmount(); }
});

for (const identityChange of ['root', 'client']) {
  test(`${identityChange} change resets state and ignores old calls, callbacks, and permission completions`, async () => {
    const read = deferred(), write = deferred(), permissions = [];
    const f = fixture({ get: () => read.promise, save: () => write.promise,
      canCurrentUserWrite: () => { const result = deferred(); permissions.push(result); return result.promise; } });
    try {
      const old = f.view.current;
      const a = start(() => old.get('a')), b = start(() => old.save('a', 1));
      const observedB = b.catch(error => error);
      const originalClient = f.boundary.client;
      if (identityChange === 'root') f.boundary.pageContext.site.absoluteUrl = ROOT + '-other';
      else f.boundary.client = {};
      f.rerender();
      assert.equal(f.view.current.isLoading, false);
      assert.equal(f.view.current.isWriting, false);
      assert.notEqual(f.view.current.get, old.get);
      await settle(async () => { permissions[1].resolve(false); permissions[0].resolve(true); read.reject(new Error('old read')); write.reject(new Error('old write')); await a; await observedB; });
      assert.equal(f.view.current.canWrite, false);
      assert.equal(f.view.current.error, undefined);
      assert.equal(f.view.current.writeError, undefined);
      // Retained callbacks remain bound to the old collection/client.
      await settle(async () => { await old.get('old'); await assert.rejects(() => old.save('old', 1)); });
      const lastRead = methodCalls(f, 'get').at(-1), lastWrite = methodCalls(f, 'save').at(-1);
      assert.equal(lastRead.args[1], ROOT);
      assert.equal(lastWrite.args[2], ROOT);
      assert.equal(lastWrite.client, originalClient);
      assert.equal(f.view.current.error, undefined);
      assert.equal(f.view.current.writeError, undefined);
      assert.equal(permissions.length, 2);
    } finally { f.view.unmount(); }
  });
}

test('normalization does not reset storage identity but scalar URL mutation does', async () => {
  const f = fixture();
  try {
    await flush();
    const old = f.view.current;
    f.boundary.pageContext.site.absoluteUrl = ' ' + ROOT + '/// ';
    f.rerender();
    assert.equal(f.view.current.get, old.get);
    assert.equal(f.view.current.canWrite, true);
    f.boundary.pageContext.site.absoluteUrl = ROOT + '-new';
    f.rerender();
    await settle(async () => { await f.view.current.get('new'); });
    assert.equal(methodCalls(f, 'get').at(-1).args[1], ROOT + '-new');
  } finally { f.view.unmount(); }
});

test('unmount ignores read/write/permission completion and starts no refresh', async () => {
  const read = deferred(), write = deferred(), permission = deferred();
  const f = fixture({ get: () => read.promise, save: () => write.promise, canCurrentUserWrite: () => permission.promise });
  const warnings = [];
  const originalError = console.error;
  console.error = (...args) => warnings.push(args.join(' '));
  try {
    const a = start(() => f.view.current.get('a')), b = start(() => f.view.current.save('a', 1));
    f.view.unmount();
    await settle(async () => { read.resolve(undefined); write.resolve(undefined); permission.resolve(true); await a; await b; });
    assert.equal(methodCalls(f, 'canCurrentUserWrite').length, 1);
    assert.equal(warnings.length, 0);
  } finally { console.error = originalError; }
});

for (const method of ['get', 'save']) {
  test(`${method} latest success suppresses earlier failure`, async () => {
    const first = deferred(), second = deferred();
    let index = 0;
    const f = fixture({ [method]: () => [first, second][index++].promise });
    try {
      const a = start(() => f.view.current[method]('a', 1)).catch(error => error);
      const b = start(() => f.view.current[method]('b', 2));
      await settle(async () => { second.resolve(undefined); await b; });
      await settle(async () => { first.reject(new Error('obsolete failure')); await a; });
      assert.equal(f.view.current[method === 'get' ? 'error' : 'writeError'], undefined);
    } finally { f.view.unmount(); }
  });
}

test('successful old write does not refresh permission or affect the new collection', async () => {
  const write = deferred();
  const f = fixture({ save: () => write.promise, canCurrentUserWrite: async url => url === ROOT });
  try {
    await flush();
    const old = f.view.current;
    const a = start(() => old.save('a', 1));
    f.boundary.pageContext.site.absoluteUrl = ROOT + '-next';
    f.rerender();
    await flush();
    await settle(async () => { write.resolve(); await a; await old.save('retained', 1); });
    assert.equal(f.view.current.canWrite, false);
    assert.equal(f.view.current.isWriting, false);
    assert.equal(methodCalls(f, 'canCurrentUserWrite').length, 2);
    assert.equal(methodCalls(f, 'save').at(-1).args[2], ROOT);
    assert.equal(methodCalls(f, 'ensureListReady').length, 0);
  } finally { f.view.unmount(); }
});

test('layout-effect reads and writes remain visible through passive initialization on mount and collection switch', async () => {
  const reads = [deferred(), deferred()], writes = [deferred(), deferred()];
  const readCalls = [], writeCalls = [];
  let readIndex = 0, writeIndex = 0;
  const f = fixture({ get: () => reads[readIndex++].promise, save: () => writes[writeIndex++].promise }, result => {
    readCalls.push(result.get('layout'));
    writeCalls.push(result.save('layout', 1));
  });
  try {
    await flush();
    assert.equal(f.view.current.isLoading, true, 'mount read must remain visible');
    assert.equal(f.view.current.isWriting, true, 'mount write must remain visible');
    f.boundary.pageContext.site.absoluteUrl = ROOT + '-layout';
    f.rerender();
    await flush();
    assert.equal(f.view.current.isLoading, true, 'new root read must remain visible');
    assert.equal(f.view.current.isWriting, true, 'new root write must remain visible');
    await settle(async () => { reads[0].resolve(); writes[0].resolve(); await readCalls[0]; await writeCalls[0]; });
    assert.equal(f.view.current.isLoading, true);
    assert.equal(f.view.current.isWriting, true);
    await settle(async () => { reads[1].resolve(); await readCalls[1]; });
    assert.equal(f.view.current.isLoading, false);
    assert.equal(f.view.current.isWriting, true);
    await settle(async () => { writes[1].resolve(); await writeCalls[1]; });
    assert.equal(f.view.current.isWriting, false);
  } finally { f.view.unmount(); }
});

test('child layout effects can start parent-provided operations before the parent layout guard', async () => {
  const ReactDOM = require('react-dom');
  const { act } = require('react-dom/test-utils');
  const reads = [deferred(), deferred()], writes = [deferred(), deferred()];
  const readCalls = [], writeCalls = [];
  let readIndex = 0, writeIndex = 0;
  function childMount(hook) {
    const container = document.createElement('div');
    document.body.appendChild(container);
    let current;
    function Child({ store }) {
      React.useLayoutEffect(() => {
        readCalls.push(store.get('child-layout'));
        writeCalls.push(store.save('child-layout', 1));
      }, [store.get, store.save]);
      return null;
    }
    function Parent() {
      current = hook();
      return React.createElement(Child, { store: current });
    }
    const render = () => act(() => { ReactDOM.render(React.createElement(Parent), container); });
    render();
    return { get current() { return current; }, render,
      unmount() { act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove(); } };
  }
  const f = fixture({ get: () => reads[readIndex++].promise, save: () => writes[writeIndex++].promise }, undefined, childMount);
  try {
    await flush();
    assert.equal(methodCalls(f, 'get').length, 1);
    assert.equal(methodCalls(f, 'save').length, 1);
    assert.equal(f.view.current.isLoading, true, 'child-started mount read is pending');
    assert.equal(f.view.current.isWriting, true, 'child-started mount write is pending');
    f.boundary.pageContext.site.absoluteUrl = ROOT + '-child';
    f.rerender();
    await flush();
    assert.equal(f.view.current.isLoading, true);
    assert.equal(f.view.current.isWriting, true);
    await settle(async () => { reads[0].resolve(); writes[0].resolve(); await readCalls[0]; await writeCalls[0]; });
    assert.equal(f.view.current.isLoading, true);
    assert.equal(f.view.current.isWriting, true);
    await settle(async () => { reads[1].resolve(); await readCalls[1]; });
    assert.equal(f.view.current.isLoading, false);
    assert.equal(f.view.current.isWriting, true);
    await settle(async () => { writes[1].resolve(); await writeCalls[1]; });
    assert.equal(f.view.current.isWriting, false);
  } finally { f.view.unmount(); }
});

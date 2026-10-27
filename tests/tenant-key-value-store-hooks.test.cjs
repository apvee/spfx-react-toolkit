const assert = require('node:assert/strict');
const test = require('node:test');
const { React, loadHook, mount, deferred, start, settle, flush, catalogBoundary } = require('./async-test-harness.cjs');

function fixture(t, service, permission) {
  const client = {}, discoverAppCatalogUrl = async () => 'https://tenant.test/catalog';
  const boundary = catalogBoundary(service, client);
  boundary['./useAppCatalogUrl.internal'] = { useAppCatalogUrl() {
    const isMountedRef = React.useRef(true);
    React.useEffect(() => () => { isMountedRef.current = false; }, []);
    return { spHttpClient: client, discoverAppCatalogUrl, checkWritePermission: permission, isMountedRef };
  } };
  const hook = loadHook('useSPFxTenantKeyValueStore', 'useSPFxTenantKeyValueStore', boundary);
  const view = mount(() => hook(), {}); t.after(view.unmount);
  return view;
}

test('TenantKV operations refresh a previously denied permission indicator without provisioning', async t => {
  let allowed = false, checks = 0, setups = 0;
  const view = fixture(t, {
    ensureListReady: async () => { setups++; }, get: async () => undefined,
    list: async () => [], save: async () => {}, remove: async () => {}
  }, async () => { checks++; return allowed; });
  await flush(); assert.equal(view.current.canWrite, false); assert.equal(view.current.isReady, true);
  allowed = true;
  for (const operation of [() => view.current.get('key'), () => view.current.list(), () => view.current.save('key', 'value'), () => view.current.remove('key')]) {
    const before = checks;
    await settle(operation); await flush();
    assert.equal(view.current.canWrite, allowed); assert.equal(checks, before + 1);
    allowed = !allowed;
  }
  assert.equal(setups, 0);
});

test('TenantKV latest permission result survives a reverse completion', async t => {
  const oldPermission = deferred(), newPermission = deferred(); let checks = 0;
  const view = fixture(t, { get: async () => undefined, ensureListReady: async () => {} },
    () => (++checks === 1 ? oldPermission.promise : newPermission.promise));
  await flush(); await settle(() => view.current.get('key'));
  await settle(async () => { newPermission.resolve(true); await newPermission.promise; });
  assert.equal(view.current.canWrite, true);
  await settle(async () => { oldPermission.resolve(false); await oldPermission.promise; });
  assert.equal(view.current.canWrite, true);
});

test('TenantKV replacement ignores old read and permission failures while exposing current read errors', async t => {
  const oldRead = deferred(), oldPermission = deferred(); let index = 0;
  const services = [{ get: () => oldRead.promise }, { get: async () => { throw new Error('Current read denied'); } }];
  const clients = [{}, {}], discoveries = [async () => 'catalog-a', async () => 'catalog-b'];
  const checks = [() => oldPermission.promise, async () => true];
  const hook = loadHook('useSPFxTenantKeyValueStore', 'useSPFxTenantKeyValueStore', {
    './useAppCatalogUrl.internal': { useAppCatalogUrl() {
      const isMountedRef = React.useRef(true);
      React.useEffect(() => () => { isMountedRef.current = false; }, []);
      return { spHttpClient: clients[index], discoverAppCatalogUrl: discoveries[index], checkWritePermission: checks[index], isMountedRef };
    } },
    '../services/spfx-tenant-key-value-store.service': { createSPFxTenantKeyValueStoreService: () => services[index] }
  });
  const view = mount(() => hook(), {}); t.after(view.unmount); await flush();
  const pending = start(() => view.current.get('key')); await flush();
  index = 1; view.render({}); await flush();
  await settle(() => view.current.get('key'));
  assert.match(view.current.error.message, /Current read denied/); assert.equal(view.current.canWrite, true);
  await settle(async () => { oldPermission.reject(new Error('Old permission denied')); oldRead.reject(new Error('Old read denied')); await pending; });
  assert.match(view.current.error.message, /Current read denied/); assert.equal(view.current.canWrite, true);
  assert.equal(view.current.isLoading, false);
});

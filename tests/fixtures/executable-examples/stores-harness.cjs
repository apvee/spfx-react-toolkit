// Existing actual provider/Shell/source-lib cache; only native SP responses are controlled.
const { createClientsFixture } = require('./clients-scenarios.cjs');
const assert = require('node:assert/strict');
const { Simulate } = require('react-dom/test-utils');
const owners = {
  tenantHook: ['packages/spfx-react-toolkit/src/hooks/useSPFxTenantKeyValueStore.ts#useSPFxTenantKeyValueStore', 'Explore tenant records and exact text persistence'],
  tenantService: ['packages/spfx-react-toolkit/src/services/spfx-tenant-key-value-store.service.ts#createSPFxTenantKeyValueStoreService', 'Round trip numeric looking text through a tenant service'],
  propertyHook: ['packages/spfx-react-toolkit/src/hooks/useSPFxTenantProperty.ts#useSPFxTenantProperty', 'Read tenant version and counter with explicit activation'],
  propertyService: ['packages/spfx-react-toolkit/src/services/spfx-tenant-property.service.ts#createSPFxTenantPropertyService', 'Read and validate a catalog StorageEntity'],
  hook: ['packages/spfx-react-toolkit/src/hooks/useSPFxSiteKeyValueStore.ts#useSPFxSiteKeyValueStore', 'Explore collection root records with guarded mutations'],
  service: ['packages/spfx-react-toolkit/src/services/spfx-site-key-value-store.service.ts#createSPFxSiteKeyValueStoreService', 'Read a collection root record before a consented save'],
};
function response(body = {}, status = 200, text = '') { return { ok: status >= 200 && status < 300, status, statusText: status === 412 ? 'Precondition Failed' : status === 403 ? 'Forbidden' : String(status), json: async () => body, text: async () => text, headers: { get: () => undefined } }; }
const schema = [{ InternalName: 'Title', FieldTypeKind: 2, Indexed: true, EnforceUniqueValues: true }, { InternalName: 'Value', FieldTypeKind: 3 }, { InternalName: 'Description', FieldTypeKind: 3 }];
function createStoreFixture(t, kind = 'hook', options = {}) {
  const [ownerKey, title] = owners[kind];
  const property = kind === 'propertyHook' || kind === 'propertyService';
  const tenant = kind === 'tenantHook' || kind === 'tenantService';
  const f = createClientsFixture(t, title, { ...options, siteStore: !property && !tenant, tenantStore: tenant, tenantProperties: property, ownerKey, compiled: options.compiled ?? process.env[property ? 'TENANT_PROPERTY_IMPLEMENTATION' : tenant ? 'TENANT_STORE_IMPLEMENTATION' : 'SITE_STORE_IMPLEMENTATION'] === 'compiled' });
  // MessageBar/Dialog native layout observation, not a layout claim.
  const nodeGlobal = Object.getOwnPropertyDescriptor(global, 'Node');
  Object.defineProperty(global, 'Node', { configurable: true, value: window.Node });
  t.after(() => { if (nodeGlobal) Object.defineProperty(global, 'Node', nodeGlobal); else delete global.Node; });
  const previous = Object.getOwnPropertyDescriptor(window, 'ResizeObserver');
  Object.defineProperty(window, 'ResizeObserver', { configurable: true, value: class { observe() {} unobserve() {} disconnect() {} } });
  t.after(() => { if (previous) Object.defineProperty(window, 'ResizeObserver', previous); else delete window.ResizeObserver; });
  function find(name) { const node = [...document.querySelectorAll('button')].find(node => node.textContent === name); assert.ok(node, name); return node; }
  function handler(name) {
    const node = find(name);
    function authored(value) {
      if (!value || typeof value !== 'object') return undefined;
      if (value.props?.children === name && value.props.onClick) return value.props.onClick;
      for (const child of Array.isArray(value) ? value : [value.props?.children]) { const found = authored(child); if (found) return found; }
      return undefined;
    }
    const direct = Object.values(f.regions ?? {}).map(authored).find(Boolean);
    if (direct) return () => direct({ currentTarget: node, target: node });
    const key = Object.keys(node).find(key => key.startsWith('__reactFiber$')); assert.ok(key);
    let fiber = node[key], callback;
    // Capture the authored Button prop, not Fluent's stable DOM event trampoline.
    while (fiber) { if (fiber.memoizedProps?.onClick) callback = fiber.memoizedProps.onClick; fiber = fiber.return; }
    assert.ok(callback); return () => callback({ currentTarget: node, target: node, type: 'click', preventDefault() {}, stopPropagation() {} });
  }
  function check(label, checked = true) { const node = [...document.querySelectorAll('input[type="checkbox"]')].find(node => node.getAttribute('aria-label') === label || node.closest('label')?.textContent.includes(label) || document.querySelector(`label[for="${node.id}"]`)?.textContent.includes(label)); assert.ok(node, label); f.h.act(() => { node.checked = checked; Simulate.change(node, { target: node }); }); }
  async function settle(index, body, status) { const request = f.calls[index]; assert.ok(request, 'native request ' + index); request.settled = true; await f.flush(() => request.resolve(response(body, status))); }
  // The protocol queue contains no subject state or conversion logic.
  async function drain(resolve) {
    for (let i = 0; i < 50; i++) {
      const request = f.calls.find(request => !request.settled); if (!request) { await f.flush(); if (!f.calls.some(r => !r.settled)) return; continue; }
      const result = resolve(request); if (result === undefined) return;
      request.settled = true; await f.flush(() => request.resolve(result));
    }
    throw new Error('Native SP queue did not drain');
  }
  return { ...f, find, handler, check, settle, drain,
    get instance() { return f.instance; }, get regions() { return f.regions; }, get entry() { return f.entry; },
    click(name) { const node = find(name); f.h.act(() => Simulate.click(node)); },
    dt(label) { const node = [...f.container.querySelectorAll('dt')].find(node => node.textContent === label); assert.ok(node, label); return node.nextElementSibling.textContent; },
  };
}
module.exports = { createStoreFixture, response, schema };

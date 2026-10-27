const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const web = 'https://tenant.test/sites/root/sub';
const collection = 'https://tenant.test/sites/root';
const tenant = '11111111-1111-4111-8111-111111111111';
const id = '22222222-2222-4222-8222-222222222222';
const option = { key: 'stable-user', text: 'Alice', kind: 'user', loginName: 'alice@tenant.test', value: { kind: 'user', loginName: 'alice@tenant.test' } };
function setup(missing = false) {
  let writes = 0;
  const h = createHarness({ '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } } });
  const context = { owner: {}, tenantId: tenant, userId: 'current', webUrl: web, siteCollectionUrl: collection,
    graphClient: { api() { const query = { version() { return query; }, get: async () => ({ id, displayName: 'Alice', userPrincipalName: 'alice@tenant.test', mail: 'alice@tenant.test' }) }; return query; } },
    spClient: { get: async url => ({ ok: !missing || url.includes('/_api/site'), status: missing && !url.includes('/_api/site') ? 404 : 200, json: async () => url.includes('/_api/site') ? { Url: collection } : { Id: 42, LoginName: 'i:0#.f|membership|alice@tenant.test', Title: 'Alice', Email: 'alice@tenant.test', PrincipalType: 1 } }), post: async (url, config, options) => { writes++; assert.match(url, /ensureuser$/); assert.deepEqual(JSON.parse(options.body), { logonName: 'alice@tenant.test' }); return { ok: true, status: 200, json: async () => ({ Id: 42, LoginName: 'i:0#.f|membership|alice@tenant.test', Title: 'Alice', Email: 'alice@tenant.test', PrincipalType: 1 }) }; } } };
  const cache = h.load('controls/pickers/picker-request-cache.internal.ts').createPickerRequestCache(context.owner);
  return { h, request: { query: '', pageSize: 50, generation: 1, context, cache, isCurrent: () => true }, get writes() { return writes; } };
}
test('Entra-only resolution preserves renderer alias and performs no SharePoint provisioning', async () => {
  const state = setup();
  const confirm = state.h.load('controls/pickers/principal-identity.internal.ts').confirmPrincipalSelection;
  const result = await confirm([option], { userIdentity: { mode: 'entra' } }, state.request, 'selection');
  assert.equal(result[0].key, option.key);
  assert.equal(result[0].value.entra.id, id);
  assert.equal(result[0].value.sharePoint, undefined);
  assert.equal(state.writes, 0);
});
test('both IDs use a verified collection root and ensureUser=false reads existing registration only', async () => {
  const state = setup();
  const confirm = state.h.load('controls/pickers/principal-identity.internal.ts').confirmPrincipalSelection;
  const result = await confirm([option], { userIdentity: { mode: 'both', ensureUser: false, webUrl: web } }, state.request, 'explicitResolution');
  assert.equal(result[0].value.entra.id, id);
  assert.deepEqual(result[0].value.sharePoint, { siteCollectionUrl: collection, id: 42 });
  assert.equal(state.writes, 0);
});
test('missing registration fails without ensureUser and explicit true can register exactly once', async () => {
  const state = setup(true);
  const confirm = state.h.load('controls/pickers/principal-identity.internal.ts').confirmPrincipalSelection;
  await assert.rejects(confirm([option], { userIdentity: { mode: 'sharePoint', ensureUser: false } }, state.request, 'selection'), /404/);
  assert.equal(state.writes, 0);
  const result = await confirm([option], { userIdentity: { mode: 'sharePoint', ensureUser: true } }, state.request, 'selection');
  assert.equal(result[0].value.sharePoint.id, 42);
  assert.equal(state.writes, 1);
});
test('cancelled resolution starts no writes and does not discard already available identities', async () => {
  const state = setup(true);
  const confirm = state.h.load('controls/pickers/principal-identity.internal.ts').confirmPrincipalSelection;
  await assert.rejects(confirm([option], { userIdentity: { mode: 'both', ensureUser: true } }, { ...state.request, isCurrent: () => false }, 'selection'), /cancel/i);
  assert.equal(state.writes, 0);
  const available = { ...option, entraId: id, entraTenantId: tenant, value: { ...option.value, entra: { tenantId: tenant, id } } };
  const result = await confirm([available], {}, state.request, 'selection');
  assert.deepEqual(result[0].value.entra, available.value.entra);
});

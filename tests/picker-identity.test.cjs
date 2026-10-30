const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const load = () => createHarness().load('controls/pickers/picker-identity.internal.ts');
const web = 'https://tenant.test/sites/hr';
const id = '11111111-1111-4111-8111-111111111111';
test('entity identity includes scope and ignores labels and renamed custom keys', () => {
  const { canonicalPickerIdentity: identity } = load();
  const item = { kind: 'item', webUrl: web, listId: id, itemId: 9 };
  assert.equal(identity(item), identity({ ...item, text: 'Renamed', customKey: { field: 'Code', value: 'new' } }));
  assert.notEqual(identity(item), identity({ ...item, webUrl: `${web}/sub` }));
  const group = { kind: 'sharePointGroup', identity: { siteCollectionUrl: web, id: 9 } };
  assert.notEqual(identity(group), identity({ ...group, identity: { siteCollectionUrl: `${web}/sub`, id: 9 } }));
  assert.notEqual(identity({ kind: 'site', url: web, siteId: id }), identity({ kind: 'site', url: `${web}/sub`, siteId: id }));
});
test('generated aliases hydrate validated scoped values but arbitrary opaque aliases cannot invent values', () => {
  const { createPickerAlias, decodePickerAlias } = load();
  const value = { kind: 'field', webUrl: web, listId: id, fieldId: id, internalName: 'Title' };
  assert.deepEqual(decodePickerAlias(createPickerAlias(value)), value);
  assert.equal(decodePickerAlias('consumer:opaque'), undefined);
  assert.equal(decodePickerAlias('picker:v1:%garbled'), undefined);
  assert.equal(decodePickerAlias('picker:v1:' + encodeURIComponent(JSON.stringify(['site', 'javascript:x']))), undefined);
});
test('retained aliases reject rebinding and survive enrichment without changing selection keys', () => {
  const { reconcilePickerAliases } = load();
  const context = { tenantId: id };
  const user = { key: 'opaque-user', text: 'User', kind: 'user', loginName: 'u@tenant.test', value: { kind: 'user', loginName: 'u@tenant.test' } };
  const first = reconcilePickerAliases(undefined, [user], context);
  const enriched = { ...user, entraId: id, entraTenantId: id, value: { ...user.value, entra: { tenantId: id, id } } };
  const next = reconcilePickerAliases(first, [enriched], context);
  assert.equal(next.options.get(user.key).value.entra.id, id);
  const rebound = reconcilePickerAliases(next, [{ ...user, loginName: 'other@tenant.test', value: { kind: 'user', loginName: 'other@tenant.test' } }], context);
  assert.equal(rebound.errors.length, 1);
  assert.equal(rebound.options.get(user.key).value.entra.id, id);
  assert.equal(reconcilePickerAliases(next, [], context).options.get(user.key).key, user.key);
});
test('generated aliases ignore display snapshots/custom-key renames and membership claims enrich the same login identity', () => {
  const { createPickerAlias, reconcilePickerAliases } = load();
  const value = { kind: 'item', webUrl: web, listId: id, itemId: 9 };
  assert.equal(createPickerAlias(value), createPickerAlias({ ...value, text: 'Renamed', customKey: { field: 'Code', value: 'New' } }));
  const user = { key: 'u', text: 'User', kind: 'user', loginName: 'user@tenant.test', value: { kind: 'user', loginName: 'user@tenant.test' } };
  const first = reconcilePickerAliases(undefined, [user], { tenantId: id });
  const next = { ...user, loginName: 'i:0#.f|membership|user@tenant.test', value: { ...user.value, loginName: 'i:0#.f|membership|user@tenant.test' } };
  assert.equal(reconcilePickerAliases(first, [next], { tenantId: id }).errors.length, 0);
});
test('verified Entra user can acquire a different collection registration without rebinding identity', () => {
  const h = createHarness();
  const { reconcilePickerAliases } = h.load('controls/pickers/picker-identity.internal.ts');
  const entra = { tenantId: '11111111-1111-4111-8111-111111111111', id: 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa' };
  const model = (url, id) => ({ kind: 'user', loginName: 'alice@example.com', key: 'alice', text: 'Alice', value: { kind: 'user', loginName: 'alice@example.com', entra, sharePoint: { siteCollectionUrl: url, id } } });
  const old = reconcilePickerAliases(undefined, [model('https://tenant.test/sites/a', 3)]);
  const next = reconcilePickerAliases(old, [model('https://tenant.test/sites/b', 9)]);
  assert.equal(next.errors.length, 0);
  assert.equal(next.options.get('alice').value.sharePoint.id, 9);
  const conflict = reconcilePickerAliases(next, [model('https://tenant.test/sites/b', 10)]);
  assert.equal(conflict.errors.length, 1);
});

test('content type identity distinguishes scopes and related IDs while preserving v1 aliases', () => {
  const { canonicalPickerIdentity: identity, createPickerAlias, decodePickerAlias } = load();
  const site = { kind: 'contentType', scope: 'site', webUrl: web, contentTypeId: '0x0101AB' };
  assert.equal(typeof identity(site), 'string');
  assert.equal(identity(site), identity({ ...site, webUrl: web.replace('/hr', '/HR') + '/', contentTypeId: '0X0101ab', text: 'Renamed' }));
  for (const other of [{ ...site, contentTypeId: '0x0101AC' }, { ...site, contentTypeId: site.contentTypeId + '01' }, { ...site, scope: 'list', listId: id }]) assert.notEqual(identity(site), identity(other));
  const list = { ...site, scope: 'list', listId: id };
  assert.notEqual(identity(list), identity({ ...list, listId: '22222222-2222-4222-8222-222222222222' }));
  assert.deepEqual(decodePickerAlias(createPickerAlias(list)), list);
});

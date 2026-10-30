const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const h = createHarness();
const load = () => h.load('controls/pickers/picker-values-validation.internal.ts');
const web = 'https://tenant.test/sites/hr';
const guid = '11111111-1111-4111-8111-111111111111';
const cases = [
  ['Site', { kind: 'site', url: web }],
  ['List', { kind: 'list', webUrl: web, listId: guid }],
  ['Field', { kind: 'field', webUrl: web, fieldId: guid, internalName: 'Title' }],
  ['View', { kind: 'view', webUrl: web, listId: guid, viewId: guid }],
  ['ListItem', { kind: 'item', webUrl: web, listId: guid, itemId: 42, customKey: { field: 'Code', value: 1 } }],
  ['Principal', { kind: 'user', loginName: 'user@tenant.test', entra: { tenantId: guid, id: guid }, sharePoint: { siteCollectionUrl: web, id: 7 } }],
];
for (const [domain, value] of cases) {
  test(`${domain} validator copies scoped JSON values and rejects malformed input`, () => {
    const validate = load()[`validate${domain}PickerValues`];
    const result = validate([value]);
    assert.deepEqual(result, [value]);
    assert.notEqual(result[0], value);
    assert.throws(() => validate([{ ...value, callback: () => {} }]), /Invalid/);
    assert.throws(() => validate([{ ...value, text: null }]), /Invalid/);
    assert.throws(() => validate(new Date()), /Invalid/);
    assert.throws(() => validate([Object.create(value)]), /Invalid/);
    assert.throws(() => validate([{ ...value, kind: 'unsupported' }]), /Invalid/);
  });
}
test('validator rejects accessors without invoking them, sparse arrays and extra array members', () => {
  let reads = 0;
  const value = { kind: 'site' };
  Object.defineProperty(value, 'url', { enumerable: true, get() { reads++; return web; } });
  const validate = load().validateSitePickerValues;
  assert.throws(() => validate([value]), /Invalid/);
  assert.equal(reads, 0);
  assert.throws(() => validate(Array(1)), /Invalid/);
  assert.throws(() => validate(Object.assign([], { extra: 1 })), /Invalid/);
});
test('numeric identities and file/custom-key boundaries reject invalid references', () => {
  const validate = load().validateListItemPickerValues;
  const value = { kind: 'item', webUrl: web, listId: guid, itemId: 42 };
  for (const itemId of [0, -1, 1.5, Infinity, NaN, Number.MAX_SAFE_INTEGER + 1]) assert.throws(() => validate([{ ...value, itemId }]), /Invalid/);
  assert.throws(() => validate([{ ...value, customKey: { field: 'Code' } }]), /Invalid/);
  assert.throws(() => validate([{ ...value, kind: 'file' }]), /Invalid/);
  assert.throws(() => validate([{ ...value, webUrl: 'javascript:alert(1)' }]), /Invalid/);
  for (const webUrl of ['https:tenant.test', ' https://tenant.test/sites/hr ', 'https://tenant.test/a/../b', 'https://tenant.test/a\\b']) assert.throws(() => validate([{ ...value, webUrl }]), /Invalid/);
  for (const serverRelativeUrl of ['javascript:alert(1)', 'https://attacker.test/a', '//attacker.test/a']) assert.throws(() => validate([{ ...value, kind: 'file', uniqueId: guid, serverRelativeUrl }]), /Invalid/);
  assert.deepEqual(validate([{ ...value, customKey: { field: 'Code', value: '1' } }])[0].customKey.value, '1');
  assert.throws(() => load().validatePrincipalPickerValues([{ kind: 'user', loginName: 'a', entra: { id: guid } }]), /Invalid/);
});
test('option metadata cannot contradict the authoritative scoped value', () => {
  const { validatePickerOption } = load();
  const option = { key: 'site', text: 'HR', url: web, value: { kind: 'site', url: web } };
  assert.equal(validatePickerOption(option).key, 'site');
  assert.throws(() => validatePickerOption({ ...option, url: 'https://tenant.test/sites/other' }), /Invalid/);
  const user = { key: 'u', text: 'User', kind: 'user', loginName: 'a@tenant.test', entraId: guid, value: { kind: 'user', loginName: 'a@tenant.test' } };
  assert.throws(() => validatePickerOption(user), /Invalid/);
  assert.throws(() => validatePickerOption({ ...option, key: '' }), /Invalid/);
  const file = { key: 'f', text: 'File', kind: 'file', webUrl: web, listId: guid, itemId: 1, keyValue: 1, fileExtension: 'exe', fields: {}, value: { kind: 'file', webUrl: web, listId: guid, itemId: 1, uniqueId: guid, extension: 'pdf' } };
  assert.throws(() => validatePickerOption(file), /Invalid/);
  const scopedUser = { key: 'u', text: 'User', kind: 'user', loginName: 'u', value: { kind: 'user', loginName: 'u', entra: { tenantId: guid, id: guid } } };
  assert.equal(validatePickerOption(scopedUser).key, 'u');
  const group = { key: 'g', text: 'Group', kind: 'securityGroup', source: 'entra', id: guid, securityEnabled: true, mailEnabled: false, value: { kind: 'securityGroup', identity: { tenantId: guid, id: guid } } };
  assert.equal(validatePickerOption(group).key, 'g');
});
test('all principal/file/folder variants copy nested identities and reject foreign domains and hidden prototype keys', () => {
  const { validatePrincipalPickerValues, validateListItemPickerValues, validateSitePickerValues } = load();
  const principals = [
    { kind: 'sharePointGroup', identity: { siteCollectionUrl: web, id: 7 } },
    ...['microsoft365Group', 'securityGroup', 'distributionList'].map(kind => ({ kind, identity: { tenantId: guid, id: guid } })),
    { kind: 'unvalidatedEmail', email: 'outside@example.com' },
  ];
  const copied = validatePrincipalPickerValues(principals);
  assert.deepEqual(copied, principals);
  assert.notEqual(copied[0].identity, principals[0].identity);
  const items = ['file', 'folder'].map(kind => ({ kind, webUrl: web, listId: guid, itemId: 2, uniqueId: guid, serverRelativeUrl: '/sites/hr/docs/a' }));
  assert.deepEqual(validateListItemPickerValues(items), items);
  assert.throws(() => validateSitePickerValues(principals), /Invalid/);
  assert.throws(() => validateSitePickerValues([JSON.parse('{"kind":"site","url":"https://tenant.test/sites/hr","__proto__":{}}')]), /Invalid/);
  const nested = { siteCollectionUrl: web };
  Object.defineProperty(nested, 'id', { enumerable: true, get() { throw new Error('Must not read'); } });
  assert.throws(() => validatePrincipalPickerValues([{ kind: 'sharePointGroup', identity: nested }]), /Invalid/);
});

test('content type references round-trip hierarchical IDs and reject malformed IDs or inconsistent scopes', () => {
  const { validateContentTypePickerValues: validate, validatePickerOption } = load();
  assert.equal(typeof validate, 'function');
  const site = { kind: 'contentType', scope: 'site', webUrl: web, contentTypeId: '0x010100A33D9AD9805788419BDAAC2CCB37509F', text: 'Document' };
  const list = { ...site, scope: 'list', listId: guid };
  const input = JSON.parse(JSON.stringify([site, list]));
  assert.deepEqual(validate(input), input); assert.notEqual(validate(input)[0], input[0]);
  assert.equal(validate([{ ...site, contentTypeId: '0x' }])[0].contentTypeId, '0x');
  for (const contentTypeId of ['0x1', '0x0100', '0x0100AB', '0101', '0x0101?x', '0x0101/AA', '0x01GG', guid]) assert.throws(() => validate([{ ...site, contentTypeId }]), /Invalid/);
  for (const value of [{ ...site, listId: guid }, { ...list, listId: undefined }, { ...site, scope: 'other' }]) assert.throws(() => validate([value]), /Invalid/);
  const option = { key: 'doc', kind: 'contentType', text: 'Document', webUrl: web, id: site.contentTypeId, name: 'Document', group: 'Documents', hidden: false, readOnly: true, sealed: true, value: site };
  assert.equal(validatePickerOption(option), option);
  assert.throws(() => validatePickerOption({ ...option, id: '0x0101' }), /Invalid/);
});

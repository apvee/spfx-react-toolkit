const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const web = 'https://tenant.test/sites/hr';
const tenant = '11111111-1111-4111-8111-111111111111';
const userId = '22222222-2222-4222-8222-222222222222';
const groupId = '33333333-3333-4333-8333-333333333333';
function setup(graphRespond, spRespond = () => ({})) {
  const graphCalls = [], spCalls = [];
  const h = createHarness({ '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } } });
  const context = { owner: {}, tenantId: tenant, userId: 'current@tenant.test', webUrl: web, siteCollectionUrl: web,
    graphClient: { api(path) { const headers = {}; const api = { header(name, value) { headers[name] = value; return api; }, version() { return api; }, async get() { graphCalls.push({ path, headers }); return graphRespond(path); }, async post(body) { graphCalls.push({ path, headers, body }); return graphRespond(path, body); } }; return api; } },
    spClient: { async get(url) { spCalls.push({ url }); return { ok: true, status: 200, json: async () => spRespond(url) }; }, async post(url, config, options) { const body = JSON.parse(options.body); spCalls.push({ url, body }); return { ok: true, status: 200, json: async () => spRespond(url, body) }; } } };
  const cache = h.load('controls/pickers/picker-request-cache.internal.ts').createPickerRequestCache(context.owner);
  return { h, context, graphCalls, spCalls, request: { query: '', pageSize: 50, generation: 0, context, cache, partition: 'principals:' } };
}
const user = { id: userId, displayName: 'Alice', userPrincipalName: 'alice@tenant.test', mail: 'alice@tenant.test' };
test('native people search sends supported query parameters and preserves substrate mode without provisioning', async () => {
  const state = setup(() => ({}), () => ({ ClientPeoplePickerSearchUser: '[]' }));
  const create = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource;
  const request = { ...state.request, query: 'fab', pageSize: 30 };
  const plain = create({}, state.context);
  await plain.discover(request);
  assert.deepEqual(state.spCalls[0].body, { queryParams: {
    AllowEmailAddresses: false, AllowMultipleEntities: false, AllUrlZones: false, MaximumEntitySuggestions: 30,
    PrincipalSource: 15, PrincipalType: 1, QueryString: 'fab', SharePointGroupID: 0, UseSubstrateSearch: false,
  } });
  const hidden = create({ showHiddenPrincipals: true }, state.context);
  assert.equal(hidden.identity, plain.identity, 'unsupported legacy visibility flag does not change the effective source');
  await hidden.discover(request);
  assert.equal(state.spCalls.length, 1);
  await create({ useSubstrateSearch: true }, state.context).discover(request);
  assert.equal(state.spCalls.length, 2); assert.equal(state.spCalls[1].body.queryParams.UseSubstrateSearch, true);
  await create({ userScope: { kind: 'web' } }, state.context).discover(request);
  assert.equal(state.spCalls[2].body.queryParams.PrincipalSource, 1);
  assert.equal(state.spCalls.some(call => call.url.includes('ensureuser')), false); assert.equal(state.graphCalls.length, 0);
});
test('native people search accepts JSON-light value and verbose named operation response envelopes', async () => {
  const rows = [{ Key: 'i:0#.f|membership|alice@tenant.test', DisplayText: 'Alice', IsResolved: true, EntityType: 'User', EntityData: { Email: 'alice@tenant.test' } }];
  for (const response of [{ value: JSON.stringify(rows) }, { value: rows }, { d: { ClientPeoplePickerSearchUser: JSON.stringify(rows) } }, { ClientPeoplePickerSearchUser: JSON.stringify(rows) }]) {
    const state = setup(() => ({}), () => response);
    const source = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource({}, state.context);
    const page = await source.discover({ ...state.request, query: 'alice' });
    assert.deepEqual(page.errors, []); assert.equal(page.completeness, 'complete'); assert.equal(page.options.length, 1);
    assert.equal(page.options[0].value.loginName, 'i:0#.f|membership|alice@tenant.test');
    assert.equal(state.spCalls.length, 1); assert.equal(state.graphCalls.length, 0);
  }
});
test('native people search distinguishes empty JSON-light results from malformed response payloads', async () => {
  for (const value of [[], '[]']) {
    const state = setup(() => ({}), () => ({ value }));
    const source = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource({}, state.context);
    const page = await source.discover({ ...state.request, query: 'nobody' });
    assert.deepEqual(page.options, []); assert.deepEqual(page.errors, []); assert.equal(page.completeness, 'complete');
  }
  for (const value of [{}, '{}', 'invalid JSON', [null], ['not a person']]) {
    const state = setup(() => ({}), () => ({ value }));
    const source = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource({}, state.context);
    const page = await source.discover({ ...state.request, query: 'alice' });
    assert.deepEqual(page.options, []); assert.equal(page.errors.length, 1); assert.equal(page.errors[0].source, 'sharePoint'); assert.equal(page.completeness, 'partial');
  }
});
test('membership-claim candidates are verified by UPN and remain resolvable with a fresh cache', async () => {
  const login = 'i:0#.f|membership|alice@tenant.test';
  const row = { Key: login, DisplayText: 'Alice', IsResolved: true, EntityType: 'User', EntityData: { Email: 'alice@tenant.test' } };
  const respond = (url, body) => {
    if (body !== undefined) return { value: JSON.stringify(['ali', 'alice@tenant.test'].includes(body.queryParams.QueryString) ? [row] : []) };
    if (url.includes('/_api/site?')) return { Url: web };
    assert.match(url, /siteusers\/getByLoginName/);
    return { Id: 42, LoginName: login, Title: 'Alice', Email: 'alice@tenant.test', PrincipalType: 1 };
  };
  const state = setup(() => { throw new Error('Graph is unnecessary'); }, respond);
  const create = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource;
  const source = create({}, state.context);
  const page = await source.discover({ ...state.request, query: 'ali' });
  const verified = await source.verifyCandidate(page.options[0], state.request);
  assert.ok(verified, 'the resolved membership identity must remain selectable');
  assert.equal(verified.value.loginName, login);
  const confirmed = await source.confirm([verified], state.request, 'selection');
  const reopened = setup(() => { throw new Error('Graph is unnecessary'); }, respond);
  const hydrated = await reopened.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource({}, reopened.context).resolve([confirmed[0].value], reopened.request);
  assert.equal(hydrated.unresolved.length, 0);
  assert.equal(hydrated.resolved[0].option.text, 'Alice');
  assert.equal(hydrated.resolved[0].value.loginName, login);
  assert.equal([...state.spCalls, ...reopened.spCalls].some(call => call.url.toLowerCase().includes('ensureuser')), false);
});
test('UPN candidate verification still rejects a different resolved login', async () => {
  const login = 'i:0#.f|membership|alice@tenant.test';
  const state = setup(() => ({}), () => ({ value: JSON.stringify([{ Key: 'i:0#.f|membership|other@tenant.test', DisplayText: 'Alice', IsResolved: true, EntityType: 'User', EntityData: { Email: 'alice@tenant.test' } }]) }));
  const source = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource({}, state.context);
  assert.equal(await source.verifyCandidate({ key: 'alice', text: 'Alice', kind: 'user', loginName: login, value: { kind: 'user', loginName: login } }, state.request), undefined);
});
test('Entra hydration, resolution and verification share directory reads while preserving aliases', async () => {
  const state = setup(() => user);
  const source = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource({}, state.context);
  const value = { kind: 'user', loginName: 'alice@tenant.test', entra: { tenantId: tenant, id: userId } };
  const [hydrated, resolved] = await Promise.all([
    source.hydrate([{ kind: 'value', key: 'consumer-alias', value }], state.request), source.resolve([value], state.request),
  ]);
  assert.equal(hydrated.resolved[0].option.key, 'consumer-alias');
  const verified = await source.verifyCandidate(hydrated.resolved[0].option, state.request);
  assert.equal(verified.key, 'consumer-alias'); assert.equal(resolved.resolved[0].value.entra.id, userId);
  assert.equal(state.graphCalls.length, 1); assert.equal(state.spCalls.length, 0);
  state.request.cache.invalidate('principals:');
  await source.resolve([value], state.request);
  assert.equal(state.graphCalls.length, 2);
});
test('personal suggestions are opt-in, keep people endpoint and resolve person identities through directory lookup', async () => {
  const state = setup(path => path.startsWith('/me/people') ? { value: [
    { id: 'aggregate-person-id', displayName: 'Alice', personType: { class: 'Person', subclass: 'OrganizationUser' }, userPrincipalName: 'alice@tenant.test' },
    { id: userId, displayName: 'Contact', personType: { class: 'Person', subclass: 'PersonalContact' }, userPrincipalName: 'contact@outside.test' },
  ] } : path.startsWith('/users/') ? user : { value: [] });
  const create = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource;
  const disabled = create({}, state.context);
  assert.equal(disabled.personalSuggestions, undefined);
  const enabled = create({ loadPersonalSuggestions: true }, state.context);
  const data = await enabled.personalSuggestions(state.request);
  assert.equal(data.people.length, 1);
  assert.equal(data.people[0].value.entra.id, userId);
  assert.equal(state.graphCalls.some(call => call.path.startsWith('/me/people')), true);
  assert.equal(state.graphCalls.some(call => call.path.startsWith('/me/memberOf')), false);
  assert.equal(state.spCalls.length, 0);
});
test('memberOf excludes roles and limited objects and distinguishes unified/mail-enabled security/distribution groups', async () => {
  const groups = [
    { '@odata.type': '#microsoft.graph.directoryRole', id: groupId, displayName: 'Role' },
    { '@odata.type': '#microsoft.graph.group', id: groupId, displayName: 'M365', groupTypes: ['Unified'], securityEnabled: true, mailEnabled: true },
    { '@odata.type': '#microsoft.graph.group', id: userId, displayName: 'Security', groupTypes: [], securityEnabled: true, mailEnabled: true },
    { '@odata.type': '#microsoft.graph.group', id: tenant, displayName: 'Distribution', groupTypes: [], securityEnabled: false, mailEnabled: true },
    { '@odata.type': '#microsoft.graph.group', id: '44444444-4444-4444-8444-444444444444', displayName: null },
  ];
  const state = setup(() => ({ value: groups }));
  const source = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource({ principalTypes: ['microsoft365Group', 'securityGroup', 'distributionList'], loadPersonalSuggestions: true }, state.context);
  const data = await source.personalSuggestions(state.request);
  assert.deepEqual(data.groups.map(o => o.kind), ['microsoft365Group', 'securityGroup', 'distributionList']);
  assert.equal(state.graphCalls.some(call => call.path.startsWith('/me/people')), false);
});
test('personal source failure leaves healthy group suggestions available and does not retry equivalent failed requests', async () => {
  let peopleCalls = 0;
  const state = setup(path => { if (path.startsWith('/me/people')) { peopleCalls++; throw Object.assign(new Error('403'), { statusCode: 403 }); } return { value: [{ '@odata.type': '#microsoft.graph.group', id: groupId, displayName: 'M365', groupTypes: ['Unified'], securityEnabled: false, mailEnabled: true }] }; });
  const source = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource({ principalTypes: ['user', 'microsoft365Group'], loadPersonalSuggestions: true }, state.context);
  const data = await source.personalSuggestions(state.request);
  assert.equal(data.errors[0].source, 'people');
  assert.equal(data.groups.length, 1);
  await source.personalSuggestions(state.request);
  assert.equal(peopleCalls, 1);
});
test('tenant user discovery uses structured native people-picker query and never ensureUser', async () => {
  const state = setup(() => ({}), (url, body) => ({ ClientPeoplePickerSearchUser: JSON.stringify([{ Key: 'i:0#.f|membership|alice@tenant.test', DisplayText: 'Alice', IsResolved: true, EntityType: 'User', EntityData: { Email: 'alice@tenant.test' } }]) }));
  const source = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource({}, state.context);
  const page = await source.discover({ ...state.request, query: 'al' });
  assert.equal(page.options[0].value.loginName, 'i:0#.f|membership|alice@tenant.test');
  assert.equal(state.spCalls[0].body.queryParams.QueryString, 'al');
  assert.equal(state.spCalls[0].body.queryParams.PrincipalType, 1);
  assert.equal(state.spCalls.some(c => c.url.includes('ensureuser')), false);
});
test('Entra-backed users cannot retain a SharePoint ID belonging to another principal', async () => {
  const state = setup(() => user, url => url.includes('/_api/site?') ? { Url: web } : { Id: 99, LoginName: 'i:0#.f|membership|bob@tenant.test', Title: 'Bob', Email: 'bob@tenant.test', PrincipalType: 1 });
  const source = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource({}, state.context);
  const value = { kind: 'user', loginName: 'alice@tenant.test', entra: { tenantId: tenant, id: userId }, sharePoint: { siteCollectionUrl: web, id: 99 } };
  const option = { key: 'bad-linked-user', text: 'Alice', kind: 'user', loginName: value.loginName, value };
  await assert.rejects(source.confirm([option], state.request, 'selection'), /match/);
  assert.ok(state.spCalls.some(call => call.url.includes('/siteusers/')));
});
test('site-user registration must expose PrincipalType=User before it can be classified as a user', async () => {
  const state = setup(() => user, url => url.includes('/_api/site?') ? { Url: web } : { Id: 99, LoginName: 'c:0t.c|tenant|group-id', Title: 'Group', Email: '' });
  const source = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource({}, state.context);
  const value = { kind: 'user', loginName: 'c:0t.c|tenant|group-id', sharePoint: { siteCollectionUrl: web, id: 99 } };
  await assert.rejects(source.confirm([{ key: 'group-misclassified', text: 'Group', kind: 'user', loginName: value.loginName, value }], state.request, 'selection'), /not a user/);
});
test('personal paging can find eligible groups beyond five already-selected source candidates', async () => {
  const groups = Array.from({ length: 6 }, (_, index) => ({ '@odata.type': '#microsoft.graph.group', id: `00000000-0000-4000-8000-${String(index + 1).padStart(12, '0')}`, displayName: `Group ${index + 1}`, groupTypes: ['Unified'], securityEnabled: false, mailEnabled: true }));
  const state = setup(path => path.includes('$skiptoken') ? { value: [groups[5]], '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#directoryObjects' } : { value: groups.slice(0, 5), '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#directoryObjects', '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/memberOf?$skiptoken=next' });
  const source = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource({ principalTypes: ['microsoft365Group'], loadPersonalSuggestions: true }, state.context);
  const data = await source.personalSuggestions(state.request);
  const selected = data.groups.slice(0, 5);
  const sections = state.h.load('controls/pickers/picker-suggestions.internal.ts').composePickerSuggestions({ groups: data.groups, selected, context: state.context });
  assert.equal(sections[0].options[0].text, 'Group 6');
  assert.equal(state.graphCalls.length, 2);
});

test('required user IDs refer to the verified destination collection, including a subweb destination', async () => {
  const state = setup(() => ({}), url => { assert.match(url, /\/_api\/site\?/); return { Url: web }; });
  const create = state.h.load('controls/pickers/principal-picker.source.internal.ts').createPrincipalPickerSource;
  const option = { kind: 'user', key: 'alice', text: 'Alice', loginName: 'alice@tenant.test', value: { kind: 'user', loginName: 'alice@tenant.test', entra: { tenantId: tenant, id: userId }, sharePoint: { siteCollectionUrl: web, id: 3 } } };
  const source = create({ userIdentity: { mode: 'both', webUrl: web + '/subweb', ensureUser: false } }, state.context);
  assert.equal(await source.needsIdentityResolution(option, state.request), false);
  assert.equal(await source.needsIdentityResolution({ ...option, value: { ...option.value, sharePoint: { siteCollectionUrl: 'https://tenant.test/sites/old', id: 3 } } }, state.request), true);
  assert.equal(state.spCalls.length, 1);
  assert.equal(state.graphCalls.length, 0);
  assert.equal(state.spCalls.some(call => call.body !== undefined), false);
  const available = create({ userIdentity: { mode: 'available' } }, state.context);
  assert.equal(await available.needsIdentityResolution(option, state.request), false);
  assert.equal(state.spCalls.length, 1);
});

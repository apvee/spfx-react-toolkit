const test = require('node:test');
const assert = require('node:assert/strict');
const { setup, flush } = require('./fixtures/access-trim-harness.cjs');
const id = n => `aaaaaaaa-0000-0000-0000-${String(n).padStart(12, '0')}`;
test('Entra All ID membership chunks deduplicated IDs at twenty without metadata', async t => {
  const f = setup(); t.after(() => f.h.close());
  const C = f.h.load('controls/entra-group-trim/entra-group-trim.tsx').EntraGroupTrim;
  const groups = Array.from({ length: 21 }, (_, n) => ({ id: id(n + 1) })); groups.push({ id: id(1) });
  const m = f.mount(C, { groups }, f.branches(C)); await flush(f.h);
  assert.equal(m.container.innerHTML, '<b>grant</b>'); assert.equal(f.graphCalls.length, 2);
  assert.deepEqual(f.graphCalls.map(c => c.body.groupIds.length), [20, 1]);
});
const group = (n, fields = {}) => ({ id: id(n), displayName: 'Team', groupTypes: [], securityEnabled: true, mailEnabled: false, onPremisesSyncEnabled: null, ...fields });
for (const relationship of ['owner', 'memberOrOwner']) test(`Entra Exchange-created mail-enabled security group ${relationship} fails before membership or owner reads`, async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/entra-group-trim/entra-group-trim.tsx').EntraGroupTrim;
  f.state.graphResponse = async (url, body) => body ? { value: body.groupIds } : url === '/me?$select=id' ? { id: id(1) } : url.includes('/owners') ? { value: [] } : group(2, { mailEnabled: true });
  const m = f.mount(C, { groups: [{ id: id(2) }], relationship }, f.branches(C)); await flush(f.h);
  assert.equal(m.container.textContent, 'error'); assert.equal(f.graphCalls.length, 1);
});
for (const selector of [{ id: id(2) }, { name: 'Team' }]) test(`Entra sparse classification arrays fail before membership: ${JSON.stringify(selector)}`, async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/entra-group-trim/entra-group-trim.tsx').EntraGroupTrim;
  f.state.graphResponse = async (url, body) => body ? { value: body.groupIds } : url.includes('/groups?') ? { value: [group(2, { groupTypes: new Array(1) })] } : group(2, { groupTypes: new Array(1) });
  const m = f.mount(C, { groups: [selector], groupType: 'Security' }, f.branches(C)); await flush(f.h);
  assert.equal(m.container.textContent, 'error'); assert.equal(f.graphCalls.length, 1);
});
test('Entra Security member keeps mail-enabled groups supported without ownership reads', async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/entra-group-trim/entra-group-trim.tsx').EntraGroupTrim;
  f.state.graphResponse = async (url, body) => body ? { value: body.groupIds } : group(2, { mailEnabled: true });
  const m = f.mount(C, { groups: [{ id: id(2) }], groupType: 'Security' }, f.branches(C)); await flush(f.h);
  assert.equal(m.container.textContent, 'grant'); assert.equal(f.graphCalls.length, 2);
});
test('Entra M365 mail-enabled groups retain supported owner checks', async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/entra-group-trim/entra-group-trim.tsx').EntraGroupTrim;
  f.state.graphResponse = async url => url === '/me?$select=id' ? { id: id(1) } : url.includes('/owners') ? { value: [{ id: id(1) }] } : group(2, { mailEnabled: true, groupTypes: ['Unified'] });
  const m = f.mount(C, { groups: [{ id: id(2) }], groupType: 'M365', relationship: 'owner' }, f.branches(C)); await flush(f.h);
  assert.equal(m.container.textContent, 'grant'); assert.equal(f.graphCalls.length, 3);
});
for (const props of [
  { groups: [] }, { groups: undefined }, { groups: [{ id: 'bad' }] }, { groups: [{ name: '' }] },
  { groups: [{ id: id(2), name: 'Team' }] }, { groups: [{ id: id(2), groupType: 'All' }] },
  { groups: [{ id: id(2), groupType: undefined }] }, { groups: [{ id: id(2), extra: true }] },
  { groups: [Object.assign(new Date(), { id: id(2) })] }, { groups: [Object.assign([], { id: id(2) })] },
  { groups: new Array(1) }, { groups: [{ id: id(2) }, ,] },
  { groups: [{ id: id(2) }], groupType: 'all' }, { groups: [{ id: id(2) }], groupType: null },
  { groups: [{ id: id(2) }], relationship: 'owners' }, { groups: [{ id: id(2) }], relationship: null },
  { groups: [{ id: id(2) }], match: 'none' },
]) test(`Entra malformed selectors/config blocks all reads: ${JSON.stringify(props)}`, async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/entra-group-trim/entra-group-trim.tsx').EntraGroupTrim;
  const m = f.mount(C, props, f.branches(C)); await flush(f.h); assert.equal(m.container.textContent, 'error'); assert.equal(f.graphCalls.length, 0);
});
for (const type of ['All', 'M365', 'Security']) test(`Entra paged name filters ${type} before ambiguity and escapes apostrophes`, async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/entra-group-trim/entra-group-trim.tsx').EntraGroupTrim;
  const name = "Team's & Guests";
  f.state.graphResponse = async (url, body) => body ? { value: body.groupIds } : url.includes('second') ? { value: [group(3, { displayName: name, groupTypes: ['Unified'] })] } : { value: [group(2, { displayName: name, groupTypes: ['DynamicMembership'] })], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/groups?second' };
  const m = f.mount(C, { groups: [{ name }], groupType: type }, f.branches(C)); await flush(f.h);
  assert.equal(m.container.textContent, type === 'All' ? 'error' : 'grant');
  assert.match(f.graphCalls[0].url, /Team%27%27s%20%26%20Guests/);
  if (type !== 'All') {
    assert.match(decodeURIComponent(f.graphCalls[0].url), /\$filter=displayName eq 'Team''s & Guests'$/);
    assert.equal(f.graphCalls.at(-1).body.groupIds[0], id(type === 'M365' ? 3 : 2));
  }
});
for (const type of ['M365', 'Security']) test(`Entra ${type} rejects mismatched IDs and zero membership requests`, async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/entra-group-trim/entra-group-trim.tsx').EntraGroupTrim;
  f.state.graphResponse = async () => group(2, { groupTypes: type === 'M365' ? [] : ['Unified'] });
  const m = f.mount(C, { groups: [{ id: id(2) }], groupType: type }, f.branches(C)); await flush(f.h);
  assert.equal(m.container.textContent, 'error'); assert.ok(f.graphCalls.every(call => !call.body));
});
for (const response of [{}, { value: [id(99)] }, { value: [2] }, { value: ['bad'] }, { value: null }]) test(`Entra invalid membership payload fails closed: ${JSON.stringify(response)}`, async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/entra-group-trim/entra-group-trim.tsx').EntraGroupTrim;
  f.state.graphResponse = async () => response;
  const m = f.mount(C, { groups: [{ id: id(2) }] }, f.branches(C)); await flush(f.h); assert.equal(m.container.textContent, 'error');
});
for (const match of ['all', 'any']) test(`Entra membership ${match} checks every selected ID`, async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/entra-group-trim/entra-group-trim.tsx').EntraGroupTrim;
  f.state.graphResponse = async () => ({ value: [id(2).toUpperCase()] });
  const m = f.mount(C, { groups: [{ id: id(2) }, { id: id(3) }], match }, f.branches(C)); await flush(f.h);
  assert.equal(m.container.textContent, match === 'all' ? 'denied' : 'grant'); assert.equal(f.graphCalls[0].body.groupIds.length, 2);
});
for (const relationship of ['owner', 'memberOrOwner']) test(`Entra ${relationship} reads /me once and bare paged owners with limited IDs`, async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/entra-group-trim/entra-group-trim.tsx').EntraGroupTrim;
  f.state.graphResponse = async (url, body) => body ? { value: [] } : url === '/me?$select=id' ? { id: id(1) } : url.includes('second') ? { value: [{ id: id(1), displayName: null }] } : url.includes('/owners') ? { value: [{ id: id(9) }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/owners?second' } : group(Number(url.match(/00000000000(\d)/)[1]));
  const m = f.mount(C, { groups: [{ id: id(2) }, { id: id(3) }], relationship }, f.branches(C)); await flush(f.h);
  assert.equal(m.container.textContent, 'grant'); assert.equal(f.graphCalls.filter(c => c.url.startsWith('/me?')).length, 1);
  assert.ok(f.graphCalls.filter(c => c.url.includes('/owners') && !c.url.includes('second')).every(c => !c.url.includes('$select')));
});
for (const reason of ['synchronized', 'distribution', 'owners-http', 'owners-bad', 'repeat', 'nextlink', 'name-missing', 'name-duplicate', 'group-bad']) test(`Entra ${reason} is error even for any/memberOrOwner`, async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/entra-group-trim/entra-group-trim.tsx').EntraGroupTrim;
  f.state.graphResponse = async (url, body) => {
    if (body) return { value: body.groupIds };
    if (url === '/me?$select=id') return { id: id(1) };
    if (url.includes('/owners')) {
      if (reason === 'owners-http' && url.includes(id(3))) throw new Error('403');
      return { value: reason === 'owners-bad' ? [{}] : [{ id: id(1) }], '@odata.nextLink': reason === 'repeat' ? url : reason === 'nextlink' ? 'https://evil.test/' : undefined };
    }
    if (reason === 'name-missing') return { value: [] };
    if (reason === 'name-duplicate') return { value: [group(2), group(3)] };
    return group(url.includes(id(3)) ? 3 : 2, reason === 'synchronized' ? { onPremisesSyncEnabled: true } : reason === 'distribution' ? { securityEnabled: false } : reason === 'group-bad' ? { groupTypes: undefined } : {});
  };
  const groups = reason.startsWith('name') ? [{ name: 'Team' }] : [{ id: id(2) }, { id: id(3) }];
  const m = f.mount(C, { groups, relationship: 'memberOrOwner', match: 'any' }, f.branches(C)); await flush(f.h); assert.equal(m.container.textContent, 'error');
});
test('Entra Graph initialization error selects error; pending initialization has no requests', async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/entra-group-trim/entra-group-trim.tsx').EntraGroupTrim;
  f.state.graphReady = false; const m = f.mount(C, { groups: [{ id: id(2) }] }, f.branches(C));
  assert.equal(m.container.textContent, 'loading'); await flush(f.h); assert.equal(f.graphCalls.length, 0);
  f.state.graphError = new Error('consent'); m.render(); assert.equal(m.container.textContent, 'error');
});

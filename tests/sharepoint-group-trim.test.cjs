const test = require('node:test');
const assert = require('node:assert/strict');
const { setup, flush } = require('./fixtures/access-trim-harness.cjs');
test('SharePoint groups use selected web current user ID and direct paged membership', async t => {
  const f = setup(); t.after(() => f.h.close());
  const C = f.h.load('controls/sharepoint-group-trim/sharepoint-group-trim.tsx').SharePointGroupTrim;
  f.state.spResponse = async url => ({ ok: true, json: async () => url.includes('currentuser') ? { Id: 27 } : url.includes('/users') ? { value: [{ Id: 17 }], '@odata.nextLink': url.endsWith('next') ? undefined : 'https://tenant.test/sites/remote/next' } : url.endsWith('next') ? { value: [{ Id: 27 }] } : { Id: 2, Title: 'Readers' } });
  const m = f.mount(C, { groups: [{ id: 2 }], webUrl: 'https://tenant.test/sites/remote' }, f.branches(C));
  await flush(f.h); assert.equal(m.container.innerHTML, '<b>grant</b>');
  assert.ok(f.spCalls.every(c => c.url.startsWith('https://tenant.test/sites/remote/')));
  assert.ok(f.spCalls.some(c => c.url.includes('currentuser')));
});
for (const props of [
  { groups: [] }, { groups: undefined }, { groups: [{ id: 0 }] }, { groups: [{ id: 2147483648 }] },
  { groups: [{ id: 1.2 }] }, { groups: [{ id: 1, name: 'Readers' }] }, { groups: [{ name: '' }] },
  { groups: [{ id: 2, groupType: 'All' }] }, { groups: [Object.assign(new Date(), { id: 2 })] },
  { groups: [Object.assign([], { id: 2 })] }, { groups: [{ id: 2 }], match: 'none' },
  { groups: new Array(1) }, { groups: [{ id: 2 }, ,] },
  { groups: [{ id: 2 }], webUrl: null }, { groups: [{ id: 2 }], webUrl: '/relative' },
]) test(`SP malformed selector/config blocks all reads: ${JSON.stringify(props)}`, async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/sharepoint-group-trim/sharepoint-group-trim.tsx').SharePointGroupTrim;
  const m = f.mount(C, props, f.branches(C)); await flush(f.h);
  assert.equal(m.container.textContent, 'error'); assert.equal(f.spCalls.length, 0);
});
test('SP exact escaped name reads all pages before ambiguity; missing name is error', async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/sharepoint-group-trim/sharepoint-group-trim.tsx').SharePointGroupTrim;
  const name = "Reader's & Guests"; let second = [{ Id: 4, Title: name }];
  f.state.spResponse = async url => ({ ok: true, json: async () => url.includes('second') ? { value: second } : { value: [{ Id: 2, Title: name }], '@odata.nextLink': 'https://tenant.test/sites/current/second' } });
  const m = f.mount(C, { groups: [{ name }] }, f.branches(C)); await flush(f.h);
  assert.equal(m.container.textContent, 'error'); assert.equal(f.spCalls.length, 2);
  assert.match(f.spCalls[0].url, /Reader%27%27s%20%26%20Guests/);
  f.state.spResponse = async () => ({ ok: true, json: async () => ({ value: [] }) });
  m.render({ groups: [{ name: 'Missing' }] }); await flush(f.h); assert.equal(m.container.textContent, 'error');
});
for (const match of ['all', 'any']) test(`SP ${match} resolves every reference; direct membership valid nonmember is denial`, async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/sharepoint-group-trim/sharepoint-group-trim.tsx').SharePointGroupTrim;
  let missing = false;
  f.state.spResponse = async url => ({ ok: !(missing && url.includes('sitegroups(3)?')), status: 404, json: async () => url.includes('currentuser') ? { Id: 17 } : url.includes('/users') ? { value: url.includes('(2)') ? [{ Id: 17 }] : [{ Id: 99 }] } : { Id: url.includes('(2)') ? 2 : 3, Title: 'Readers' } });
  const props = { groups: [{ id: 2 }, { id: 3 }], match }; const m = f.mount(C, props, f.branches(C)); await flush(f.h);
  assert.equal(m.container.textContent, match === 'all' ? 'denied' : 'grant');
  missing = true; m.render({ ...props, webUrl: 'https://tenant.test/sites/other' }); await flush(f.h);
  assert.equal(m.container.textContent, 'error');
});
for (const kind of ['http', 'bad-group', 'bad-user', 'bad-users', 'repeat', 'nextlink']) test(`SP ${kind} fails closed`, async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/sharepoint-group-trim/sharepoint-group-trim.tsx').SharePointGroupTrim;
  f.state.spResponse = async url => ({ ok: kind !== 'http', status: 403, json: async () => url.includes('currentuser') ? { Id: kind === 'bad-user' ? '17' : 17 } : url.includes('/users') ? kind === 'bad-users' ? { value: [{}] } : { value: [{ Id: 17 }], '@odata.nextLink': kind === 'repeat' ? url : kind === 'nextlink' ? 'https://evil.test/' : undefined } : { Id: kind === 'bad-group' ? 3 : 2, Title: 'Readers' } });
  const m = f.mount(C, { groups: [{ id: 2 }] }, f.branches(C)); await flush(f.h); assert.equal(m.container.textContent, 'error');
});
test('SP native v1 header and verbose paging payload are supported', async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/sharepoint-group-trim/sharepoint-group-trim.tsx').SharePointGroupTrim;
  f.state.spResponse = async url => ({ ok: true, json: async () => ({ d: url.includes('currentuser') ? { Id: 17 } : url.includes('/users') ? { results: [{ Id: 17 }] } : { Id: 2, Title: 'Readers' } }) });
  const m = f.mount(C, { groups: [{ id: 2 }] }, f.branches(C)); await flush(f.h); assert.equal(m.container.textContent, 'grant');
  assert.ok(f.spCalls.every(call => call.options.headers.Accept === 'application/json;odata.metadata=none'));
});
test('SP an explicit null paging link is malformed rather than silently omitted', async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/sharepoint-group-trim/sharepoint-group-trim.tsx').SharePointGroupTrim;
  f.state.spResponse = async url => ({ ok: true, json: async () => url.includes('currentuser') ? { Id: 17 } : url.includes('/users') ? { value: [{ Id: 17 }], '@odata.nextLink': null } : { Id: 2, Title: 'Readers' } });
  const m = f.mount(C, { groups: [{ id: 2 }] }, f.branches(C)); await flush(f.h); assert.equal(m.container.textContent, 'error');
});

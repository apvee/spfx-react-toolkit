const test = require('node:test');
const assert = require('node:assert/strict');
const { setup, flush, Permission } = require('./fixtures/access-trim-harness.cjs');
for (const permissions of [[{}], [{ value: { High: 0, Low: -1 } }], [{ value: { High: '0', Low: 1 } }], new Array(1)]) {
  test(`malformed native-compatible requested mask cannot inherit full-mask grant: ${JSON.stringify(permissions)}`, async t => {
    const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/sharepoint-permission-trim/sharepoint-permission-trim.tsx').SharePointPermissionTrim;
    f.state.page.web.permissions = { hasPermission: () => true };
    const m = f.mount(C, { permissions }, f.branches(C)); await flush(f.h);
    assert.equal(m.container.textContent, 'error'); assert.equal(f.spCalls.length, 0);
  });
}
test('valid requested mask shape accepts a native-compatible peer instance', async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/sharepoint-permission-trim/sharepoint-permission-trim.tsx').SharePointPermissionTrim;
  f.state.page.web.permissions = { hasPermission: permission => permission.value.Low === 1 };
  const m = f.mount(C, { permissions: [{ value: { High: 0, Low: 1 } }] }, f.branches(C)); await flush(f.h);
  assert.equal(m.container.textContent, 'grant'); assert.equal(f.spCalls.length, 0);
});
test('permissions mount only current branch with no wrapper and support all/any', async t => {
  const f = setup(); t.after(() => f.h.close());
  const C = f.h.load('controls/sharepoint-permission-trim/sharepoint-permission-trim.tsx').SharePointPermissionTrim;
  const m = f.mount(C, { permissions: [new Permission({ Low: 1 }), new Permission({ Low: 2 })] }, f.branches(C));
  assert.equal(m.container.textContent, 'loading'); await flush(f.h); assert.equal(m.container.innerHTML, '<i>denied</i>');
  m.render({ permissions: [new Permission({ Low: 1 }), new Permission({ Low: 2 })], match: 'any' }); await flush(f.h);
  assert.equal(m.container.innerHTML, '<b>grant</b>');
});
for (const props of [
  { permissions: [] }, { permissions: undefined }, { permissions: [1] },
  { permissions: [new Permission({ Low: 1 })], match: 'bad' },
  { permissions: [new Permission({ Low: 1 })], match: null },
  { permissions: [new Permission({ Low: 1 })], scope: 'other' },
  { permissions: [new Permission({ Low: 1 })], scope: 'list' },
  { permissions: [new Permission({ Low: 1 })], scope: 'item', listTitle: 'Tasks', itemId: 0 },
]) test(`permission malformed input selects error with zero reads: ${JSON.stringify(props)}`, async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/sharepoint-permission-trim/sharepoint-permission-trim.tsx').SharePointPermissionTrim;
  const m = f.mount(C, props, f.branches(C)); await flush(f.h);
  assert.equal(m.container.textContent, 'error'); assert.equal(f.spCalls.length, 0);
});
test('permission native current mask skips HTTP; resource replacement revokes before effects', async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/sharepoint-permission-trim/sharepoint-permission-trim.tsx').SharePointPermissionTrim;
  f.state.page.web.permissions = new Permission({ Low: 1 });
  const props = { permissions: [new Permission({ Low: 1 })] }; const m = f.mount(C, props, f.branches(C));
  assert.equal(m.container.textContent, 'grant'); assert.equal(f.spCalls.length, 0);
  m.render({ ...props, scope: 'list', listTitle: 'Tasks' });
  assert.equal(m.observed.at(-1), 'loading'); await flush(f.h); assert.equal(m.container.textContent, 'grant');
  assert.equal(f.spCalls.length, 1); assert.match(f.spCalls[0].url, /getByTitle.*effectivebasepermissions$/);
});
test('permission denial and HTTP failure select distinct branches; omitted branches render empty', async t => {
  const f = setup(); t.after(() => f.h.close()); const C = f.h.load('controls/sharepoint-permission-trim/sharepoint-permission-trim.tsx').SharePointPermissionTrim;
  f.state.spResponse = async () => ({ ok: true, json: async () => ({ High: 0, Low: 0 }) });
  const props = { permissions: [new Permission({ Low: 1 })] }; const m = f.mount(C, props, f.branches(C));
  await flush(f.h); assert.equal(m.container.textContent, 'denied');
  f.state.spResponse = async () => ({ ok: false, status: 403, json: async () => { throw new Error('must not decode'); } });
  m.render({ ...props, webUrl: 'https://tenant.test/sites/other' }); await flush(f.h); assert.equal(m.container.textContent, 'error');
  m.render({ ...props, webUrl: 'https://tenant.test/sites/empty' }, []); await flush(f.h); assert.equal(m.container.innerHTML, '');
});

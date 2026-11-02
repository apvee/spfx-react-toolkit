const assert = require('node:assert/strict');
const { test } = require('node:test');
const { createStoreFixture, response, schema } = require('./fixtures/executable-examples/stores-harness.cjs');
const catalog = 'https://tenant.test/sites/catalog-a';
const key = 'spfx-toolkit-config-demo-disposable';
const row = (Value = '50', Id = 7, Title = key, Description = 'destination metadata') => ({ Id, Title, Value, Description, '@odata.etag': '"3"' });
function protocol(rows = [], override = () => undefined) { return r => override(r) ?? (r.url.endsWith('SP_TenantSettings_Current') ? response({ CorporateCatalogUrl: catalog }) : r.url.includes('/fields?') ? response({ value: schema }) : r.url.includes('/items?') ? response({ value: rows }) : response({ Id: 'list' })); }
const approve = f => f.check('I consent to these exact disposable mutations');
async function manual(f, action) { f.change('Operation', action); if (action === 'save' || action === 'remove') approve(f); f.click('Run'); if (action === 'remove') f.click('Confirm removal'); }
for (const action of ['get', 'list', 'save', 'remove']) test(`emitted configuration ${action} catches physical rejection without absence or success`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); await manual(f, action); await f.drain(() => response({}, 403));
  assert.match(f.container.textContent, /Forbidden/); assert.ok(f.container.querySelector('[data-demo-configuration]').contains(f.container.querySelector('[role="alert"]')), 'actual caught operation error belongs to the configuration result region'); assert.doesNotMatch(f.container.textContent, /not found\.|Saved |Removed |Loaded /);
});
for (const mode of ['site', 'tenant', 'layered']) test(`${mode} selects only enabled scopes without edit/mount I/O`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.change('Mode', mode); f.change('Operation', 'list'); assert.equal(f.calls.length, 0);
  if (mode !== 'layered') { assert.equal(f.container.querySelector('[aria-label="Physical scope"]').disabled, true); f.change('Physical scope', mode === 'site' ? 'tenant' : 'site'); assert.equal(f.calls.length, 0); }
  f.click('Run'); await f.drain(protocol([row()])); assert.match(f.container.textContent, new RegExp(`Loaded 1 ${mode === 'tenant' ? 'tenant' : 'site'} physical`));
  assert.ok(f.calls.every(r => mode === 'tenant' ? !r.url.includes('SiteKeyValueStore') : !r.url.includes('TenantKeyValueStore') && !r.url.endsWith('SP_TenantSettings_Current')));
  const count = f.calls.length; f.change('Key', 'another'); assert.equal(f.calls.length, count); assert.ok(!f.container.textContent.includes('Loaded 1'));
});
test('pending read hides obsolete target completion and later explicit read succeeds', async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.click('Run'); await f.drain(r => r.url.includes('/items?') ? undefined : protocol()(r));
  const old = f.calls.at(-1); f.change('Key', 'another'); const count = f.calls.length;
  old.settled = true; await f.flush(() => old.resolve(response({ value: [row('"obsolete-value"')] }))); assert.equal(f.calls.length, count); assert.ok(!f.container.textContent.includes('obsolete-value'));
  f.click('Run'); await f.drain(protocol()); assert.match(f.container.textContent, /another not found\./);
});
for (const mode of ['site', 'tenant']) test(`${mode} irrelevant URL scalar preserves current read ownership`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.change('Mode', mode); f.click('Run'); await f.drain(r => r.url.includes('/items?') ? undefined : protocol()(r));
  const old = f.calls.at(-1); if (mode === 'site') f.instance.context.pageContext.web.absoluteUrl += '/other'; else f.instance.context.pageContext.site.absoluteUrl += '/other'; f.render();
  old.settled = true; await f.flush(() => old.resolve(response({ value: [row()] })));
  assert.equal(f.dt('Source'), mode); assert.equal(f.dt('Selected value'), '50');
});
for (const field of ['site', 'web']) test(`layered live ${field} scalar rejects captured commands before provider rerender`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); const old = f.handler('Run'); f.instance.context.pageContext[field].absoluteUrl += '/other';
  await f.flush(() => old()); assert.equal(f.calls.length, 0); f.render(); f.click('Run'); await f.drain(protocol([row()]));
  assert.equal(f.dt('Source'), 'site');
});
const formats = [
  ['text', '', '', '""'], ['text', '123', '123', '"123"'], ['text', 'true', 'true', '"true"'], ['text', 'null', 'null', '"null"'], ['text', '{"enabled":true}', '{"enabled":true}', '"{\\"enabled\\":true}"'],
  ['bigint', '9007199254740993', '9007199254740993', '"9007199254740993"'], ['date', '2026-01-02T03:04:05Z', '2026-01-02T03:04:05.000Z', '2026-01-02T03:04:05.000Z'], ['json', '{"enabled":true}', { enabled: true }, '{"enabled":true}'],
];
for (const [format, input, expected, encoded] of formats) test(`actual configuration ${format} preserves ${JSON.stringify(input)} and manual Save never auto-reads`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.change('Operation', 'save'); f.change('Value format', format); f.change('Value', input); approve(f); f.click('Run'); let raw;
  await f.drain(protocol([], r => { if (r.method === 'POST') { raw = JSON.parse(r.init.body); return response({}); } }));
  assert.equal(raw.Value, encoded); assert.match(f.container.textContent, /Saved /); assert.equal(f.calls.filter(r => r.url.includes('/items?')).length, 1);
  f.change('Operation', 'get'); f.click('Run'); await f.drain(protocol([row(raw.Value)])); assert.equal(f.dt('Selected value'), JSON.stringify(expected) + (expected === '' ? ' (empty string)' : '')); assert.equal(f.dt('Value type'), typeof expected);
});
for (const scope of ['site', 'tenant']) test(`${scope} description omission preserves only its destination and explicit empty clears`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.change('Operation', 'save'); f.change('Physical scope', scope);
  for (const replace of [false, true]) { if (replace) { f.check('Replace/clear destination description'); f.change('Description', ''); } approve(f); f.click('Run'); await f.drain(protocol([row()]));
    const post = f.calls.filter(r => r.method === 'POST').at(-1); assert.equal(JSON.parse(post.init.body).Description, replace ? '' : 'destination metadata'); assert.ok(post.url.includes(scope === 'site' ? 'SiteKeyValueStore' : 'TenantKeyValueStore')); }
});
for (const [format, value] of [['json', '{invalid'], ['bigint', '1.2'], ['date', 'not a date']]) test(`invalid ${format} never dispatches configuration mutation or discovery`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.change('Operation', 'save'); f.change('Value format', format); f.change('Value', value); approve(f); f.click('Run'); await f.flush(); assert.equal(f.calls.length, 0); assert.ok(f.container.querySelector('[role="alert"]'));
});
for (const action of ['save', 'remove']) test(`${action} HTTP412 rejects with the physical observed ETag`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); await manual(f, action); await f.drain(protocol([row()], r => r.method === 'POST' ? response({}, 412) : undefined));
  assert.match(f.container.textContent, /Precondition Failed/); assert.equal(f.calls.find(r => r.method === 'POST').init.headers['If-Match'], '"3"'); assert.ok(!f.container.textContent.includes(action === 'save' ? 'Saved ' : 'Removed '));
});
test('manual Remove captures exact token/scope/key, locks fields and refuses old confirmation during a newer Dialog', async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.change('Operation', 'remove'); approve(f); const trigger = f.find('Run'); const nativeFocus = trigger.focus; const focusCalls = []; t.mock.method(trigger, 'focus', function() { focusCalls.push({ disabled: this.disabled, connected: this.isConnected }); return nativeFocus.call(this); }); f.click('Run'); const old = f.handler('Confirm removal');
  assert.equal(f.calls.length, 0); for (const label of ['Mode', 'Physical scope', 'Operation', 'Key']) assert.equal(f.container.querySelector(`[aria-label="${label}"]`).disabled, true);
  f.click('Cancel'); await f.flush(); assert.ok(focusCalls.length > 0 && focusCalls.every(call => !call.disabled && call.connected), 'return focus waits for a connected focusable command'); assert.ok(document.activeElement === trigger, 'current connected Run trigger regains focus after closed commit'); approve(f); f.click('Run'); await f.flush(() => old()); assert.equal(f.calls.length, 0);
  f.click('Confirm removal'); await f.drain(protocol([row()])); assert.match(f.container.textContent, /Removed /); assert.ok(f.calls.every(r => r.url.includes('SiteKeyValueStore')));
});
for (const kind of ['configurationHook', 'configurationService']) test(`${kind} recipe keeps saved step facts when later read fails`, async t => {
  const f = createStoreFixture(t, kind); approve(f); f.click(kind === 'configurationHook' ? 'Run default and override recipe' : 'Run service recipe'); let writes = 0;
  await f.drain(protocol([], r => { if (r.method === 'POST') { writes++; return response({}); } if (writes === 2 && r.url.includes('/items?')) return response({}, 403); }));
  assert.equal(writes, 2); assert.match(f.container.textContent, /Saved tenant default 20/); assert.match(f.container.textContent, /Saved site override 50/); assert.match(f.container.textContent, /Forbidden/); assert.ok(!f.container.textContent.includes('Removed only'));
});
for (const kind of ['configurationHook', 'configurationService']) test(`${kind} portable fallback renders commands, result and bounded diagnostics`, async t => {
  const f = createStoreFixture(t, kind, { fallback: true }); assert.ok(f.container.querySelector('section')); const command = kind === 'configurationHook' ? 'Run' : 'Read effective value'; f.click(command); await f.drain(protocol([row()])); assert.equal(f.dt('Source'), 'site'); assert.ok(f.container.querySelector('[tabindex="0"]'));
});
test('missing item ETag rejects a configuration mutation before any POST', async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.change('Operation', 'save'); approve(f); f.click('Run'); const missing = { ...row(), '@odata.etag': undefined };
  await f.drain(protocol([missing], r => r.url.includes('/items(7)?') ? response(missing) : undefined)); assert.match(f.container.textContent, /ETag/); assert.equal(f.calls.filter(r => r.method === 'POST').length, 0);
});
test('physical inventories retain same IDs per scope and display first ten with complete payload', async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.change('Operation', 'list');
  for (const scope of ['site', 'tenant']) { f.change('Physical scope', scope); f.click('Run'); await f.drain(protocol(Array.from({ length: 12 }, (_, i) => row(String(i), i + 1, `${scope}-${i}`))));
    assert.match(f.container.textContent, /Inventory count: 12/); assert.equal(f.container.querySelector('[aria-label="First ten configuration records"]').querySelectorAll('dt').length, 10);
    assert.ok(f.container.querySelector('[aria-label="Configuration payload"]').textContent.includes(`${scope}-11`)); assert.ok(f.container.querySelector('[aria-label="First ten configuration records"]').textContent.includes(`${scope} / 1`)); }
});
for (const ending of ['reset', 'host']) test(`retained Remove confirmation cannot dispatch after ${ending}`, async t => {
  const f = createStoreFixture(t, 'configurationHook'); f.change('Operation', 'remove'); approve(f); f.click('Run'); const confirm = f.handler('Confirm removal');
  if (ending === 'reset') f.reset(); else f.changeScope(f.host('configuration-confirm-other'));
  await f.flush(() => confirm()); assert.equal(f.calls.length, 0);
});
for (const kind of ['configurationHook', 'configurationService']) test(`${kind} recipe Cancel restores connected command focus after sequence unlock without removing data`, async t => {
  const f = createStoreFixture(t, kind), command = kind === 'configurationHook' ? 'Run default and override recipe' : 'Run service recipe'; approve(f); const trigger = f.find(command); f.click(command); let writes = 0;
  await f.drain(protocol([row()], r => { if (r.method === 'POST') { writes++; return response({}); } })); assert.equal(writes, 2);
  f.click('Cancel'); await f.flush(); assert.ok(document.activeElement === trigger, 'focus follows committed sequence unlock'); assert.equal(writes, 2); assert.match(f.container.textContent, /Site removal cancelled/);
});
for (const kind of ['configurationHook', 'configurationService']) test(`${kind} Escape cancels exact recipe confirmation and returns focus without restoring consent`, async t => {
  const { Simulate } = require('react-dom/test-utils');
  const f = createStoreFixture(t, kind), command = kind === 'configurationHook' ? 'Run default and override recipe' : 'Run service recipe'; approve(f); const trigger = f.find(command); trigger.focus(); f.click(command); let writes = 0;
  await f.drain(protocol([row()], r => { if (r.method === 'POST') { writes++; return response({}); } })); const old = f.handler('Confirm removal'); assert.equal(writes, 2);
  const surface = document.querySelector('[role="dialog"]'); assert.ok(surface); f.h.act(() => Simulate.keyDown(surface, { key: 'Escape', code: 'Escape', keyCode: 27 })); await f.flush();
  assert.equal(document.querySelector('[role="dialog"]'), null); assert.ok(document.activeElement === trigger); assert.equal(trigger.getAttribute('aria-disabled'), 'true'); assert.equal(trigger.disabled, false, 'Fluent disabledFocusable keeps the unavailable command focusable');
  const count = f.calls.length; await f.flush(() => old()); f.click(command); assert.equal(f.calls.length, count); assert.equal(writes, 2); assert.equal(f.container.querySelector('input[type="checkbox"]').checked, false);
});

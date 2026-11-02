const assert = require('node:assert/strict');
const { test } = require('node:test');
const { createStoreFixture, response, schema } = require('./fixtures/executable-examples/stores-harness.cjs');
const key = 'spfx-toolkit-site-demo-disposable';
const row = (value = '"persisted-value"', id = 12, title = key, description = 'saved metadata') => ({ Id: id, Title: title, Value: value, Description: description, '@odata.etag': '"3"' });
function protocol(rows = [], override = () => undefined) { return r => override(r) ?? (r.url.endsWith('/effectiveBasePermissions') ? response({ High: '0', Low: '14' }) : r.url.includes('/fields?') ? response({ value: schema }) : r.url.includes('/items?') ? response({ value: rows }) : response({})); }
async function active(t, options) { const f = createStoreFixture(t, 'hook', options); f.click('Activate site store'); await f.drain(protocol()); return f; }
for (const [action, message] of [['get', 'not found'], ['list', 'Loaded 0']]) test(`${action} current HTTP failure never becomes absent/empty on rerender and retry succeeds`, async t => {
  const f = await active(t); f.change('Operation', action); f.click('Run'); await f.drain(protocol([], r => r.url.includes('/items?') ? response({}, 403) : undefined));
  assert.match(f.container.textContent, /Forbidden/); assert.ok(!f.container.textContent.includes(message)); f.render(); assert.ok(!f.container.textContent.includes(message));
  f.click('Run'); await f.drain(protocol()); assert.ok(f.container.textContent.includes(message)); assert.ok(!f.container.textContent.includes('Forbidden'));
});
for (const action of ['save', 'remove']) test(`${action} rejects HTTP 412 without success and uses the observed ETag`, async t => {
  const f = await active(t); f.change('Operation', action); f.check('I consent to this exact disposable mutation'); f.click('Run');
  if (action === 'remove') f.click('Confirm removal');
  await f.drain(protocol([row()], r => r.method === 'POST' ? response({}, 412) : undefined));
  assert.match(f.container.textContent, /Precondition Failed/); assert.ok(!f.container.textContent.includes(action === 'save' ? 'Saved ' : 'Removed '));
  assert.equal(f.calls.find(r => r.method === 'POST').init.headers['If-Match'], '"3"');
});
test('root/subsite, selected key/value/type and complete current records are readable with full collapsed diagnostics', async t => {
  const f = await active(t); f.instance.context.pageContext.web.absoluteUrl += '/sub'; f.render();
  assert.equal(f.find('Activate site store').disabled, false); f.click('Activate site store'); await f.drain(protocol());
  f.click('Run'); await f.drain(protocol([row()]));
  assert.equal(f.dt('Selected key'), key); assert.equal(f.dt('Selected value'), '"persisted-value"'); assert.equal(f.dt('Value type'), 'string');
  assert.match(f.dt('Current web URL'), /\/sub$/); assert.ok(!f.dt('Collection root URL').endsWith('/sub'));
  assert.ok(f.calls.every(r => !r.url.includes('/sub/_api/')));
  const configuration = f.container.querySelector('[data-demo-configuration]'); assert.ok(configuration.querySelector('input'));
  const commands = [...configuration.querySelectorAll('[role="group"]')].find(n => n.getAttribute('aria-label')?.endsWith(' commands'));
  assert.equal(commands.querySelectorAll('input').length, 0); assert.equal(commands.textContent, 'Run');
  assert.ok(configuration.querySelector('[data-demo-result]')); assert.ok([...configuration.querySelectorAll('button')].some(n => n.textContent === 'Advanced diagnostics' && n.getAttribute('aria-expanded') === 'false'));
});
test('list walks actual later pages in order with full count and first-ten readable rows', async t => {
  const f = await active(t); f.change('Operation', 'list'); const root = f.instance.context.pageContext.site.absoluteUrl;
  const next = `${root}/_api/web/lists/getByTitle('SiteKeyValueStore')/items?$skiptoken=2`;
  f.click('Run'); await f.drain(protocol([], r => r.url.includes('/items?') ? response(r.url === next ? { value: [row('"last"', 12, 'last')] } : { value: Array.from({ length: 11 }, (_, i) => row('"value"', i + 1, 'key-' + i)), '@odata.nextLink': next }) : undefined));
  assert.match(f.container.textContent, /Loaded 12/); assert.match(f.container.textContent, /Inventory count: 12/);
  const rows = f.container.querySelector('[aria-label="First ten site records"]'); assert.equal(rows.querySelectorAll('dt').length, 10); assert.equal(rows.querySelector('dt').textContent, 'key-0');
  f.click('Advanced diagnostics'); f.click('Full site payload');
  assert.ok(f.container.querySelector('[aria-label="Full site payload"]').textContent.includes('"last"'));
});
test('later paging error never reports partial inventory', async t => {
  const f = await active(t); f.change('Operation', 'list'); const next = `${f.instance.context.pageContext.site.absoluteUrl}/_api/web/lists/getByTitle('SiteKeyValueStore')/items?$skiptoken=2`;
  f.click('Run'); await f.drain(protocol([], r => r.url.includes('/items?') ? r.url === next ? response({}, 403) : response({ value: [row()], 'odata.nextLink': next }) : undefined));
  assert.match(f.container.textContent, /Forbidden/); assert.ok(!f.container.textContent.includes('Loaded ')); assert.equal(f.container.querySelector('[aria-label="First ten site records"]'), null);
});
test('target edits and collection change invalidate pending read without another business request', async t => {
  const f = await active(t); f.click('Run'); await f.drain(r => r.url.includes('/items?') ? undefined : protocol()(r));
  const old = f.calls.at(-1); const count = f.calls.length; f.change('Key', 'new-target'); assert.equal(f.calls.length, count);
  old.settled = true; await f.flush(() => old.resolve(response({ value: [row('"obsolete-value"')] })));
  await f.drain(protocol()); assert.ok(!f.container.textContent.includes('obsolete-value'));
  f.click('Run'); f.instance.context.pageContext.site.absoluteUrl += '-other'; f.render(); await f.drain(protocol([row('"obsolete-other"')]));
  assert.ok(!f.container.textContent.includes('obsolete-other')); assert.equal(f.find('Activate site store').disabled, false);
});
test('unmount suppresses pending result and permission refresh without React warnings', async t => {
  const f = await active(t); f.click('Run'); const old = f.calls.at(-1); f.unmount(); const before = f.calls.length;
  const errors = t.mock.method(console, 'error', () => {}); await f.flush(() => old.resolve(response({}, 403)));
  assert.equal(f.calls.length, before); assert.equal(errors.mock.callCount(), 0);
});
const formats = [
  ['text', '', ''], ['text', '123', '123'], ['text', 'true', 'true'], ['text', 'null', 'null'], ['text', '{"enabled":true}', '{"enabled":true}'],
  ['bigint', '9007199254740993', '9007199254740993'], ['date', '2026-01-02T03:04:05Z', '2026-01-02T03:04:05.000Z'], ['json', '{"enabled":true}', { enabled: true }],
];
for (const [format, input, expected] of formats) test(`actual ${format} save/read preserves exact ${JSON.stringify(input)}`, async t => {
  const f = await active(t); f.change('Operation', 'save'); f.change('Value format', format); f.change('Value', input);
  f.check('I consent to this exact disposable mutation'); f.click('Run'); let raw;
  await f.drain(protocol([], r => { if (r.method === 'POST') { raw = JSON.parse(r.init.body); return response({}); } }));
  assert.equal(raw.Title, key); f.change('Operation', 'get'); f.click('Run'); await f.drain(protocol([row(raw.Value)]));
  assert.equal(f.dt('Selected value'), JSON.stringify(expected) + (expected === '' ? ' (empty string)' : '')); assert.equal(f.dt('Value type'), typeof expected);
});
test('description omission preserves and explicit empty replacement clears destination metadata', async t => {
  const f = await active(t); f.change('Operation', 'save');
  for (const replace of [false, true]) {
    if (replace) { f.check('Replace/clear destination description'); f.change('Description', ''); }
    f.check('I consent to this exact disposable mutation'); f.click('Run'); await f.drain(protocol([row()]));
    assert.equal(JSON.parse(f.calls.filter(r => r.method === 'POST').at(-1).init.body).Description, replace ? '' : 'saved metadata');
  }
});
for (const [format, value] of [['json', '{invalid'], ['bigint', '1.2'], ['date', 'not a date']]) test(`invalid ${format} produces local error before any mutation/setup`, async t => {
  const f = await active(t); f.change('Operation', 'save'); f.change('Value format', format); f.change('Value', value); f.check('I consent to this exact disposable mutation'); const count = f.calls.length; f.click('Run'); await f.flush();
  assert.equal(f.calls.length, count); assert.ok(f.container.querySelector('[role="alert"]')); assert.ok(!f.container.textContent.includes('Saved '));
});

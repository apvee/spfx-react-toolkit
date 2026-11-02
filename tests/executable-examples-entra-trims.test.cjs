const test = require('node:test');
const assert = require('node:assert/strict');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const { createTrimsFixture, groupId } = require('./fixtures/executable-examples/trims-scenarios.cjs');
const owner = 'packages/spfx-react-toolkit/src/controls/entra-group-trim/entra-group-trim.tsx#EntraGroupTrim';
const title = 'Stage Entra relationships and observe protected lifetimes';
const me = 'cccccccc-1111-4111-8111-111111111111';
const metadata = (id = groupId, displayName = 'Project Readers') => ({ id, displayName, groupTypes: [], securityEnabled: true, mailEnabled: false, onPremisesSyncEnabled: null }); // Graph legacy nullable field.
test('Entra exact executable availability', () => {
  const selected = scenario(title, owner);
  assert.deepEqual(selected.entry.metadata, { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Run staged Entra group checks, then observe protected content only after the selected membership or ownership relationship allows it; drafts and Reset do not start checks.' });
});
for (const compiled of [false, true]) test(`${compiled ? 'COMPILED' : 'SOURCE'} Entra native checkpoint held factory, full relationship, lifetime, Reset and valid Denied`, async t => {
  const f = createTrimsFixture(t, title, { compiled });
  const original = f.action();
  assert.equal(f.initializations.length, 0); assert.equal(f.calls.length, 0);
  f.change('Relationship', 'memberOrOwner');
  const field = f.container.querySelector('[aria-label="Group values"]'); field.focus();
  f.click(); await f.flush();
  assert.equal(f.initializations.length, 1); assert.equal(f.calls.length, 0);
  assert.ok(f.container.querySelector('[data-trim-state="loading"]')); assert.equal(f.diagnostics.mounts, 0);
  await f.initialize();
  await f.settle(f.calls[0], { value: [metadata()] });
  assert.equal(f.calls[1].phase, 'member'); await f.settle(f.calls[1], { value: [] });
  assert.equal(f.calls[2].phase, 'me'); await f.settle(f.calls[2], { id: me });
  assert.equal(f.calls[3].phase, 'owners'); await f.settle(f.calls[3], { value: [{ id: me }] });
  assert.equal(f.diagnostics.mounts, 1); assert.equal(f.diagnostics.active, true); assert.equal(f.diagnostics.attached, true);
  const callbacks = f.callbacks(); f.click('Protected counter: 0'); assert.equal(f.diagnostics.counter, 1);
  f.change('Group values', 'Draft only'); assert.equal(f.diagnostics.groups[0].name, 'Project Readers'); assert.equal(f.calls.length, 4); assert.equal(f.diagnostics.counter, 1);
  f.setRevision(2); const baseline = f.diagnostics;
  f.h.act(() => { original(); callbacks.increment({}); callbacks.ref(callbacks.node); callbacks.observe('mount', callbacks.lifetime); callbacks.observe('unmount', callbacks.lifetime); });
  await f.flush(); assert.deepEqual(f.diagnostics, baseline); assert.equal(f.calls.length, 4); assert.equal(f.initializations.length, 1);
  assert.ok(f.container.querySelector('[data-trim-demo="idle"]')); assert.equal(f.container.querySelector('[aria-label="Group values"]'), field); assert.equal(document.activeElement, field);
  f.click(); await f.flush(); await f.initialize();
  await f.settle(f.calls[4], { value: [{ id: groupId, displayName: 'Project Readers' }] }); await f.settle(f.calls[5], { value: [] });
  assert.ok(f.container.querySelector('[data-trim-state="denied"]')); assert.equal(f.diagnostics.mounts, 0);
  f.origins();
});
const { entraIds } = require('./fixtures/executable-examples/trims-scenarios.cjs');
const pending = (f, phase) => f.calls.filter(r => !r.settled && r.phase === phase);
const branch = (f, state) => assert.ok(f.container.querySelector(`[data-trim-state="${state}"]`), state);
async function start(f) { f.click(); await f.flush(); await f.initialize(); }
async function references(f, type = 'All') {
  for (const r of f.calls.filter(r => !r.settled && ['name', 'metadata'].includes(r.phase))) {
    const id = r.phase === 'metadata' ? new URL(r.url, 'https://graph.microsoft.com').pathname.split('/')[2] : decodeURIComponent(r.url).includes('Other Readers') ? entraIds[1] : groupId;
    const name = decodeURIComponent(r.url).includes('Other Readers') ? 'Other Readers' : 'Project Readers';
    const row = { ...metadata(id, name), groupTypes: type === 'M365' ? ['Unified'] : [] };
    await f.settle(r, r.phase === 'name' ? { value: [row] } : row);
  }
}
async function finish(f, { type = 'All', members = true, owners = true, exclude = [] } = {}) {
  await references(f, type);
  while (pending(f, 'member').some(r => !exclude.includes(r))) { const r = pending(f, 'member').find(r => !exclude.includes(r)); await f.settle(r, { value: members === true ? r.body.groupIds : members === false ? [] : members }); }
  for (const r of pending(f, 'me')) await f.settle(r, { id: me });
  for (const r of pending(f, 'owners')) await f.settle(r, { value: owners ? [{ id: me }] : [] });
}
for (const compiled of [false, true]) {
  const mode = compiled ? 'COMPILED' : 'SOURCE';
  for (const selector of ['name', 'id']) for (const type of ['All', 'M365', 'Security']) for (const relationship of ['member', 'owner', 'memberOrOwner']) test(`${mode} Entra ${selector}/${type}/${relationship} all and any apply all scalars and complete every group`, async t => {
    const f = createTrimsFixture(t, title, { compiled });
    f.change('Group selector', selector); f.change('Group values', selector === 'id' ? `${groupId};${entraIds[1]}` : 'Project Readers;Other Readers');
    f.change('Group type', type); f.change('Relationship', relationship);
    assert.equal(f.initializations.length, 0); assert.equal(f.calls.length, 0);
    for (const match of ['all', 'any']) {
      f.change('Match', match); await start(f);
      await finish(f, { type, members: [groupId], owners: false });
      if (relationship === 'owner' || match === 'all') branch(f, 'denied'); else assert.equal(f.diagnostics.active, true);
      assert.equal(f.diagnostics.groupType, type); assert.equal(f.diagnostics.relationship, relationship); assert.equal(f.diagnostics.match, match);
      assert.deepEqual(f.diagnostics.groups, selector === 'id' ? [{ id: groupId }, { id: entraIds[1] }] : [{ name: 'Project Readers' }, { name: 'Other Readers' }]);
    }
    if (type === 'All' && selector === 'id' && relationship === 'member') assert.ok(f.calls.every(r => r.phase === 'member'), 'metadata bypass');
    if (relationship === 'owner') assert.equal(f.calls.filter(r => r.phase === 'member').length, 0);
    if (relationship !== 'member') assert.equal(f.calls.filter(r => r.phase === 'owners').length, 4);
  });
  test(`${mode} Entra All ID deduplication and sequential 21-ID chunks enforce chunk-local results`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); f.change('Group selector', 'id');
    f.change('Group values', [...entraIds, entraIds[0].toUpperCase()].join(';')); await start(f);
    assert.equal(f.calls.length, 1); assert.equal(f.calls[0].body.groupIds.length, 20); await f.settle(f.calls[0], { value: f.calls[0].body.groupIds });
    assert.equal(f.calls.length, 2); assert.equal(f.calls[1].body.groupIds.length, 1); await f.settle(f.calls[1], { value: f.calls[1].body.groupIds }); assert.equal(f.diagnostics.active, true);
    await start(f); const first = f.calls[2]; await f.settle(first, { value: [entraIds[20]] }); branch(f, 'error'); assert.equal(f.calls.length, 3, 'valid GUID in another chunk is not admitted');
    for (const invalid of ['bad', me, 17]) { await start(f); await f.settle(f.calls.at(-1), { value: [invalid] }); branch(f, 'error'); }
    await start(f); await f.settle(f.calls.at(-1), { value: [] }); await f.settle(f.calls.at(-1), { value: [] }); branch(f, 'denied');
  });
  for (const type of ['All', 'M365', 'Security']) test(`${mode} Entra escaped paged names parse every row and classify ${type} before ambiguity`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); const name = "O'Brien & Readers";
    f.change('Group values', ` ${name};${name} `); f.change('Group type', type); await start(f);
    const first = f.calls[0]; assert.equal(first.url, `/groups?$select=${type === 'All' ? 'id,displayName' : 'id,displayName,groupTypes,securityEnabled,mailEnabled,onPremisesSyncEnabled'}&$filter=displayName%20eq%20%27O%27%27Brien%20%26%20Readers%27`);
    assert.equal(f.calls.length, 1, 'UI trims entries, native deduplicates');
    await f.settle(first, { value: [metadata(groupId, name)], '@odata.nextLink': '?'+new URL(first.url, 'https://graph.microsoft.com').searchParams.toString()+'&$skiptoken=names2' });
    const later = f.calls[1]; assert.ok(later.url.startsWith('https://graph.microsoft.com/v1.0/groups?')); branch(f, 'loading');
    await f.settle(later, { value: [{ ...metadata(entraIds[0], name), groupTypes: ['Unified'] }] });
    if (type === 'All') branch(f, 'error'); else { const r = f.calls[2]; assert.deepEqual(r.body.groupIds, [type === 'M365' ? entraIds[0] : groupId]); await f.settle(r, { value: r.body.groupIds }); assert.equal(f.diagnostics.active, true); }
    await start(f); await f.settle(f.calls.at(-1), { value: [metadata(groupId, name), { id: 'bad', displayName: 'unrelated' }] }); branch(f, 'error');
  });
  for (const relationship of ['owner', 'memberOrOwner']) test(`${mode} Entra ${relationship} uses actual Graph me and completes later owners pages`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); f.instance.native.page.aadInfo = { userId: { toString: () => entraIds[0] } };
    f.change('Relationship', relationship); await start(f); await references(f);
    if (relationship === 'memberOrOwner') await f.settle(pending(f, 'member')[0], { value: [] });
    await f.settle(pending(f, 'me')[0], { id: me });
    const first = pending(f, 'owners')[0]; assert.equal(first.url, `/groups/${groupId}/owners`);
    await f.settle(first, { value: [{ id: entraIds[0] }], '@odata.nextLink': '?$skiptoken=owners2' }); branch(f, 'loading');
    assert.equal(f.diagnostics.mounts, 0); await f.settle(pending(f, 'owners')[0], { value: [{ id: me, displayName: null }] }); assert.equal(f.diagnostics.active, true); // Limited native owner IDs suffice.
    await start(f); await finish(f, { members: false, owners: false }); branch(f, 'denied');
  });
  for (const reason of ['sync', 'mail', 'distribution', 'undefined-sync', 'wrong-type', 'wrong-id', 'missing', 'ambiguous', 'bad-classification']) test(`${mode} Entra ${reason} rejects before relationship reads after valid positive control`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); f.change('Relationship', 'memberOrOwner'); await start(f); await finish(f); assert.equal(f.diagnostics.active, true);
    if (['wrong-type', 'wrong-id'].includes(reason)) { f.change('Group selector', 'id'); f.change('Group values', groupId); f.change('Group type', 'Security'); }
    await start(f); const r = f.calls.at(-1), row = metadata();
    if (reason === 'sync') row.onPremisesSyncEnabled = true;
    if (reason === 'mail') row.mailEnabled = true;
    if (reason === 'distribution') row.securityEnabled = false;
    if (reason === 'undefined-sync') delete row.onPremisesSyncEnabled;
    if (reason === 'wrong-type') row.groupTypes = ['Unified'];
    if (reason === 'wrong-id') row.id = entraIds[0];
    if (reason === 'bad-classification') row.groupTypes = ['DynamicMembership', 17];
    await f.settle(r, r.phase === 'metadata' ? row : { value: reason === 'missing' ? [] : reason === 'ambiguous' ? [row, metadata(entraIds[0])] : [row] });
    branch(f, 'error'); assert.equal(f.calls.at(-1), r, 'no member/me/owners request');
    f.change('Relationship', 'member'); f.change('Group type', 'All'); await start(f); await finish(f); assert.equal(f.diagnostics.active, true, 'ownership constraints do not prohibit member-only behavior');
  });
  for (const match of ['all', 'any']) for (const failure of ['reference', 'me', 'owner', 'later-owner', 'later-chunk']) test(`${mode} Entra ${match} complete-read ${failure} failure cannot be masked by a positive`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); f.change('Match', match);
    if (failure === 'later-chunk') { f.change('Group selector', 'id'); f.change('Group values', entraIds.join(';')); }
    else { f.change('Relationship', 'memberOrOwner'); f.change('Group values', 'Project Readers;Other Readers'); }
    await start(f);
    if (failure === 'reference') { const refs = pending(f, 'name'); await f.settle(refs[0], { value: [metadata(entraIds[1], 'Other Readers')] }); await f.reject(refs[1]); }
    else if (failure === 'later-chunk') { await f.settle(f.calls[0], { value: f.calls[0].body.groupIds }); branch(f, 'loading'); await f.reject(f.calls[1]); }
    else { await references(f); const member = pending(f, 'member')[0]; await f.settle(member, { value: member.body.groupIds }); branch(f, 'loading');
      if (failure === 'me') await f.reject(pending(f, 'me')[0]);
      else { await f.settle(pending(f, 'me')[0], { id: me }); const owners = pending(f, 'owners'); await f.settle(owners[0], { value: [{ id: me }] });
        if (failure === 'owner') await f.reject(owners[1]);
        else { await f.settle(owners[1], { value: [{ id: me }], '@odata.nextLink': '?$skiptoken=owners2' }); branch(f, 'loading'); await f.settle(pending(f, 'owners')[0], { value: [{ id: 'bad' }] }); }
      }
    }
    branch(f, 'error'); assert.equal(f.diagnostics.mounts, 0);
    await start(f); await finish(f); assert.equal(f.diagnostics.active, true, 'independent valid recovery');
  });
  for (const next of [null, '', 17, 'http://graph.microsoft.com/groups', 'https://evil.test/groups', 'https://graph.microsoft.us/groups', 'https://user@graph.microsoft.com/groups', 'https://graph.microsoft.com:444/groups', 'https://graph.microsoft.com/groups#fragment', 'cycle']) test(`${mode} Entra malformed or outside paging ${String(next)} fails independently after positive`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); await start(f); await finish(f); assert.equal(f.diagnostics.active, true);
    await start(f); const r = f.calls.at(-1); await f.settle(r, { value: [metadata()], '@odata.nextLink': next === 'cycle' ? 'https://graph.microsoft.com/v1.0'+r.url : next }); branch(f, 'error'); assert.equal(f.calls.at(-1), r);
  });
  test(`${mode} Entra allowed continuation outside groups and v1 path is not an invented public restriction`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); await start(f); const r = f.calls[0];
    await f.settle(r, { value: [], '@odata.nextLink': 'https://graph.microsoft.com/beta/directoryObjects'+new URL(r.url, 'https://graph.microsoft.com').search+'&$skiptoken=outsidePath' });
    await f.settle(f.calls[1], { value: [metadata()] }); await f.settle(f.calls[2], { value: [groupId] }); assert.equal(f.diagnostics.active, true);
  });
  test(`${mode} Entra invalid drafts are preserved for public validation; initialization failure remains Error`, async t => {
    const f = createTrimsFixture(t, title, { compiled });
    for (const [selector, value] of [['name', ''], ['name', 'Project Readers;'], ['id', 'bad'], ['id', groupId+';']]) {
      f.change('Group selector', selector); f.change('Group values', value); await start(f); branch(f, 'error'); assert.equal(f.calls.length, 0); assert.equal(f.diagnostics.groups.length, value.split(';').length);
    }
    f.reset(); f.click(); await f.flush(); branch(f, 'loading'); await f.initialize(undefined, new Error('factory unavailable')); branch(f, 'error'); assert.equal(f.calls.length, 0);
    await start(f); await finish(f); assert.equal(f.diagnostics.active, true);
  });
  for (const fallback of [false, true]) test(`${mode} Entra ${fallback ? 'standalone' : 'regions'} retains real caller DOM/classes/ref and immutable drafts; settled repeat and nonmonotonic Reset`, async t => {
    const f = createTrimsFixture(t, title, { compiled, fallback }); await start(f); await finish(f); assert.ok(f.container.querySelector('[data-trim-demo="protected"] strong')); assert.ok(f.callbacks().node.className); assert.equal(f.diagnostics.attached, true);
    f.click('Protected counter: 0'); const before = f.diagnostics, calls = f.calls.length, initializations = f.initializations.length;
    f.change('Group selector', 'id'); f.change('Group values', entraIds[0]); f.change('Group type', 'M365'); f.change('Relationship', 'owner'); f.change('Match', 'any'); f.render();
    assert.deepEqual(f.diagnostics, before); assert.equal(f.calls.length, calls); assert.equal(f.initializations.length, initializations);
    const callbacks = f.callbacks(); f.setRevision(3); assert.equal(f.diagnostics.mounts, 0); assert.equal(f.diagnostics.counter, 0); assert.equal(f.diagnostics.attached, false); assert.equal(f.calls.length, calls);
    f.h.act(() => { callbacks.increment({}); callbacks.ref(callbacks.node); callbacks.observe('mount', callbacks.lifetime); callbacks.observe('unmount', callbacks.lifetime); }); assert.equal(f.diagnostics.mounts, 0);
    assert.equal(f.container.querySelector('[aria-label="Group values"]').value, 'Project Readers'); assert.equal(f.container.querySelector('[aria-label="Group type"]').value, 'All'); assert.equal(f.container.querySelector('[aria-label="Relationship"]').value, 'member'); assert.equal(f.container.querySelector('[aria-label="Match"]').value, 'all');
    await start(f); await finish(f); f.click('Protected counter: 0'); await start(f); branch(f, 'loading'); assert.equal(f.diagnostics.unmounts, 1); await finish(f); assert.equal(f.diagnostics.counter, 0); assert.equal(f.diagnostics.mounts, 2);
  });
}
for (const compiled of [false, true]) {
  const mode = compiled ? 'COMPILED' : 'SOURCE';
  for (const terminal of ['success', 'error']) for (const order of ['before', 'after']) for (const next of ['grant', 'error']) test(`${mode} Entra old ${terminal} ${order} newer ${next} cannot publish`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); await start(f); await references(f); const old = pending(f, 'member')[0], run = f.action();
    f.change('Group values', 'Other Readers'); await start(f); const newer = f.calls.at(-1);
    const settleOld = () => { assert.equal(old.settled, false, 'old native decision is still individually held'); return terminal === 'success' ? f.settle(old, { value: [groupId] }) : f.reject(old); };
    if (order === 'before') await settleOld();
    if (next === 'error') await f.reject(newer); else await finish(f, { exclude: [old] });
    if (order === 'after') await settleOld();
    const current = f.diagnostics; f.h.act(run); await f.flush(); assert.deepEqual(f.diagnostics, current); assert.equal(current.active, next === 'grant'); assert.deepEqual(current.groups, [{ name: 'Other Readers' }]); assert.equal(current.attempt, 2);
    assert.equal(old.capture.client, f.instance.native.graph); assert.equal(old.capture.factory, f.instance.native.factory); assert.equal(old.capture.page, f.instance.native.page); assert.equal(old.capture.login, 'trims-one'); assert.equal(old.capture.method, 'POST'); assert.deepEqual(old.capture.body, { groupIds: [groupId] }); assert.equal(Object.isFrozen(old.capture.body.groupIds), true);
  });
  for (const terminal of ['success', 'error']) test(`${mode} Entra Reset and departure retire original commands and held initialization or native settlement ${terminal}`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); f.click(); await f.flush(); const init = f.initializations[0], unarmedRun = f.action();
    f.reset(); f.h.act(unarmedRun); await f.initialize(init, terminal === 'error' ? new Error('late init') : undefined); assert.equal(f.calls.length, 0); assert.ok(f.container.querySelector('[data-trim-demo="idle"]'));
    await start(f); await references(f); const held = pending(f, 'member')[0], run = f.action(); f.setRevision(1); const baseline = f.diagnostics; f.h.act(run);
    if (terminal === 'success') await f.settle(held, { value: [groupId] }); else await f.reject(held);
    assert.deepEqual(f.diagnostics, baseline); assert.equal(f.calls.length, 2);
    await start(f); await finish(f); const callbacks = f.callbacks(); f.click('Protected counter: 0');
    await start(f); await references(f); const departed = pending(f, 'member')[0], departedRun = f.action(), count = f.calls.length;
    f.unmount(); f.h.act(() => { departedRun(); callbacks.increment({}); callbacks.ref(callbacks.node); callbacks.observe('mount', callbacks.lifetime); callbacks.observe('unmount', callbacks.lifetime); });
    if (terminal === 'success') await f.settle(departed, { value: [groupId] }); else await f.reject(departed);
    assert.equal(f.calls.length, count);
  });
  test(`${mode} Entra provider-local inflight sharing and two actual providers sharing native client do not share decisions`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); const peer = f.mountPeer(f.instance);
    await start(f); peer.click(); await f.flush(); await f.initialize(); assert.equal(f.calls.length, 2);
    f.click(); await f.flush(); await f.initialize(); assert.equal(f.calls.length, 2, 'equivalent pending work shares only current provider');
    assert.equal(f.calls[0].client, f.calls[1].client);
    await f.settle(f.calls[0], { value: [metadata()] }); const firstMember = f.calls.at(-1); await f.settle(firstMember, { value: [groupId] }); assert.equal(f.diagnostics.active, true); assert.equal(peer.container.querySelector('[data-trim-demo="protected"]'), null);
    await f.settle(f.calls[1], { value: [metadata()] }); await f.settle(f.calls.at(-1), { value: [] }); assert.ok(peer.container.querySelector('[data-trim-state="denied"]')); assert.equal(f.diagnostics.mounts, 1);
  });
  for (const scalar of ['login', 'email', 'aad']) test(`${mode} Entra actual ${scalar} refresh retains editor, hides grant and retires original Run/callbacks`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); await start(f); await finish(f); const run = f.action(), callbacks = f.callbacks(); f.click('Protected counter: 0');
    f.change('Group values', 'Other Readers'); const field = f.container.querySelector('[aria-label="Group values"]'); field.focus();
    if (scalar === 'login') f.instance.native.page.user.loginName = 'changed';
    if (scalar === 'email') f.instance.native.page.user.email = 'changed@tenant.test';
    if (scalar === 'aad') f.instance.native.page.aadInfo = { userId: { toString: () => me } };
    f.render(); assert.equal(f.container.querySelector('[data-trim-demo="protected"]'), null); assert.equal(f.container.querySelector('[aria-label="Group values"]'), field); assert.equal(document.activeElement, field); assert.equal(field.value, 'Other Readers');
    await f.flush(); f.h.act(run); assert.equal(f.diagnostics.attempt, 1); await finish(f, { members: false }); branch(f, 'denied');
    const before = f.diagnostics; f.h.act(() => { callbacks.increment({}); callbacks.ref(callbacks.node); callbacks.observe('mount', callbacks.lifetime); callbacks.observe('unmount', callbacks.lifetime); }); assert.deepEqual(f.diagnostics, before); assert.equal(f.diagnostics.unmounts, 1); assert.equal(f.diagnostics.attached, false);
  });
  for (const terminal of ['success', 'error']) test(`${mode} Entra same-scope context retains editor and mounted native caches; explicit Run consumes offered factory/page while old ${terminal} stays retired`, async t => {
    const f = createTrimsFixture(t, title, { compiled, directProvider: true }); await start(f); await finish(f); f.click('Protected counter: 0');
    const callbacks = f.callbacks(), field = f.container.querySelector('[aria-label="Group values"]'); field.focus();
    f.change('Group values', 'Other Readers'); f.change('Match', 'any'); f.change('Group type', 'Security'); f.change('Relationship', 'memberOrOwner'); const run = f.action();
    const original = f.instance, offered = f.host(original.instanceId), scope = original.context.serviceScope;
    // This finite offered map diagnoses per-hook caching, not a native service-map subscription.
    original.native.services.set(f.keys.page, offered.native.page); original.native.services.set(f.keys.graph, offered.native.factory);
    f.replace({ ...original, context: { ...original.context, pageContext: offered.native.page } });
    assert.equal(f.instance.context.serviceScope, scope); assert.equal(f.container.querySelector('[aria-label="Group values"]'), field); assert.equal(document.activeElement, field); assert.equal(field.value, 'Other Readers'); assert.equal(f.container.querySelector('[aria-label="Relationship"]').value, 'memberOrOwner');
    assert.deepEqual(f.diagnostics.groups, [{ name: 'Project Readers' }]); assert.equal(f.diagnostics.match, 'all'); assert.equal(f.diagnostics.attempt, 1); assert.equal(f.container.querySelector('[data-trim-demo="protected"]'), null); f.h.act(run);
    await f.flush(); const refresh = f.calls.at(-1); assert.equal(refresh.client, original.native.graph); assert.equal(refresh.factory, original.native.factory); assert.equal(refresh.page, original.native.page); assert.equal(f.initializations.length, 1); await finish(f); assert.equal(f.diagnostics.mounts, 2); assert.equal(f.diagnostics.counter, 0);
    const now = f.diagnostics; f.h.act(() => { callbacks.increment({}); callbacks.ref(callbacks.node); callbacks.observe('mount', callbacks.lifetime); callbacks.observe('unmount', callbacks.lifetime); }); assert.deepEqual(f.diagnostics, now);
    await start(f); const fresh = f.calls.at(-1); assert.equal(f.initializations.at(-1).factory, offered.native.factory); assert.equal(fresh.client, offered.native.graph); assert.equal(fresh.page, offered.native.page); await references(f, 'Security'); const held = pending(f, 'member')[0], retainedRun = f.action();
    f.replace({ ...f.instance, context: { ...f.instance.context } }); await f.flush();
    assert.equal(f.container.querySelector('[aria-label="Group values"]'), field); assert.equal(document.activeElement, field); assert.equal(f.diagnostics.attempt, 2); assert.deepEqual(f.diagnostics.groups, [{ name: 'Other Readers' }]); assert.equal(f.diagnostics.relationship, 'memberOrOwner'); assert.equal(f.calls.at(-1).client, offered.native.graph);
    // A same-identity pending native call can be shared. Give replacement a distinct real AAD identity.
    offered.native.page.aadInfo = { userId: { toString: () => me } }; f.render(); await f.flush();
    const replacement = f.calls.at(-1); assert.notEqual(replacement, held); await finish(f, { type: 'Security', members: false, owners: false, exclude: [held] }); branch(f, 'denied');
    const denied = f.diagnostics; f.h.act(retainedRun); assert.equal(held.settled, false); if (terminal === 'success') await f.settle(held, { value: [entraIds[1]] }); else await f.reject(held);
    // Obsolete already-started chains may proceed; settle them without claiming transport cancellation.
    for (const r of pending(f, 'me')) await f.settle(r, { id: me }); for (const r of pending(f, 'owners')) await f.settle(r, { value: [{ id: me }] });
    assert.deepEqual(f.diagnostics, denied);
  });
  for (const terminal of ['success', 'error']) test(`${mode} Entra genuine scope factory/client/page replacement with same instanceId retires provider subtree; held ${terminal} cannot restore`, async t => {
    const f = createTrimsFixture(t, title, { compiled, directProvider: true }); await start(f); await finish(f); f.click('Protected counter: 0'); const field = f.container.querySelector('[aria-label="Group values"]'), callbacks = f.callbacks();
    f.change('Group values', 'Other Readers'); await start(f); await references(f); const old = pending(f, 'member')[0], run = f.action(), original = f.instance, next = f.host(original.instanceId);
    assert.notEqual(next.context.serviceScope, original.context.serviceScope); assert.notEqual(next.native.graph, original.native.graph); assert.notEqual(next.native.factory, original.native.factory); assert.notEqual(next.native.page, original.native.page);
    f.replace(next); const replacement = f.container.querySelector('[aria-label="Group values"]'); assert.notEqual(replacement, field); assert.equal(field.isConnected, false); assert.equal(replacement.value, 'Project Readers'); branchIdle(f);
    const baseline = f.diagnostics, count = f.calls.length, inits = f.initializations.length;
    f.h.act(() => { run(); callbacks.increment({}); callbacks.ref(callbacks.node); callbacks.observe('mount', callbacks.lifetime); callbacks.observe('unmount', callbacks.lifetime); }); await f.flush(); assert.deepEqual(f.diagnostics, baseline); assert.equal(f.calls.length, count); assert.equal(f.initializations.length, inits);
    await start(f); assert.equal(f.initializations.at(-1).factory, next.native.factory); assert.equal(f.calls.at(-1).client, next.native.graph); assert.equal(f.calls.at(-1).page, next.native.page); await finish(f, { exclude: [old] }); f.click('Protected counter: 0'); const granted = f.diagnostics;
    assert.equal(old.settled, false); if (terminal === 'success') await f.settle(old, { value: [entraIds[1]] }); else await f.reject(old); f.h.act(run); await f.flush(); assert.deepEqual(f.diagnostics, granted); assert.equal(f.container.querySelector('[aria-label="Group values"]'), replacement);
  });
  test(`${mode} Entra whole Shell host retirement and obsolete factory settlement remain distinct from mounted cache behavior`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); f.click(); await f.flush(); const old = f.initializations[0], run = f.action(); f.replace(f.host('new-provider')); branchIdle(f); const before = f.diagnostics; f.h.act(run); await f.initialize(old); assert.deepEqual(f.diagnostics, before); assert.equal(f.calls.length, 0);
    await start(f); await finish(f); assert.equal(f.diagnostics.mounts, 1); assert.equal(f.calls[0].capture.login, 'new-provider');
  });
}
function branchIdle(f) { assert.ok(f.container.querySelector('[data-trim-demo="idle"]')); assert.equal(f.diagnostics.attempt, undefined); assert.equal(f.diagnostics.mounts, 0); assert.equal(f.diagnostics.counter, 0); }
for (const compiled of [false, true]) test(`${compiled ? 'COMPILED' : 'SOURCE'} Entra real normalizer preserves exact names and own-key/undefined distinction`, async t => {
  const f = createTrimsFixture(t, title, { compiled });
  const path = require('node:path');
  const { normalizeEntraGroups } = f.h.load(path.join(f.root, 'services/spfx-entra-groups.internal' + (compiled ? '.js' : '.ts')));
  assert.deepEqual(normalizeEntraGroups([{ name: ' Project Readers ', id: undefined }]), [{ name: ' Project Readers ' }], 'service preserves exact name independently of UI trimming');
  assert.deepEqual(normalizeEntraGroups([{ id: groupId.toUpperCase(), name: undefined }]), [{ id: groupId }]);
  for (const input of [[], [{}], [{ id: groupId, name: 'Project Readers' }], [{ id: groupId, groupType: undefined }], [{ id: groupId, unknown: undefined }], [Object.assign({ id: groupId }, { [Symbol('unknown')]: undefined })], [Object.assign(new Date(), { id: groupId })], new Array(1)]) assert.throws(() => normalizeEntraGroups(input));
  assert.equal(f.calls.length, 0); assert.equal(f.initializations.length, 0);
  await start(f); await finish(f); assert.equal(f.diagnostics.active, true, 'valid native positive control after validation negatives');
});
for (const compiled of [false, true]) test(`${compiled ? 'COMPILED' : 'SOURCE'} Entra native attempted malformed body is recorded before independent admission failure`, async t => {
  const f = createTrimsFixture(t, title, { compiled, expectedProtocolErrors: 1 }); await start(f); await finish(f); assert.equal(f.diagnostics.active, true);
  const count = f.nativeAttempts.length;
  await assert.rejects(f.instance.native.graph.api('/me/checkMemberGroups').post({ groupIds: undefined }));
  assert.equal(f.nativeAttempts.length, count + 1); assert.equal(f.protocolErrors.length, 1);
  const attempt = f.nativeAttempts.at(-1); assert.equal(attempt.phase, 'member'); assert.equal(attempt.method, 'POST'); assert.equal(attempt.client, f.instance.native.graph); assert.equal(attempt.factory, f.instance.native.factory); assert.equal(attempt.body.groupIds, undefined); assert.equal(Object.isFrozen(attempt), true);
});

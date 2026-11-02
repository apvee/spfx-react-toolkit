const test = require('node:test');
const assert = require('node:assert/strict');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const { createTrimsFixture } = require('./fixtures/executable-examples/trims-scenarios.cjs');
const owner = 'packages/spfx-react-toolkit/src/controls/sharepoint-group-trim/sharepoint-group-trim.tsx#SharePointGroupTrim';
const title = 'Stage direct SharePoint group membership and observe protected lifetimes';
test('SPGroup exact executable availability', () => {
  const selected = scenario(title, owner);
  assert.deepEqual(selected.entry.metadata, { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Run staged direct SharePoint group checks for the selected web, then observe protected content only when the current membership decision allows it; drafts and Reset do not start checks.' });
});
for (const compiled of [false, true]) test(`${compiled ? 'COMPILED' : 'SOURCE'} SPGroup native checkpoint remote-user complete pages, protected lifetimes and valid nonmember`, async t => {
  const f = createTrimsFixture(t, title, { compiled });
  const originalRun = f.action();
  assert.equal(f.calls.length, 0); assert.equal(f.initializations.length, 0);
  f.change('Web URL', 'https://tenant.test/sites/remote');
  const field = f.container.querySelector('[aria-label="Web URL"]'); field.focus();
  f.click(); await f.flush();
  assert.equal(f.calls[0].phase, 'group-name');
  await f.settle(f.calls[0], { value: [{ Id: 2, Title: 'Project Readers' }] });
  assert.equal(f.calls[1].url, 'https://tenant.test/sites/remote/_api/web/currentuser?$select=Id');
  await f.settle(f.calls[1], { Id: 91 });
  assert.equal(f.calls[2].phase, 'group-users');
  assert.ok(f.container.querySelector('[data-trim-state="loading"]'));
  assert.equal(f.diagnostics.mounts, 0);
  await f.settle(f.calls[2], { value: [{ Id: 17 }], '@odata.nextLink': '?$select=Id&$skiptoken=members2' });
  assert.ok(f.container.querySelector('[data-trim-state="loading"]'));
  await f.settle(f.calls[3], { d: { results: [{ Id: 91 }] } });
  assert.equal(f.diagnostics.mounts, 1); assert.equal(f.diagnostics.active, true); assert.equal(f.diagnostics.attached, true);
  const callbacks = f.callbacks(); f.click('Protected counter: 0');
  assert.equal(f.diagnostics.counter, 1);
  const applied = f.diagnostics.groups;
  f.change('Group values', 'Draft only');
  assert.deepEqual(f.diagnostics.groups, applied); assert.equal(f.calls.length, 4); assert.equal(f.diagnostics.counter, 1);
  f.change('Group values', 'Project Readers'); f.click(); await f.flush();
  assert.equal(f.calls.length, 5); assert.equal(f.diagnostics.unmounts, 1);
  await f.settle(f.calls[4], { value: [{ Id: 2, Title: 'Project Readers' }] });
  await f.settle(f.calls[5], { Id: 91 });
  await f.settle(f.calls[6], { value: [{ Id: 17 }] });
  assert.ok(f.container.querySelector('[data-trim-state="denied"]'));
  assert.equal(f.diagnostics.mounts, 1); assert.equal(f.initializations.length, 0);
  f.setRevision(2); const baseline = f.diagnostics;
  f.h.act(() => { originalRun(); callbacks.increment({}); callbacks.ref(callbacks.node); callbacks.observe('mount', callbacks.lifetime); callbacks.observe('unmount', callbacks.lifetime); });
  await f.flush(); assert.deepEqual(f.diagnostics, baseline); assert.equal(f.calls.length, 7);
  assert.equal(f.container.querySelector('[aria-label="Web URL"]'), field); assert.equal(document.activeElement, field);
  f.origins();
});

async function references(f, { ids = [2], names = ['Project Readers'], user = 17 } = {}) {
  const pending = f.calls.filter(request => !request.settled && ['group-name', 'group-id'].includes(request.phase));
  assert.equal(pending.length, ids.length);
  for (const [index, request] of pending.entries()) {
    const group = { Id: ids[index], Title: names[index] ?? 'Other Readers' };
    await f.settle(request, request.phase === 'group-name' ? { value: [group] } : group);
  }
  const current = f.calls.find(request => !request.settled && request.phase === 'group-currentuser');
  assert.ok(current); await f.settle(current, { Id: user });
  return f.calls.filter(request => !request.settled && request.phase === 'group-users');
}
for (const compiled of [false,true]) test(`${compiled?'COMPILED':'SOURCE'} coupled AND AAD-only refresh retires the original Run`,async t=>{
  const {nested,groupId}=require('./fixtures/executable-examples/trims-scenarios.cjs');
  const f=createTrimsFixture(t,nested,{compiled,mask:1});f.click();await f.flush();await f.initialize();
  await f.settle(f.calls[0],{value:[{id:groupId,displayName:'Project Readers'}]});await f.settle(f.calls[1],{value:[groupId]});
  assert.equal(f.diagnostics.active,true);const run=f.action();
  f.instance.native.page.aadInfo={userId:{toString:()=> '22222222-3333-4444-8555-666666666666'}};f.render();
  assert.equal(f.container.querySelector('[data-trim-demo="protected"]'),null);await f.flush();
  f.h.act(run);assert.equal(f.diagnostics.attempt,1,'original AND Run must retire with the actual inner AAD identity');
  await f.settle(f.calls[2],{value:[{id:groupId,displayName:'Project Readers'}]});await f.settle(f.calls[3],{value:[]});
  assert.ok(f.container.querySelector('[data-trim-level="inner"][data-trim-state="denied"]'));assert.equal(f.diagnostics.mounts,1);
});
async function decision(f, member = true) {
  const users = await references(f); assert.equal(users.length, 1);
  await f.settle(users[0], { value: member ? [{ Id: 17 }] : [] });
}
for (const compiled of [false, true]) {
  const mode = compiled ? 'COMPILED' : 'SOURCE';
  for (const selector of ['name', 'id']) for (const match of ['all', 'any']) test(`${mode} SPGroup ${selector} semicolon ${match} complete independent memberships`, async t => {
    const f = createTrimsFixture(t, title, { compiled });
    f.change('Group selector', selector); f.change('Group values', selector === 'id' ? ' 2 ; 3 ' : ' Project Readers ; Other Readers '); f.change('Match', match);
    assert.equal(f.calls.length, 0); f.click(); await f.flush();
    assert.deepEqual(f.diagnostics.groups, selector === 'id' ? [{ id: 2 }, { id: 3 }] : [{ name: 'Project Readers' }, { name: 'Other Readers' }]);
    // The normalizer sorts name selectors; settle by independently observed query.
    for (const request of [...f.calls]) await f.settle(request, request.phase === 'group-id' ? { Id: request.url.includes('(2)') ? 2 : 3, Title: 'Resolved' } : { value: [{ Id: request.url.includes('Other') ? 3 : 2, Title: request.url.includes('Other') ? 'Other Readers' : 'Project Readers' }] });
    await f.settle(f.calls.at(-1), { Id: 17 });
    const users = f.calls.filter(request => request.phase === 'group-users'); assert.equal(users.length, 2);
    await f.settle(users.find(request => request.url.includes('(2)')), { value: [{ Id: 17 }] });
    assert.ok(f.container.querySelector('[data-trim-state="loading"]'), 'one positive cannot settle before every read');
    await f.settle(users.find(request => request.url.includes('(3)')), { value: [] });
    assert.equal(f.diagnostics.active, match === 'any'); assert.equal(f.diagnostics.mounts, match === 'any' ? 1 : 0);
    assert.equal(f.initializations.length, 0); assert.ok(f.calls.every(request => request.kind === 'sp'));
  });
  for (const match of ['all', 'any']) test(`${mode} SPGroup ${match} positive cannot mask second membership transport failure`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); f.change('Group selector', 'id'); f.change('Group values', '2;3'); f.change('Match', match); f.click(); await f.flush();
    const users = await references(f, { ids: [2,3] });
    await f.settle(users[0], { value: [{ Id: 17 }] }); assert.ok(f.container.querySelector('[data-trim-state="loading"]'));
    await f.reject(users[1]); assert.ok(f.container.querySelector('[data-trim-state="error"]')); assert.equal(f.diagnostics.mounts, 0);
  });
  test(`${mode} SPGroup escaped name all-page resolution and duplicate selector/group deduplication`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); f.change('Group values', "O'Brien & Readers;O'Brien & Readers"); f.click(); await f.flush();
    assert.equal(f.calls.length, 1); assert.equal(f.calls[0].url, "https://tenant.test/sites/team/_api/web/sitegroups?$select=Id,Title&$filter=Title%20eq%20%27O%27%27Brien%20%26%20Readers%27");
    await f.settle(f.calls[0], { d: { results: [{ Id: 2, Title: "O'Brien & Readers" }], __next: "?$select=Id,Title&$filter=Title%20eq%20%27O%27%27Brien%20%26%20Readers%27&$skiptoken=names2" } });
    assert.equal(f.calls[1].phase, 'group-name');
    await f.settle(f.calls[1], { value: [{ Id: 2, Title: "O'Brien & Readers" }] });
    await f.settle(f.calls[2], { Id: 17 }); await f.settle(f.calls[3], { value: [{ Id: 17 }] }); assert.equal(f.diagnostics.mounts, 1);
    f.change('Group values', 'Project Readers;Other Readers'); f.click(); await f.flush();
    const users = await references(f, { ids: [2,2], names: ['Other Readers','Project Readers'] });
    assert.equal(users.length, 1, 'different names resolving to same ID have one membership read'); await f.settle(users[0], { value: [] }); assert.ok(f.container.querySelector('[data-trim-state="denied"]'));
  });
  for (const failure of ['missing', 'ambiguous', 'mismatch', 'bad-id', 'bad-title', 'http', 'collection', 'null-next', 'outside-next', 'repeated-next']) test(`${mode} SPGroup ${failure} is Error after independent positive control`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); f.click(); await f.flush(); await decision(f); assert.equal(f.diagnostics.active, true);
    if (failure === 'mismatch') { f.change('Group selector', 'id'); f.change('Group values', '2'); }
    f.click(); await f.flush(); const request = f.calls.at(-1);
    if (failure === 'http') await f.settle(request, {}, 403);
    else if (failure === 'mismatch') await f.settle(request, { Id: 3, Title: 'Other Readers' });
    else if (failure === 'missing') await f.settle(request, { value: [] });
    else if (failure === 'ambiguous') await f.settle(request, { value: [{ Id: 2, Title: 'Project Readers' }, { Id: 3, Title: 'Project Readers' }] });
    else if (failure === 'bad-id') await f.settle(request, { value: [{ Id: 0, Title: 'Project Readers' }] });
    else if (failure === 'bad-title') await f.settle(request, { value: [{ Id: 2, Title: '' }] });
    else if (failure === 'collection') await f.settle(request, { value: {} });
    else await f.settle(request, { value: [{ Id: 2, Title: 'Project Readers' }], '@odata.nextLink': failure === 'null-next' ? null : failure === 'outside-next' ? 'https://tenant.test/sites/other/_api/web/sitegroups' : request.url });
    assert.ok(f.container.querySelector('[data-trim-state="error"]')); assert.equal(f.diagnostics.active, false); assert.equal(f.diagnostics.mounts, 1);
  });
  test(`${mode} SPGroup invalid drafts reach public validation without filtering or native dispatch`, async t => {
    const f = createTrimsFixture(t, title, { compiled });
    for (const [selector, values] of [['name',''], ['name','Project Readers;'], ['id','2;'], ['id','NaN'], ['id','0'], ['id','2147483648']]) {
      f.change('Group selector', selector); f.change('Group values', values); f.click(); await f.flush();
      assert.equal(f.calls.length, 0); assert.equal(f.nativeAttempts.length, 0); assert.ok(f.container.querySelector('[data-trim-state="error"]')); assert.equal(f.diagnostics.mounts, 0);
      assert.equal(f.diagnostics.groups.length, values.split(';').length, 'every entry remains observable');
    }
    f.change('Group values', '2'); f.change('Web URL', 'https://tenant.test/sites/team?invalid=1'); f.click(); await f.flush(); assert.equal(f.nativeAttempts.length, 0); assert.ok(f.container.querySelector('[data-trim-state="error"]'));
    f.change('Web URL',''); f.click(); await f.flush(); await decision(f); assert.equal(f.diagnostics.active,true,'valid control after invalid selectors');
  });
  test(`${mode} SPGroup selected-user and later user-page malformed responses never turn positive into grant`, async t => {
    const f = createTrimsFixture(t, title, { compiled });
    for (const failure of ['user-id','user-shape','late-user-id','late-http']) {
      f.click(); await f.flush(); const name = f.calls.at(-1); await f.settle(name,{ value:[{Id:2,Title:'Project Readers'}] });
      const user = f.calls.at(-1);
      if(failure === 'user-id') await f.settle(user,{Id:0});
      else { await f.settle(user,{Id:17}); const users=f.calls.at(-1);
        if(failure==='user-shape') await f.settle(users,{value:'invalid'});
        else {await f.settle(users,{value:[{Id:17}],'odata.nextLink':'?$select=Id&$skiptoken=members2'});assert.ok(f.container.querySelector('[data-trim-state="loading"]'));if(failure==='late-http')await f.settle(f.calls.at(-1),{},500);else await f.settle(f.calls.at(-1),{value:[{Id:'17'}]});}
      }
      assert.ok(f.container.querySelector('[data-trim-state="error"]')); assert.equal(f.diagnostics.mounts,0);
    }
    f.click();await f.flush();await decision(f);assert.equal(f.diagnostics.mounts,1);
  });
  for (const terminal of ['success', 'error']) for (const order of ['before', 'after']) for (const nextOutcome of ['grant', 'error']) test(`${mode} SPGroup old ${terminal} ${order} newer ${nextOutcome} cannot publish`, async t => {
    const f = createTrimsFixture(t, title, { compiled }); f.click(); await f.flush();
    const users = await references(f), old = users[0], run = f.action();
    f.change('Web URL','https://tenant.test/sites/new'); f.click(); await f.flush(); const name = f.calls.at(-1);
    const oldFinish = () => terminal === 'success' ? f.settle(old,{value:[{Id:17}]}) : f.reject(old);
    if(order==='before')await oldFinish();
    if(nextOutcome==='error') await f.reject(name);
    else {await f.settle(name,{value:[{Id:2,Title:'Project Readers'}]});await f.settle(f.calls.at(-1),{Id:91});await f.settle(f.calls.at(-1),{value:[{Id:91}]});}
    if(order==='after')await oldFinish();
    const before=f.diagnostics;f.h.act(run);await f.flush();assert.deepEqual(f.diagnostics,before);assert.equal(f.diagnostics.webUrl,'https://tenant.test/sites/new');assert.equal(f.diagnostics.active,nextOutcome==='grant');assert.equal(f.diagnostics.mounts,nextOutcome==='grant'?1:0);
    assert.equal(old.capture.login,'trims-one');assert.equal(old.capture.client,f.instance.native.sp);assert.equal(old.capture.page,f.instance.native.page);assert.equal(old.capture.method,'GET');assert.deepEqual(old.capture.query,[['$select','Id']]);assert.deepEqual(old.capture.init.headers,{Accept:'application/json;odata.metadata=none'});
  });
  for(const terminal of ['success','error']) test(`${mode} SPGroup reset and departure retire held final settlement and original callbacks ${terminal}`,async t=>{
    const f=createTrimsFixture(t,title,{compiled});f.click();await f.flush();const users=await references(f),old=users[0],run=f.action();
    f.setRevision(3);const before=f.diagnostics;f.h.act(run);if(terminal==='success')await f.settle(old,{value:[{Id:17}]});else await f.reject(old);assert.deepEqual(f.diagnostics,before);assert.equal(f.calls.length,3);assert.ok(f.container.querySelector('[data-trim-demo="idle"]'));
    f.click();await f.flush();const departed=(await references(f))[0],retained=f.action();f.unmount();f.h.act(retained);if(terminal==='success')await f.settle(departed,{value:[{Id:17}]});else await f.reject(departed);assert.equal(f.calls.length,6);
  });
  test(`${mode} SPGroup provider-local inflight sharing and actual two-provider isolation`,async t=>{
    const f=createTrimsFixture(t,title,{compiled});const peer=f.mountPeer(f.instance);f.click();await f.flush();peer.click();await f.flush();assert.equal(f.calls.length,2);assert.equal(f.calls[0].client,f.calls[1].client);f.click();await f.flush();assert.equal(f.calls.length,2);
    await f.settle(f.calls[0],{value:[{Id:2,Title:'Project Readers'}]});await f.settle(f.calls[2],{Id:17});await f.settle(f.calls[3],{value:[{Id:17}]});assert.equal(f.diagnostics.mounts,1);assert.equal(peer.container.querySelector('[data-trim-demo="protected"]'),null);
    await f.settle(f.calls[1],{value:[{Id:2,Title:'Project Readers'}]});await f.settle(f.calls[4],{Id:17});await f.settle(f.calls[5],{value:[]});assert.ok(peer.container.querySelector('[data-trim-state="denied"]'));assert.equal(f.diagnostics.active,true);
  });
  test(`${mode} SPGroup native user client page and provider replacements reject old host actions`,async t=>{
    const f=createTrimsFixture(t,title,{compiled});f.click();await f.flush();await decision(f);const callbacks=f.callbacks(),run=f.action();
    f.instance.native.page.user.loginName='replacement-user';f.render();assert.equal(f.container.querySelector('[data-trim-demo="protected"]'),null);await f.flush();
    f.h.act(run);assert.equal(f.diagnostics.attempt,1);await decision(f,false);const before=f.diagnostics;f.h.act(()=>{callbacks.increment({});callbacks.ref(callbacks.node);callbacks.observe('mount',callbacks.lifetime);callbacks.observe('unmount',callbacks.lifetime);});assert.deepEqual(f.diagnostics,before);
    f.click();await f.flush();const old=(await references(f))[0];f.replace(f.host('new-provider'));assert.ok(f.container.querySelector('[data-trim-demo="idle"]'));await f.settle(old,{value:[{Id:17}]});assert.equal(f.diagnostics.mounts,0);
    f.click();await f.flush();assert.notEqual(f.calls.at(-1).client,old.client);assert.notEqual(f.calls.at(-1).page,old.page);await decision(f);assert.equal(f.diagnostics.mounts,1);assert.equal(f.calls.at(-1).capture.login,'new-provider');f.origins();
  });
  test(`${mode} SPGroup actual AAD user identity refresh retires the original Run`,async t=>{
    const f=createTrimsFixture(t,title,{compiled});f.click();await f.flush();await decision(f);const run=f.action();
    f.instance.native.page.aadInfo={userId:{toString:()=> '11111111-2222-4333-8444-555555555555'}};f.render();
    assert.equal(f.container.querySelector('[data-trim-demo="protected"]'),null);await f.flush();
    f.h.act(run);assert.equal(f.diagnostics.attempt,1,'retained Run cannot replace the current refreshed check');
    await decision(f,false);assert.ok(f.container.querySelector('[data-trim-state="denied"]'));
  });
  for(const fallback of [false,true]) test(`${mode} SPGroup renderer ${fallback?'standalone':'regions'} has actual caller DOM refs formatting and reset baseline`,async t=>{
    const f=createTrimsFixture(t,title,{compiled,fallback});assert.equal(f.calls.length,0);f.click();await f.flush();await decision(f);assert.ok(f.container.querySelector('[data-trim-demo="protected"] strong'));assert.ok(f.container.querySelector('[data-trim-demo="protected"]').className);assert.equal(f.diagnostics.attached,true);f.click('Protected counter: 0');assert.equal(f.diagnostics.counter,1);f.render();assert.equal(f.calls.length,3);assert.equal(f.diagnostics.counter,1);f.reset();assert.equal(f.diagnostics.mounts,0);assert.equal(f.diagnostics.unmounts,0);assert.equal(f.diagnostics.counter,0);assert.equal(f.diagnostics.attached,false);assert.equal(f.container.querySelector('[aria-label="Group values"]').value,'Project Readers');
  });
}

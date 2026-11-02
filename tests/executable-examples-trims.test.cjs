const test = require('node:test');
const assert = require('node:assert/strict');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const owner = 'packages/spfx-react-toolkit/src/controls/sharepoint-permission-trim/sharepoint-permission-trim.tsx#SharePointPermissionTrim';
const single = 'Stage resource permissions and observe protected lifetimes';
const nested = 'Compose permission and Entra checks with nested AND';
const observations = new Map([[single, 'Run a staged resource permission check and observe that protected content mounts only after a real grant; draft edits preserve the applied check and Reset returns to idle.'], [nested, 'Run a staged permission check, then observe that the inner Entra check starts only after the outer grant and protected content requires both decisions.']]);
for (const title of [single, nested]) test(`permission executable availability: ${title}`, () => {
  const selected = scenario(title, owner);
  assert.equal(selected.entry.metadata.host, 'webPart');
  assert.equal(selected.entry.metadata.kind, 'read');
  assert.equal(selected.entry.metadata.contentPlacement, 'configuration');
  assert.equal(selected.entry.metadata.observe, observations.get(title));
  assert.equal(Object.hasOwn(selected.entry.metadata, 'hostAdapter'), false);
});

const { createTrimsFixture, groupId } = require('./fixtures/executable-examples/trims-scenarios.cjs');
for (const compiled of [false, true]) {
  const mode = compiled ? 'COMPILED' : 'SOURCE';
  test(`${mode} native permission grant mounts real child; original Run/reset callbacks retire`, async t => {
    const f = createTrimsFixture(t, single, { compiled }); await f.flush();
    const originalRun = f.action();
    assert.equal(f.calls.length, 0); assert.equal(f.initializations.length, 0);
    f.change('Web URL', 'https://tenant.test/sites/remote'); assert.equal(f.calls.length, 0);
    const field = f.container.querySelector('[aria-label="Web URL"]'); field.focus();
    f.click(); await f.flush(); assert.equal(f.calls.length, 1);
    assert.equal(f.calls[0].phase, 'mask'); assert.equal(f.container.querySelector('[data-trim-state]').dataset.trimState, 'loading');
    assert.equal(f.container.querySelector('[data-trim-demo="protected"]'), null);
    await f.settle(f.calls[0], { High: '0', Low: '1' });
    assert.equal(f.diagnostics.mounts, 1); assert.equal(f.diagnostics.active, true); assert.equal(f.diagnostics.attached, true);
    f.click('Protected counter: 0'); assert.equal(f.diagnostics.counter, 1);
    const applied = f.diagnostics.target;
    f.change('Web URL', 'https://tenant.test/sites/new'); assert.deepEqual(f.diagnostics.target, applied); assert.equal(f.diagnostics.counter, 1); assert.equal(f.calls.length, 1);
    f.reset(); assert.equal(f.container.querySelector('[aria-label="Web URL"]'), field); assert.equal(document.activeElement, field);
    assert.equal(f.diagnostics.mounts, 0); assert.equal(f.diagnostics.unmounts, 0); assert.equal(f.diagnostics.active, false); assert.equal(f.diagnostics.attempt, undefined);
    f.h.act(originalRun); await f.flush(); assert.equal(f.calls.length, 1); assert.equal(f.container.querySelector('[data-trim-demo="protected"]'), null);
    f.origins();
  });
  test(`${mode} actual outer deny/error issue zero inner initialization; grant starts independent Graph phases`, async t => {
    const f = createTrimsFixture(t, nested, { compiled }); f.click(); await f.flush();
    assert.equal(f.initializations.length, 0);
    await f.settle(f.calls[0], { High: '0', Low: '0' }); assert.equal(f.initializations.length, 0); assert.equal(f.diagnostics.mounts, 0);
    f.click(); await f.flush(); await f.settle(f.calls[1], {}, 503); assert.equal(f.initializations.length, 0); assert.equal(f.container.querySelector('[data-trim-state]').dataset.trimState, 'error');
    f.click(); await f.flush(); await f.settle(f.calls[2], { High: '0', Low: '1' });
    assert.equal(f.initializations.length, 1); assert.equal(f.calls.length, 3); assert.equal(f.diagnostics.mounts, 0);
    await f.initialize(); assert.equal(f.calls[3].phase, 'name'); assert.equal(f.diagnostics.mounts, 0);
    await f.settle(f.calls[3], { value: [{ id: groupId, displayName: 'Project Readers' }] }); assert.equal(f.calls[4].phase, 'member'); assert.deepEqual(f.calls[4].body, { groupIds: [groupId] });
    await f.settle(f.calls[4], { value: [groupId] }); assert.equal(f.diagnostics.mounts, 1); assert.equal(f.diagnostics.active, true);
    f.reset(); assert.equal(f.diagnostics.mounts, 0); assert.equal(f.diagnostics.active, false); console.log('# 5c1 native origins ' + JSON.stringify({ mode: compiled ? 'lib' : 'src', ...f.origins() }));
  });
}
const { listId, web } = require('./fixtures/executable-examples/trims-scenarios.cjs');
for (const compiled of [false, true]) {
  const mode = compiled ? 'COMPILED' : 'SOURCE';
  test(`${mode} native both-word masks and all/any permission decisions have zero HTTP`, async t => {
    const f = createTrimsFixture(t, single, { compiled, mask: 0 });
    for (const mask of [0, 1, 4, 5]) for (const permission of ['view', 'edit', 'both']) for (const match of ['all', 'any']) {
      f.instance.native.page.web.permissions = new f.pageNative.SPPermission({ High: 0, Low: mask }); f.render();
      f.change('Permissions', permission); f.change('Match', match); f.click();
      const expected = permission === 'view' ? !!(mask & 1) : permission === 'edit' ? !!(mask & 4) : match === 'all' ? mask === 5 : mask !== 0;
      assert.equal(!!f.container.querySelector('[data-trim-demo="protected"]'), expected, `${mask}/${permission}/${match}`);
      assert.equal(f.container.querySelector('[data-trim-state="loading"]'), null);
      assert.equal(f.calls.length, 0); assert.equal(f.initializations.length, 0);
    }
    const nativeHigh = new f.pageNative.SPPermission({ High: 1, Low: 1 });
    assert.equal(new f.pageNative.SPPermission({ High: 0, Low: 1 }).hasPermission(nativeHigh), false, 'installed native constructor checks high word too');
    assert.equal(new f.pageNative.SPPermission({ High: 1, Low: 1 }).hasPermission(nativeHigh), true);
  });
  for (const scope of ['list','item','folder']) for (const selector of ['title','id','path']) test(`${mode} actual ${scope}/${selector} endpoint and parent-list authority`, async t => {
    const f = createTrimsFixture(t, single, { compiled });
    f.change('Resource scope', scope); f.change('List selector', selector); f.change('List value', selector === 'id' ? listId : selector === 'path' ? 'Lists/Tasks' : "O'Brien");
    if (scope === 'item') f.change('Item ID', '7');
    assert.equal(f.calls.length, 0); f.click(); await f.flush();
    assert.equal(f.calls.length, 1); const call = f.calls[0];
    const endpoint = selector === 'id' ? `lists(guid'${listId}')` : selector === 'path' ? "GetListUsingPath(DecodedUrl='%2Fsites%2Fteam%2FLists%2FTasks')" : "lists/getByTitle('O%27%27Brien')";
    assert.equal(call.url, `${web}/_api/web/${endpoint}${scope === 'folder' ? '?$select=Id' : scope === 'item' ? '/items(7)/effectivebasepermissions' : '/effectivebasepermissions'}`);
    if (scope === 'folder') { await f.settle(call, { Id: listId }); assert.equal(f.calls[1].phase, 'folder'); await f.settle(f.calls[1], { ParentList: { Id: listId }, EffectiveBasePermissions: { High: '0', Low: '1' } }); }
    else await f.settle(call, { High: '0', Low: '1' });
    assert.equal(f.diagnostics.active, true); assert.equal(f.diagnostics.mounts, 1);
  });
  test(`${mode} invalid URL GUID item and folder parent failures select Error, never Denied`, async t => {
    const f = createTrimsFixture(t, single, { compiled });
    for (const edit of [() => f.change('Web URL','invalid URL'), () => { f.change('Web URL',''); f.change('Resource scope','list'); f.change('List selector','id'); f.change('List value','invalid'); }, () => { f.change('Resource scope','item'); f.change('List selector','title'); f.change('List value','Tasks'); f.change('Item ID','NaN'); }]) {
      edit(); const before = f.calls.length; f.click(); await f.flush(); assert.equal(f.calls.length,before); assert.equal(f.nativeAttempts.length,before,'invalid input never crosses the native transport boundary'); assert.ok(f.container.querySelector('[data-trim-state="error"]')); assert.equal(f.diagnostics.mounts,0);
    }
    f.change('Resource scope','folder'); f.click(); await f.flush(); await f.settle(f.calls.at(-1), { Id: listId }); await f.settle(f.calls.at(-1), { ParentList: { Id: groupId }, EffectiveBasePermissions: { High: 0, Low: 5 } }); assert.ok(f.container.querySelector('[data-trim-state="error"]'));
    f.change('Resource scope','web'); f.click(); await f.flush(); await f.settle(f.calls.at(-1), { High: 'bad', Low: '1' }); assert.ok(f.container.querySelector('[data-trim-state="error"]')); assert.equal(f.diagnostics.active,false);
  });
  test(`${mode} settled repeats create attempts; pending equivalent requests share only same provider`, async t => {
    const f = createTrimsFixture(t, single, { compiled }); const peer = f.mountPeer(f.instance);
    f.click(); await f.flush(); peer.click(); await f.flush(); assert.equal(f.calls.length,2); assert.equal(f.calls[0].client,f.calls[1].client,'same native client, separate providers');
    const oldRun=f.action(); f.click(); await f.flush(); assert.equal(f.calls.length,2,'equivalent pending request can share provider-local promise');
    await f.settle(f.calls[0], { High:0, Low:1 }); assert.equal(f.diagnostics.active,true); assert.equal(peer.container.querySelector('[data-trim-demo="protected"]'),null);
    await f.settle(f.calls[1], { High:0, Low:0 }); assert.ok(peer.container.querySelector('[data-trim-state="denied"]'));
    f.h.act(oldRun); await f.flush(); assert.equal(f.calls.length,2);
    f.click(); await f.flush(); assert.equal(f.calls.length,3); assert.equal(f.diagnostics.active,false); await f.settle(f.calls[2], { High:0, Low:0 }); assert.equal(f.diagnostics.unmounts,1);
  });
  for (const terminal of ['success','error']) for (const order of ['before','after']) test(`${mode} old ${terminal} settles ${order} newer grant and cannot restore state`, async t => {
    const f=createTrimsFixture(t,single,{compiled}); f.click(); await f.flush(); const old=f.calls[0]; f.change('Web URL','https://tenant.test/sites/new'); f.click(); await f.flush(); const next=f.calls[1];
    const finishOld=()=>terminal==='success'?f.settle(old,{High:0,Low:0}):f.reject(old);
    if(order==='before') await finishOld(); await f.settle(next,{High:0,Low:1}); if(order==='after') await finishOld();
    assert.equal(f.diagnostics.target.webUrl,'https://tenant.test/sites/new'); assert.equal(f.diagnostics.active,true); assert.equal(f.diagnostics.mounts,1); assert.equal(f.container.querySelector('[data-trim-state]'),null);
    assert.equal(old.capture.web,web); assert.equal(old.capture.client,f.instance.native.sp);
  });
  test(`${mode} reset pending nonmonotonic revisions and departure retire original actions and transport`, async t => {
    const f=createTrimsFixture(t,single,{compiled}); const field=f.container.querySelector('[aria-label="Web URL"]'); field.focus(); const run=f.action(); f.click(); await f.flush(); const old=f.calls[0];
    f.setRevision(2); assert.equal(f.container.querySelector('[aria-label="Web URL"]'),field); assert.equal(document.activeElement,field); f.h.act(run); await f.settle(old,{High:0,Low:1}); assert.equal(f.diagnostics.mounts,0); assert.equal(f.calls.length,1);
    f.click(); await f.flush(); const departed=f.calls[1]; const retained=f.action(); f.unmount(); f.h.act(retained); await f.reject(departed); assert.equal(f.calls.length,2);
  });
  test(`${mode} actual user replacement hides grant and retires old Run`, async t => {
    const f=createTrimsFixture(t,single,{compiled}); f.click(); await f.flush(); await f.settle(f.calls[0],{High:0,Low:1}); const run=f.action();
    f.instance.native.page.user.loginName='new-user'; f.render(); assert.equal(f.container.querySelector('[data-trim-demo="protected"]'),null); await f.flush(); assert.equal(f.calls.length,2);
    f.h.act(run); await f.flush(); assert.equal(f.diagnostics.attempt,1,'old identity Run cannot replace active attempt');
    await f.settle(f.calls[1],{High:0,Low:0}); assert.equal(f.diagnostics.active,false);
  });
  for (const outcome of ['denied','error']) test(`${mode} nested inner ${outcome} never mounts final content and late inner after reset stays retired`, async t => {
    const f=createTrimsFixture(t,nested,{compiled}); f.click(); await f.flush(); await f.settle(f.calls[0],{High:0,Low:1}); await f.initialize(); await f.settle(f.calls[1],{value:[{id:groupId,displayName:'Project Readers'}]});
    if(outcome==='denied') await f.settle(f.calls[2],{value:[]}); else await f.reject(f.calls[2]);
    assert.ok(f.container.querySelector(`[data-trim-level="inner"][data-trim-state="${outcome}"]`)); assert.equal(f.diagnostics.mounts,0);
    f.click(); await f.flush(); await f.settle(f.calls[3],{High:0,Low:1}); await f.initialize(); const late=f.calls[4]; f.reset();
    assert.equal(f.calls.length,5,'Reset itself dispatches no new check');
    if(outcome==='denied') { await f.settle(late,{value:[{id:groupId,displayName:'Project Readers'}]}); assert.equal(f.calls[5].phase,'member','already-started service continuation is not transport-aborted'); await f.settle(f.calls[5],{value:[groupId]}); } else await f.reject(late);
    assert.equal(f.diagnostics.mounts,0); assert.ok(f.container.querySelector('[data-trim-demo="idle"]'));
  });
  for (const title of [single,nested]) test(`${mode} standalone renderer fallback ${title}`, async t => {
    const f=createTrimsFixture(t,title,{compiled,fallback:true}); assert.ok(f.container.querySelector('[data-trim-demo="idle"]')); assert.equal(f.calls.length,0); f.click(); await f.flush(); assert.ok(f.container.querySelector('[data-trim-level="outer"][data-trim-state="loading"]')); f.reset(); assert.equal(f.diagnostics.mounts,0);
  });
}

for (const compiled of [false,true]) test(`${compiled?'COMPILED':'SOURCE'} original protected ref/increment callbacks stay inert after user refresh`, async t => {
  const f=createTrimsFixture(t,single,{compiled}); f.click(); await f.flush(); await f.settle(f.calls[0],{High:0,Low:1}); const callbacks=f.callbacks();
  f.instance.native.page.user.loginName='replacement'; f.render(); await f.flush(); await f.settle(f.calls[1],{High:0,Low:0});
  const before=f.diagnostics; f.h.act(()=>{callbacks.ref(callbacks.node);callbacks.increment({});}); assert.deepEqual(f.diagnostics,before);
});
for (const compiled of [false,true]) {
  const mode=compiled?'COMPILED':'SOURCE';
  for(const terminal of ['success','error']) for(const order of ['before','after']) test(`${mode} held old ${terminal} ${order} newer error cannot restore original grant`,async t=>{
    const f=createTrimsFixture(t,single,{compiled}); f.click(); await f.flush(); const old=f.calls[0];
    f.change('Web URL','https://tenant.test/sites/new'); f.click(); await f.flush(); const next=f.calls[1];
    const release=()=>terminal==='success'?f.settle(old,{High:0,Low:1}):f.reject(old);
    if(order==='before') await release(); await f.settle(next,{},403); if(order==='after') await release();
    assert.ok(f.container.querySelector('[data-trim-state="error"]')); assert.equal(f.diagnostics.mounts,0); assert.equal(f.diagnostics.target.webUrl,'https://tenant.test/sites/new');
    assert.equal(old.capture.url,web+'/_api/web/effectivebasepermissions'); assert.equal(old.capture.method,'GET'); assert.deepEqual(old.capture.query,[]);
  });
  test(`${mode} original mount cleanup increment ref and Run callbacks cannot repopulate reset baseline`,async t=>{
    const f=createTrimsFixture(t,single,{compiled});f.click();await f.flush();await f.settle(f.calls[0],{High:0,Low:1});const callbacks=f.callbacks(),run=f.action();
    f.click('Protected counter: 0'); f.reset(); const before=f.diagnostics;
    f.h.act(()=>{callbacks.increment({});callbacks.ref(callbacks.node);callbacks.ref(null);callbacks.observe('mount',callbacks.lifetime);callbacks.observe('unmount',callbacks.lifetime);run();});
    await f.flush();assert.deepEqual(f.diagnostics,before);assert.equal(f.calls.length,1);assert.equal(callbacks.current(),false);
  });
  test(`${mode} Graph initialization failure never mounts final protected content`,async t=>{
    const f=createTrimsFixture(t,nested,{compiled}); f.click();await f.flush();await f.settle(f.calls[0],{High:0,Low:1});
    assert.equal(f.initializations[0].capture.factory,f.instance.native.factory);await f.initialize(undefined,new Error('factory unavailable'));assert.ok(f.container.querySelector('[data-trim-level="inner"][data-trim-state="error"]'));assert.equal(f.calls.length,1);assert.equal(f.diagnostics.mounts,0);
  });
  test(`${mode} both grants then outer revoke tears down inner and protected subtree`,async t=>{
    const f=createTrimsFixture(t,nested,{compiled});f.click();await f.flush();await f.settle(f.calls[0],{High:0,Low:1});await f.initialize();await f.settle(f.calls[1],{value:[{id:groupId,displayName:'Project Readers'}]});await f.settle(f.calls[2],{value:[groupId]});
    const callbacks=f.callbacks();f.change('Nested Entra group name','Draft only');assert.equal(f.diagnostics.group,'Project Readers');assert.equal(f.calls.length,3);
    f.change('Nested Entra group name','Project Readers');f.click();await f.flush();assert.equal(f.diagnostics.active,false);assert.equal(f.diagnostics.unmounts,1);await f.settle(f.calls[3],{High:0,Low:0});assert.equal(f.initializations.length,1);assert.equal(f.calls.length,4);
    const before=f.diagnostics;f.h.act(()=>{callbacks.ref(callbacks.node);callbacks.increment({});callbacks.observe('mount',callbacks.lifetime);});assert.deepEqual(f.diagnostics,before);
  });
  test(`${mode} provider/client/page replacement retires old captured identity`,async t=>{
    const f=createTrimsFixture(t,single,{compiled}); f.click();await f.flush();const old=f.calls[0],run=f.action();
    f.replace(f.host('replacement'));assert.ok(f.container.querySelector('[data-trim-demo="idle"]'));f.h.act(run);await f.settle(old,{High:0,Low:1});assert.equal(f.diagnostics.mounts,0);assert.equal(f.calls.length,1);
    f.click();await f.flush();assert.notEqual(f.calls[1].client,old.client);assert.notEqual(f.calls[1].page,old.page);assert.equal(f.calls[1].capture.login,'replacement');await f.settle(f.calls[1],{High:0,Low:0});assert.equal(f.diagnostics.active,false);
  });
  test(`${mode} origin witness rejects actually loaded opposite implementation`,t=>{
    const f=createTrimsFixture(t,single,{compiled}); f.origins();
    f.h.load(require('node:path').resolve('packages/spfx-react-toolkit',compiled?'src/services/spfx-scoped-permissions.internal.ts':'lib/services/spfx-scoped-permissions.internal.js'));
    assert.throws(()=>f.origins(),/observed opposite implementation is forbidden/);
  });
}
for(const compiled of [false,true]) test(`${compiled?'COMPILED':'SOURCE'} fresh native positive mask preserves actual granted child counter and observations`,async t=>{
  const f=createTrimsFixture(t,single,{compiled,mask:1});f.click();f.click('Protected counter: 0');const node=f.container.querySelector('[data-trim-demo="protected"]');
  f.instance.native.page.web.permissions=new f.pageNative.SPPermission({High:0,Low:5});f.render();
  assert.equal(f.container.querySelector('[data-trim-demo="protected"]'),node);assert.equal(f.diagnostics.counter,1);assert.equal(f.diagnostics.active,true);assert.equal(f.diagnostics.mounts,1);assert.equal(f.calls.length,0);
});

for(const compiled of [false,true]) test(`${compiled?'COMPILED':'SOURCE'} native synchronous user refresh retains current child lifetime but retires original Run`,t=>{
  const f=createTrimsFixture(t,single,{compiled,mask:1});f.click();f.click('Protected counter: 0');const node=f.container.querySelector('[data-trim-demo="protected"]'),run=f.action();
  f.instance.native.page.user.loginName='fresh-native-user';f.render();assert.equal(f.container.querySelector('[data-trim-demo="protected"]'),node);assert.equal(f.diagnostics.counter,1);assert.equal(f.diagnostics.active,true);assert.equal(f.diagnostics.mounts,1);
  f.h.act(run);assert.equal(f.diagnostics.attempt,1);assert.equal(f.calls.length,0);
});
for(const compiled of [false,true]) for(const terminal of ['success','error']) for(const order of ['before','after']) test(`${compiled?'COMPILED':'SOURCE'} original inner ${terminal} ${order} replaced outer denial cannot restore content`,async t=>{
  const f=createTrimsFixture(t,nested,{compiled});f.click();await f.flush();await f.settle(f.calls[0],{High:0,Low:1});await f.initialize();await f.settle(f.calls[1],{value:[{id:groupId,displayName:'Project Readers'}]});const old=f.calls[2];
  f.instance.native.page.user.loginName='next-user';f.render();await f.flush();const next=f.calls[3];assert.equal(next.kind,'sp');assert.equal(f.initializations.length,1);
  const settle=()=>terminal==='success'?f.settle(old,{value:[groupId]}):f.reject(old);
  if(order==='before')await settle();await f.settle(next,{High:0,Low:0});if(order==='after')await settle();
  assert.ok(f.container.querySelector('[data-trim-level="outer"][data-trim-state="denied"]'));assert.equal(f.diagnostics.mounts,0);assert.equal(f.initializations.length,1);assert.equal(old.capture.login,'trims-one');assert.equal(next.capture.login,'next-user');
});
for(const compiled of [false,true]) test(`${compiled?'COMPILED':'SOURCE'} normalized explicit current web uses actual native mask without fake Loading`,t=>{
  const f=createTrimsFixture(t,single,{compiled,mask:1});f.change('Web URL','https://TENANT.test/sites/team///');f.click();assert.ok(f.container.querySelector('[data-trim-demo="protected"]'));assert.equal(f.calls.length,0);assert.equal(f.container.querySelector('[data-trim-state="loading"]'),null);assert.ok(f.container.querySelector('[data-trim-demo="protected"]').className);assert.ok(f.container.querySelector('[data-trim-demo="protected"] strong'));
});
for(const compiled of [false,true]) test(`${compiled?'COMPILED':'SOURCE'} synchronous settled repeat reports actual new caller ref and cumulative child lifetimes`,t=>{
  const f=createTrimsFixture(t,single,{compiled,mask:1});f.click();f.click('Protected counter: 0');const old=f.callbacks();f.click();
  assert.equal(f.diagnostics.mounts,2);assert.equal(f.diagnostics.unmounts,1);assert.equal(f.diagnostics.counter,0);assert.equal(f.diagnostics.active,true);assert.equal(f.diagnostics.attached,true);assert.notEqual(f.container.querySelector('[data-trim-demo="protected"]'),old.node);
});

for(const compiled of [false,true]) test(`${compiled?'COMPILED':'SOURCE'} native route rejection is recorded independently of product Error and cannot masquerade as zero reads`,async t=>{
 const f=createTrimsFixture(t,single,{compiled,expectedProtocolErrors:1});f.change('Resource scope','list');f.change('List value','Not admitted');f.click();await f.flush();assert.equal(f.nativeAttempts.length,1);assert.equal(f.calls.length,0);assert.equal(f.protocolErrors.length,1);assert.ok(f.container.querySelector('[data-trim-state="error"]'));assert.equal(f.diagnostics.mounts,0);
});
for(const compiled of [false,true]) test(`${compiled?'COMPILED':'SOURCE'} actual matching provider metadata does not leak into the opposite implementation context`,t=>{
 const f=createTrimsFixture(t,single,{compiled}),path=require('node:path');f.origins();
 const actual=f.h.load(path.join(f.root,'core/context.internal'+(compiled?'.js':'.tsx')));
 const opposite=f.h.load(path.join(path.dirname(f.root),compiled?'src':'lib','core/context.internal'+(compiled?'.tsx':'.js')));
 assert.notEqual(actual.SPFxContext,opposite.SPFxContext);let observed;
 function Probe(){const matching=actual.useSPFxContext();const mismatched=f.h.React.useContext(opposite.SPFxContext);observed={matching,mismatched:mismatched??undefined};return f.h.React.createElement('output',undefined,'Observed actual context identities');}
 const peer=f.h.mountComponent(()=>f.h.React.createElement(f.toolkit.SPFxWebPartProvider,{instance:f.instance},f.h.React.createElement(Probe)));
 assert.equal(observed.matching.spfxContext,f.instance.context);assert.equal(observed.mismatched,undefined,'a real other implementation context has no matching provider value');assert.equal(f.calls.length,0);assert.equal(f.initializations.length,0);
 assert.throws(()=>f.origins(),/observed opposite implementation is forbidden/);peer.unmount();
});

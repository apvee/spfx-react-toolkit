const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const h = createHarness();
const load = () => h.load('controls/file-folder-picker/file-folder-picker-session.internal.ts');
const tick = () => new Promise(resolve => setImmediate(resolve));
function host(overrides = {}) {
  const listeners = new Set(), sent = [], calls = [];
  const ownerWindow = { addEventListener: (_, fn) => listeners.add(fn), removeEventListener: (_, fn) => listeners.delete(fn) };
  const contentWindow = { closed: false, close() { this.closed = true; } };
  const port = { onmessage: undefined, start() {}, close() { this.closed = true; }, postMessage: message => sent.push(message) };
  const session = load().createFileFolderPickerSession({ ownerWindow, contentWindow, trustedOrigin: 'https://tenant.sharepoint.com', channelId: 'one',
    allowedResourceOrigins: ['https://tenant.sharepoint.com'], requestToken: async command => { calls.push(command); return 'secret'; },
    validatePick: async items => items.map(item => ({ text: item.id })), ...overrides });
  return {session, sent, calls, listeners, contentWindow, port,
    initialize(extra = {}) { for (const fn of listeners) fn({ origin: 'https://tenant.sharepoint.com', source: contentWindow, data: {type:'initialize',channelId:'one'}, ports:[port], ...extra }); },
    command(id, data) { port.onmessage?.({data:{type:'command',id,data}}); } };
}
test('authenticates only exact source/origin/channel and processes each command ID once', async () => {
  const x = host(); const result = x.session.run();
  x.initialize({origin:'https://evil.sharepoint.com'}); assert.equal(x.port.onmessage, undefined);
  x.initialize({source:{}}); assert.equal(x.port.onmessage, undefined);
  x.initialize({data:{type:'initialize',channelId:'other'}}); assert.equal(x.port.onmessage, undefined);
  x.initialize(); assert.deepEqual(x.sent, [{type:'activate'}]);
  x.command('a',{command:'authenticate',resource:'https://tenant.sharepoint.com',type:'SharePoint'});
  x.command('a',{command:'authenticate',resource:'https://tenant.sharepoint.com',type:'SharePoint'}); await tick();
  assert.equal(x.calls.length,1); assert.equal(x.sent.filter(m=>m.type==='acknowledge').length,1);
  assert.equal(x.sent.find(m=>m.type==='result').data.token,'secret');
  x.command('b',{command:'pick',items:[{id:'FILE'}]}); await tick();
  assert.deepEqual(await result,{status:'picked',items:[{text:'FILE'}]}); assert.equal(x.listeners.size,0); assert.equal(x.port.closed,true);
});
test('foreign token resource and claims never reach provider; unsupported commands get explicit errors', async () => {
  const x=host(); const result=x.session.run(); x.initialize();
  for (const [id,data] of [['a',{command:'authenticate',resource:'https://graph.microsoft.com',type:'Graph'}],['b',{command:'authenticate',resource:'https://tenant.sharepoint.com',claims:'challenge'}],['c',{command:'delete'}]]) x.command(id,data);
  await tick(); assert.equal(x.calls.length,0);
  assert.equal(x.sent.filter(m=>m.type==='result'&&m.data.result==='error').length,3);
  x.session.cancel(); assert.deepEqual(await result,{status:'cancelled'});
});

test('optional auth diagnostics are bounded, deduplicated and cannot break protocol or leak provider details', async(t)=>{
 const events=[];const x=host({onAuthenticationDiagnostic:event=>{events.push(event);throw new Error('observer');},requestToken:async()=>{throw new Error('PRIVATE TOKEN https://secret.test');}});
 t.after(()=>x.session.cancel());const result=x.session.run();x.initialize();
 for(let i=0;i<30;i++)x.command('foreign'+i,{command:'authenticate',resource:'https://foreign.test/private?token=SECRET',type:'PRIVATE'});
 x.command('provider',{command:'authenticate',resource:'https://tenant.sharepoint.com'});await tick();
 assert.equal(events.filter(event=>event.stage==='resource-invalid-shape').length,1);
 assert.equal(events.some(event=>event.stage==='provider-failed'),true);
 assert.doesNotMatch(JSON.stringify(events),/PRIVATE|SECRET|secret|https:/);
 assert.equal(x.sent.filter(m=>m.type==='result').length,31);
 x.session.cancel();await result;
});

test('late provider failures after cancel cannot publish diagnostics',async(t)=>{
 let rejectToken;const events=[];const x=host({onAuthenticationDiagnostic:event=>events.push(event),requestToken:()=>new Promise((_,reject)=>rejectToken=reject)});
 t.after(()=>x.session.cancel());const result=x.session.run();x.initialize();x.command('a',{command:'authenticate',resource:'https://tenant.sharepoint.com'});await tick();
 const before=events.length;x.session.cancel();rejectToken(new Error('PRIVATE'));await tick();await result;assert.equal(events.length,before);
});

test('diagnostic budget is per opening and obsolete ports cannot publish late provider failures',async(t)=>{
 const events=[];const x=host({onAuthenticationDiagnostic:event=>events.push(event)});t.after(()=>x.session.cancel());const result=x.session.run();x.initialize();
 let id=0;
 for(const origin of ['https://tenant.sharepoint.com','https://graph.microsoft.com','https://foreign.test'])
  for(const path of ['/private','/?private=1','/#private','/private?private=1#private'])
   for(const type of [undefined,'SharePoint','Graph','PRIVATE'])x.command(++id,{command:'authenticate',resource:origin+path,type});
 await tick();assert.equal(events.length,16);assert.equal(x.calls.length,0);x.session.cancel();await result;
 let rejectToken;const late=[];const y=host({onAuthenticationDiagnostic:event=>late.push(event),requestToken:()=>new Promise((_,reject)=>rejectToken=reject)});t.after(()=>y.session.cancel());const outcome=y.session.run();y.initialize();y.command('old',{command:'authenticate',resource:'https://tenant.sharepoint.com'});await tick();
 y.initialize({ports:[{start(){},close(){},postMessage(){}}]});const count=late.length;rejectToken(new Error('PRIVATE'));await tick();assert.equal(late.length,count);y.session.cancel();await outcome;
});

test('personal navigation authenticates only an independently resolved exact audience and coalesces discovery',async(t)=>{
 let resolutions=0;const recoveries=[],origins=Object.freeze(['https://tenant.sharepoint.com']);const x=host({allowedResourceOrigins:origins,onAuthenticationSuccess:()=>recoveries.push(true),resolveAuxiliaryAuthenticationOrigins:async()=>{resolutions++;return ['https://tenant-my.sharepoint.com'];}});t.after(()=>x.session.cancel());const result=x.session.run();x.initialize();
 for(const id of ['a','b'])x.command(id,{command:'authenticate',resource:'https://tenant-my.sharepoint.com',type:'SharePoint_MSAL'});
 await tick();assert.equal(resolutions,1);assert.equal(x.calls.length,2);assert.equal(x.calls[0].resource,'https://tenant-my.sharepoint.com');assert.deepEqual(recoveries,[]);assert.deepEqual(origins,['https://tenant.sharepoint.com']);
 for(const [id,resource,type] of [['foreign','https://other-my.sharepoint.com','SharePoint_MSAL'],['path','https://tenant-my.sharepoint.com/private','SharePoint_MSAL'],['unknown','https://tenant-my.sharepoint.com','PRIVATE'],['graph','https://graph.microsoft.com','Graph']])x.command(id,{command:'authenticate',resource,type});
 await tick();assert.equal(resolutions,1);assert.equal(x.calls.length,2);x.session.cancel();await result;
});

test('auxiliary auth never reports or clears configured-source failures and invalid claims do not discover',async(t)=>{
 const failures=[],recoveries=[],events=[];let discoveries=0;
 const x=host({claimsChallengeEnabled:false,resolveAuxiliaryAuthenticationOrigins:async()=>{discoveries++;return ['https://tenant-my.sharepoint.com'];},onError:e=>failures.push(e),onAuthenticationSuccess:()=>recoveries.push(true),onAuthenticationDiagnostic:e=>events.push(e),requestToken:async r=>{if(r.resource.includes('-my.'))throw new Error('PRIVATE');return 'source';}});
 t.after(()=>x.session.cancel());const result=x.session.run();x.initialize();
 x.command('claims',{command:'authenticate',resource:'https://tenant-my.sharepoint.com',type:'SharePoint_MSAL',claims:'PRIVATE'});await tick();assert.equal(discoveries,0);
 x.command('personal',{command:'authenticate',resource:'https://tenant-my.sharepoint.com',type:'SharePoint_MSAL'});await tick();assert.equal(discoveries,1);assert.deepEqual(failures,[]);assert.deepEqual(recoveries,[]);
 assert.equal(events.some(e=>e.stage==='accepted'&&e.resource==='tenant-personal-candidate'),true);
 x.command('source',{command:'authenticate',resource:'https://tenant.sharepoint.com',type:'SharePoint'});await tick();assert.deepEqual(recoveries,[true]);x.session.cancel();await result;
});

test('auxiliary audience resolution failure, mismatch and cancellation never reach token provider',async(t)=>{
 for(const resolve of [async()=>{throw new Error('PRIVATE');},async()=>['https://other-my.sharepoint.com']]){
  const x=host({resolveAuxiliaryAuthenticationOrigins:resolve});t.after(()=>x.session.cancel());const result=x.session.run();x.initialize();x.command('a',{command:'authenticate',resource:'https://tenant-my.sharepoint.com',type:'SharePoint_MSAL'});await tick();assert.equal(x.calls.length,0);assert.equal(x.sent.find(m=>m.type==='result').data.result,'error');assert.doesNotMatch(JSON.stringify(x.sent),/PRIVATE/);x.session.cancel();await result;
 }
 let finish;const x=host({resolveAuxiliaryAuthenticationOrigins:()=>new Promise(resolve=>finish=resolve)});t.after(()=>x.session.cancel());const result=x.session.run();x.initialize();x.command('a',{command:'authenticate',resource:'https://tenant-my.sharepoint.com',type:'SharePoint_MSAL'});await tick();x.session.cancel();finish(['https://tenant-my.sharepoint.com']);await tick();assert.equal(x.calls.length,0);await result;
});

test('synchronous cancellation prevents auxiliary discovery from starting in its queued microtask',async()=>{
 let discoveries=0;const x=host({resolveAuxiliaryAuthenticationOrigins:async()=>{discoveries++;return ['https://tenant-my.sharepoint.com'];}});const result=x.session.run();x.initialize();
 x.command('a',{command:'authenticate',resource:'https://tenant-my.sharepoint.com',type:'SharePoint_MSAL'});x.session.cancel();await tick();await result;assert.equal(discoveries,0);assert.equal(x.calls.length,0);
});
test('cancel/dispose suppresses late auth and pending validated picks and closes owned resources', async () => {
  let resolveToken, resolvePick;
  const x=host({requestToken:()=>new Promise(r=>resolveToken=r),validatePick:()=>new Promise(r=>resolvePick=r)});
  const result=x.session.run(); x.initialize(); x.command('a',{command:'authenticate',resource:'https://tenant.sharepoint.com'});
  x.command('b',{command:'pick',items:[{id:'FILE'}]}); await tick(); x.session.dispose();
  resolveToken('private'); resolvePick([{text:'late'}]); await tick();
  assert.deepEqual(await result,{status:'cancelled'}); assert.equal(x.sent.filter(m=>m.type==='result').length,0);
  assert.equal(x.contentWindow.closed,true); assert.equal(x.listeners.size,0);
});
test('validation errors are atomic and actionable but do not confirm a subset', async () => {
  const x=host({validatePick:async()=>{throw new Error('Outside configured folder.');}}); const result=x.session.run(); x.initialize();
  x.command('a',{command:'pick',items:[{id:'inside'},{id:'outside'}]}); await tick();
  const response=x.sent.find(m=>m.type==='result'); assert.equal(response.data.result,'error'); assert.equal(response.data.error.code,'unusableItem');
  x.command('b',{command:'close'}); assert.deepEqual(await result,{status:'cancelled'});
});
test('handshake and command timeout fail and cleanup; manual window close cancels', async () => {
  const x=host({handshakeTimeoutMs:10}); await assert.rejects(x.session.run(),{ code: 'timeout' }); assert.equal(x.listeners.size,0);
  const y=host({commandTimeoutMs:10,requestToken:()=>new Promise(()=>{})}); const result=y.session.run(); y.initialize(); y.command('a',{command:'authenticate',resource:'https://tenant.sharepoint.com'});
  await assert.rejects(result,{ code: 'timeout' }); assert.equal(y.port.closed,true);
  const z=host({closePollMs:5}); const cancel=z.session.run(); z.contentWindow.closed=true; assert.deepEqual(await cancel,{status:'cancelled'});
});
test('reinitialization replaces port and invalidates late auth on old port', async () => {
  let resolveToken;const x=host({requestToken:()=>new Promise(r=>resolveToken=r)});const result=x.session.run();x.initialize();const old=x.port.onmessage;
  x.command('a',{command:'authenticate',resource:'https://tenant.sharepoint.com'});await tick();
  const secondSent=[];const second={close(){this.closed=true},start(){},postMessage(m){secondSent.push(m)}};x.initialize({ports:[second]});assert.equal(x.port.closed,true);assert.equal(second.closed,undefined);
  resolveToken('private');await tick();assert.equal(secondSent.some(m=>m.data?.token),false);assert.equal(x.sent.some(m=>m.data?.token),false);
  old({data:{type:'command',id:'late',data:{command:'authenticate',resource:'https://tenant.sharepoint.com'}}});await tick();assert.equal(x.calls.length,0);
  x.session.cancel();await result;
});
test('only the owning pending pick can release confirmation lock', async()=>{
  let resolvePick;let validations=0;const x=host({validatePick:()=>{validations++;return new Promise(r=>resolvePick=r)}});const result=x.session.run();x.initialize();
  x.command('a',{command:'pick',items:[{id:'A'}]});x.command('b',{command:'pick',items:[{id:'B'}]});await tick();x.command('c',{command:'pick',items:[{id:'C'}]});await tick();assert.equal(validations,1);
  resolvePick([{text:'A'}]);assert.equal((await result).status,'picked');
});
test('current actionable pick failures reach the host without exposing authentication errors',async(t)=>{
 const failures=[];const x=host({onError:error=>failures.push(error.code),validatePick:async()=>{throw new Error('Missing selected-item SharePoint provenance metadata.')}});const result=x.session.run();t.after(()=>x.session.cancel());x.initialize();x.command('p',{command:'pick',items:[{id:'FILE'}]});await tick();assert.deepEqual(failures,['selection-identity']);x.session.cancel();await result;
});

test('auth failures distinguish unsupported resource from provider rejection and recover after a valid token', async t => {
  const failures = [], recoveries = [];
  const x = host({ onError: error => failures.push(error), onAuthenticationSuccess: () => recoveries.push(true), requestToken: async () => { throw new Error('private-provider-details'); } });
  t.after(() => x.session.cancel());
  x.session.run(); x.initialize();
  x.command('blocked', { command: 'authenticate', resource: 'https://evil.sharepoint.com', type: 'SharePoint' });
  await tick();
  assert.equal(failures.length, 0, 'a denied optional audience must not mark the configured source as failed');
  assert.equal(x.sent.find(m => m.type === 'result').data.error.code, 'unableToObtainToken');
  x.command('provider', { command: 'authenticate', resource: 'https://tenant.sharepoint.com', type: 'SharePoint' });
  await tick();
  assert.equal(failures[0].code, 'auth-provider');
  assert.ok(failures.every(error => !error.message.includes('private')));
  assert.equal(recoveries.length, 0);
  const y = host({ onAuthenticationSuccess: () => recoveries.push(true) });
  t.after(() => y.session.cancel()); y.session.run(); y.initialize();
  y.command('good', { command: 'authenticate', resource: 'https://tenant.sharepoint.com', type: 'SharePoint' });
  await tick(); assert.equal(recoveries.length, 1);
});

test('page-loaded notifies once per opening and close acknowledges success before disposal', async () => {
 const ready=[]; const x=host({onReady:()=>ready.push(true)}); const result=x.session.run(); x.initialize();
 assert.equal(ready.length,0);
 x.port.onmessage({data:{type:'notification',data:{notification:'page-loaded'}}});
 x.port.onmessage({data:{type:'notification',data:{notification:'page-loaded'}}});
 x.port.onmessage({data:{type:'notification',data:{notification:'unknown'}}});
 await tick(); assert.equal(ready.length,1);
 x.command('close',{command:'close'}); await result;
 assert.ok(x.sent.some(m=>m.id==='close'&&m.type==='result'&&m.data.result==='success'));
});
test('claims retry is bounded per resource and Graph failure does not break baseline tokens', async () => {
 const diagnostics=[];const x=host({graphEnabled:true,claimsChallengeEnabled:true,onAuthenticationDiagnostic:event=>diagnostics.push(event)});const result=x.session.run();x.initialize();
 const challenge={claims:'{"access_token":{}}'};
 x.command('c1',{command:'authenticate',resource:'https://tenant.sharepoint.com',claimsChallenge:challenge});await tick();
 assert.equal(x.calls[0].claims,challenge.claims);
 x.command('c2',{command:'authenticate',resource:'https://tenant.sharepoint.com',claimsChallenge:challenge});await tick();
 assert.equal(x.calls.length,1);
 assert.equal(diagnostics.some(event=>event.stage==='claims-denied'),true);
 assert.equal(x.sent.find(m=>m.id==='c2'&&m.type==='result').data.result,'error');
 x.command('normal',{command:'authenticate',resource:'https://tenant.sharepoint.com'});await tick();
 assert.equal(x.sent.find(m=>m.id==='normal'&&m.type==='result').data.result,'token');
 x.session.cancel();await result;
});

test('optional Graph token failure stays in its command and cannot report or clear SharePoint source errors', async t => {
 const failures=[],recoveries=[];
 const x=host({graphEnabled:true,onError:e=>failures.push(e),onAuthenticationSuccess:()=>recoveries.push(true),requestToken:async r=>{if(r.type==='Graph') throw new Error('private graph failure');return 'sp-token';}});
 t.after(()=>x.session.cancel());
 const result=x.session.run();x.initialize();
 x.command('g',{command:'authenticate',resource:'https://graph.microsoft.com',type:'Graph'});await tick();
 assert.deepEqual(failures,[]);assert.deepEqual(recoveries,[]);
 const reply=x.sent.find(m=>m.id==='g'&&m.type==='result');assert.equal(reply.data.result,'error');assert.match(reply.data.error.message,/Graph/);
 x.command('sp',{command:'authenticate',resource:'https://tenant.sharepoint.com'});await tick();
 assert.equal(x.sent.find(m=>m.id==='sp'&&m.type==='result').data.token,'sp-token');
 x.session.cancel();await result;
});

test('invalid optional Graph schema is command-local and never acquires or recovers source tokens', async t => {
 const failures=[],recoveries=[];const x=host({graphEnabled:true,claimsChallengeEnabled:true,onError:e=>failures.push(e.code),onAuthenticationSuccess:()=>recoveries.push(true)});
 t.after(()=>x.session.cancel());const result=x.session.run();x.initialize();
 for(const [id,extra] of [['type',{type:'SharePoint'}],['nested',{type:'Graph',claimsChallenge:{claims:'PRIVATE'}}],['direct',{type:'Graph',claims:'PRIVATE'}]])x.command(id,{command:'authenticate',resource:'https://graph.microsoft.com',...extra});
 await tick();assert.equal(x.calls.length,0);assert.deepEqual(failures,[]);assert.deepEqual(recoveries,[]);
 assert.equal(x.sent.filter(m=>m.type==='result'&&m.data.result==='error').length,3);assert.doesNotMatch(JSON.stringify(x.sent),/PRIVATE/);
 x.command('source',{command:'authenticate',resource:'https://tenant.sharepoint.com',type:'SharePoint'});await tick();assert.equal(x.calls.length,1);assert.deepEqual(recoveries,[true]);x.session.cancel();await result;
});

test('replacement before auxiliary discovery microtask shares opening discovery without obsolete replies', async t => {
 let discoveries=0;const events=[];const x=host({resolveAuxiliaryAuthenticationOrigins:async()=>{discoveries++;return ['https://tenant-my.sharepoint.com'];},onAuthenticationDiagnostic:e=>events.push(e)});
 t.after(()=>x.session.cancel());const result=x.session.run();x.initialize();
 const data={command:'authenticate',resource:'https://tenant-my.sharepoint.com',type:'SharePoint_MSAL'};
 x.command('old',data);const sent=[];const second={start(){},close(){},postMessage:m=>sent.push(m)};x.initialize({ports:[second]});const previousEvents=events.length;
 second.onmessage({data:{type:'command',id:'new',data}});await tick();
 assert.equal(discoveries,1);assert.equal(x.calls.length,1);assert.equal(sent.find(m=>m.id==='new'&&m.type==='result').data.token,'secret');
 assert.equal(x.sent.some(m=>m.type==='result'),false);assert.equal(sent.some(m=>m.id==='old'),false);
 assert.equal(events.slice(previousEvents).filter(e=>e.stage==='token-supplied').length,1);x.session.cancel();await result;
});

test('optional authentication timeout preserves source session and suppresses all late completion effects',async t=>{
 for(const resource of ['https://graph.microsoft.com','https://tenant-my.sharepoint.com']){
  let finish;const failures=[],recoveries=[],events=[];
  const x=host({graphEnabled:true,commandTimeoutMs:10,resolveAuxiliaryAuthenticationOrigins:async()=>['https://tenant-my.sharepoint.com'],onError:e=>failures.push(e.code),onAuthenticationSuccess:()=>recoveries.push(true),onAuthenticationDiagnostic:e=>events.push(e),requestToken:r=>r.resource===resource?new Promise(resolve=>finish=resolve):Promise.resolve('source')});
  t.after(()=>x.session.cancel());const result=x.session.run();result.catch(()=>undefined);x.initialize();x.command('optional',{command:'authenticate',resource,type:resource.includes('graph')?'Graph':'SharePoint_MSAL'});
  await new Promise(resolve=>setTimeout(resolve,25));
  assert.equal(x.port.closed,undefined);assert.deepEqual(failures,[]);assert.deepEqual(recoveries,[]);assert.equal(x.sent.find(m=>m.id==='optional'&&m.type==='result').data.result,'error');
  const count=events.length;finish('LATE-PRIVATE');await tick();assert.equal(events.length,count);assert.equal(x.sent.some(m=>m.data?.token==='LATE-PRIVATE'),false);
  x.command('source',{command:'authenticate',resource:'https://tenant.sharepoint.com'});await tick();assert.equal(x.sent.find(m=>m.id==='source'&&m.type==='result').data.token,'source');x.session.cancel();assert.deepEqual(await result,{status:'cancelled'});
 }
});

test('optional discovery timing out cannot acquire a late token or publish source errors',async t=>{
 let resolveDiscovery;const failures=[],events=[];const x=host({commandTimeoutMs:10,resolveAuxiliaryAuthenticationOrigins:()=>new Promise(resolve=>resolveDiscovery=resolve),onError:e=>failures.push(e),onAuthenticationDiagnostic:e=>events.push(e)});
 t.after(()=>x.session.cancel());const result=x.session.run();result.catch(()=>undefined);x.initialize();x.command('personal',{command:'authenticate',resource:'https://tenant-my.sharepoint.com',type:'SharePoint_MSAL'});
 await new Promise(resolve=>setTimeout(resolve,25));const count=events.length;assert.equal(x.sent.find(m=>m.id==='personal'&&m.type==='result').data.result,'error');
 resolveDiscovery(['https://tenant-my.sharepoint.com']);await tick();assert.equal(x.calls.length,0);assert.deepEqual(failures,[]);assert.equal(events.length,count);
 x.command('source',{command:'authenticate',resource:'https://tenant.sharepoint.com'});await tick();assert.equal(x.calls.length,1);x.session.cancel();await result;
});

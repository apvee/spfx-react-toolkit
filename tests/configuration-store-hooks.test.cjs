const assert = require('node:assert/strict');
const test = require('node:test');
const { React, loadHook, mount, deferred, start, settle, flush } = require('./async-test-harness.cjs');

function fixture(t, implementations={}) {
  let client={},mode='layered';
  const context={site:{absoluteUrl:'https://tenant.test/sites/root'},web:{absoluteUrl:'https://tenant.test/sites/root/sub'}};
  const creations=[], calls=[];
  const hook=loadHook('useSPFxConfigurationStore','useSPFxConfigurationStore',{
    './useSPFxSPHttpClient':{useSPFxSPHttpClient:()=>({client})},
    './useSPFxPageContext':{useSPFxPageContext:()=>context},
    '../services/spfx-configuration-store.service':{createSPFxConfigurationStoreService:(c,options)=>{
      creations.push({client:c,options});
      return Object.fromEntries(['get','listScope','save','remove'].map(method=>[method,async(...args)=>{
        calls.push({method,args,options});
        return implementations[method] ? implementations[method](...args) : method==='listScope'?[]:undefined;
      }]));
    }}
  });
  const view=mount(()=>hook({mode}),{});t.after(view.unmount);
  return {...view,get current(){return view.current;},context,creations,calls,setClient:c=>{client=c;view.render({});},setMode:m=>{mode=m;view.render({});},rerender:()=>view.render({})};
}
test('mount is request-free and equivalent inline options preserve callbacks',async t=>{
  const f=fixture(t);await flush();assert.deepEqual(f.calls,[]);assert.equal(f.current.isReady,true);
  const get=f.current.get;f.rerender();assert.equal(f.current.get,get);assert.equal(f.creations.length,1);
});
for(const method of ['get','listScope','save','remove'])test(method+' rejects original failure and publishes only its channel',async t=>{
  const error=new Error('denied');const f=fixture(t,{[method]:async()=>{throw error;}});
  await settle(async()=>assert.equal(await f.current[method]('a',1,{scope:'site'}).catch(e=>e),error));
  const write=method==='save'||method==='remove';
  assert.equal(f.current[write?'writeError':'error'],error);assert.equal(f.current[write?'error':'writeError'],undefined);
  assert.equal(f.current.isLoading,false);assert.equal(f.current.isWriting,false);
});
test('read/write counters are independent and old failure cannot replace newer success',async t=>{
  const old=deferred(),fresh=deferred(),write=deferred();let n=0;
  const f=fixture(t,{get:()=>++n===1?old.promise:fresh.promise,save:()=>write.promise});
  const a=start(()=>f.current.get('a').catch(e=>e)),b=start(()=>f.current.get('b')),c=start(()=>f.current.save('a',1,{scope:'site'}));
  assert.equal(f.current.isLoading,true);assert.equal(f.current.isWriting,true);
  await settle(async()=>{fresh.resolve(undefined);await b;});assert.equal(f.current.isLoading,true);
  await settle(async()=>{old.reject(new Error('old'));await a;});assert.equal(f.current.error,undefined);assert.equal(f.current.isLoading,false);
  await settle(async()=>{write.resolve();await c;});assert.equal(f.current.isWriting,false);
});
test('non Error rejection is retained while error state is normalized',async t=>{
  const f=fixture(t,{get:async()=>{throw 'denied';}});
  await settle(async()=>assert.equal(await f.current.get('a').catch(e=>e),'denied'));
  assert.equal(f.current.error.message,'denied');
});
test('missing local dependencies reject in the right channel without dispatch',async t=>{
  const f=fixture(t);f.setClient(undefined);assert.equal(f.current.isReady,false);
  await settle(async()=>{await assert.rejects(f.current.get('a'));await assert.rejects(f.current.save('a',1,{scope:'site'}));});
  assert.ok(f.current.error);assert.ok(f.current.writeError);assert.deepEqual(f.calls,[]);
  f.setClient({});f.context.site.absoluteUrl='///';f.rerender();assert.equal(f.current.isReady,false);
});
test('mode-specific URL dependencies ignore irrelevant navigation',async t=>{
  const f=fixture(t);f.setMode('site');const siteGet=f.current.get;
  f.context.web.absoluteUrl='https://tenant.test/other';f.rerender();assert.equal(f.current.get,siteGet);
  f.setMode('tenant');const tenantGet=f.current.get;
  f.context.site.absoluteUrl='https://tenant.test/sites/other';f.rerender();assert.equal(f.current.get,tenantGet);
  assert.equal('siteCollectionUrl' in f.creations.at(-1).options,false);
});
test('retained callbacks keep old URL identity and obsolete completions cannot publish',async t=>{
  const old=deferred();let count=0;
  const f=fixture(t,{get:()=>++count===1?old.promise:Promise.resolve(undefined)});
  const oldGet=f.current.get;const pending=start(()=>oldGet('a').catch(e=>e));
  f.context.site.absoluteUrl='https://tenant.test/sites/new';f.context.web.absoluteUrl='https://tenant.test/sites/new/child';f.rerender();
  assert.equal(f.current.isLoading,false);assert.equal(f.current.error,undefined);
  await settle(async()=>{await f.current.get('b');old.reject(new Error('obsolete'));await pending;});assert.equal(f.current.error,undefined);
  await settle(()=>oldGet('old'));assert.equal(f.calls.at(-1).options.siteCollectionUrl,'https://tenant.test/sites/root');
});
test('save remains successful when a later explicit read fails and does not automatically read',async t=>{
  const f=fixture(t,{get:async()=>{throw new Error('read failed');}});
  await settle(()=>f.current.save('a',1,{scope:'site'}));assert.deepEqual(f.calls.map(c=>c.method),['save']);
  await settle(()=>f.current.get('a').catch(()=>{}));assert.equal(f.current.writeError,undefined);assert.equal(f.current.error.message,'read failed');
});
test('newer write success wins over an older write failure and instances have independent counters',async t=>{
  const old=deferred(),fresh=deferred();let n=0;
  const f=fixture(t,{save:()=>++n===1?old.promise:fresh.promise}),other=fixture(t);
  const a=start(()=>f.current.save('a',1,{scope:'site'}).catch(e=>e)),b=start(()=>f.current.save('b',2,{scope:'tenant'}));
  assert.equal(other.current.isWriting,false);
  await settle(async()=>{fresh.resolve();await b;});assert.equal(f.current.isWriting,true);
  await settle(async()=>{old.reject(new Error('old write'));await a;});
  assert.equal(f.current.isWriting,false);assert.equal(f.current.writeError,undefined);assert.equal(other.current.writeError,undefined);
});
test('client/mode replacement and unmount suppress pending state publication',async t=>{
  const d=deferred();const f=fixture(t,{save:()=>d.promise});const pending=start(()=>f.current.save('a',1,{scope:'site'}).catch(e=>e));
  const previous=f.current.save;f.setClient({});f.setMode('site');assert.notEqual(f.current.save,previous);assert.equal(f.current.isWriting,false);
  f.unmount();await settle(async()=>{d.reject(new Error('old write'));await pending;});
});
test('child layout effects can invoke operations before the parent layout effect',async t=>{
  let current,pending;const d=deferred(),client={};
  const hook=loadHook('useSPFxConfigurationStore','useSPFxConfigurationStore',{
    './useSPFxSPHttpClient':{useSPFxSPHttpClient:()=>({client})},
    './useSPFxPageContext':{useSPFxPageContext:()=>({site:{absoluteUrl:'https://tenant.test/root'},web:{absoluteUrl:'https://tenant.test/root'}})},
    '../services/spfx-configuration-store.service':{createSPFxConfigurationStoreService:()=>({get:()=>d.promise})}
  });
  function Child(){React.useLayoutEffect(()=>{pending=current.get('a');},[]);return React.createElement('span');}
  const ReactDOM=require('react-dom'),{act}=require('react-dom/test-utils');
  const container=document.createElement('div');document.body.appendChild(container);
  function Parent(){current=hook({mode:'site'});return React.createElement(Child);}
  act(()=>ReactDOM.render(React.createElement(Parent),container));
  t.after(()=>{act(()=>ReactDOM.unmountComponentAtNode(container));container.remove();});
  assert.equal(current.isLoading,true);await settle(async()=>{d.resolve(undefined);await pending;});assert.equal(current.isLoading,false);
});

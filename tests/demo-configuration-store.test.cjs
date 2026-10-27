const assert=require('node:assert/strict');
const path=require('node:path');
const test=require('node:test');
const {Simulate}=require('react-dom/test-utils');
const {FluentProvider,webLightTheme}=require('@fluentui/react-components');
const {createHarness}=require('./services-test-harness.cjs');
const panel=path.resolve(__dirname,'../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/ConfigurationPanel.tsx');
function fixture(t,overrides={}){
  const calls=[],data={site:new Map(),tenant:new Map()},client={};
  const context={web:{absoluteUrl:'https://tenant.test/root/child'},site:{absoluteUrl:'https://tenant.test/root'}};
  const boundaries={
    './useSPFxSPHttpClient':{useSPFxSPHttpClient:()=>({client})},'./useSPFxPageContext':{useSPFxPageContext:()=>context},
    './spfx-app-catalog.service':{createSPFxAppCatalogService:()=>({discoverUrl:async()=>{calls.push(['discover']);return 'https://tenant.test/catalog';}})}
  };
  for(const scope of ['tenant','site']) boundaries[`./spfx-${scope}-key-value-store.service`]={
    [`createSPFx${scope==='site'?'Site':'Tenant'}KeyValueStoreService`]:()=>({
      get:async key=>{calls.push([scope,'get',key]);if(overrides.get)return overrides.get();return data[scope].get(key);},
      list:async()=>{calls.push([scope,'list']);if(overrides.list)return overrides.list();return [...data[scope].values()];},
      save:async(key,value,url,description)=>{calls.push([scope,'save',key,value]);if(overrides.save)return overrides.save();data[scope].set(key,{key,value,description,id:7});},
      remove:async key=>{calls.push([scope,'remove',key]);if(overrides.remove)return overrides.remove();data[scope].delete(key);}
    })
  };
  const h=createHarness(boundaries);t.after(h.close);
  boundaries['../shared']=Object.assign({},...['ActionResult','DemoCard','InfoGrid','JsonDetails','StatusBadge'].map(name=>h.load(path.resolve(path.dirname(panel),'../shared',name+'.tsx'))));
  boundaries['@apvee/spfx-react-toolkit']={useSPFxConfigurationStore:h.load('hooks/useSPFxConfigurationStore.ts').useSPFxConfigurationStore,useSPFxPageContext:()=>context};
  window.ResizeObserver=class{observe(){}unobserve(){}disconnect(){}};
  t.mock.method(window,'setTimeout',(fn,delay)=>setTimeout(fn,delay));
  const Panel=h.load(panel).default;
  const view=h.mountComponent(()=>h.React.createElement(FluentProvider,{theme:webLightTheme},h.React.createElement(Panel)));
  const select=(selector,value)=>{const e=view.container.querySelector(selector);assert.ok(e,selector);h.act(()=>Simulate.change(e,{target:{value}}));};
  const input=(label,value)=>{const e=view.container.querySelector(`input[aria-label="${label}"]`);assert.ok(e,label);h.act(()=>Simulate.change(e,{target:{value}}));};
  const run=async()=>h.act(async()=>{const e=[...view.container.querySelectorAll('button')].find(b=>b.textContent==='Run');assert.equal(e.disabled,false);Simulate.click(e);await new Promise(resolve=>setImmediate(resolve));});
  return {...view,calls,data,context,select,input,run,flush:()=>h.act(async()=>{await new Promise(resolve=>setImmediate(resolve));}),rerender:()=>h.act(()=>require('react-dom').render(h.React.createElement(FluentProvider,{theme:webLightTheme},h.React.createElement(Panel)),view.container))};
}
test('configuration demo executes default, override and restoration using the real hook/service facade',async t=>{
  const f=fixture(t);assert.deepEqual(f.calls,[]);
  f.select('[data-demo-variant="ConfigurationPanel"]','save');f.select('[data-config-store-scope]','tenant');f.input('Value','20');await f.run();
  f.select('[data-config-store-scope]','site');f.input('Value','50');await f.run();
  f.select('[data-demo-variant="ConfigurationPanel"]','get');await f.run();assert.match(f.container.textContent,/Source.*site/);assert.match(f.container.textContent,/Selected value.*50/);
  f.select('[data-demo-variant="ConfigurationPanel"]','remove');await f.run();
  f.select('[data-demo-variant="ConfigurationPanel"]','get');await f.run();assert.match(f.container.textContent,/Source.*tenant/);assert.match(f.container.textContent,/Selected value.*20/);
});
for(const action of ['get','list','save','remove'])test('configuration demo '+action+' catches failures without absence/success messages',async t=>{
  const method=action==='list'?'list':action;const f=fixture(t,{[method]:async()=>{throw new Error('Access denied');}});
  f.select('[data-demo-variant="ConfigurationPanel"]',action);await f.run();
  assert.match(f.container.textContent,/Access denied/);assert.doesNotMatch(f.container.textContent,/not found\.|Saved |Removed |Loaded /);
});
test('mode and input changes hide obsolete results and do not automatically dispatch',async t=>{
  const f=fixture(t);await f.run();assert.match(f.container.textContent,/not found\./);const count=f.calls.length;
  f.input('Key','another');assert.doesNotMatch(f.container.textContent,/not found\./);
  f.select('[data-config-store-mode]','site');assert.equal(f.calls.length,count);
  f.select('[data-demo-variant="ConfigurationPanel"]','save');
  assert.equal(f.container.querySelector('[data-config-store-scope] option[value="tenant"]').disabled,true);
});
test('physical inventory is labelled and shows a fetched count without merging stores',async t=>{
  const f=fixture(t);f.data.site.set('a',{key:'a',value:1,id:7,description:undefined});
  f.select('[data-demo-variant="ConfigurationPanel"]','list');await f.run();
  assert.match(f.container.textContent,/Loaded 1 site physical item/);assert.ok(f.calls.every(c=>c[0]==='site'));
});
test('editing a pending read hides its eventual obsolete response',async t=>{
  let resolve;const pending=new Promise(done=>{resolve=done;});
  const f=fixture(t,{get:()=>pending});await f.run();f.input('Key','new-key');
  resolve({key:'old-key',value:'obsolete-value',id:7,description:undefined});await f.flush();
  assert.doesNotMatch(f.container.textContent,/obsolete-value|Found /);
});
test('pending writes disable target edits and invalid JSON fails before a mutation',async t=>{
  let resolve;const pending=new Promise(done=>{resolve=done;});
  const f=fixture(t,{save:()=>pending});f.select('[data-demo-variant="ConfigurationPanel"]','save');
  f.select('[data-config-store-format]','json');f.input('Value','{invalid');await f.run();
  assert.ok(f.calls.every(c=>c[1]!=='save'));assert.doesNotMatch(f.container.textContent,/Saved /);
  f.select('[data-config-store-format]','text');await f.run();
  for(const selector of ['[data-config-store-mode]','[data-demo-variant="ConfigurationPanel"]','[data-config-store-scope]','input[aria-label="Key"]']) assert.equal(f.container.querySelector(selector).disabled,true);
  resolve();await f.flush();assert.match(f.container.textContent,/Saved /);
});

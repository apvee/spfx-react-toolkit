const assert = require('node:assert/strict');
const test = require('node:test');
const { createHarness } = require('./services-test-harness.cjs');
const { createScriptedSPHttpClient } = require('./fixtures/key-value-store-transport.cjs');

const root = 'https://tenant.test/sites/root';
const web = root + '/child';
const catalog = 'https://tenant.test/sites/catalog';
const row = value => ({ key: 'StoredKey', value, description: 'local', id: 7 });
function deferred() { let resolve, reject; const promise = new Promise((a,b) => { resolve=a; reject=b; }); return { promise, resolve, reject }; }
function fixture(t, options = { mode: 'layered', discoveryWebUrl: web, siteCollectionUrl: root }, overrides = {}) {
  const calls = [], records = { tenant: new Map(), site: new Map() };
  const stores = {};
  for (const scope of ['tenant','site']) stores[scope] = {
    get: async (key,url) => { calls.push([scope,'get',key,url]); return records[scope].get(key); },
    list: async url => { calls.push([scope,'list',url]); return [...records[scope].values()]; },
    save: async (key,value,url,description) => {
      calls.push([scope,'save',key,value,url,description]);
      const previous=records[scope].get(key);
      records[scope].set(key,{ key,value,description: description === undefined ? previous?.description : description || undefined,id: previous?.id ?? 7 });
    },
    remove: async (key,url) => { calls.push([scope,'remove',key,url]); records[scope].delete(key); },
    ...overrides[scope]
  };
  let discovery = async () => catalog;
  const h = createHarness({
    './spfx-site-key-value-store.service': { createSPFxSiteKeyValueStoreService: () => stores.site },
    './spfx-tenant-key-value-store.service': { createSPFxTenantKeyValueStoreService: () => stores.tenant },
    './spfx-app-catalog.service': { createSPFxAppCatalogService: (_client,context) => ({ discoverUrl: () => { calls.push(['discover',context.web.absoluteUrl]); return discovery(); } }) }
  });
  t.after(h.close);
  const create = h.load('services/spfx-configuration-store.service.ts').createSPFxConfigurationStoreService;
  return { store:create({},options), create, calls, records, stores, discover: fn => { discovery=fn; } };
}

for (const mode of ['tenant','site','layered']) test(mode+' constructs without I/O and restricts physical scope operations',async t => {
  const options=mode==='site'?{mode,siteCollectionUrl:root}:mode==='tenant'?{mode,discoveryWebUrl:web}:{mode,discoveryWebUrl:web,siteCollectionUrl:root};
  const f=fixture(t,options); assert.deepEqual(f.calls,[]);
  for (const scope of ['tenant','site']) {
    if (mode!=='layered' && scope!==mode) {
      for (const run of [()=>f.store.listScope(scope),()=>f.store.save('a',1,{scope}),()=>f.store.remove('a',{scope})]) await assert.rejects(run,/scope/i);
      continue;
    }
    await f.store.save('a',1,{scope});
    assert.equal((await f.store.listScope(scope))[0].source,scope);
    await f.store.remove('a',{scope});
    assert.equal(f.records[scope].size,0);
  }
  if(mode==='site') assert.ok(f.calls.every(call=>call[0]==='site'));
});
for (const value of [false,0,'',[],{},null,undefined,{nested:{y:2}}]) test('layered site presence wins for '+JSON.stringify(value),async t=>{
  const f=fixture(t); const original=row(value); f.records.site.set('key',original);
  f.records.tenant.set('key',{key:'tenant-key',value:{a:1,nested:{x:1}},description:'tenant',id:7});
  assert.deepEqual(await f.store.get('key'),{key:'StoredKey',value,description:'local',id:7,source:'site'});
  assert.equal('source' in original,false);
  assert.deepEqual(f.calls,[['site','get','key',root]]);
});
test('layered falls through only confirmed absence and retains winner casing and metadata',async t=>{
  const f=fixture(t); f.records.tenant.set('key',{key:'FEATURE',value:20,description:undefined,id:9});
  assert.deepEqual(await f.store.get('key'),{key:'FEATURE',value:20,description:undefined,id:9,source:'tenant'});
  f.records.tenant.clear(); assert.equal(await f.store.get('key'),undefined);
});
for (const message of ['403 denied','500 failed','invalid schema','duplicate key']) test('site failure does not enable tenant fallback: '+message,async t=>{
  const error=new Error(message); const f=fixture(t,undefined,{site:{get:async()=>{throw error;}}});
  await assert.rejects(f.store.get('key'),e=>e===error); assert.deepEqual(f.calls,[]);
});
test('single-scope get never consults another store',async t=>{
  for(const mode of ['site','tenant']) {
    const f=fixture(t,mode==='site'?{mode,siteCollectionUrl:root}:{mode,discoveryWebUrl:web});
    f.records[mode].set('key',row(1)); assert.equal((await f.store.get('key')).source,mode);
    assert.ok(f.calls.every(call=>call[0]===mode || (mode==='tenant'&&call[0]==='discover')));
  }
});
test('all invalid targets, keys and descriptions fail before discovery or requests',async t=>{
  const f=fixture(t);
  for(const run of [()=>f.store.get(1),()=>f.store.save(1,2,{scope:'tenant'}),()=>f.store.remove(1,{scope:'site'}),
    ()=>f.store.save('a',1),()=>f.store.remove('a'),()=>f.store.listScope('other'),
    ()=>f.store.save('a',1,{scope:'other'}),()=>f.store.save('a',1,{scope:'tenant',description:1})]) await assert.rejects(run);
  assert.deepEqual(f.calls,[]);
  for(const options of [{mode:'other'},{mode:'site',siteCollectionUrl:''},{mode:'site',siteCollectionUrl:'///'},{mode:'tenant',discoveryWebUrl:'  '},{mode:'layered',discoveryWebUrl:web}]) assert.throws(()=>f.create({},options));
});
test('destinations are snapshotted before await and descriptions stay destination-local',async t=>{
  const options={mode:'layered',discoveryWebUrl:' '+web+'/// ',siteCollectionUrl:' '+root+'/// '};
  const f=fixture(t,options), d=deferred(); f.discover(()=>d.promise);
  f.records.tenant.set('a',{...row(20),description:'tenant'});
  const target={scope:'tenant',description:'original'};
  const pending=f.store.save('a',30,target); target.scope='site'; target.description='changed';
  options.discoveryWebUrl='https://other.test'; options.siteCollectionUrl='https://other.test';
  d.resolve(catalog); await pending;
  assert.deepEqual(f.calls.at(-1),['tenant','save','a',30,catalog,'original']);
  await f.store.save('a',50,{scope:'site'}); assert.equal(f.records.site.get('a').description,undefined);
  await f.store.save('a',60,{scope:'site',description:'site'});
  await f.store.save('a',70,{scope:'site'}); assert.equal(f.records.site.get('a').description,'site');
  await f.store.save('a',80,{scope:'site',description:''}); assert.equal(f.records.site.get('a').description,undefined);
  assert.equal(f.calls.find(call=>call[0]==='discover')[1],web);
});
test('removing site reveals tenant and removing tenant leaves site untouched',async t=>{
  const f=fixture(t); await f.store.save('a',20,{scope:'tenant'}); await f.store.save('a',50,{scope:'site'});
  await f.store.remove('a',{scope:'site'}); assert.equal((await f.store.get('a')).value,20);
  await f.store.save('a',50,{scope:'site'}); await f.store.remove('a',{scope:'tenant'});
  assert.equal((await f.store.get('a')).value,50);
});
test('concurrent tenant operations share pending discovery and retry after rejection; site remains available',async t=>{
  const f=fixture(t), d=deferred(); f.discover(()=>d.promise);
  const a=f.store.get('a'), b=f.store.listScope('tenant');
  await new Promise(resolve=>setImmediate(resolve));
  const error=new Error('catalog denied'); d.reject(error);
  assert.equal((await Promise.allSettled([a,b])).every(r=>r.status==='rejected'&&r.reason===error),true);
  assert.equal(f.calls.filter(c=>c[0]==='discover').length,1);
  await f.store.save('site',1,{scope:'site'});
  f.discover(async()=>catalog); await f.store.get('a');
  assert.equal(f.calls.filter(c=>c[0]==='discover').length,2);
});
test('physical list preserves casing, duplicates, ordering and scope-local IDs without other reads',async t=>{
  const rows=[{...row(1),key:'Feature'},{...row(2),key:'feature'}];
  const f=fixture(t,undefined,{site:{list:async()=>rows}});
  assert.deepEqual(await f.store.listScope('site'),rows.map(item=>({...item,source:'site'})));
  assert.deepEqual(f.calls,[]);
});
test('list and mutation errors propagate without retry or failover',async t=>{
  const error=new Error('Precondition Failed');
  const f=fixture(t,undefined,{site:{list:async()=>{throw error;},save:async()=>{throw error;},remove:async()=>{throw error;}}});
  for(const run of [()=>f.store.listScope('site'),()=>f.store.save('a',1,{scope:'site'}),()=>f.store.remove('a',{scope:'site'})]) await assert.rejects(run,e=>e===error);
  assert.deepEqual(f.calls,[]);
});
test('real services treat absent lists as read/remove absence without provisioning and discover lazily',async t=>{
  const f=createScriptedSPHttpClient();
  const h=createHarness({'@microsoft/sp-http':{SPHttpClient:{configurations:{v1:{}}}},'@microsoft/sp-page-context':{SPPermission:{}}});t.after(h.close);
  const store=h.load('services/spfx-configuration-store.service.ts').createSPFxConfigurationStoreService(f.client,{mode:'layered',siteCollectionUrl:root,discoveryWebUrl:web});
  const siteList=root+"/_api/web/lists/getByTitle('SiteKeyValueStore')?$select=Id";
  const tenantList=catalog+"/_api/web/lists/getByTitle('TenantKeyValueStore')?$select=Id";
  f.enqueue('GET',siteList,{status:404}); f.enqueue('GET',web+'/_api/SP_TenantSettings_Current',{body:{CorporateCatalogUrl:catalog}});f.enqueue('GET',tenantList,{status:404});
  assert.equal(await store.get('a'),undefined);
  f.enqueue('GET',siteList,{status:404});await store.remove('a',{scope:'site'});
  f.enqueue('GET',tenantList,{status:404});assert.deepEqual(await store.listScope('tenant'),[]);
  assert.ok(f.calls.every(c=>c.method==='GET'));f.assertDrained();
});

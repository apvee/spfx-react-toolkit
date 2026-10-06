const test=require('node:test');
const assert=require('node:assert/strict');
const H=require('/Users/fabiofranzini/.codex/worktrees/spfx-toolkit-monorepo/spfx-react-toolkit/tests/async-test-harness.cjs');
const {loadHook,mount,deferred,start,settle,pnpBoundary,listPage}=H;
test('completed old CRUD must not refetch superseded query',async t=>{
 const write=deferred(), calls=[];
 const A=x=>x, B=x=>x;
 const service={query:async builder=>{calls.push(builder===A?'A':'B');return listPage([builder===A?'A':'B'],1,true)},create:()=>write.promise};
 const hook=loadHook('useSPFxPnPList','useSPFxPnPList',pnpBoundary(()=>service,'list'));
 const h=mount(()=>hook('list',{pageSize:1}),{});t.after(()=>h.unmount());
 await settle(()=>h.current.query(A));
 const p=start(()=>h.current.create({Title:'new'}));
 await settle(()=>h.current.query(B));
 await settle(async()=>{write.resolve(1);await p});
 await settle(()=>new Promise(resolve=>setTimeout(resolve,120)));
 console.log('query calls',calls,'items',h.current.items);
 assert.deepEqual(h.current.items,['B']);
});
test('failed refetch must preserve current successful pagination offset',async t=>{
 let n=0;const skips=[];
 const service={query:async()=>{if(n++)throw new Error('refetch fails');return listPage(['0','1'],2,true)},loadMore:async(_builder,_size,skip)=>{skips.push(skip);return listPage(['2'],3,false)}};
 const hook=loadHook('useSPFxPnPList','useSPFxPnPList',pnpBoundary(()=>service,'list'));
 const h=mount(()=>hook('list',{pageSize:2}),{});t.after(()=>h.unmount());
 await settle(()=>h.current.query());
 await settle(()=>h.current.refetch().catch(()=>{}));
 await settle(()=>h.current.loadMore());
 console.log('pagination offsets after failed refetch',skips);
 assert.deepEqual(skips,[2]);
});

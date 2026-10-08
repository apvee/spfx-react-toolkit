const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./provider-test-harness.cjs');
test('existing container hook observes owner windows, replacement, isolation and unsupported cleanup', () => {
  const h=createHarness(); const original=h.window.ResizeObserver;
  const iframe=document.createElement('iframe');document.body.appendChild(iframe);
  const observations=[];
  function Observer(realm){return class{constructor(fn){this.fn=fn;this.realm=realm;this.disconnected=false;observations.push(this);}observe(target){this.target=target;}disconnect(){this.disconnected=true;}};}
  h.window.ResizeObserver=Observer('host');iframe.contentWindow.ResizeObserver=Observer('iframe');
  const {useSPFxContainerInfo}=h.load('hooks/useSPFxContainerInfo.ts');
  const latest={};
  function Consumer({id}){latest[id]=useSPFxContainerInfo();return h.React.createElement('span',undefined,latest[id].size?.width);}
  try{
    const a=h.instance('a'),b=h.instance('b');b.domElement=iframe.contentDocument.createElement('div');iframe.contentDocument.body.appendChild(b.domElement);
    const fixtureA=h.mount(a,Consumer,{id:'a'}),fixtureB=h.mount(b,Consumer,{id:'b'});
    const active=target=>observations.findLast(o=>o.target===target&&!o.disconnected);
    const oa=active(a.domElement),ob=active(b.domElement);assert.equal(oa.realm,'host');assert.equal(ob.realm,'iframe');
    h.act(()=>oa.fn([{contentRect:{width:160,height:24}}]));h.act(()=>ob.fn([{contentRect:{width:220,height:40}}]));
    assert.deepEqual(latest.a.size,{width:160,height:24});assert.deepEqual(latest.b.size,{width:220,height:40});
    assert.equal(latest.a.element,a.domElement);
    const next=iframe.contentDocument.createElement('div');iframe.contentDocument.body.appendChild(next);fixtureA.render({...a,domElement:next});
    assert.equal(oa.disconnected,true);const nextObserver=active(next);assert.equal(nextObserver.realm,'iframe');
    h.act(()=>nextObserver.fn([{contentRect:{width:300,height:60}}]));assert.deepEqual(latest.a.size,{width:300,height:60});assert.deepEqual(latest.b.size,{width:220,height:40});
    delete iframe.contentWindow.ResizeObserver;
    fixtureA.render({...a,domElement:iframe.contentDocument.createElement('div')});assert.equal(nextObserver.disconnected,true);assert.equal(latest.a.size,undefined);
    fixtureB.unmount();assert.equal(ob.disconnected,true);
  }finally{h.close();h.window.ResizeObserver=original;iframe.remove();}
});

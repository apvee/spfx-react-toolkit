const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let dom, React, ReactDOM, act, useResizeObserver;
test.before(async()=>{ dom=harness.createSxDom(); React=require('react'); ReactDOM=require('react-dom'); ({act}=require('react-dom/test-utils')); ({useResizeObserver}=await harness.loadModule('packages/spfx-react-toolkit/src/utils/resize-observer.internal.ts')); });
test.after(async()=>{ dom.window.close(); await harness.cleanup(); });
test('resize subscription uses element owner window, delivers observations and disconnects',()=>{
 const host=document.createElement('div'); document.body.appendChild(host);
 const iframe=document.createElement('iframe'); document.body.appendChild(iframe);
 const element=iframe.contentDocument.createElement('div'); iframe.contentDocument.body.appendChild(element);
 let callback, observed, disconnected=0; const sizes=[];
 iframe.contentWindow.ResizeObserver=class { constructor(fn){callback=fn;} observe(target){observed=target;} disconnect(){disconnected++;} };
 const onResize=size=>sizes.push(size);
 function Consumer({target}){useResizeObserver(target,onResize); return null;}
 try { act(()=>{ReactDOM.render(React.createElement(Consumer,{target:element}),host);}); assert.equal(observed,element); assert.equal(sizes.length,0);
 act(()=>callback([{contentRect:{width:160,height:24}}])); assert.deepEqual(sizes,[{width:160,height:24}]);
 act(()=>{ReactDOM.render(React.createElement(Consumer,{}),host);}); assert.equal(disconnected,1); assert.equal(sizes.at(-1),undefined);
 }finally{act(()=>{ReactDOM.unmountComponentAtNode(host);});host.remove();iframe.remove();}
});

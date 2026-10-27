const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const harness = require('./fixtures/formatted-text-harness.cjs');
const entry = path.resolve(__dirname, '../packages/spfx-react-toolkit/src/controls/auto-skeleton/auto-skeleton.tsx');
let React, ReactDOM, act, AutoSkeleton, dom;
test.before(async () => {
  assert.ok(fs.existsSync(entry), 'AutoSkeleton implementation missing');
  dom = harness.createSxDom();
  for (const name of ['Node','Element','HTMLElement','NodeFilter','MutationObserver','Event']) global[name] = window[name];
  global.getComputedStyle = window.getComputedStyle.bind(window);
  document.createRange = () => ({ selectNodeContents() {}, getClientRects() { return []; } });
  ({ AutoSkeleton } = await harness.loadModule(entry));
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
});
test.after(async () => { if (dom) dom.window.close(); await harness.cleanup(); });
test('empty loading preserves content and native root contract', () => {
  const host = document.createElement('div'); document.body.appendChild(host);
  const ref = React.createRef();
  act(() => { ReactDOM.render(React.createElement(AutoSkeleton, { loading: true, ref, id: 'profile', className: 'consumer' }, 'Unmeasurable'), host); });
  try {
    assert.equal(ref.current.id, 'profile'); assert.equal(ref.current.getAttribute('aria-busy'), 'true');
    assert.equal(host.textContent, 'Unmeasurable'); assert.equal(host.querySelector('.fui-Skeleton'), null);
    assert.equal(host.querySelector('[data-auto-skeleton-content]').hasAttribute('inert'), false);
    assert.equal(ref.current.classList.contains('consumer'), true); assert.equal(ref.current.hasAttribute('style'), false);
  } finally { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); }
});
function rect(x, y, width, height) { return { x, y, left: x, top: y, right: x + width, bottom: y + height, width, height, toJSON() {} }; }
function mount(props, children) {
  const host = document.createElement('div'); document.body.appendChild(host);
  const render = (next = props, nextChildren = children) => { props = next; children = nextChildren; act(() => { ReactDOM.render(React.createElement(AutoSkeleton, props, children), host); }); };
  try { render(); } catch (error) { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); throw error; }
  return { host, render, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
test('real component keeps children mounted, circle/rectangle items, native inert and empty-scan clearing', () => {
  const prototype = window.HTMLElement.prototype;
  const originalRect = prototype.getBoundingClientRect;
  const inert = Object.getOwnPropertyDescriptor(prototype, 'inert');
  Object.defineProperty(prototype, 'inert', { configurable: true, get() { return this.hasAttribute('inert'); } });
  prototype.getBoundingClientRect = function () {
    const shape = this.getAttribute('data-skeleton-shape');
    return shape === 'circle' ? rect(20, 10, 40, 40) : shape === 'rectangle' ? rect(70, 15, 160, 24) : rect(0, 0, 300, 100);
  };
  let mounts = 0, unmounts = 0;
  function Children() { React.useEffect(() => { mounts++; return () => unmounts++; }, []); return React.createElement('section', {},
    React.createElement('button', { 'data-skeleton-shape': 'rectangle', id: 'inner-button' }, 'Action'),
    React.createElement('img', { 'data-skeleton-shape': 'circle', alt: '' })); }
  const f = mount({ loading: false }, React.createElement(Children));
  try {
    const button = f.host.querySelector('button'); button.focus();
    f.render({ loading: true });
    const items = [...f.host.querySelectorAll('[data-auto-skeleton-shape]')];
    assert.deepEqual(items.map(item => item.getAttribute('data-auto-skeleton-shape')), ['rectangle', 'circle']);
    assert.equal(f.host.querySelector('button'), button); assert.equal(mounts, 1); assert.equal(unmounts, 0);
    const content = f.host.querySelector('[data-auto-skeleton-content]');
    assert.equal(content.getAttribute('aria-hidden'), 'true'); assert.equal(content.hasAttribute('inert'), true);
    assert.notEqual(document.activeElement, button);
    assert.equal(f.host.querySelector('[data-auto-skeleton-overlay]').getAttribute('role'), 'presentation');
    for (const item of items) assert.equal(item.hasAttribute('style'), false);
    f.render({ loading: false });
    assert.equal(content.hasAttribute('inert'), false); assert.equal(content.hasAttribute('aria-hidden'), false);
    assert.equal(Boolean(f.host.querySelector('[data-auto-skeleton-overlay]')), false);
    f.render({ loading: true }, false);
    assert.equal(Boolean(f.host.querySelector('[data-auto-skeleton-overlay]')), false); assert.equal(content.hasAttribute('inert'), false);
  } finally { f.close(); prototype.getBoundingClientRect = originalRect; if (inert) Object.defineProperty(prototype, 'inert', inert); else delete prototype.inert; }
});
test('invalid scan configuration reports errors instead of hiding content', () => {
  const original = console.error; console.error = () => {};
  try { for (const props of [{ minSize: NaN }, { minSize: 0 }, { minSize: Infinity }, { ignoreSelectors: ['['] }]) assert.throws(() => mount({ loading: true, ...props }, React.createElement('p', {}, 'Content'))); }
  finally { console.error = original; }
});
test('layout notifications coalesce, refresh during loading and cancel on unmount', () => {
  const prototype = window.HTMLElement.prototype;
  const originalRect = prototype.getBoundingClientRect;
  const original = { request: window.requestAnimationFrame, cancel: window.cancelAnimationFrame, mutation: window.MutationObserver, resize: window.ResizeObserver };
  const inert = Object.getOwnPropertyDescriptor(prototype, 'inert');
  Object.defineProperty(prototype, 'inert', { configurable:true, get(){ return this.hasAttribute('inert'); } });
  let width = 160, measurements = 0, token = 0, disconnected = 0;
  const frames = new Map(), mutations = [], resizes = [];
  window.requestAnimationFrame = fn => { frames.set(++token,fn); return token; };
  window.cancelAnimationFrame = id => frames.delete(id);
  window.MutationObserver = class { constructor(fn){ this.fn=fn; mutations.push(this); } observe(target){this.target=target;} disconnect(){disconnected++;} };
  window.ResizeObserver = class { constructor(fn){this.fn=fn; resizes.push(this);} observe(target){this.target=target;} disconnect(){disconnected++;} };
  prototype.getBoundingClientRect = function(){ if(this.hasAttribute('data-skeleton-shape')){ measurements++; return rect(10,10,width,24); } return rect(0,0,300,100); };
  const child = React.createElement('button',{'data-skeleton-shape':'rectangle'},'Action');
  const f = mount({loading:true},child);
  try {
    const baseline=measurements, item=f.host.querySelector('[data-auto-skeleton-shape]'); assert.ok(item);
    width=220;
    for(let i=0;i<8;i++) window.dispatchEvent(new window.Event('resize'));
    const mutation=mutations[0]; mutation.fn([{target:f.host.querySelector('button'),type:'attributes',attributeName:'class'}]);
    assert.equal(frames.size,1); assert.equal(measurements,baseline);
    const [id,fn]=frames.entries().next().value; frames.delete(id); act(()=>{fn(0);});
    assert.equal(measurements,baseline+1); assert.equal(window.getComputedStyle(f.host.querySelector('[data-auto-skeleton-shape]')).width,'220px');
    assert.equal(f.host.querySelector('[data-auto-skeleton-content]').hasAttribute('inert'),true);
    mutation.fn([{target:mutation.target,type:'attributes',attributeName:'class'}]); assert.equal(frames.size,0);
    window.dispatchEvent(new window.Event('resize')); assert.equal(frames.size,1);
    f.close(); assert.equal(frames.size,0); assert.equal(disconnected,2);
  } finally {
    if(f.host.isConnected) f.close(); prototype.getBoundingClientRect=originalRect;
    if(inert) Object.defineProperty(prototype,'inert',inert); else delete prototype.inert;
    window.requestAnimationFrame=original.request; window.cancelAnimationFrame=original.cancel; window.MutationObserver=original.mutation; window.ResizeObserver=original.resize;
  }
});
test('unsupported inert leaves measurable loading content accessible', () => {
  const prototype=window.HTMLElement.prototype, original=prototype.getBoundingClientRect;
  prototype.getBoundingClientRect=function(){return rect(0,0,160,24);};
  const f=mount({loading:true,style:{color:'red'},dangerouslySetInnerHTML:{__html:'Injected'}},React.createElement('button',{'data-skeleton-shape':'rectangle'},'Action'));
  try {assert.equal(f.host.querySelector('[data-auto-skeleton-overlay]'),null);assert.equal(f.host.querySelector('[data-auto-skeleton-content]').hasAttribute('aria-hidden'),false);assert.equal(f.host.textContent,'Action');assert.equal(f.host.firstChild.style.color,'red');}
  finally{f.close();prototype.getBoundingClientRect=original;}
});

const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const { pilotModules, button } = require('./fixtures/executable-examples/pilot-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const { createDOMRenderer } = require('@griffel/core');
const { RendererProvider } = require('@griffel/react');
const ReactDOM = require('react-dom');
function fixture(t, title, hostDirection = 'ltr') {
  const h = createHarness({ '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 } } });
  t.after(h.close);
  const modules = pilotModules(h);
  const { Demo } = modules.scenario(title);
  const renderer = createDOMRenderer(null);
  let revision = 7, direction = hostDirection;
  const render = regions => h.React.createElement('div', {}, regions.controls, regions.commands, regions.status, regions.children, regions.diagnostics);
  function Root() { return h.React.createElement(RendererProvider, { renderer }, h.React.createElement(FluentProvider, { theme: webLightTheme, dir: direction }, h.React.createElement(Demo, { resetRevision: revision, render }))); }
  const view = h.mountComponent(Root);
  const update = () => h.act(() => { ReactDOM.render(h.React.createElement(Root), view.container); });
  const node = selector => { const result = view.container.querySelector(selector); assert.ok(result, selector); return result; };
  const click = text => h.act(() => Simulate.click(button(view.container, text)));
  const rules = element => {
    const classes = new Set(element.className.split(/\s+/));
    return Object.keys(renderer.insertionCache).filter(rule => [...rule.matchAll(/\.([a-zA-Z0-9_-]+)/g)].some(match => classes.has(match[1]))).join('\n');
  };
  return { h, view, renderer, node, click, rules, reset() { revision++; update(); },
    direction(value) { direction = value; update(); },
    select(selector, value) { h.act(() => Simulate.change(node(selector), { target: { value } })); },
    check(id, checked) { const input = node(`[data-pagination-demo="${id}"]`); h.act(() => { input.checked = checked; Simulate.change(input); }); },
    rows: () => [...view.container.querySelectorAll('[data-order-id]')].map(item => item.textContent),
  };
}
test('source-owned advanced Pagination covers actual mode rows and consumer resets', t => {
  const f = fixture(t, 'Browse orders across pagination modes');
  assert.deepEqual(f.rows(), ['Order 01 · Hardware','Order 02 · Software','Order 03 · Hardware','Order 04 · Software','Order 05 · Hardware']);
  const select = (name, value) => f.select(`select[data-pagination-demo="${name}"]`, value);
  for (const mode of ['known','unknown','cursor','loadMore']) {
    select('mode', mode);
    assert.equal(f.rows().length, 5);
    f.click(mode === 'loadMore' ? 'Load additional orders' : 'Next page');
    assert.equal(f.rows().length, mode === 'loadMore' ? 10 : 5);
    assert.equal(f.rows()[0], mode === 'loadMore' ? 'Order 01 · Hardware' : 'Order 06 · Software');
    const request = JSON.parse(f.node('[data-pagination-demo="request"]').textContent.replace('Last request: ', ''));
    assert.deepEqual(request, mode === 'known' || mode === 'unknown' ? { type: 'pages', pageIndex: 1, pageSize: 5 } : mode === 'cursor' ? { type: 'cursor', direction: 'next' } : { type: 'loadMore' });
  }
  select('page-size','3'); assert.equal(f.rows().length,3);
  select('filter','Hardware'); assert.equal(f.rows()[0],'Order 01 · Hardware');
  f.check('reduce-total', true); assert.equal(f.rows().length,3); assert.ok(button(f.view.container,'Load additional orders').disabled);
  select('filter','missing'); assert.deepEqual(f.rows(),[]); assert.match(f.view.container.textContent,/No matching orders/);
  const nav = f.node('nav'); f.reset(); assert.equal(f.node('nav'),nav); assert.equal(f.rows().length,5);
  assert.equal(f.node('select[data-pagination-demo="mode"]').value,'known');
});
test('source-owned import adapters retain callback state and mixed descriptor composition', t => {
  const f = fixture(t, 'Compare root domain and historical imports');
  for (const adapter of ['root','domain','legacy']) {
    f.select('select[data-demo-variant="imports"]',adapter);
    for (const other of ['root','domain','legacy'].filter(value => value !== adapter)) assert.equal(f.view.container.querySelector(`[data-imports-demo="${other}"]`),null);
    f.click('Increment counter'); f.click('Increment counter'); f.click('Invoke captured callback');
    assert.match(f.view.container.textContent,/Observed counter: 2/); assert.match(f.view.container.textContent,/Same callback identity: yes/);
    const preview = f.node(`[data-imports-demo="${adapter}-preview"]`);
    const baseline = preview.className;
    f.click('Apply alternate width'); assert.notEqual(preview.className,baseline);
    f.click('Remove alternate width'); assert.equal(preview.className,baseline);
    const retained = button(f.view.container,'Invoke captured callback');
    f.reset(); assert.equal(button(f.view.container,'Invoke captured callback'),retained);
    f.click('Invoke captured callback'); assert.match(f.view.container.textContent,/Observed counter: 0/); assert.match(f.view.container.textContent,/Same callback identity: yes/);
  }
});
test('advanced Pagination owns exact-once pending completion, reset invalidation and native presentation', t => {
  const f = fixture(t, 'Browse orders across pagination modes');
  const nav = f.node('nav');
  f.check('manual-loading',true); f.click('Next page');
  assert.equal(f.rows()[0],'Order 01 · Hardware'); assert.equal(nav.getAttribute('aria-busy'),'true');
  assert.ok([...nav.querySelectorAll('button')].every(node => node.disabled));
  f.click('Inspect navigation ref'); assert.match(f.node('[data-pagination-demo="ref"]').textContent,/NAV; hidden=false; busy=true/);
  const complete = button(f.view.container,'Complete loading');
  f.h.act(() => { Simulate.click(complete); Simulate.click(complete); });
  assert.equal(f.rows()[0],'Order 06 · Software'); assert.ok(complete.disabled);
  f.click('Next page'); f.select('select[data-pagination-demo="filter"]','Software');
  f.h.act(() => Simulate.click(complete)); assert.equal(f.rows()[0],'Order 02 · Software');
  f.click('Next page'); f.reset(); f.h.act(() => Simulate.click(complete));
  assert.equal(f.rows()[0],'Order 01 · Hardware'); assert.equal(f.node('nav'),nav);
  f.select('select[data-pagination-demo="size"]','small');
  f.select('select[data-pagination-demo="appearance"]','outline');
  assert.ok([...nav.querySelectorAll('button')].every(button => button.type === 'button'));
  f.check('narrow',true); assert.match(f.rules(nav.parentElement),/320px/);
  f.check('hidden',true); f.click('Inspect navigation ref');
  assert.equal(nav.hidden,true); assert.match(f.node('[data-pagination-demo="ref"]').textContent,/NAV; hidden=true; busy=absent/);
  f.check('hidden',false); f.select('select[data-pagination-demo="mode"]','loadMore');
  assert.equal(button(f.view.container,'Load additional orders').disabled,false);
  f.check('localization',false); assert.equal(button(f.view.container,'Load more').disabled,false);
  f.check('rtl',true); assert.equal(f.node('[data-pagination-demo="provider"]').getAttribute('dir'),'rtl');
  f.direction('rtl'); f.check('rtl',false); assert.equal(f.node('[data-pagination-demo="provider"]').getAttribute('dir'),'rtl');
  f.direction('ltr'); assert.equal(f.node('[data-pagination-demo="provider"]').getAttribute('dir'),'ltr');
});
test('interoperability preserves real typography, ordered widths, scroll roots and host or explicit direction', t => {
  const f = fixture(t, 'Compare root domain and historical imports');
  const classes = [];
  for (const adapter of ['root','domain','legacy']) {
    f.select('select[data-demo-variant="imports"]',adapter);
    const preview = f.node(`[data-imports-demo="${adapter}-preview"]`);
    const scroll = f.node(`[data-imports-demo="${adapter}-scroll"]`);
    classes.push(preview.className);
    assert.match(f.rules(preview),/240px/); assert.match(f.rules(preview),/font-size/); assert.match(f.rules(preview),/padding-inline-start/);
    assert.equal(scroll.getAttribute('role'),'region'); assert.equal(scroll.tabIndex,0); assert.equal(preview.parentElement,scroll);
    assert.match(f.rules(scroll),/100%/); assert.match(f.rules(scroll),/overflow-x/); assert.match(f.rules(scroll),/auto/);
    f.click('Apply alternate width'); assert.match(f.rules(preview),/360px/); assert.doesNotMatch(f.rules(preview),/240px/);
    f.click('Remove alternate width'); assert.match(f.rules(preview),/240px/); assert.doesNotMatch(f.rules(preview),/360px/);
  }
  assert.equal(new Set(classes).size,1,'all adapters use the same shared style implementation');
  const preview = f.node('[data-imports-demo="legacy-preview"]');
  const ltr = preview.className;
  f.click('Override direction to RTL'); assert.equal(f.node('[data-imports-demo="provider"]').getAttribute('dir'),'rtl'); assert.notEqual(preview.className,ltr);
  f.direction('rtl'); f.click('Use host direction'); assert.equal(f.node('[data-imports-demo="provider"]').getAttribute('dir'),'rtl');
  f.direction('ltr'); assert.equal(preview.className,ltr);
  f.click('Increment counter'); f.direction('rtl'); f.click('Invoke captured callback');
  assert.match(f.view.container.textContent,/Observed counter: 1/); assert.match(f.view.container.textContent,/Same callback identity: yes/);
});

const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, ReactDOM, act, StateMessage, FluentProvider, webLightTheme, webDarkTheme, RendererProvider, renderer, dom;
test.before(async () => {
  dom = harness.createSxDom();
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  ({ FluentProvider, webLightTheme, webDarkTheme } = require('@fluentui/react-components'));
  ({ RendererProvider } = require('@griffel/react'));
  // JSDOM 15 cannot parse container query rules; layout is verified in a real browser.
  renderer = require('@griffel/core').createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  ({ StateMessage } = await harness.loadModule('packages/spfx-react-toolkit/src/controls/state-message/state-message.tsx'));
});
test.after(async () => { dom?.window.close(); await harness.cleanup(); });
function mount(props, theme = webLightTheme) {
  const host = document.createElement('div'); document.body.appendChild(host);
  const render = (next = props, nextTheme = theme) => {
    props = next; theme = nextTheme;
    act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer },
      React.createElement(FluentProvider, { theme }, React.createElement(StateMessage, props))), host); });
  };
  render();
  return { host, render, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
test('title remains literal h3 while description composes existing Markdown and native root', () => {
  const ref = React.createRef();
  const f = mount({ title: '**No results**', description: 'Try **filters** and [help](/help)', ref, id: 'state', className: 'consumer' });
  try {
    assert.equal(f.host.querySelector('h3').textContent, '**No results**');
    assert.equal(f.host.querySelector('h3 strong'), null);
    assert.equal(f.host.querySelector('p strong').textContent, 'filters');
    assert.equal(f.host.querySelector('p a').getAttribute('href'), '/help');
    assert.equal(ref.current.id, 'state'); assert.equal(ref.current.tagName, 'DIV');
    assert.ok(ref.current.classList.contains('consumer')); assert.equal(ref.current.hasAttribute('style'), false);
    assert.equal(f.host.querySelector('svg'), null);
    f.render({ title: 'Heading', headingLevel: 2, description: React.createElement('em', { id: 'owned-jsx' }, '**literal**') });
    assert.equal(f.host.querySelector('h2').textContent, 'Heading');
    assert.equal(f.host.querySelector('#owned-jsx').textContent, '**literal**');
  } finally { f.close(); }
});
test('preset roles contain only the text and alert requires explicit selection', () => {
  for (const preset of ['generic','empty','error','success','welcome']) {
    const f = mount({ preset, title: preset, primaryAction: { label: 'Retry', onClick() {} } });
    try {
      const region = f.host.querySelector('[role="status"]');
      assert.equal(Boolean(region), ['empty','error','success'].includes(preset));
      if (region) { assert.equal(region.contains(f.host.querySelector('button')), false); assert.equal(region.contains(f.host.querySelector('svg')), false); }
      assert.equal(f.host.querySelector('[role="alert"]'), null);
      f.render({ preset, title: preset, role: 'alert', visual: false });
      assert.ok(f.host.querySelector('[role="alert"] h3'));
      assert.equal(f.host.querySelector('svg'), null);
      assert.equal(f.host.querySelector('[data-state-message-visual]'), null);
    } finally { f.close(); }
  }
});
test('custom visual overrides artwork and keeps its own accessible semantics', () => {
  const f = mount({ title: 'Welcome', preset: 'welcome', visual: React.createElement('span', { role: 'img', 'aria-label': 'Custom artwork' }, 'art') });
  try {
    assert.equal(f.host.querySelector('svg'), null);
    assert.equal(f.host.querySelector('[role="img"]').getAttribute('aria-hidden'), null);
    f.render({ title: 'No image', preset: 'empty', visual: true });
    assert.equal(f.host.querySelector('[data-state-message-visual]'), null);
  } finally { f.close(); }
});
test('actions preserve button/link semantics, order and callbacks without submitting', () => {
  let clicks = 0, rootClicks = 0;
  const f = mount({ title: 'Actions', onClick: () => rootClicks++,
    primaryAction: { label: 'Retry', onClick: () => clicks++ },
    secondaryAction: { label: 'Disabled', disabled: true, onClick: () => clicks++ },
    tertiaryAction: { label: 'Help', href: '/help', target: '_self' } });
  try {
    const actions = f.host.querySelector('[data-state-message-actions]');
    assert.deepEqual([...actions.children].map(e=>e.textContent), ['Retry','Disabled','Help']);
    const button=actions.querySelector('button'); assert.equal(button.type,'button');
    act(()=>button.click()); assert.equal(clicks,1); assert.equal(rootClicks,1);
    assert.equal(actions.querySelectorAll('button')[1].disabled,true);
    assert.equal(actions.querySelector('a').getAttribute('href'),'/help');
    assert.equal(actions.querySelector('a').getAttribute('target'),'_self');
  } finally { f.close(); }
});
test('owned content ignores raw HTML input and stays mounted during theme/preset updates', () => {
  const f = mount({ title: 'Safe', children: 'unexpected', dangerouslySetInnerHTML: { __html: '<img id="injected">' },
    primaryAction: { label: 'Retry', onClick() {}, dangerouslySetInnerHTML: { __html: 'injected action' } } });
  try {
    assert.equal(f.host.querySelector('#injected'),null); assert.equal(f.host.textContent.includes('unexpected'),false);
    const heading=f.host.querySelector('h3'), button=f.host.querySelector('button');
    button.focus();
    f.render({ title: 'Safe', preset: 'success', primaryAction: { label: 'Retry', onClick() {} } },webDarkTheme);
    assert.equal(f.host.querySelector('h3'),heading); assert.equal(f.host.querySelector('button'),button);
    assert.equal(document.activeElement,button); assert.equal(button.textContent,'Retry');
    assert.equal(f.host.querySelector('svg').getAttribute('aria-hidden'),'true');
    assert.equal(f.host.querySelector('svg').getAttribute('focusable'),'false');
  } finally { f.close(); }
});

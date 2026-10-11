const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, ReactDOM, act, stage, Fluent, renderer, dom;
test.before(async () => {
  process.env.NODE_ENV = 'test';
  dom = harness.createSxDom();
  // Bind real browser constructors needed by the installed native positioning/Tabster code.
  for (const name of ['Node', 'NodeFilter', 'HTMLElement', 'Element', 'Event', 'CustomEvent']) global[name] = dom.window[name];
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  Fluent = require('@fluentui/react-components');
  renderer = require('@griffel/core').createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  stage = await harness.loadModule('tests/fixtures/stage-popover-entry.tsx');
});
test.after(async () => { dom?.window.close(); await harness.cleanup(); });
function mount(element, providerProps = {}) {
  const host = document.createElement('div'); document.body.appendChild(host);
  const { RendererProvider } = require('@griffel/react');
  function render(next = element) {
    element = next;
    act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer },
      React.createElement(Fluent.FluentProvider, { theme: Fluent.webLightTheme, ...providerProps }, element)), host); });
  }
  render();
  return { host, render, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
function tree(options = {}, props = {}) {
  return React.createElement(stage.PopoverOwner, { options: {
    presentation: 'popover', trigger: React.createElement('button', { 'data-trigger': '' }, 'Details'),
    children: React.createElement(React.Fragment, {}, React.createElement(stage.StageHeader,
      { title: '**Literal title**', description: '*Literal description*' }),
      React.createElement(stage.StageBody, {}, React.createElement('input', { 'aria-label': 'Value', defaultValue: 'Initial' }))),
    ...options,
  }, ...props });
}
function surface() { return document.querySelector('.fui-PopoverSurface'); }
function click(element) { act(() => { element.click(); }); }
function escape(element) {
  const event = new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
  act(() => { element.dispatchEvent(event); }); return event;
}
function cssFor(node) {
  const classes = [...node.classList];
  return Array.from(document.styleSheets).flatMap(sheet => Array.from(sheet.cssRules).map(rule => rule.cssText))
    .filter(rule => classes.some(name => rule.includes(`.${name}`))).join('\n');
}

test('native trigger toggles the controlled primitive once with associated nonmodal literal content and refs', () => {
  const triggerRef = React.createRef(), surfaceRef = React.createRef(), requests = []; let consumerClicks = 0;
  const trigger = React.createElement('button', { ref: triggerRef, 'data-trigger': '', 'aria-haspopup': 'dialog', onClick: () => consumerClicks++ }, 'Details');
  const f = mount(tree({ trigger, onOpenChange: (_, data) => requests.push(data) }, {
    surfaceRef, surfaceProps: { id: 'details-surface', className: 'consumer-surface', style: { color: 'red' }, 'data-owner': 'details' },
  }));
  try {
    const button = triggerRef.current;
    assert.equal(surface(), null); assert.equal(button.getAttribute('aria-expanded'), 'false');
    assert.equal(button.type, 'button'); assert.equal(button.hasAttribute('aria-haspopup'), false);
    assert.equal(button.getAttribute('aria-controls'), 'details-surface');
    click(button);
    const root = surface();
    assert.equal(surfaceRef.current, root); assert.equal(root.id, 'details-surface');
    assert.equal(root.getAttribute('role'), 'group'); assert.equal(root.hasAttribute('aria-modal'), false);
    assert.equal(root.dataset.owner, 'details'); assert.ok(root.classList.contains('consumer-surface')); assert.equal(root.style.color, 'red');
    assert.equal(document.getElementById(root.getAttribute('aria-labelledby')).textContent, '**Literal title**');
    assert.equal(document.getElementById(root.getAttribute('aria-describedby')).textContent, '*Literal description*');
    assert.equal(root.querySelector('h2 strong'), null); assert.equal(root.querySelector('.fui-CardHeader__description em'), null);
    assert.equal(button.getAttribute('aria-expanded'), 'true'); assert.equal(triggerRef.current, button);
    click(button); assert.equal(surface(), null); assert.equal(consumerClicks, 2);
    assert.deepEqual(requests, [{ open: true, reason: 'triggerClick' }, { open: false, reason: 'triggerClick' }]);
  } finally { f.close(); }
});

test('Fluent Button refs, native callbacks, disabled and consumer cancellation survive enhancement', () => {
  for (const makeTrigger of [props => React.createElement('button', props, 'Details'), props => React.createElement(Fluent.Button, props, 'Details')]) {
    const ref = React.createRef(), requests = []; let callbacks = 0;
    const render = props => tree({ trigger: makeTrigger({ ref, 'data-trigger': '', onClick: event => { callbacks++; if (props.cancel) event.preventDefault(); },
      disabled: props.disabled }), onOpenChange: (_, data) => requests.push(data) });
    const f = mount(render({ disabled: true }));
    try {
      const button = ref.current; assert.equal(button.tagName, 'BUTTON'); assert.equal(button.disabled, true);
      click(button); assert.equal(surface(), null); assert.equal(callbacks, 0);
      f.render(render({ cancel: true })); click(button); assert.equal(callbacks, 1); assert.equal(surface(), null); assert.deepEqual(requests, []);
      f.render(render({})); click(button); assert.equal(ref.current, button); assert.equal(callbacks, 2);
      assert.deepEqual(requests, [{ open: true, reason: 'triggerClick' }]); assert.ok(surface());
    } finally { f.close(); }
  }
});

test('Stage-owned trigger expanded state overrides conflicting native and Fluent Button values across open changes', () => {
  for (const makeTrigger of [props => React.createElement('button', props, 'Details'), props => React.createElement(Fluent.Button, props, 'Details')]) {
    const ref = React.createRef();
    const render = open => tree({ open, trigger: makeTrigger({ ref, 'aria-expanded': !open }) });
    const f = mount(render(true));
    try {
      const button = ref.current;
      assert.ok(surface()); assert.equal(button.getAttribute('aria-expanded'), 'true');
      f.render(render(false)); assert.equal(surface(), null); assert.equal(button.getAttribute('aria-expanded'), 'false');
      f.render(render(true)); assert.ok(surface()); assert.equal(button.getAttribute('aria-expanded'), 'true'); assert.equal(ref.current, button);
    } finally { f.close(); }
  }
});

test('real inline and portaled Menu consume their first Escape before the ordinary Stage Escape', () => {
  for (const inline of [true, false]) {
    const stageRequests = [], menuRequests = []; let surfaceKeydowns = 0;
    const menu = React.createElement(Fluent.Menu, { inline, defaultOpen: true, onOpenChange: (_, data) => menuRequests.push(data.open) },
      React.createElement(Fluent.MenuTrigger, { disableButtonEnhancement: true }, React.createElement('button', { 'data-menu-trigger': '' }, 'Actions')),
      React.createElement(Fluent.MenuPopover, {}, React.createElement(Fluent.MenuList, {}, React.createElement(Fluent.MenuItem, {}, 'Nested action'))));
    const f = mount(tree({ defaultOpen: true, onOpenChange: (_, data) => stageRequests.push(data),
      children: React.createElement(React.Fragment, {}, React.createElement(stage.StageHeader, { title: 'Actions' }),
        React.createElement(stage.StageBody, {}, menu)) }, { surfaceProps: { onKeyDown: () => surfaceKeydowns++ } }));
    try {
      const root = surface(), item = document.querySelector('[role="menuitem"]');
      assert.ok(item); assert.equal(root.contains(item), inline);
      assert.equal(escape(item).defaultPrevented, true);
      assert.deepEqual(menuRequests, [false]); assert.deepEqual(stageRequests, []); assert.equal(surface(), root);
      assert.equal(document.querySelector('[role="menuitem"]'), null); assert.equal(surfaceKeydowns, 1);
      assert.equal(escape(root.querySelector('[data-menu-trigger]')).defaultPrevented, true);
      assert.equal(surface(), null); assert.deepEqual(stageRequests, [{ open: false, reason: 'escapeKeyDown' }]);
      assert.equal(surfaceKeydowns, 2);
    } finally { f.close(); }
  }
});

test('standard normalizes already-prevented native Escape and outside click once; explicit filters ambient dismissal', () => {
  for (const dismissMode of ['standard', 'explicit']) {
    const requests = [], prevented = [];
    const f = mount(tree({ defaultOpen: true, dismissMode, onOpenChange: (event, data) => { requests.push(data); prevented.push(event.defaultPrevented); } }));
    try {
      const input = surface().querySelector('input');
      assert.equal(escape(input).defaultPrevented, true);
      if (dismissMode === 'standard') {
        assert.equal(surface(), null); assert.deepEqual(requests, [{ open: false, reason: 'escapeKeyDown' }]); assert.deepEqual(prevented, [true]);
        click(f.host.querySelector('[data-trigger]')); requests.length = 0;
        click(document.body); assert.deepEqual(requests, [{ open: false, reason: 'outsideClick' }]); assert.equal(surface(), null);
      } else {
        assert.ok(surface()); click(document.body); assert.ok(surface()); assert.deepEqual(requests, []);
        click(f.host.querySelector('[data-trigger]')); assert.equal(surface(), null);
        assert.deepEqual(requests, [{ open: false, reason: 'triggerClick' }]);
      }
    } finally { f.close(); }
  }
});

test('controlled dirty refusal retains open surface and input until the consumer accepts close', () => {
  const requests = []; const options = { open: true, onOpenChange: (_, data) => requests.push(data) };
  const f = mount(tree(options));
  try {
    const root = surface(), input = root.querySelector('input'); input.value = 'Draft';
    escape(input); click(document.body); click(root.querySelector('[aria-label="Close"]'));
    assert.equal(surface(), root); assert.equal(input.value, 'Draft');
    assert.deepEqual(requests, [{ open: false, reason: 'escapeKeyDown' }, { open: false, reason: 'outsideClick' }, { open: false, reason: 'closeButton' }]);
    f.render(tree({ ...options, open: false })); assert.equal(surface(), null);
  } finally { f.close(); }
});

test('retention begins after first open, hides and inerts the same subtree, and default close resets it', () => {
  for (const unmountOnClose of [false, true]) {
    let mounts = 0, cleanups = 0;
    function Content() {
      React.useEffect(() => { mounts++; return () => cleanups++; }, []);
      return React.createElement('input', { 'aria-label': 'Draft', defaultValue: 'Initial' });
    }
    const options = { unmountOnClose, children: React.createElement(React.Fragment, {}, React.createElement(stage.StageHeader, { title: 'Draft' }),
      React.createElement(stage.StageBody, {}, React.createElement(Content))) };
    const f = mount(tree(options));
    try {
      assert.equal(surface(), null); assert.equal(mounts, 0);
      const trigger = f.host.querySelector('[data-trigger]'); click(trigger);
      const root = surface(), input = root.querySelector('input'); input.value = 'Edited'; assert.equal(mounts, 1);
      f.render(tree({ ...options, size: 'large' })); assert.equal(surface(), root); assert.equal(surface().querySelector('input'), input);
      click(trigger);
      if (!unmountOnClose) {
        assert.equal(surface(), root); assert.equal(root.hidden, true); assert.equal(root.getAttribute('aria-hidden'), 'true'); assert.equal(root.hasAttribute('inert'), true);
        assert.equal(cleanups, 0); click(trigger); assert.equal(surface().querySelector('input'), input); assert.equal(input.value, 'Edited');
        assert.equal(root.hidden, false); assert.equal(root.hasAttribute('inert'), false); assert.equal(root.hasAttribute('aria-hidden'), false); assert.equal(mounts, 1);
      } else {
        assert.equal(surface(), null); assert.equal(cleanups, 1); click(trigger); assert.notEqual(surface().querySelector('input'), input);
        assert.equal(surface().querySelector('input').value, 'Initial'); assert.equal(mounts, 2);
      }
    } finally { f.close(); }
    assert.equal(cleanups, mounts);
  }
});

test('disabling Popover retention after completed close immediately removes retained content and cleans its effects', () => {
  let mounts = 0, cleanups = 0;
  function Content() {
    React.useEffect(() => { mounts++; return () => { cleanups++; }; }, []);
    return React.createElement('input', { defaultValue: 'Initial' });
  }
  const children = React.createElement(React.Fragment, {}, React.createElement(stage.StageHeader, { title: 'Draft' }),
    React.createElement(stage.StageBody, {}, React.createElement(Content)));
  const render = unmountOnClose => tree({ defaultOpen: true, unmountOnClose, children });
  const f = mount(render(false));
  try {
    const root = surface(), input = root.querySelector('input'); input.value = 'Edited';
    click(f.host.querySelector('[data-trigger]')); assert.equal(surface(), root); assert.equal(root.hidden, true); assert.equal(cleanups, 0);
    f.render(render(true));
    assert.equal(surface() === null, true, 'policy-only change after closed exit must release retained native surface');
    assert.equal(cleanups, 1); assert.equal(mounts, 1);
    click(f.host.querySelector('[data-trigger]')); assert.equal(surface().querySelector('input').value, 'Initial'); assert.equal(mounts, 2);
    click(f.host.querySelector('[data-trigger]')); assert.equal(surface(), null); assert.equal(cleanups, 2);
  } finally { f.close(); }
});

test('changing Popover retention while open preserves child DOM and state until its next native close', () => {
  const render = unmountOnClose => tree({ defaultOpen: true, unmountOnClose });
  const f = mount(render(false));
  try {
    const root = surface(), input = root.querySelector('input'); input.value = 'Edited';
    f.render(render(true)); assert.equal(surface(), root); assert.equal(root.querySelector('input'), input); assert.equal(input.value, 'Edited');
    f.render(render(false)); assert.equal(surface(), root); assert.equal(root.querySelector('input'), input);
    click(f.host.querySelector('[data-trigger]')); assert.equal(surface(), root); assert.equal(root.hidden, true);
    f.render(render(true)); assert.equal(surface() === null, true);
  } finally { f.close(); }
});

test('preset widths and padding stay class-based while native positioning retains default or bounded placement and auto-size', async () => {
  for (const [size, pixels] of [[undefined, 240], ['medium', 320], ['large', 480]]) {
    const positioning = size === 'medium' ? { position: 'above', align: 'end' } : undefined; const seen = [];
    const f = mount(React.createElement(Fluent.PositioningConfigurationProvider, { value: data => { seen.push(data.options); return data.options; } },
      tree({ defaultOpen: true, size, positioning })));
    try {
      await act(async () => {});
      const css = cssFor(surface());
      assert.match(css, new RegExp(`width: ${pixels}px`)); assert.match(css, /--apvee-sx-paddingBlockStart-base: 12px;/);
      assert.ok(seen.length > 0); const options = seen[seen.length - 1];
      assert.equal(options.position, positioning?.position ?? 'below'); assert.equal(options.align, positioning?.align ?? 'start');
      assert.equal(options.autoSize, true); assert.notEqual(options.pinned, true);
    } finally { f.close(); }
  }
});

test('native portal inherits Fluent theme, direction and targetDocument and permits an explicit destination', () => {
  const secondaryDom = new (require('jsdom').JSDOM)('<!doctype html><html><head></head><body></body></html>');
  const secondary = secondaryDom.window.document;
  Object.defineProperty(secondaryDom.window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  const theme = { ...Fluent.webDarkTheme, colorNeutralForeground1: '#abcdef' };
  const f = mount(tree({ defaultOpen: true }), { targetDocument: secondary, dir: 'rtl', theme });
  try {
    assert.equal(surface(), null);
    const root = secondary.querySelector('.fui-PopoverSurface'); assert.ok(root); assert.equal(root.ownerDocument, secondary);
    assert.equal(root.parentElement.getAttribute('dir'), 'rtl');
    assert.notEqual(root.parentElement, secondary.body);
    const themeClass = [...root.parentElement.classList].find(name => name.startsWith('fui-FluentProvider'));
    assert.ok(themeClass);
    const css = Array.from(secondary.styleSheets).flatMap(sheet => Array.from(sheet.cssRules).map(rule => rule.cssText)).join('\n');
    assert.match(css, /--colorNeutralForeground1:\s*#abcdef/);
  } finally { f.close(); secondaryDom.window.close(); }
  const destination = document.createElement('div'); destination.id = 'explicit-portal'; document.body.appendChild(destination);
  const explicit = mount(tree({ defaultOpen: true, mountNode: destination }));
  try { assert.equal(surface().parentElement, destination); }
  finally { explicit.close(); destination.remove(); }
});

test('native outside iframe focus maps to focusOutside and is filtered in explicit mode', () => {
  for (const dismissMode of ['standard', 'explicit']) {
    const requests = [], iframe = document.createElement('iframe'); document.body.appendChild(iframe);
    const f = mount(tree({ defaultOpen: true, dismissMode, onOpenChange: (_, data) => requests.push(data) }));
    try {
      // Exercise Fluent's public native iframe-focus event boundary, without claiming real browser focus behavior.
      act(() => { iframe.dispatchEvent(new window.FocusEvent('fuiframefocus', { bubbles: true })); });
      if (dismissMode === 'standard') { assert.equal(surface(), null); assert.deepEqual(requests, [{ open: false, reason: 'focusOutside' }]); }
      else { assert.ok(surface()); assert.deepEqual(requests, []); }
    } finally { f.close(); iframe.remove(); }
  }
});

test('invalid JavaScript Popover full and absent or unsupported trigger fail clearly', () => {
  class Boundary extends React.Component {
    constructor(props) { super(props); this.state = { message: undefined }; }
    static getDerivedStateFromError(error) { return { message: error.message }; }
    render() { return this.state.message ? React.createElement('p', { 'data-error': '' }, this.state.message) : this.props.children; }
  }
  for (const [options, expected] of [[{ size: 'full' }, /Popover.*full/], [{ trigger: undefined }, /Popover.*trigger/],
    [{ trigger: React.createElement('a', { href: '#' }, 'Details') }, /button|Button/]]) {
    const original = console.error; console.error = () => {};
    let f;
    try { f = mount(React.createElement(Boundary, {}, tree(options))); assert.match(f.host.querySelector('[data-error]').textContent, expected); }
    finally { f?.close(); console.error = original; }
  }
});

test('two Popovers preserve independent IDs, state and close callbacks', () => {
  const first = [], second = [];
  const f = mount(React.createElement(React.Fragment, {}, tree({ onOpenChange: (_, data) => first.push(data) }),
    tree({ onOpenChange: (_, data) => second.push(data) })));
  try {
    const triggers = f.host.querySelectorAll('[data-trigger]');
    assert.notEqual(triggers[0].getAttribute('aria-controls'), triggers[1].getAttribute('aria-controls'));
    click(triggers[0]); assert.equal(triggers[1].getAttribute('aria-expanded'), 'false'); assert.deepEqual(second, []);
    click(surface().querySelector('[aria-label="Close"]')); click(triggers[1]);
    assert.deepEqual(first, [{ open: true, reason: 'triggerClick' }, { open: false, reason: 'closeButton' }]);
    assert.deepEqual(second, [{ open: true, reason: 'triggerClick' }]);
    assert.equal(surface().id, triggers[1].getAttribute('aria-controls'));
  } finally { f.close(); }
});

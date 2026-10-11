const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, ReactDOM, act, parts, Fluent, renderer, dom;
test.before(async () => {
  process.env.NODE_ENV = 'test';
  dom = harness.createSxDom();
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  Fluent = require('@fluentui/react-components');
  renderer = require('@griffel/core').createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  parts = await harness.loadModule('tests/fixtures/stage-parts-entry.tsx');
});
test.after(async () => { dom?.window.close(); await harness.cleanup(); });
function mount(element) {
  const host = document.createElement('div'); document.body.appendChild(host);
  const { RendererProvider } = require('@griffel/react');
  function render(next = element) {
    element = next;
    act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer },
      React.createElement(Fluent.FluentProvider, { theme: Fluent.webLightTheme }, element)), host); });
  }
  render();
  return { host, render, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
const controller = { open: true, presentation: 'dialog', size: 'small', requestOpenChange() {} };
function Owner({ children, request = controller.requestOpenChange }) {
  const value = parts.useStageContextValue({ ...controller, requestOpenChange: request });
  return React.createElement(parts.StageContext.Provider, { value },
    React.createElement('div', { 'data-description-present': String(value.hasDescription) }, children));
}
test('header uses literal semantic text, real CardHeader slots and stable caller JSX', () => {
  const ref = React.createRef(); let clicks = 0;
  const title = React.createElement('em', { ref, className: 'caller-title', onClick: () => clicks++ }, '**Title**');
  const tree = description => React.createElement(Owner, {}, React.createElement(parts.StageHeader,
    { title, description, headingLevel: 4, icon: React.createElement('svg', { 'data-icon': '' }) }));
  const f = mount(tree('*Description*'));
  try {
    const child = ref.current, heading = f.host.querySelector('h4');
    assert.equal(heading.firstElementChild, child); assert.equal(child.className, 'caller-title');
    assert.equal(heading.textContent, '**Title**'); assert.equal(heading.querySelector('strong'), null);
    const description = f.host.querySelector('.fui-CardHeader__description');
    assert.equal(description.textContent, '*Description*'); assert.equal(description.querySelector('em'), null);
    assert.ok(heading.parentElement.id); assert.ok(description.id);
    assert.equal(f.host.querySelector('.fui-CardHeader__image').getAttribute('aria-hidden'), 'true');
    assert.equal(f.host.querySelector('.fui-CardHeader__image').style.length, 0);
    assert.equal(f.host.querySelector('[data-description-present]').dataset.descriptionPresent, 'true');
    act(() => child.click()); assert.equal(clicks, 1);
    f.render(tree('New description')); assert.equal(ref.current, child);
    f.render(tree(React.createElement(React.Fragment, {}, false, '')));
    assert.equal(f.host.querySelector('.fui-CardHeader__description'), null);
    assert.equal(f.host.querySelector('[data-description-present]').dataset.descriptionPresent, 'false');
    f.render(tree('*Description*')); assert.equal(ref.current, child);
    f.render(React.createElement(Owner));
    assert.equal(f.host.querySelector('[data-description-present]').dataset.descriptionPresent, 'false');
  } finally { f.close(); }
});
test('close is last, defaults to Close, is localizable and alone requests dismissal', () => {
  const requests = []; let actionClicks = 0, footerClicks = 0;
  const tree = options => React.createElement(Owner, { request: (_, data) => requests.push(data) },
    React.createElement(parts.StageHeader, { title: '**Literal title**', actions: React.createElement('button',
      { id: 'help', onClick: () => actionClicks++ }, 'Help'), ...options }),
    React.createElement(parts.StageBody, {}, 'Body'),
    React.createElement(parts.StageFooter, {}, React.createElement('button', { id: 'save', onClick: () => footerClicks++ }, 'Save')));
  const f = mount(tree({}));
  try {
    assert.equal(f.host.querySelector('h2').textContent, '**Literal title**');
    assert.equal(f.host.querySelector('h2 strong'), null);
    assert.equal(f.host.querySelector('.fui-CardHeader__action').lastElementChild.getAttribute('aria-label'), 'Close');
    act(() => f.host.querySelector('#help').click()); act(() => f.host.querySelector('#save').click());
    assert.equal(actionClicks, 1); assert.equal(footerClicks, 1); assert.deepEqual(requests, []);
    act(() => f.host.querySelector('[aria-label="Close"]').click());
    assert.deepEqual(requests, [{ open: false, reason: 'closeButton' }]);
    f.render(tree({ closeLabel: 'Chiudi' })); assert.ok(f.host.querySelector('[aria-label="Chiudi"]'));
    f.render(tree({ showCloseButton: false })); assert.equal(f.host.querySelector('[aria-label="Chiudi"]'), null);
    assert.equal(f.host.querySelector('.fui-CardHeader__action').children.length, 1);
    f.render(React.createElement(Owner, {}, React.createElement(parts.StageHeader, { title: 'Title', showCloseButton: false }),
      React.createElement(parts.StageBody, {}, 'Body')));
    assert.equal(f.host.querySelector('.fui-CardHeader__action'), null);
    assert.equal(f.host.querySelector('[data-stage-footer]'), null);
  } finally { f.close(); }
});
test('part roots forward native refs/events/hidden/style and discard untyped raw HTML', () => {
  const refs = [React.createRef(), React.createRef(), React.createRef()]; let clicks = 0;
  const f = mount(React.createElement(Owner, {}, ...[parts.StageHeader, parts.StageBody, parts.StageFooter].map((Part, index) =>
    React.createElement(Part, { key: index, ref: refs[index], id: `part-${index}`, hidden: true, className: 'caller',
      style: { color: 'red' }, onClick: () => clicks++, dangerouslySetInnerHTML: { __html: '<img id="injected">' },
      ...(index === 0 ? { title: 'Safe', showCloseButton: false } : { children: 'Safe' }) }))));
  try {
    refs.forEach((ref, index) => {
      assert.equal(ref.current, f.host.querySelector(`#part-${index}`)); assert.equal(ref.current.hidden, true);
      assert.ok(ref.current.classList.contains('caller')); assert.equal(ref.current.style.color, 'red');
      act(() => ref.current.click());
    });
    assert.equal(clicks, 3); assert.equal(f.host.querySelector('#injected'), null);
  } finally { f.close(); }
});
test('Body and Footer retain caller DOM, refs and events across same-part rerenders', () => {
  const inputRef = React.createRef(), actionRef = React.createRef(); let clicks = 0;
  const tree = className => React.createElement(Owner, {},
    React.createElement(parts.StageBody, { className }, React.createElement('input', { ref: inputRef, defaultValue: 'Initial' })),
    React.createElement(parts.StageFooter, { className }, React.createElement('button', { ref: actionRef,
      onClick: () => clicks++ }, 'Save')));
  const f = mount(tree('before'));
  try {
    const input = inputRef.current, action = actionRef.current; input.value = 'Edited';
    f.render(tree('after'));
    assert.equal(inputRef.current, input); assert.equal(input.value, 'Edited'); assert.equal(actionRef.current, action);
    act(() => action.click()); assert.equal(clicks, 1);
  } finally { f.close(); }
});
test('Header activation handlers preserve a neutral Card in hover, pressed and forced-color CSS', () => {
  const ref = React.createRef(), callerRef = React.createRef(); let clicks = 0, mouseDowns = 0, pointerDowns = 0;
  const f = mount(React.createElement(Owner, {}, React.createElement(parts.StageHeader, {
    ref, title: 'Title', onClick: () => clicks++, onMouseDown: () => mouseDowns++, onPointerDown: () => pointerDowns++,
    actions: React.createElement('button', { ref: callerRef, className: 'caller-action', style: { color: 'red' } }, 'Action'),
  })));
  try {
    const header = f.host.querySelector('[data-stage-header]');
    assert.equal(ref.current, header);
    act(() => { header.click(); header.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true }));
      header.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true })); });
    assert.equal(clicks, 1); assert.equal(mouseDowns, 1); assert.equal(pointerDowns, 1);
    assert.equal(callerRef.current.className, 'caller-action'); assert.equal(callerRef.current.style.color, 'red');
    const neutral = f.host.querySelector('.fui-Card');
    const classes = [...new Set([...header.classList, ...neutral.classList])].filter(name => name.startsWith('f'));
    const rules = Array.from(document.styleSheets).flatMap(sheet => Array.from(sheet.cssRules).map(rule => rule.cssText));
    const matching = rules.filter(rule => classes.some(name => rule.includes(`.${name}`))).join('\n');
    assert.doesNotMatch(matching, /cursor:\s*pointer/);
    assert.doesNotMatch(matching, /var\(--colorNeutralForeground1Hover\)/);
    assert.doesNotMatch(matching, /var\(--colorSubtleBackground(?:Hover|Pressed)\)/);
    // Fluent's interactive high-contrast rules also target descendant :active; caller JSX must stay untouched.
    assert.doesNotMatch(matching, /background-color:\s*Highlight[;}]/);
    assert.doesNotMatch(matching, /color:\s*HighlightText[;}]/);
    assert.doesNotMatch(matching, /forced-color-adjust:\s*none[;}]/);
    assert.equal(neutral.getAttribute('role'), 'presentation'); assert.equal(neutral.hasAttribute('tabindex'), false);
  } finally { f.close(); }
});
test('portaled header isolates selectable ancestor Card labels and DOM ref across mount/unmount', () => {
  const ref = React.createRef(), portal = document.createElement('div'); document.body.appendChild(portal);
  const tree = open => React.createElement(Fluent.Card, { ref, selected: true, checkbox: {}, 'aria-label': 'Original card' },
    React.createElement(Fluent.CardHeader, { header: React.createElement('span', { id: 'original-title' }, 'Original title') }),
    open ? ReactDOM.createPortal(React.createElement(Owner, {}, React.createElement(parts.StageHeader, { title: 'Stage title' })), portal) : undefined);
  const f = mount(tree(false));
  try {
    const ancestor = ref.current, checkbox = ancestor.querySelector('input[type="checkbox"]');
    const label = checkbox.getAttribute('aria-labelledby'), accessibleLabel = ancestor.getAttribute('aria-label');
    assert.ok(label);
    f.render(tree(true)); assert.equal(ref.current, ancestor);
    assert.equal(checkbox.getAttribute('aria-labelledby'), label); assert.equal(ancestor.getAttribute('aria-label'), accessibleLabel);
    const neutral = portal.querySelector('.fui-Card');
    assert.equal(neutral.getAttribute('role'), 'presentation'); assert.equal(neutral.hasAttribute('tabindex'), false);
    assert.equal(neutral.hasAttribute('aria-checked'), false); assert.equal(neutral.hasAttribute('aria-selected'), false);
    f.render(tree(false)); assert.equal(ref.current, ancestor);
    assert.equal(checkbox.getAttribute('aria-labelledby'), label); assert.equal(ancestor.getAttribute('aria-label'), accessibleLabel);
  } finally { f.close(); portal.remove(); }
});
test('shared geometry and part layout insert host CSS classes with exact presets and no authored inline style', () => {
  function Probe() {
    const styles = parts.useStageStyles();
    return React.createElement('div', {}, Object.entries(styles).map(([key, className]) => React.createElement('div',
      { key, 'data-style': key, className })));
  }
  const f = mount(React.createElement(Owner, {}, React.createElement(Probe),
    React.createElement(parts.StageHeader, { title: 'Title', description: 'Description', icon: React.createElement('svg') }),
    React.createElement(parts.StageBody, {}, 'Body'), React.createElement(parts.StageFooter, {}, 'Footer')));
  try {
    const rules = Array.from(document.styleSheets).flatMap(sheet => Array.from(sheet.cssRules).map(rule => rule.cssText)).join('\n');
    function cssForNode(node) {
      const classes = node.className.split(/\s+/).filter(Boolean);
      return rules.split('\n').filter(rule => classes.some(name => rule.includes(`.${name}`))).join('\n');
    }
    const cssFor = key => cssForNode(f.host.querySelector(`[data-style="${key}"]`));
    assert.ok(f.host.querySelector('[data-style="backdrop"]'), 'Retained native backdrop requires an owned hidden-state class');
    assert.match(cssFor('backdrop'), /\[hidden\]\s*\{display: none;/);
    assert.match(cssFor('backdrop'), /\[hidden\]\s*\{pointer-events: none;/);
    for (const [key, pixels] of Object.entries({ dialogSmall: 640, dialogMedium: 800, dialogLarge: 1120,
      drawerSmall: 320, drawerMedium: 592, drawerLarge: 940, popoverSmall: 240, popoverMedium: 320, popoverLarge: 480 })) {
      assert.match(cssFor(key), new RegExp(`width: ${pixels}px`));
    }
    assert.match(cssFor('dialogFull'), /100vw/); assert.match(cssFor('dialogFull'), /100vh/);
    assert.match(cssFor('dialogFull'), /100dvh/); assert.match(cssFor('dialogFull'), /clamp\(8px, 1vw, 16px\)/);
    // Fluent's reset-class media query sets max-width:100vw on narrow screens.
    // The composed surface must keep Stage's gap with higher selector specificity.
    assert.match(cssFor('dialog'), /\.fui-DialogSurface\s*\{max-width: calc\(100vw - 2 \* clamp\(8px, 1vw, 16px\)\);/);
    // Descriptor values are inserted into class-scoped CSS variables by the real useSx runtime.
    const descriptor = (key, property, value) => assert.match(cssFor(key), new RegExp(`--apvee-sx-${property}-base: ${value};`));
    descriptor('surface', 'minHeight', '0px'); descriptor('surface', 'flexDirection', 'column');
    descriptor('body', 'flexGrow', '1'); descriptor('body', 'overflowY', 'auto');
    descriptor('header', 'flexShrink', '0'); descriptor('footer', 'flexShrink', '0');
    descriptor('actions', 'flexWrap', 'wrap'); assert.match(cssFor('icon'), /24px/);
    descriptor('neutralCard', 'backgroundColor', 'var\\(--colorTransparentBackground\\)');
    assert.match(cssFor('neutralCard'), /::after/);
    descriptor('text', 'rowGap', 'var\\(--spacingVerticalNone\\)');
    const titleCSS = cssForNode(f.host.querySelector('h2'));
    const descriptionCSS = cssForNode(f.host.querySelector('.fui-CardHeader__description span'));
    assert.match(titleCSS, /--apvee-sx-fontSize-base: var\(--fontSizeBase400\);/);
    assert.match(titleCSS, /--apvee-sx-fontWeight-base: var\(--fontWeightSemibold\);/);
    assert.match(titleCSS, /--apvee-sx-color-base: var\(--colorNeutralForeground1\);/);
    assert.match(descriptionCSS, /--apvee-sx-fontSize-base: var\(--fontSizeBase200\);/);
    assert.match(descriptionCSS, /--apvee-sx-color-base: var\(--colorNeutralForeground2\);/);
    assert.match(titleCSS, /--apvee-sx-marginBlockStart-base: var\(--spacingVerticalNone\);/);
    assert.match(descriptionCSS, /--apvee-sx-marginBlockEnd-base: var\(--spacingVerticalNone\);/);
    descriptor('modalPadding', 'paddingBlockStart', '24px'); descriptor('popoverPadding', 'paddingBlockStart', '12px');
    for (const part of ['header', 'body', 'footer']) descriptor(part, 'paddingBlockStart', 'var\\(--spacingVerticalNone\\)');
    assert.equal(f.host.querySelectorAll('[style]').length, 0);
  } finally { f.close(); }
});

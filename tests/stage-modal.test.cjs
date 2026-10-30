const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, ReactDOM, act, fixture, Fluent, renderer, dom;
test.before(async () => {
  process.env.NODE_ENV = 'test';
  dom = harness.createSxDom();
  // Tabster uses the realm's Node constants when traversing the actual modal DOM.
  global.Node = dom.window.Node;
  global.NodeFilter = dom.window.NodeFilter;
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  Fluent = require('@fluentui/react-components');
  renderer = require('@griffel/core').createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  fixture = await harness.loadModule('tests/fixtures/stage-modal-entry.tsx');
});
test.after(async () => { act(() => {}); dom?.window.close(); await harness.cleanup(); delete global.Node; delete global.NodeFilter; });
function mount(initial, dir = 'ltr', targetDocument = document) {
  let props = initial;
  const host = targetDocument.createElement('div'); targetDocument.body.appendChild(host);
  const activeRenderer = targetDocument === document ? renderer : require('@griffel/core').createDOMRenderer(targetDocument, {
    unstable_filterCSSRule: rule => !rule.startsWith('@container'),
  });
  const { RendererProvider } = require('@griffel/react');
  function render(next = props) {
    props = next;
    act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer: activeRenderer, targetDocument },
      React.createElement(Fluent.FluentProvider, { theme: Fluent.webLightTheme, dir, targetDocument },
        React.createElement(fixture.ModalFixture, props))), host); });
  }
  render();
  return { host, render, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
const options = (presentation, extra = {}) => ({ presentation, children: undefined, ...extra });
function click(node) { act(() => { node.click(); }); }
function escape(node) { act(() => { node.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })); }); }
function surface(id) { return document.getElementById(id); }
function rulesFor(node) {
  return Array.from(document.styleSheets).flatMap(sheet => Array.from(sheet.cssRules).map(rule => rule.cssText))
    .filter(rule => Array.from(node.classList).some(name => rule.includes(`.${name}`))).join('\n');
}
for (const presentation of ['dialog', 'drawer']) {
  test(`${presentation}: native modal semantics, associations, ref, attributes and size updates preserve children`, () => {
    const ref = React.createRef(); let clicks = 0;
    const initial = { options: options(presentation, { defaultOpen: true }), surfaceRef: ref, description: 'Details',
      surfaceProps: { id: `${presentation}-native`, className: 'consumer', style: { color: 'red' }, onClick: () => clicks++,
        role: 'none', 'aria-modal': false, 'aria-labelledby': 'wrong', 'aria-describedby': 'wrong',
        dangerouslySetInnerHTML: { __html: '<img id="injected">' } } };
    const f = mount(initial);
    try {
      const node = surface(`${presentation}-native`), input = node.querySelector('input');
      assert.equal(ref.current, node); assert.equal(node.getAttribute('role'), 'dialog');
      assert.equal(node.getAttribute('aria-modal'), 'true'); assert.ok(node.classList.contains('consumer'));
      assert.equal(node.style.color, 'red'); assert.equal(node.querySelector('#injected'), null);
      assert.equal(document.getElementById(node.getAttribute('aria-labelledby')).textContent, 'Editor');
      assert.equal(document.getElementById(node.getAttribute('aria-describedby')).textContent, 'Details');
      assert.equal(node.hasAttribute('presentation'), false); assert.equal(node.hasAttribute('size'), false);
      assert.equal(node.hasAttribute('dismissMode'), false); click(node); assert.equal(clicks, 1);
      input.value = 'Edited';
      const widths = presentation === 'dialog' ? { small: 640, medium: 800, large: 1120 } : { small: 320, medium: 592, large: 940 };
      for (const [size, pixels] of Object.entries(widths)) {
        f.render({ ...initial, options: options(presentation, { defaultOpen: true, size }), description: undefined });
        assert.equal(ref.current, node); assert.equal(node.querySelector('input'), input); assert.equal(input.value, 'Edited');
        assert.equal(node.hasAttribute('aria-describedby'), false); assert.match(rulesFor(node), new RegExp(`width: ${pixels}px`));
        assert.match(rulesFor(node), /--apvee-sx-paddingBlockStart-base: 24px;/);
      }
      f.render({ ...initial, surfaceProps: { ...initial.surfaceProps, hidden: true }, options: options(presentation, { defaultOpen: true, size: 'full' }) });
      assert.equal(node.hidden, true); assert.match(rulesFor(node), /100vw/);
    } finally { f.close(); }
  });
  test(`${presentation}: trigger preserves native/Fluent refs, events, name and cancellation/disabled`, () => {
    for (const Type of ['button', Fluent.Button]) {
      const ref = React.createRef(); let clicks = 0, mouseDowns = 0;
      const id = `${presentation}-trigger`, requests = [];
      const trigger = extra => React.createElement(Type, { ref, onClick: event => { clicks++; if (extra.cancel) event.preventDefault(); },
        onMouseDown: () => mouseDowns++, disabled: extra.disabled }, 'Open editor');
      const tree = extra => ({ options: options(presentation, { trigger: trigger(extra), onOpenChange: (_, data) => requests.push(data) }), surfaceProps: { id } });
      const f = mount(tree({ cancel: true }));
      try {
        const button = ref.current; assert.equal(button.textContent, 'Open editor');
        assert.equal(button.getAttribute('aria-controls'), id); assert.equal(button.getAttribute('aria-haspopup'), 'dialog');
        assert.equal(button.getAttribute('aria-expanded'), 'false'); click(button); assert.equal(surface(id), null); assert.equal(clicks, 1);
        f.render(tree({ disabled: true })); assert.equal(ref.current, button); click(button); assert.equal(surface(id), null);
        f.render(tree({})); act(() => { button.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true })); });
        assert.equal(mouseDowns, 1); click(button); assert.ok(surface(id)); assert.equal(button.getAttribute('aria-expanded'), 'true');
        assert.deepEqual(requests, [{ open: true, reason: 'triggerClick' }]);
        const focusAttributes = JSON.parse(button.getAttribute('data-tabster'));
        assert.equal(focusAttributes.restorer.type, 1);
      } finally { f.close(); assert.equal(ref.current, null); }
    }
  });
  test(`${presentation}: callback refs/native portals and child effects follow modal lifecycle`, () => {
    for (const unmountOnClose of [true, false]) {
      const portal = document.createElement('div'); document.body.appendChild(portal);
      const id = `${presentation}-effects`, assignments = [], triggerAssignments = [];
      let mounts = 0, cleanups = 0;
      function Stateful() {
        const [count, setCount] = React.useState(0);
        React.useEffect(() => { mounts++; return () => { cleanups++; }; }, []);
        return React.createElement('button', { id: `${id}-count`, onClick: () => setCount(value => value + 1) }, count);
      }
      const tree = open => ({ options: options(presentation, { open, unmountOnClose, mountNode: portal,
        trigger: React.createElement('button', { ref: node => triggerAssignments.push(node) }, 'Open') }),
        surfaceProps: { id }, surfaceRef: node => assignments.push(node), children: React.createElement(Stateful) });
      const f = mount(tree(false));
      try {
        assert.equal(mounts, 0); f.render(tree(true));
        const node = surface(id), count = node.querySelector(`#${id}-count`);
        assert.ok(portal.contains(node)); assert.equal(f.host.contains(node), false); assert.equal(assignments.at(-1), node);
        assert.equal(mounts, 1); assert.equal(JSON.parse(node.getAttribute('data-tabster')).restorer.type, 0);
        click(count); assert.equal(count.textContent, '1'); f.render(tree(false));
        assert.equal(cleanups, unmountOnClose ? 1 : 0);
        f.render(tree(true)); assert.equal(mounts, unmountOnClose ? 2 : 1);
        assert.equal(surface(id).querySelector(`#${id}-count`).textContent, unmountOnClose ? '0' : '1');
      } finally {
        f.close(); portal.remove(); assert.equal(assignments.at(-1), null); assert.equal(triggerAssignments.at(-1), null);
        assert.equal(cleanups, mounts);
      }
    }
  });
  test(`${presentation}: child-consumed Escape stays local; callback prevention does not veto a native close request`, () => {
    const id = `${presentation}-prevented`, requests = [];
    const f = mount({ options: options(presentation, { defaultOpen: true,
      onOpenChange: (event, data) => { event.preventDefault(); requests.push(data); } }),
      surfaceProps: { id }, children: React.createElement('input', { onKeyDown: event => event.preventDefault() }) });
    try {
      escape(surface(id).querySelector('input')); assert.deepEqual(requests, []); assert.ok(surface(id));
      escape(surface(id)); assert.deepEqual(requests, [{ open: false, reason: 'escapeKeyDown' }]); assert.equal(surface(id), null);
    }
    finally { f.close(); }
  });
  test(`${presentation}: standard Escape/backdrop and explicit Close normalize once, controlled refusal remains open`, () => {
    const id = `${presentation}-policy`, requests = [];
    const tree = extra => ({ options: options(presentation, { open: true, onOpenChange: (_, data) => requests.push(data), ...extra }), surfaceProps: { id } });
    const f = mount(tree({}));
    try {
      escape(surface(id)); assert.deepEqual(requests, [{ open: false, reason: 'escapeKeyDown' }]); assert.ok(surface(id));
      const backdrop = document.querySelector(presentation === 'dialog' ? '.fui-DialogSurface__backdrop' : '.fui-OverlayDrawer__backdrop');
      click(backdrop); assert.deepEqual(requests.at(-1), { open: false, reason: 'backdropClick' }); assert.equal(requests.length, 2);
      f.render(tree({ dismissMode: 'explicit' })); escape(surface(id)); click(backdrop); assert.equal(requests.length, 2);
      click(surface(id).querySelector('[aria-label="Close"]')); assert.equal(requests.length, 3);
      assert.deepEqual(requests.at(-1), { open: false, reason: 'closeButton' }); assert.ok(surface(id));
      f.render(tree({ open: false })); assert.equal(surface(id), null);
    } finally { f.close(); }
  });
  test(`${presentation}: default closure resets state; retention starts at first open and closes inert and hidden`, () => {
    for (const unmountOnClose of [true, false]) {
      const id = `${presentation}-retention`, trigger = React.createElement('button', { id: `${id}-button` }, 'Open');
      const initial = { options: options(presentation, { unmountOnClose, trigger }), surfaceProps: { id } };
      const f = mount(initial);
      try {
        assert.equal(surface(id), null); click(document.getElementById(`${id}-button`));
        const node = surface(id), input = node.querySelector('input'), label = node.getAttribute('aria-labelledby'); input.value = 'Edited';
        click(node.querySelector('[aria-label="Close"]'));
        if (unmountOnClose) assert.equal(surface(id), null);
        else {
          assert.equal(surface(id), node); assert.equal(node.getAttribute('aria-hidden'), 'true');
          assert.equal(node.hasAttribute('inert'), true); assert.equal(node.hidden, true); assert.equal(node.tabIndex, -1);
        }
        click(document.getElementById(`${id}-button`));
        const reopened = surface(id); assert.equal(reopened.getAttribute('aria-labelledby'), label);
        assert.equal(reopened.hasAttribute('inert'), false); assert.equal(reopened.hidden, false);
        assert.equal(reopened.querySelector('input').value, unmountOnClose ? 'Initial' : 'Edited');
        if (!unmountOnClose) assert.equal(reopened.querySelector('input'), input);
      } finally { f.close(); }
    }
  });
  test(`${presentation}: closed retained surface remains inert when the caller changes portal destination`, () => {
    const first = document.createElement('div'), second = document.createElement('div');
    document.body.appendChild(first); document.body.appendChild(second);
    const id = `${presentation}-portal-change`;
    const tree = (open, mountNode) => ({ options: options(presentation, { open, unmountOnClose: false, mountNode }), surfaceProps: { id } });
    const f = mount(tree(true, first));
    try {
      f.render(tree(false, first)); f.render(tree(false, second));
      const node = surface(id); assert.ok(second.contains(node));
      assert.equal(node.getAttribute('aria-hidden'), 'true'); assert.equal(node.hidden, true); assert.equal(node.hasAttribute('inert'), true);
    } finally { f.close(); first.remove(); second.remove(); }
  });
  test(`${presentation}: retained content survives native backdrop exit; changing retention while closed cannot leave a blocking backdrop`, () => {
    const id = `${presentation}-backdrop-retention`, trigger = React.createElement('button', { id: `${id}-trigger` }, 'Open');
    const tree = unmountOnClose => ({ options: options(presentation, { trigger, unmountOnClose }), surfaceProps: { id } });
    const backdropSelector = presentation === 'dialog' ? '.fui-DialogSurface__backdrop' : '.fui-OverlayDrawer__backdrop';
    const f = mount(tree(false));
    try {
      click(document.getElementById(`${id}-trigger`));
      const node = surface(id), input = node.querySelector('input'); input.value = 'Edited';
      assert.ok(document.querySelector(backdropSelector));
      click(node.querySelector('[aria-label="Close"]'));
      assert.equal(surface(id), node); assert.equal(node.querySelector('input'), input); assert.equal(input.value, 'Edited');
      assert.equal(node.hidden, true);
      const closedBackdrop = document.querySelector(backdropSelector);
      if (closedBackdrop) {
        assert.equal(closedBackdrop.hidden, true, 'native backdrop exit hides its hit-test target independently of retained content');
        assert.equal(window.getComputedStyle(closedBackdrop).display, 'none');
        assert.equal(window.getComputedStyle(closedBackdrop).pointerEvents, 'none');
      }
      f.render(tree(true));
      assert.equal(document.querySelector(backdropSelector) === null, true, 'changing retention while already closed cannot restore a blocker');
      assert.equal(surface(id) === null, true, 'disabling retention after a completed close releases the retained subtree immediately');
      click(document.getElementById(`${id}-trigger`));
      assert.equal(surface(id) === node, false); assert.equal(surface(id).querySelector('input').value, 'Initial');
      assert.ok(document.querySelector(backdropSelector));
      click(surface(id).querySelector('[aria-label="Close"]'));
      assert.equal(surface(id), null); assert.equal(document.querySelector(backdropSelector), null);
    } finally { f.close(); }
  });
  test(`${presentation}: changing retention while open preserves the child subtree through its next native close`, () => {
    const id = `${presentation}-live-retention`;
    const tree = unmountOnClose => ({ options: options(presentation, { defaultOpen: true, unmountOnClose }), surfaceProps: { id } });
    const f = mount(tree(false));
    try {
      const node = surface(id), input = node.querySelector('input'); input.value = 'Edited';
      f.render(tree(true)); assert.equal(surface(id), node); assert.equal(node.querySelector('input'), input); assert.equal(input.value, 'Edited');
      f.render(tree(false)); assert.equal(surface(id), node); assert.equal(node.querySelector('input'), input);
      click(node.querySelector('[aria-label="Close"]')); assert.equal(surface(id), node); assert.equal(node.hidden, true);
      f.render(tree(true)); assert.equal(surface(id) === null, true);
    } finally { f.close(); }
  });
  test(`${presentation}: default-unmount closing keeps native focus ownership until source removal`, () => {
    const id = `${presentation}-native-focus-exit`;
    const f = mount({ options: options(presentation, { defaultOpen: true }), surfaceProps: { id } });
    const node = surface(id), inertMutations = [];
    const observer = new window.MutationObserver(records => inertMutations.push(...records));
    observer.observe(node, { attributes: true, attributeFilter: ['inert'] });
    try {
      click(node.querySelector('[aria-label="Close"]')); inertMutations.push(...observer.takeRecords());
      assert.equal(surface(id) === null, true);
      assert.equal(inertMutations.length, 0, 'authored inert must not move focus out before the native source is removed');
    } finally { observer.disconnect(); f.close(); }
  });
  test(`${presentation}: a closed retained surface releases native Tabster modal trapping and restores it on reopen`, () => {
    const id = `${presentation}-retained-modalizer`, trigger = React.createElement('button', { id: `${id}-trigger` }, 'Open');
    const f = mount({ options: options(presentation, { defaultOpen: true, unmountOnClose: false, trigger }), surfaceProps: { id } });
    try {
      const node = surface(id); assert.ok(JSON.parse(node.getAttribute('data-tabster')).modalizer);
      click(node.querySelector('[aria-label="Close"]'));
      assert.equal(surface(id), node); assert.equal(node.hidden, true);
      assert.equal(JSON.parse(node.getAttribute('data-tabster')).modalizer === undefined, true, 'closed retained source must not trap page keyboard focus');
      click(document.getElementById(`${id}-trigger`));
      assert.equal(surface(id), node); assert.ok(JSON.parse(node.getAttribute('data-tabster')).modalizer);
    } finally { f.close(); }
  });
  test(`${presentation}: retained mouse focus repair targets only the connected enabled own-document trigger after source focus loss`, () => {
    for (const mode of ['restore', 'otherFocus', 'disabled', 'removed', 'sourceNotFocused']) {
      const id = `${presentation}-retained-focus-${mode}`, ref = React.createRef();
      const tree = trigger => ({ options: options(presentation, { defaultOpen: true, unmountOnClose: false, trigger }), surfaceProps: { id } });
      const trigger = React.createElement('button', { ref }, 'Open');
      const outside = document.createElement('button'); document.body.appendChild(outside);
      const f = mount(tree(trigger));
      try {
        let node = surface(id), button = ref.current;
        if (mode === 'removed') { f.render(tree(undefined)); node = surface(id); assert.equal(ref.current, null); }
        if (mode === 'disabled') button.disabled = true;
        node.querySelector('input').focus();
        if (mode === 'sourceNotFocused') outside.focus();
        // JSDOM lacks inert's native focus-loss side effect. Model only that platform
        // boundary on this actual source; Fluent components, refs and motion stay real.
        const nativeSetAttribute = node.setAttribute.bind(node);
        node.setAttribute = (name, value) => {
          nativeSetAttribute(name, value);
          if (name === 'inert') {
            if (mode === 'otherFocus') outside.focus();
            else if (node.contains(document.activeElement)) document.activeElement.blur();
          }
        };
        click(node.querySelector('[aria-label="Close"]'));
        assert.equal(surface(id), node); assert.equal(node.hidden, true);
        if (mode === 'restore') assert.equal(document.activeElement === button, true);
        else if (mode === 'otherFocus' || mode === 'sourceNotFocused') assert.equal(document.activeElement, outside);
        else assert.equal(document.activeElement, document.body);
      } finally { f.close(); outside.remove(); }
    }
  });
  test(`${presentation}: retained focus repair uses the actual iframe owner document and preserves consumer trigger ref`, () => {
    const frame = document.createElement('iframe'); document.body.appendChild(frame);
    const owner = frame.contentDocument;
    Object.defineProperty(owner.defaultView, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
    const ref = React.createRef(), id = `${presentation}-frame-focus`;
    const f = mount({ options: options(presentation, { defaultOpen: true, unmountOnClose: false,
      trigger: React.createElement('button', { ref }, 'Open') }), surfaceProps: { id } }, 'rtl', owner);
    try {
      const node = owner.getElementById(id), button = ref.current, nativeSetAttribute = node.setAttribute.bind(node);
      node.querySelector('input').focus();
      node.setAttribute = (name, value) => { nativeSetAttribute(name, value); if (name === 'inert' && node.contains(owner.activeElement)) owner.activeElement.blur(); };
      click(node.querySelector('[aria-label="Close"]'));
      assert.equal(button.ownerDocument, owner); assert.equal(owner.activeElement === button, true);
      assert.equal(document.activeElement === button, false);
    } finally { f.close(); assert.equal(ref.current, null); frame.remove(); }
  });
  test(`${presentation}: retained iframe close preserves newly focused parent-document control`, () => {
    const frame = document.createElement('iframe'), outside = document.createElement('button');
    document.body.appendChild(frame); document.body.appendChild(outside);
    const owner = frame.contentDocument;
    Object.defineProperty(owner.defaultView, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
    const id = `${presentation}-frame-parent-focus`;
    const f = mount({ options: options(presentation, { defaultOpen: true, unmountOnClose: false,
      trigger: React.createElement('button', {}, 'Open') }), surfaceProps: { id } }, 'rtl', owner);
    try {
      const node = owner.getElementById(id), nativeSetAttribute = node.setAttribute.bind(node);
      node.querySelector('input').focus();
      // Model inert's platform blur, then the user's actual focus transition to
      // another document while the native closing motion is still completing.
      node.setAttribute = (name, value) => {
        nativeSetAttribute(name, value);
        if (name === 'inert') { if (node.contains(owner.activeElement)) owner.activeElement.blur(); outside.focus(); }
      };
      click(node.querySelector('[aria-label="Close"]'));
      assert.equal(document.activeElement === outside, true, 'retained restoration must not steal focus from the parent document');
      assert.equal(owner.activeElement, owner.body);
    } finally { f.close(); frame.remove(); outside.remove(); }
  });
}
test('Drawer logical position follows provider and an explicit native dir override through native classes', () => {
  const cases = [{ dir: 'ltr', native: undefined, position: undefined, edge: 'right' },
    { dir: 'rtl', native: undefined, position: undefined, edge: 'left' },
    { dir: 'ltr', native: 'rtl', position: 'end', edge: 'left' },
    { dir: 'rtl', native: 'ltr', position: 'start', edge: 'left' }];
  for (const item of cases) {
    const f = mount({ options: options('drawer', { open: true, drawerPosition: item.position }), surfaceProps: { id: 'position', dir: item.native } }, item.dir);
    try { assert.match(rulesFor(surface('position')), new RegExp(`${item.edge}: 0;`)); }
    finally { f.close(); }
  }
});

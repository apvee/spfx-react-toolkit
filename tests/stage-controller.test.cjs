const test = require('node:test');
const assert = require('node:assert/strict');
const { createStageHarness } = require('./fixtures/stage-harness.cjs');
let harness;
test.before(async () => { harness = await createStageHarness(); });
test.after(async () => { await harness?.close(); });
const props = overrides => ({ presentation: 'dialog', children: 'Body', ...overrides });

test('uncontrolled requests update the state and notify once; defaultOpen is initialization only', () => {
  const changes = [];
  const f = harness.mountController(props({ onOpenChange: (event, data) => changes.push([event, data]) }));
  try {
    assert.equal(f.value.open, false); assert.equal(f.value.size, 'small');
    const event = f.request({ open: true, reason: 'triggerClick' });
    assert.equal(f.value.open, true); assert.deepEqual(changes, [[event, { open: true, reason: 'triggerClick' }]]);
    f.render(props({ defaultOpen: false, onOpenChange: (_, data) => changes.push(data) }));
    assert.equal(f.value.open, true); assert.equal(changes.length, 1);
  } finally { f.close(); }
  const initialOpen = harness.mountController(props({ defaultOpen: true }));
  try { assert.equal(initialOpen.value.open, true); } finally { initialOpen.close(); }
});

test('controlled requests remain informational and prop updates never emit callbacks', () => {
  const changes = []; const onOpenChange = (_, data) => changes.push(data);
  const f = harness.mountController(props({ open: true, onOpenChange }));
  try {
    f.request({ open: false, reason: 'closeButton' });
    assert.equal(f.value.open, true); assert.deepEqual(changes, [{ open: false, reason: 'closeButton' }]);
    f.render(props({ open: false, onOpenChange })); assert.equal(f.value.open, false); assert.equal(changes.length, 1);
    f.render(props({ open: true, onOpenChange })); assert.equal(f.value.open, true); assert.equal(changes.length, 1);
  } finally { f.close(); }
});

test('explicit dismissal filters external interactions while trigger and close button remain available', () => {
  const changes = [];
  const f = harness.mountController(props({ defaultOpen: true, dismissMode: 'explicit', onOpenChange: (_, data) => changes.push(data) }));
  try {
    for (const reason of ['escapeKeyDown', 'backdropClick', 'outsideClick', 'focusOutside']) {
      f.request({ open: false, reason }); assert.equal(f.value.open, true);
    }
    assert.equal(changes.length, 0);
    f.request({ open: false, reason: 'triggerClick' }); assert.equal(f.value.open, false);
    f.request({ open: true, reason: 'triggerClick' }); f.request({ open: false, reason: 'closeButton' });
    assert.equal(f.value.open, false); assert.equal(changes.length, 3);
  } finally { f.close(); }
});

test('standard accepts every normalized dismissal and native prevented Escape; cancelled trigger is distinct', () => {
  const changes = [];
  const f = harness.mountController(props({ presentation: 'popover', trigger: harness.React.createElement('button'), defaultOpen: true,
    onOpenChange: (_, data) => changes.push(data) }));
  try {
    for (const reason of ['escapeKeyDown', 'outsideClick', 'focusOutside', 'backdropClick']) {
      f.request({ open: true, reason: 'triggerClick' });
      const event = new window.KeyboardEvent('keydown', { key: 'Escape', cancelable: true }); event.preventDefault();
      f.request({ open: false, reason }, event); assert.equal(f.value.open, false);
    }
    const event = new window.MouseEvent('click', { cancelable: true }); event.preventDefault();
    const before = changes.length; f.request({ open: true, reason: 'triggerClick' }, event);
    assert.equal(f.value.open, false); assert.equal(changes.length, before);
  } finally { f.close(); }
});

test('open presentation stays latched through refused close and updates next session, while size is live', () => {
  const f = harness.mountController(props({ open: true }));
  try {
    f.render(props({ presentation: 'drawer', size: 'large', open: true }));
    assert.equal(f.value.presentation, 'dialog'); assert.equal(f.value.size, 'large');
    f.request({ open: false, reason: 'closeButton' });
    f.render(props({ presentation: 'drawer', size: 'medium', open: true }));
    assert.equal(f.value.presentation, 'dialog'); assert.equal(f.value.size, 'medium');
    f.render(props({ presentation: 'drawer', open: false }));
    f.render(props({ presentation: 'drawer', open: true })); assert.equal(f.value.presentation, 'drawer');
  } finally { f.close(); }
});

test('active Popover keeps its last compatible size when pending Drawer requests full', () => {
  const trigger = harness.React.createElement('button');
  const f = harness.mountController(props({ presentation: 'popover', trigger, defaultOpen: true, size: 'medium' }));
  try {
    f.render(props({ presentation: 'drawer', size: 'large' })); assert.equal(f.value.size, 'large');
    f.render(props({ presentation: 'drawer', size: 'full' }));
    assert.equal(f.value.presentation, 'popover'); assert.equal(f.value.size, 'large');
    f.request({ open: false, reason: 'closeButton' }); f.request({ open: true, reason: 'triggerClick' });
    assert.equal(f.value.presentation, 'drawer'); assert.equal(f.value.size, 'full');
  } finally { f.close(); }
});

test('context IDs isolate roots and include description association only while registered', () => {
  const { React, StageContext, useStageController, useStageContextValue, useStageContext } = harness;
  function Description() {
    const context = useStageContext();
    React.useLayoutEffect(() => context.registerDescription(true), [context.registerDescription]);
    return React.createElement('p', { id: context.descriptionId }, 'Description');
  }
  function Surface({ description, id }) {
    const context = useStageContextValue(useStageController(props()), id);
    return React.createElement(StageContext.Provider, { value: context },
      React.createElement('div', { id: context.surfaceId, 'aria-labelledby': context.titleId,
        'aria-describedby': context.hasDescription ? context.descriptionId : undefined },
      React.createElement('h2', { id: context.titleId }, 'Title'), description ? React.createElement(Description) : undefined));
  }
  const tree = description => React.createElement(React.Fragment, {}, React.createElement(Surface, { description, id: 'consumer-id' }), React.createElement(Surface, {}));
  const f = harness.mount(tree(true));
  try {
    const surface = f.host.querySelector('#consumer-id'); const second = surface.nextElementSibling;
    assert.notEqual(surface.id, second.id); assert.notEqual(surface.getAttribute('aria-labelledby'), second.getAttribute('aria-labelledby'));
    const titleId = surface.getAttribute('aria-labelledby'); const descriptionId = surface.getAttribute('aria-describedby');
    assert.equal(document.getElementById(descriptionId).textContent, 'Description');
    assert.equal(second.hasAttribute('aria-describedby'), false);
    f.render(tree(false)); assert.equal(surface.hasAttribute('aria-describedby'), false);
    f.render(tree(true)); assert.equal(surface.getAttribute('aria-describedby'), descriptionId); assert.equal(surface.getAttribute('aria-labelledby'), titleId);
  } finally { f.close(); }
});

test('invalid JavaScript variants and parts outside Stage fail with actionable errors', () => {
  const { React } = harness;
  class Boundary extends React.Component {
    state = { error: undefined };
    static getDerivedStateFromError(error) { return { error }; }
    render() { return this.state.error ? React.createElement('output', {}, this.state.error.message) : this.props.children; }
  }
  function ControllerProbe({ options }) { harness.useStageController(options); return undefined; }
  function ContextProbe() { harness.useStageContext(); return undefined; }
  for (const [element, expected] of [
    [React.createElement(ControllerProbe, { options: props({ presentation: 'auto' }) }), /presentation/],
    [React.createElement(ControllerProbe, { options: props({ size: 'huge' }) }), /size/],
    [React.createElement(ControllerProbe, { options: props({ presentation: 'popover', size: 'full', trigger: React.createElement('button') }) }), /Popover.*full|popover.*full/],
    [React.createElement(ContextProbe), /inside Stage/],
  ]) {
    // React logs expected caught render errors; keep unrelated test output visible.
    const original = console.error; console.error = () => {};
    let f;
    try { f = harness.mount(React.createElement(Boundary, {}, element)); assert.match(f.host.querySelector('output').textContent, expected); }
    finally { console.error = original; f?.close(); }
  }
});

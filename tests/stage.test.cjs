const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let dom, React, ReactDOM, act, api, FluentProvider, RendererProvider, renderer, lightTheme, darkTheme;

test.before(async () => {
  process.env.NODE_ENV = 'test';
  dom = harness.createSxDom();
  for (const name of ['Node', 'NodeFilter', 'Element', 'HTMLElement', 'Event', 'CustomEvent', 'MutationObserver']) global[name] = dom.window[name];
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  ({ FluentProvider, webLightTheme: lightTheme, webDarkTheme: darkTheme } = require('@fluentui/react-components'));
  ({ RendererProvider } = require('@griffel/react'));
  renderer = require('@griffel/core').createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  api = await harness.loadModule('tests/fixtures/stage-root-entry.tsx');
});
test.after(async () => { dom?.window.close(); await harness.cleanup(); });

async function mount(element, { dark = false, dir = 'ltr' } = {}) {
  const host = document.createElement('div'); document.body.appendChild(host);
  async function render(next = element, options = {}) {
    element = next;
    dark = options.dark ?? dark; dir = options.dir ?? dir;
    await act(async () => { ReactDOM.render(React.createElement(RendererProvider, { renderer },
      React.createElement(FluentProvider, { theme: dark ? darkTheme : lightTheme, dir }, element)), host); });
  }
  await render();
  return { host, render, async close() { await act(async () => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
function parts({ title = '**Profile**', description = '*Details*', children = 'Body', close = true } = {}) {
  return [
    React.createElement(api.Stage.Header, { key: 'header', title, description, showCloseButton: close, closeLabel: 'Close profile' }),
    React.createElement(api.Stage.Body, { key: 'body' }, children),
    React.createElement(api.Stage.Footer, { key: 'footer' }, React.createElement('button', { type: 'button' }, 'Free action')),
  ];
}
function surface() { return document.querySelector('[data-stage-test="surface"]'); }
async function click(element) { await act(async () => { element.click(); }); }

test('root forwards native surface ID/ref and associates compound header without leaking configuration', async () => {
  const ref = React.createRef();
  const f = await mount(React.createElement(api.Stage, { presentation: 'dialog', open: true, ref, id: 'profile-surface',
    size: 'medium', 'data-stage-test': 'surface', 'aria-label': undefined }, parts()));
  try {
    const root = surface(); assert.equal(ref.current, root); assert.equal(root.id, 'profile-surface');
    assert.equal(root.getAttribute('role'), 'dialog'); assert.equal(root.getAttribute('aria-modal'), 'true');
    assert.equal(document.getElementById(root.getAttribute('aria-labelledby')).textContent, '**Profile**');
    assert.equal(document.getElementById(root.getAttribute('aria-describedby')).textContent, '*Details*');
    assert.equal(root.querySelector('h2 strong'), null);
    for (const attribute of ['presentation', 'size', 'dismissmode', 'unmountonclose', 'defaultopen']) assert.equal(root.hasAttribute(attribute), false);
  } finally { await f.close(); }
});

test('controlled close is an informational request and never changes the authoritative prop', async () => {
  const requests = [];
  const f = await mount(React.createElement(api.Stage, { presentation: 'dialog', open: true, 'data-stage-test': 'surface',
    onOpenChange: (_, data) => requests.push(data) }, parts()));
  try {
    const root = surface(); await click(root.querySelector('[aria-label="Close profile"]'));
    assert.deepEqual(requests, [{ open: false, reason: 'closeButton' }]); assert.equal(surface(), root);
    assert.equal(root.getAttribute('aria-hidden'), null);
  } finally { await f.close(); }
});

test('size and theme updates preserve a mounted body while description associations update', async () => {
  const inputRef = React.createRef(); let mounts = 0;
  function Content() { React.useEffect(() => { mounts++; }, []); return React.createElement('input', { ref: inputRef, defaultValue: 'Initial' }); }
  const tree = (size, description) => React.createElement(api.Stage, { presentation: 'dialog', open: true, size,
    'data-stage-test': 'surface' }, parts({ description, children: React.createElement(Content) }));
  const f = await mount(tree('small', '*Details*'));
  try {
    const input = inputRef.current; input.value = 'Changed';
    await f.render(tree('full', undefined), { dark: true, dir: 'rtl' });
    assert.equal(inputRef.current, input); assert.equal(input.value, 'Changed'); assert.equal(mounts, 1);
    // The explicit empty description removes the label; default helper copy is avoided here.
    await f.render(tree('large', ''), { dark: false });
    assert.equal(surface().hasAttribute('aria-describedby'), false);
  } finally { await f.close(); }
});

test('latched Popover keeps its active trigger when a valid pending Dialog has no trigger', async () => {
  const triggerRef = React.createRef();
  const popover = React.createElement(api.Stage, { presentation: 'popover', open: true,
    trigger: React.createElement('button', { ref: triggerRef, type: 'button' }, 'Context trigger'),
    positioning: { position: 'above', align: 'start' }, withArrow: true, 'data-stage-test': 'surface' }, parts());
  const f = await mount(popover);
  try {
    const root = surface(); const trigger = triggerRef.current;
    const arrow = root.firstElementChild;
    assert.equal(window.getComputedStyle(arrow).position, 'absolute', 'native arrow is present');
    await f.render(React.createElement(api.Stage, { presentation: 'dialog', open: true, size: 'full',
      'data-stage-test': 'surface' }, parts()));
    assert.equal(surface(), root); assert.equal(triggerRef.current, trigger); assert.equal(trigger.isConnected, true);
    assert.equal(root.firstElementChild, arrow, 'pending mode keeps the active Popover arrow');
    assert.equal(root.getAttribute('aria-modal'), null);
    await f.render(React.createElement(api.Stage, { presentation: 'dialog', open: false, size: 'full',
      'data-stage-test': 'surface' }, parts()));
    await f.render(React.createElement(api.Stage, { presentation: 'dialog', open: true, size: 'full',
      'data-stage-test': 'surface' }, parts()));
    assert.equal(surface().getAttribute('aria-modal'), 'true');
  } finally { await f.close(); }
});

test('Popover arrow toggles live without remounting content or leaking configuration', async () => {
  const tree = withArrow => React.createElement(api.Stage, { presentation: 'popover', open: true, withArrow,
    trigger: React.createElement('button', { type: 'button' }, 'Details'), 'data-stage-test': 'surface' },
  parts({ children: React.createElement('input', { defaultValue: 'Initial' }) }));
  const f = await mount(tree(undefined));
  try {
    const root = surface(), input = root.querySelector('input'); input.value = 'Changed';
    const nativeArrow = () => [...root.children].find(child => window.getComputedStyle(child).position === 'absolute');
    assert.equal(nativeArrow(), undefined, 'default omits arrow');
    await f.render(tree(true));
    const arrow = root.firstElementChild;
    assert.equal(window.getComputedStyle(arrow).position, 'absolute', 'Fluent native arrow is rendered');
    assert.equal(root.hasAttribute('witharrow'), false);
    assert.equal(root.querySelector('input'), input); assert.equal(input.value, 'Changed');
    await f.render(tree(false));
    assert.equal(surface(), root); assert.equal(nativeArrow(), undefined);
    assert.equal(root.querySelector('input'), input); assert.equal(input.value, 'Changed');
  } finally { await f.close(); }
});

test('two independent roots isolate IDs and controlled state requests', async () => {
  const first = [], second = [];
  const f = await mount(React.createElement(React.Fragment, {},
    React.createElement(api.Stage, { presentation: 'dialog', open: true, id: 'first-stage', onOpenChange: (_, data) => first.push(data) }, parts()),
    React.createElement(api.Stage, { presentation: 'dialog', open: true, id: 'second-stage', onOpenChange: (_, data) => second.push(data) }, parts())));
  try {
    const one = document.getElementById('first-stage'), two = document.getElementById('second-stage');
    assert.notEqual(one.getAttribute('aria-labelledby'), two.getAttribute('aria-labelledby'));
    await click(one.querySelector('[aria-label="Close profile"]'));
    assert.equal(first.length, 1); assert.equal(second.length, 0); assert.equal(document.getElementById('second-stage'), two);
  } finally { await f.close(); }
});

test('Form.Root wrapper keeps footer Submit associated with its native form and submission gate', async () => {
  const Editor = api.createForm(); let methods, saves = 0; const closeRequests = [];
  function Content() {
    methods = Editor.useForm({ defaultValues: { name: 'Initial' } });
    // RHF computes/subscribes to proxy state when read during render, not only in a callback.
    const dirty = methods.formState.isDirty;
    const sx = api.useSx();
    return React.createElement(api.Stage, { presentation: 'dialog', open: true, 'data-stage-test': 'surface',
      onOpenChange: (_, data) => { closeRequests.push({ ...data, dirty }); } },
    React.createElement(Editor.Root, { methods, onSubmit: async () => { saves++; },
      className: sx(api.flex.column, api.flexItem.grow, api.minHeight.zero, api.width.full) },
      React.createElement(api.Stage.Header, { title: 'Editor', closeLabel: 'Close editor' }),
      React.createElement(api.Stage.Body, {}, React.createElement(Editor.Input, { name: 'name', label: 'Name' })),
      React.createElement(api.Stage.Footer, {}, React.createElement(Editor.Submit, { appearance: 'primary' }, 'Save editor'))));
  }
  const f = await mount(React.createElement(Content));
  try {
    const form = surface().querySelector('form');
    const submit = Array.from(form.querySelectorAll('button')).find(button => button.textContent === 'Save editor');
    assert.equal(submit.type, 'submit'); assert.equal(submit.form, form);
    await act(async () => { methods.setValue('name', 'Changed', { shouldDirty: true }); });
    await click(surface().querySelector('[aria-label="Close editor"]'));
    assert.equal(closeRequests.length, 1); assert.equal(closeRequests[0].dirty, true); assert.ok(surface());
    await act(async () => {
      form.dispatchEvent(new window.Event('submit', { bubbles: true, cancelable: true }));
      form.dispatchEvent(new window.Event('submit', { bubbles: true, cancelable: true }));
    });
    assert.equal(saves, 1, 'Footer submit retains the existing Root gate');
  } finally { await f.close(); }
});

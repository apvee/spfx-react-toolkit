const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./fixtures/web-part-shell/title-editor-harness.cjs');

test('actual editor publishes one atomic title patch through real provider preserving app properties', async t => {
  const h = createHarness(); t.after(() => h.close());
  const host = h.instance('editor'); host.displayMode = 2; host.properties.listId = 'keep';
  let refreshes = 0; host.context.propertyPane.refresh = () => refreshes++;
  const fixture = h.mount(host, 'Child'); await h.settle();
  const editor = h.editor(fixture); assert.ok(editor, 'Edit mounts the real engine');
  assert.equal(refreshes, 0);
  h.act(() => editor.commands.insertContent('new '));
  assert.equal(host.properties.title, 'new initial');
  assert.equal(host.properties.titleHTML, '<h2>new initial</h2>');
  assert.equal(host.properties.listId, 'keep'); assert.equal(refreshes, 1);
});

async function mounted(t, properties = {}) {
  const h = createHarness(); t.after(() => h.close());
  const host = h.instance('sync'); host.displayMode = 2; Object.assign(host.properties, properties);
  let writes = 0; host.context.propertyPane.refresh = () => writes++;
  const fixture = h.mount(host, 'Child'); await h.settle();
  return { h, host, fixture, editor: h.editor(fixture), writes: () => writes };
}

test('canonical echo preserves view, model caret and undo; selection-only transactions publish nothing', async t => {
  const { h, host, fixture, editor, writes } = await mounted(t);
  const view = editor.view;
  h.act(() => { editor.commands.setTextSelection(4); editor.commands.insertContent('X'); });
  const caret = editor.state.selection.from; const count = writes();
  fixture.render(host, { dir: 'rtl' });
  assert.equal(h.editor(fixture), editor); assert.equal(editor.view, view); assert.equal(editor.state.selection.from, caret);
  h.act(() => editor.commands.setTextSelection(2)); assert.equal(writes(), count);
  h.act(() => editor.commands.undo()); assert.equal(host.properties.title, 'initial');
});

test('external content uses existing view with clamped selection and a new history base', async t => {
  const { h, host, fixture, editor, writes } = await mounted(t);
  h.act(() => editor.commands.insertContent('local '));
  const view = editor.view; const count = writes();
  host.properties.titleHTML = '<h3>X</h3>'; fixture.render();
  assert.equal(editor.view, view); assert.equal(editor.getHTML(), '<h3>X</h3>'); assert.equal(writes(), count);
  assert.ok(editor.state.selection.from <= 2);
  h.act(() => editor.commands.undo()); assert.equal(editor.getHTML(), '<h3>X</h3>');
  h.act(() => editor.commands.insertContent('Y')); h.act(() => editor.commands.undo());
  assert.equal(editor.getHTML(), '<h3>X</h3>');
});

test('fallback/normalization and defined empty mount publish nothing; empty edit stays authoritative', async t => {
  const { h, host, fixture, editor, writes } = await mounted(t, { titleHTML: '<h2><b>initial</b></h2>' });
  assert.equal(writes(), 0);
  h.act(() => editor.commands.clearContent());
  assert.equal(host.properties.title, ''); assert.equal(host.properties.titleHTML, '');
  host.displayMode = 1; fixture.render(); assert.equal(fixture.container.querySelector('h2'), null);
  host.properties.title = 'old'; host.displayMode = 2; fixture.render(); await h.settle();
  assert.equal(h.editor(fixture).getText(), ''); assert.equal(writes(), 1);
});

test('composition defers replacements, local final document wins and no-local composition accepts latest external', async t => {
  const { h, host, fixture, editor, writes } = await mounted(t);
  const event = type => editor.view.dom.dispatchEvent(new h.window.CompositionEvent(type, { bubbles: true }));
  h.act(() => event('compositionstart'));
  host.properties.titleHTML = '<h2>external</h2>'; fixture.render();
  assert.equal(editor.getText(), 'initial');
  h.act(() => editor.commands.insertContent('IME ')); assert.equal(writes(), 0);
  h.act(() => event('compositionend')); await h.settle();
  assert.equal(editor.getText(), 'IME initial'); assert.equal(host.properties.title, 'IME initial'); assert.equal(writes(), 1);
  h.act(() => event('compositionstart'));
  host.properties.titleHTML = '<p>first</p>'; fixture.render();
  host.properties.titleHTML = '<p>last</p>'; fixture.render();
  h.act(() => event('compositionend')); await h.settle();
  assert.equal(editor.getHTML(), '<p>last</p>'); assert.equal(writes(), 1);
});

test('blur, Read and unmount flush current composition draft; destroyed callbacks cannot publish', async t => {
  const { h, host, fixture, editor, writes } = await mounted(t);
  h.act(() => { editor.view.dom.dispatchEvent(new h.window.CompositionEvent('compositionstart', { bubbles: true })); editor.commands.insertContent('final '); });
  assert.equal(writes(), 0);
  h.act(() => editor.emit('blur', { editor, event: new h.window.FocusEvent('blur') }));
  assert.equal(host.properties.title, 'final initial'); assert.equal(writes(), 1);
  h.act(() => editor.commands.insertContent('read ')); host.displayMode = 1; fixture.render();
  assert.equal(host.properties.title, 'final read initial'); assert.equal(writes(), 2); assert.equal(editor.isDestroyed, true);
  assert.equal(fixture.container.querySelector('h2').textContent, 'final read initial');
  editor.emit('update', { editor, transaction: {} }); assert.equal(writes(), 2);
  host.displayMode = 2; fixture.render(); await h.settle(); const next = h.editor(fixture);
  h.act(() => { next.view.dom.dispatchEvent(new h.window.CompositionEvent('compositionstart', { bubbles: true })); next.commands.insertContent('unmount '); });
  fixture.unmount(); assert.equal(host.properties.title, 'unmount final read initial'); assert.equal(writes(), 3);
});

test('oversized canonical draft is retained with Fluent error and no partial patch, then recovers', async t => {
  const { h, host, fixture, editor, writes } = await mounted(t);
  h.act(() => editor.commands.setContent({ type: 'doc', content: [{ type: 'heading', attrs: { level: 2 }, content: [{ type: 'text', text: '&'.repeat(14000) }] }] }));
  assert.equal(editor.getText().length, 14000); assert.equal(writes(), 0); assert.equal(host.properties.title, 'initial');
  assert.ok(fixture.container.textContent.includes('Title is too large'));
  h.act(() => editor.commands.setContent('<h2>Fits</h2>'));
  assert.equal(host.properties.titleHTML, '<h2>Fits</h2>'); assert.equal(host.properties.title, 'Fits'); assert.equal(writes(), 1);
  assert.ok(!fixture.container.textContent.includes('Title is too large'));
});

test('latest callback receives final flush and stale callbacks are ignored after unmount', async t => {
  const h = createHarness(); t.after(() => h.close());
  const ReactDOM = require('react-dom');
  const { WebPartTitleEditor } = h.load('controls/web-part-shell/title/web-part-title-editor.internal.tsx');
  const doc = { type: 'heading', level: 2, content: [{ type: 'text', text: 'start' }] };
  const container = document.createElement('div'); document.body.appendChild(container); const calls = [];
  const render = callback => h.act(() => ReactDOM.render(h.React.createElement(WebPartTitleEditor, { document: doc, onDocumentChange: callback, ariaLabel: 'Title', placeholder: 'Add' }), container));
  render(() => calls.push('old')); const editor = container.querySelector('[contenteditable]').editor;
  render(document => calls.push(document.content[0].text));
  h.act(() => { editor.view.dom.dispatchEvent(new h.window.CompositionEvent('compositionstart', { bubbles: true })); editor.commands.insertContent('new '); });
  h.act(() => ReactDOM.unmountComponentAtNode(container)); container.remove();
  assert.deepEqual(calls, ['new start']); editor.emit('update', { editor, transaction: {} }); assert.equal(calls.length, 1);
});

test('lazy loader is unused in Read/hidden Edit; chunk failure and retry stay local preserving children', async t => {
  const h = createHarness(); t.after(() => h.close());
  const module = h.load('controls/web-part-shell/title/web-part-title-editor.internal.tsx');
  const real = module.WebPartTitleEditor; let loads = 0;
  Object.defineProperty(module, 'WebPartTitleEditor', { configurable: true, get() { loads++; throw new Error('test chunk failure'); } });
  let mounts = 0;
  function Child() { h.React.useEffect(() => { mounts++; }, []); return h.React.createElement('input', { defaultValue: 'draft' }); }
  const host = h.instance('lazy'); const fixture = h.mount(host, h.React.createElement(Child));
  const child = fixture.container.querySelector('input'); child.value = 'retained';
  assert.equal(loads, 0);
  host.displayMode = 2; host.properties.showTitle = false; fixture.render(); await h.settle(); assert.equal(loads, 0);
  const oldError = console.error; console.error = (...args) => { if (!args.some(arg => String(arg).includes('test chunk failure') || String(arg).includes('The above error occurred'))) oldError(...args); };
  try {
    host.properties.showTitle = true; fixture.render();
    assert.ok(fixture.container.textContent.includes('Loading title editor')); await h.settle();
    assert.ok(fixture.container.textContent.includes('Title editor could not load')); assert.equal(loads, 1);
    assert.equal(fixture.container.querySelector('input'), child); assert.equal(child.value, 'retained'); assert.equal(mounts, 1);
    Object.defineProperty(module, 'WebPartTitleEditor', { configurable: true, value: real });
    h.act(() => Array.from(fixture.container.querySelectorAll('button')).find(button => button.textContent === 'Retry').click());
    await h.settle(); assert.ok(h.editor(fixture)); assert.equal(mounts, 1);
  } finally { console.error = oldError; }
});

test('application and Read-renderer errors are not swallowed by the editor boundary', t => {
  const h = createHarness(); t.after(() => h.close());
  const ReactDOM = require('react-dom'); const oldError = console.error; console.error = () => {};
  try {
    const host = h.instance('error'); function Child() { throw new Error('app failure'); }
    assert.throws(() => h.mount(host, h.React.createElement(Child)), /app failure/);
    const renderer = h.load('controls/web-part-shell/title/web-part-title-renderer.internal.tsx');
    renderer.WebPartTitleReadRenderer = () => { throw new Error('renderer failure'); };
    assert.throws(() => h.mount(host, 'Child'), /renderer failure/);
  } finally { console.error = oldError; }
  // Keep actual failures visible to React and the host; no property repair/write was performed.
  assert.ok(ReactDOM);
});

test('composition arbitration waits for final engine input in the compositionend microtask', async t => {
  const { h, host, fixture, editor, writes } = await mounted(t);
  h.act(() => editor.view.dom.dispatchEvent(new h.window.CompositionEvent('compositionstart', { bubbles: true })));
  host.properties.titleHTML = '<h2>obsolete external</h2>'; fixture.render();
  h.act(() => editor.view.dom.dispatchEvent(new h.window.CompositionEvent('compositionend', { bubbles: true })));
  queueMicrotask(() => h.act(() => editor.commands.insertContent('final ')));
  await h.settle(); assert.equal(host.properties.title, 'final initial'); assert.equal(writes(), 1);
});

test('engine document/view belongs to the public web part ownerDocument', async t => {
  const h = createHarness(); t.after(() => h.close());
  const frame = document.createElement('iframe'); document.body.appendChild(frame);
  const doc = frame.contentDocument; doc.getSelection = () => null;
  const host = h.instance('frame'); host.displayMode = 2; host.context.domElement = doc.createElement('div');
  const fixture = h.mount(host, 'Child', {}, { document: doc }); await h.settle();
  const editor = h.editor(fixture); assert.ok(editor); assert.equal(editor.view.dom.ownerDocument, doc);
  h.act(() => editor.commands.insertContent('frame ')); assert.equal(host.properties.title, 'frame initial');
  fixture.unmount(); frame.remove();
});

test('latest external echo cancels a superseded composition candidate without replacing the view state', async t => {
  const { h, host, fixture, editor, writes } = await mounted(t);
  h.act(() => editor.view.dom.dispatchEvent(new h.window.CompositionEvent('compositionstart', { bubbles: true })));
  const state = editor.state; const view = editor.view;
  host.properties.titleHTML = '<h2>intermediate</h2>'; fixture.render();
  host.properties.titleHTML = '<h2>initial</h2>'; fixture.render();
  h.act(() => editor.view.dom.dispatchEvent(new h.window.CompositionEvent('compositionend', { bubbles: true })));
  await h.settle();
  assert.equal(editor.getHTML(), '<h2>initial</h2>'); assert.equal(host.properties.titleHTML, '<h2>initial</h2>');
  assert.equal(editor.view, view); assert.equal(editor.state, state); assert.equal(writes(), 0);
});

test('Select All external replacement uses valid inline endpoints and insertion preserves the incoming heading', async t => {
  const { h, host, fixture, editor, writes } = await mounted(t);
  h.act(() => editor.commands.selectAll());
  assert.equal(editor.state.selection.constructor.name, 'AllSelection');
  const warnings = []; const oldWarn = console.warn; console.warn = (...args) => warnings.push(args.map(String).join(' '));
  try {
    host.properties.titleHTML = '<h3>X</h3>'; fixture.render();
    const selection = editor.state.selection;
    const endpointsInline = selection.$anchor.parent.inlineContent && selection.$head.parent.inlineContent;
    h.act(() => editor.commands.insertContent('Y'));
    assert.deepEqual({ endpointsInline, html: editor.getHTML(), persisted: host.properties.titleHTML, warnings },
      { endpointsInline: true, html: '<h3>Y</h3>', persisted: '<h3>Y</h3>', warnings: [] });
    assert.equal(writes(), 1);
  } finally { console.warn = oldWarn; }
});

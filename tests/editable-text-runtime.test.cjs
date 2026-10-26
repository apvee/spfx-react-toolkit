const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./fixtures/editable-text/runtime-harness.cjs');

test('standalone Read defaults to a literal paragraph, enforces policy and never loads the engine', t => {
  const h = createHarness(); t.after(() => h.close());
  const f = h.mount({ fallbackText: '<literal>\n text' });
  assert.equal(f.root.querySelector('p').textContent, '<literal>\n text');
  assert.equal(f.root.querySelector('literal'), null);
  assert.equal(h.engineLoads, 0);
  f.render({ value: '<h2>Saved</h2>', allowedBlocks: ['p'] });
  assert.equal(f.root.querySelector('p').textContent, 'Saved');
  f.render({ value: '', fallbackText: 'Must not replace empty HTML' });
  assert.equal(f.root.textContent, ''); assert.equal(f.calls.length, 0); assert.equal(h.engineLoads, 0);
});

test('one native root keeps refs, attributes, events and consumer styles through mode changes', async t => {
  const h = createHarness(); t.after(() => h.close()); const ref = h.React.createRef(); let clicks = 0;
  const f = h.mount({ ref, className: 'consumer', style: { marginTop: 7 }, id: 'region', tabIndex: 3,
    title: 'Native title', onClick: () => clicks++, 'data-record': '42', 'aria-describedby': 'help' });
  const root = f.root, before = f.container.querySelector('[data-before]'), after = f.container.querySelector('[data-after]');
  assert.equal(ref.current, root); assert.equal(root.tagName, 'DIV'); assert.ok(root.classList.contains('consumer'));
  assert.equal(root.style.marginTop, '7px'); assert.equal(root.id, 'region'); assert.equal(root.tabIndex, 3);
  assert.equal(root.title, 'Native title'); assert.equal(root.getAttribute('data-record'), '42');
  assert.equal(root.getAttribute('aria-describedby'), 'help');
  h.act(() => root.dispatchEvent(new h.window.MouseEvent('click', { bubbles: true }))); assert.equal(clicks, 1);
  f.render({ mode: 'edit', placeholder: 'Write here', 'aria-label': 'Independent name' }); await h.settle();
  const e = f.editor; assert.ok(e); assert.equal(ref.current, root); assert.equal(f.root, root);
  assert.equal(e.view.dom.getAttribute('aria-label'), 'Independent name');
  assert.equal(e.view.dom.getAttribute('aria-describedby'), 'help'); assert.equal(root.hasAttribute('aria-label'), false);
  f.render({ mode: 'read' }); assert.equal(e.isDestroyed, true); assert.equal(f.root, root);
  assert.equal(f.container.querySelector('[data-before]'), before); assert.equal(f.container.querySelector('[data-after]'), after);
  assert.equal(root.getAttribute('aria-label'), 'Independent name'); f.unmount(); assert.equal(ref.current, null);
});

test('untyped callers cannot replace owned content, semantics, direction or engine configuration', async t => {
  const h = createHarness(); t.after(() => h.close());
  const f = h.mount({ children: 'INJECTED', dangerouslySetInnerHTML: { __html: '<img src=x onerror="alert(1)">' },
    contentEditable: true, suppressContentEditableWarning: true, defaultValue: 'raw', role: 'button', dir: 'rtl',
    theme: 'override', instance: 'host', ownerDocument: 'foreign', editor: 'engine' });
  assert.equal(f.root.textContent, 'initial'); assert.equal(f.root.querySelector('img'), null);
  for (const attr of ['role', 'contenteditable', 'defaultvalue', 'dir', 'theme', 'instance', 'ownerdocument', 'editor']) assert.equal(f.root.hasAttribute(attr), false, attr);
  f.render({ mode: 'edit' }); await h.settle(); assert.equal(f.editor.view.dom.getAttribute('role'), 'textbox');
  assert.equal(f.root.hasAttribute('contenteditable'), false); assert.equal(f.root.textContent.includes('INJECTED'), false);
});

test('JS Edit validates a nonempty accessible name, callback and finite configuration', t => {
  const h = createHarness(); t.after(() => h.close());
  // Validation happens in the function render, before any editor loader is invoked.
  const render = props => h.React.createElement(h.load('controls/editable-text/editable-text.tsx').EditableText, props);
  const ReactDOMServer = require('react-dom/server');
  const previousError = console.error; console.error = () => {};
  try {
    for (const props of [
      { mode: 'edit', 'aria-label': 'Name' }, { mode: 'edit', onChange() {} },
      { mode: 'edit', 'aria-label': '  ', 'aria-labelledby': '\n', onChange() {} },
      { mode: 'write' }, { mode: 'read', onChange: 'bad' }, { value: 3 },
      { allowedBlocks: [] }, { allowedBlocks: ['h5'] }, { allowedBlocks: ['h2'], defaultBlock: 'p' }
    ]) assert.throws(() => ReactDOMServer.renderToStaticMarkup(render(props)), { name: 'TypeError' });
  } finally { console.error = previousError; }
  assert.equal(h.engineLoads, 0);
});

test('Edit fallback and localized placeholder remain distinct from the textbox accessible name', async t => {
  const h = createHarness(); t.after(() => h.close());
  const f = h.mount({ mode: 'edit', value: '', locale: 'it-IT', 'aria-label': undefined, 'aria-labelledby': 'label-id' }); await h.settle();
  const e = f.editor; assert.equal(e.getHTML(), '<p></p>'); assert.equal(e.view.dom.getAttribute('aria-labelledby'), 'label-id');
  const placeholder = e.view.dom.querySelector('[data-placeholder]')?.getAttribute('data-placeholder'); assert.ok(placeholder);
  f.render({ placeholder: 'Literal placeholder' }); assert.equal(e.view.dom.querySelector('[data-placeholder]').getAttribute('data-placeholder'), 'Literal placeholder');
  assert.equal(e.view.dom.getAttribute('aria-label'), null); assert.equal(f.calls.length, 0);
});

test('effective changes derive canonical HTML and search text from one document and use the latest callback', async t => {
  const h = createHarness(); t.after(() => h.close()); const next = [];
  const f = h.mount({ mode: 'edit', value: '<h2>A<br>B</h2>', allowedBlocks: ['p'] }); await h.settle();
  const e = f.editor; assert.equal(e.getHTML(), '<p>A<br>B</p>'); assert.equal(f.calls.length, 0);
  f.render({ onChange: data => next.push(data) }); h.act(() => e.commands.insertContent('X '));
  assert.deepEqual(next, [{ html: '<p>X A<br>B</p>', text: 'X A B' }]); assert.equal(f.calls.length, 0);
  f.render({ value: '<h4>External</h4>' }); assert.equal(e.getHTML(), '<p>External</p>'); assert.equal(next.length, 1);
  h.act(() => e.commands.selectAll()); h.act(() => e.commands.deleteSelection());
  assert.deepEqual(next.at(-1), { html: '', text: '' });
});

test('configuration rerenders reconcile live draft without stale-source replacement, callbacks or editor recreation', async t => {
  const h = createHarness(); t.after(() => h.close());
  const f = h.mount({ mode: 'edit', value: '<p>Initial</p>' }); await h.settle(); const e = f.editor, view = e.view;
  h.act(() => e.commands.insertContent('Draft ')); const count = f.calls.length;
  f.render({ allowedBlocks: ['h4', 'h2'], locale: 'fr-FR', placeholder: 'New' });
  assert.equal(f.editor, e); assert.equal(e.view, view); assert.equal(e.getHTML(), '<h2>Draft Initial</h2>'); assert.equal(f.calls.length, count);
  f.render({ allowedBlocks: ['h2', 'h4', 'h2'], defaultBlock: 'h4' });
  assert.equal(e.getHTML(), '<h2>Draft Initial</h2>'); assert.equal(f.calls.length, count);
  f.render({ value: '<p>Changed source</p>' }); assert.equal(e.getHTML(), '<h4>Changed source</h4>'); assert.equal(f.calls.length, count);
});

test('oversized live drafts remain visible with local error and never publish partial payloads', async t => {
  const h = createHarness(); t.after(() => h.close()); const f = h.mount({ mode: 'edit' }); await h.settle();
  h.act(() => f.editor.commands.setContent({ type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: '&'.repeat(14000) }] }] }));
  assert.equal(f.editor.getText().length, 14000); assert.ok(f.root.textContent.includes('too large')); assert.equal(f.calls.length, 0);
  f.render({ allowedBlocks: ['h4'] }); assert.equal(f.editor.getText().length, 14000); assert.equal(f.calls.length, 0);
});

test('invalid saved markup shows the literal fallback warning only in Edit without publishing', async t => {
  const h = createHarness(); t.after(() => h.close());
  const f = h.mount({ value: '<div>'.repeat(33) + '<img src=x onerror="alert(1)">' + '</div>'.repeat(33), fallbackText: '<Safe fallback>' });
  assert.equal(f.root.querySelector('p').textContent, '<Safe fallback>'); assert.equal(f.root.querySelector('img'), null);
  f.render({ mode: 'edit' }); await h.settle();
  assert.equal(f.editor.getText(), '<Safe fallback>'); assert.ok(f.root.textContent.includes('formatting'));
  assert.equal(f.calls.length, 0);
  f.render({ mode: 'read' }); assert.equal(f.root.textContent, '<Safe fallback>'); assert.equal(f.calls.length, 0);
});

test('Read exit and unmount flush IME once with the latest policy and committed callback', async t => {
  const h = createHarness(); t.after(() => h.close()); const latest = [];
  const f = h.mount({ mode: 'edit', value: '<p>Initial</p>' }); await h.settle(); const e = f.editor;
  h.act(() => e.view.dom.dispatchEvent(new h.window.CompositionEvent('compositionstart', { bubbles: true })));
  h.act(() => e.commands.insertContent('Final '));
  f.render({ allowedBlocks: ['h4'], onChange: data => latest.push(data) });
  assert.equal(f.calls.length, 0); assert.equal(latest.length, 0);
  f.render({ mode: 'read' });
  assert.equal(e.isDestroyed, true); assert.deepEqual(latest, [{ html: '<h4>Final Initial</h4>', text: 'Final Initial' }]);
  f.unmount(); assert.equal(latest.length, 1);
});

for (const localInput of [false, true]) test(`same-render Read exit commits current policy/callback with ${localInput ? 'one final' : 'no configuration-only'} publication`, async t => {
  const h = createHarness(); t.after(() => h.close()); const latest = [];
  const f = h.mount({ value: '<p>Initial</p>' }); const root = f.root;
  assert.equal(h.engineLoads, 0);
  f.render({ mode: 'edit' }); await h.settle(); const e = f.editor;
  h.act(() => e.view.dom.dispatchEvent(new h.window.CompositionEvent('compositionstart', { bubbles: true })));
  if (localInput) h.act(() => e.commands.insertContent('Final '));
  f.render({ mode: 'read', allowedBlocks: ['h4'], onChange: data => latest.push(data) });
  assert.equal(e.isDestroyed, true); assert.equal(f.root, root);
  assert.equal(root.querySelector('h4').textContent, 'Initial');
  assert.deepEqual(f.calls, [], 'the obsolete callback must not receive the final change');
  assert.deepEqual(latest, localInput ? [{ html: '<h4>Final Initial</h4>', text: 'Final Initial' }] : []);
  await h.settle(); f.unmount(); assert.equal(latest.length, localInput ? 1 : 0);
});

test('same-render Read exit without a callback flushes through the last committed Edit callback and latest policy', async t => {
  const h = createHarness(); t.after(() => h.close());
  const f = h.mount({ mode: 'edit', value: '<p>Initial</p>' }); await h.settle(); const e = f.editor, root = f.root;
  const committed = []; f.render({ onChange: data => committed.push(data) });
  h.act(() => e.view.dom.dispatchEvent(new h.window.CompositionEvent('compositionstart', { bubbles: true })));
  h.act(() => e.commands.insertContent('Final '));
  f.render({ mode: 'read', allowedBlocks: ['h4'], onChange: undefined });
  assert.equal(e.isDestroyed, true); assert.equal(f.root, root); assert.deepEqual(f.calls, []);
  assert.deepEqual(committed, [{ html: '<h4>Final Initial</h4>', text: 'Final Initial' }]);
  await h.settle(); f.unmount(); assert.equal(committed.length, 1);
});

test('two standalone regions keep draft, history, callback and local load recovery independent', async t => {
  const h = createHarness({ failEngineLoads: 1 }); t.after(() => h.close());
  const left = h.mount({ mode: 'edit', fallbackText: 'Left' }); const right = h.mount({ mode: 'edit', fallbackText: 'Right' });
  const previousError = console.error; console.error = () => {};
  try { await h.settle(); } finally { console.error = previousError; }
  const leftRetry = Array.from(left.root.querySelectorAll('button')).find(button => button.textContent === 'Retry');
  h.act(() => leftRetry.click()); await h.settle(); assert.ok(left.editor); assert.equal(right.editor, undefined);
  const rightRetry = Array.from(right.root.querySelectorAll('button')).find(button => button.textContent === 'Retry');
  h.act(() => rightRetry.click()); await h.settle(); assert.ok(right.editor); assert.notEqual(left.editor, right.editor);
  h.act(() => left.editor.commands.insertContent('Edited ')); assert.equal(right.editor.getText(), 'Right');
  assert.equal(left.calls.length, 1); assert.equal(right.calls.length, 0);
  h.act(() => left.editor.commands.undo()); assert.equal(left.editor.getText(), 'Left'); assert.equal(right.editor.getText(), 'Right');
});

test('lazy chunk failure is local and Retry mounts the real editor with current callback and labels', async t => {
  const h = createHarness({ failEngineLoads: 1 }); t.after(() => h.close());
  const f = h.mount({ mode: 'edit' }); const root = f.root;
  const previousError = console.error; console.error = () => {};
  try { await h.settle(); } finally { console.error = previousError; }
  assert.ok(root.textContent.includes('Retry')); assert.equal(f.editor, undefined); assert.equal(f.container.querySelector('[data-after]').textContent, 'After');
  const calls = []; f.render({ onChange: data => calls.push(data), 'aria-label': 'Latest name' });
  const retry = Array.from(root.querySelectorAll('button')).find(button => button.textContent === 'Retry');
  h.act(() => retry.click()); await h.settle(); assert.equal(f.root, root); assert.ok(f.editor);
  assert.equal(f.editor.view.dom.getAttribute('aria-label'), 'Latest name'); h.act(() => f.editor.commands.insertContent('X'));
  assert.equal(calls.length, 1); assert.equal(f.calls.length, 0);
});

test('real secondary Document owns root, editor and Fluent popup without SPFx or a global-document dependency', async t => {
  const h = createHarness(); t.after(() => h.close()); const second = h.secondaryDocument(); t.after(() => second.window.close());
  const target = second.window.document; let usedTarget = false; const original = target.createElement.bind(target);
  target.createElement = (...args) => { if (args[0] === 'template') usedTarget = true; return original(...args); };
  const f = h.mount({ mode: 'edit', value: '<h3>Second</h3>', 'aria-label': 'Second document' }, { document: target }); await h.settle();
  assert.equal(f.root.ownerDocument, target); assert.equal(f.editor.view.dom.ownerDocument, target); assert.equal(usedTarget, true);
  h.act(() => f.editor.view.dom.dispatchEvent(new second.window.FocusEvent('focusin', { bubbles: true }))); await h.settle();
  const dropdown = target.querySelector('[aria-label="Heading"]'); assert.ok(dropdown);
  h.act(() => dropdown.click()); await h.settle();
  assert.ok(target.querySelector('[role="listbox"]')); assert.equal(document.querySelector('[role="listbox"]'), null);
  f.unmount();
});

test('without a Fluent document context, the mounted root supplies the parser document', t => {
  const h = createHarness(); t.after(() => h.close()); const target = document;
  const original = target.createElement;
  let parsed = false;
  target.createElement = function (...args) {
    if (args[0] === 'template') {
      assert.ok(target.querySelector('[data-editable-fixture]'), 'root must be mounted before parsing without context');
      parsed = true;
    }
    return original.apply(this, args);
  };
  try {
    const f = h.mount({ value: '<h4>Root only</h4>' }, { context: { dir: 'ltr', targetDocument: undefined } });
    assert.equal(parsed, true); assert.equal(f.root.querySelector('h4').textContent, 'Root only');
    assert.equal(f.root.ownerDocument, target); assert.equal(h.engineLoads, 0);
    f.unmount();
  } finally { target.createElement = original; }
});

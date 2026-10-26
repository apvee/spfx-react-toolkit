const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const base = 'controls/editable-text/';
async function fixture(t, html) {
  const h = createHarness();
  const names = ['Node', 'HTMLElement', 'Element', 'MutationObserver', 'getComputedStyle', 'innerHeight'];
  const previous = names.map(name => Object.getOwnPropertyDescriptor(global, name));
  names.forEach(name => { global[name] = name === 'getComputedStyle' ? window[name].bind(window) : window[name]; });
  const oldSelection = document.getSelection;
  document.getSelection = () => undefined;
  const { Editor } = require('@tiptap/core');
  const editor = new Editor({ element: document.createElement('div'),
    extensions: h.load(base + 'editable-text-extensions.internal.ts').createEditableTextExtensions('Title', () => h.load(base + 'editable-text-policy.internal.ts').resolveEditableTextBlockPolicy(undefined, 'h2')),
    content: html ?? '<h2>initial</h2>', injectCSS: false });
  const adapter = h.load(base + 'editable-text-editor-adapter.internal.ts');
  const codec = h.load(base + 'editable-text-codec.internal.ts');
  const canonicalHTML = () => codec.serializeEditableTextDocument(adapter.editorContentToEditableTextDocument(editor.getJSON()));
  t.after(() => { editor.destroy(); h.close(); document.getSelection = oldSelection;
    names.forEach((name, i) => { if (previous[i]) Object.defineProperty(global, name, previous[i]); else delete global[name]; }); });
  return { h, editor, canonicalHTML, format: h.load(base + 'editable-text-formatting.internal.ts') };
}
test('selection distinguishes Automatic from explicit primary even though defaults may visually match', async t => {
  const { h, editor, format, canonicalHTML } = await fixture(t, '<h2>A<span class="wps-color-primary">B</span>C</h2>');
  h.act(() => { editor.commands.setTextSelection({ from: 1, to: 2 }); });
  assert.equal(format.readEditableTextSelectionFormat(editor).color, undefined);
  h.act(() => { editor.commands.setTextSelection({ from: 2, to: 3 }); });
  assert.equal(format.readEditableTextSelectionFormat(editor).color, 'primary');
  h.act(() => { editor.commands.setTextSelection({ from: 1, to: 4 }); });
  assert.equal(format.readEditableTextSelectionFormat(editor).color, 'mixed');
  assert.equal(canonicalHTML(), '<h2>A<span class="wps-color-primary">B</span>C</h2>');
});
test('all color roles are read from ranges and caret stored marks, Automatic removal is undoable', async t => {
  const { h, editor, format, canonicalHTML } = await fixture(t);
  const roles = ['primary', 'subtle', 'brand', 'neutralMuted', 'brandAlt', 'red', 'orange', 'yellow', 'green', 'teal', 'blue', 'purple', 'pink', 'brown'];
  for (const color of roles) {
    h.act(() => { editor.commands.setTextSelection({ from: 1, to: 8 }); editor.commands.setMark('titleStyle', { color }); });
    assert.equal(format.readEditableTextSelectionFormat(editor).color, color);
    assert.equal(canonicalHTML(), `<h2><span class="wps-color-${color}">initial</span></h2>`);
  }
  // Isolate the removal's undo event from preceding formatting commands.
  const { closeHistory } = require('@tiptap/pm/history');
  h.act(() => { editor.view.dispatch(closeHistory(editor.state.tr)); });
  h.act(() => { editor.commands.setMark('titleStyle', { color: undefined }); });
  assert.equal(format.readEditableTextSelectionFormat(editor).color, undefined);
  assert.equal(canonicalHTML(), '<h2>initial</h2>');
  h.act(() => { editor.commands.undo(); });
  assert.equal(format.readEditableTextSelectionFormat(editor).color, 'brown');
  assert.equal(canonicalHTML(), '<h2><span class="wps-color-brown">initial</span></h2>');
  for (const color of roles) {
    h.act(() => { editor.commands.setTextSelection(4); editor.commands.setMark('titleStyle', { color }); });
    assert.equal(format.readEditableTextSelectionFormat(editor).color, color);
  }
});

test('whole-title caret style changes preserve each run, Automatic and clear have separate undo events', async t => {
  const { h, editor, format, canonicalHTML } = await fixture(t, '<h3 class="wps-line-20"><strong><span class="wps-color-brand wps-size-20">AB</span></strong><em><span class="wps-color-red wps-size-37">CD</span></em></h3>');
  const original = canonicalHTML();
  h.act(() => { editor.commands.setTextSelection(2); });
  assert.equal(format.readEditableTextToolbarFormat(editor).size, 'mixed');
  assert.equal(format.readEditableTextToolbarFormat(editor).color, 'mixed');
  h.act(() => assert.equal(format.setEditableTextFontSize(editor, '28'), true));
  assert.equal(canonicalHTML(), '<h3 class="wps-line-20"><strong><span class="wps-color-brand wps-size-28">AB</span></strong><em><span class="wps-color-red wps-size-28">CD</span></em></h3>');
  h.act(() => { format.setEditableTextColor(editor, undefined); });
  assert.equal(canonicalHTML(), '<h3 class="wps-line-20"><strong><span class="wps-size-28">AB</span></strong><em><span class="wps-size-28">CD</span></em></h3>');
  assert.equal(editor.state.selection.from, 2); assert.equal(editor.state.selection.to, 2);
  h.act(() => { format.clearEditableTextFormat(editor); }); assert.equal(canonicalHTML(), '<h3>ABCD</h3>');
  assert.deepEqual(editor.state.storedMarks, []);
  h.act(() => { editor.commands.undo(); });
  assert.equal(canonicalHTML(), '<h3 class="wps-line-20"><strong><span class="wps-size-28">AB</span></strong><em><span class="wps-size-28">CD</span></em></h3>');
  h.act(() => { editor.commands.undo(); });
  assert.match(canonicalHTML(), /wps-color-brand wps-size-28/);
  h.act(() => { editor.commands.undo(); }); assert.equal(canonicalHTML(), original);
});

test('title formatting preserves partial ranges, handles end carets and ignores breaks in whole-title state', async t => {
  const { h, editor, format, canonicalHTML } = await fixture(t, '<h2>AB<br>CD</h2>');
  h.act(() => { editor.commands.setTextSelection(6); format.toggleEditableTextMark(editor, 'bold'); });
  assert.equal(canonicalHTML(), '<h2><strong>AB</strong><br><strong>CD</strong></h2>');
  assert.equal(format.readEditableTextToolbarFormat(editor).bold, true);
  assert.equal(editor.state.selection.from, 6);
  h.act(() => { editor.commands.setTextSelection({ from: 1, to: 2 }); format.toggleEditableTextMark(editor, 'italic'); });
  assert.equal(canonicalHTML(), '<h2><strong><em>A</em></strong><strong>B</strong><br><strong>CD</strong></h2>');
  assert.equal(editor.state.selection.from, 1); assert.equal(editor.state.selection.to, 2);
  h.act(() => { editor.commands.setTextSelection(6); });
  assert.equal(format.readEditableTextToolbarFormat(editor).italic, 'mixed');
});

test('empty-title commands prepare future typing and invalid/composing commands never dispatch', async t => {
  const { h, editor, format, canonicalHTML } = await fixture(t, '<h4></h4>');
  let documentTransactions = 0;
  editor.on('transaction', ({ transaction }) => { if (transaction.docChanged) documentTransactions++; });
  h.act(() => { format.setEditableTextFontSize(editor, '37'); format.setEditableTextColor(editor, 'brand'); format.toggleEditableTextMark(editor, 'italic'); });
  assert.equal(canonicalHTML(), ''); assert.equal(documentTransactions, 0);
  assert.equal(format.readEditableTextToolbarFormat(editor).size, 37);
  assert.equal(format.readEditableTextToolbarFormat(editor).color, 'brand');
  assert.equal(format.readEditableTextToolbarFormat(editor).italic, true);
  for (const value of ['', '7', '97', '20.5', 'no']) assert.equal(format.setEditableTextFontSize(editor, value), false);
  assert.equal(format.setEditableTextColor(editor, '#fff'), false);
  const descriptor = Object.getOwnPropertyDescriptor(editor.view, 'composing');
  Object.defineProperty(editor.view, 'composing', { configurable: true, value: true });
  assert.equal(format.toggleEditableTextMark(editor, 'bold'), false);
  assert.equal(format.setEditableTextFontSize(editor, '28'), false);
  assert.equal(format.setEditableTextColor(editor, 'red'), false);
  assert.equal(format.clearEditableTextFormat(editor), false);
  if (descriptor) Object.defineProperty(editor.view, 'composing', descriptor); else delete editor.view.composing;
  h.act(() => { editor.commands.insertContent('Title'); });
  assert.equal(canonicalHTML(), '<h4><em><span class="wps-color-brand wps-size-37">Title</span></em></h4>');
  assert.equal(documentTransactions, 1);
  h.act(() => { editor.commands.undo(); }); assert.equal(canonicalHTML(), '');
});

const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./fixtures/web-part-shell/title-editor-harness.cjs');
const base = 'controls/web-part-shell/title/';

test('adapter round-trips finite block attributes, breaks, marks and rejects invalid engine input', t => {
  const h = createHarness(); t.after(() => h.close());
  const adapter = h.load(base + 'web-part-title-editor-adapter.internal.ts');
  const doc = { type: 'heading', level: 4, lineSpacing: 2.4, spaceBefore: true, spaceAfter: true,
    content: [{ type: 'text', text: 'Title', marks: { bold: true, italic: true, underline: true, size: 96, color: 'brand' } }, { type: 'hardBreak' }] };
  assert.deepEqual(adapter.editorContentToTitleDocument(adapter.titleDocumentToEditorContent(doc)), doc);
  for (const content of [{ type: 'doc', content: [] }, { type: 'doc', content: [{ type: 'heading', attrs: { level: 5 } }] },
    { type: 'doc', content: [{ type: 'paragraph', attrs: { lineSpacing: 1.5 } }] },
    { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'X', marks: [{ type: 'titleStyle', attrs: { color: '#fff' } }] }] }] }]) {
    assert.throws(() => adapter.editorContentToTitleDocument(content), /Invalid/);
  }
});

test('real engine formats undoably, instances isolate, Enter and ShiftEnter stay in one block', async t => {
  const h = createHarness(); t.after(() => h.close());
  const one = h.instance('one'); one.displayMode = 2; const two = h.instance('two'); two.displayMode = 2;
  const a = h.mount(one, 'A'); const b = h.mount(two, 'B'); await h.settle();
  const editor = h.editor(a); const other = h.editor(b);
  h.act(() => { editor.commands.setTextSelection({ from: 1, to: 8 }); editor.commands.toggleBold(); });
  assert.equal(one.properties.titleHTML, '<h2><strong>initial</strong></h2>');
  assert.equal(other.getHTML(), '<h2>initial</h2>');
  h.act(() => editor.commands.undo()); assert.equal(editor.getHTML(), '<h2>initial</h2>');
  for (const shiftKey of [false, true]) {
    h.act(() => { editor.commands.setTextSelection(3); editor.view.dom.dispatchEvent(new h.window.KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', shiftKey, bubbles: true, cancelable: true })); });
    assert.equal(editor.getJSON().content.length, 1);
  }
  assert.equal(editor.getJSON().content[0].content.filter(n => n.type === 'hardBreak').length, 2);
  assert.equal(document.querySelector('style[data-tiptap-style]'), null);
});

test('HTML paste flattens blocks, imports semantic styles, preserves first block and is one undo step', async t => {
  const h = createHarness(); t.after(() => h.close());
  const host = h.instance('paste'); host.displayMode = 2;
  const fixture = h.mount(host, 'Child'); await h.settle(); const editor = h.editor(fixture);
  h.act(() => editor.commands.setTextSelection({ from: 1, to: 8 }));
  const event = new h.window.Event('paste', { bubbles: true, cancelable: true });
  Object.defineProperty(event, 'clipboardData', { value: { getData: kind => kind === 'text/html' ? '<h4 class="wps-line-24 wps-space-before"><span class="wps-size-28 fontColorBlue"><b>One</b></span></h4><p><u>Two</u></p>' : 'One\nTwo' } });
  h.act(() => editor.view.dom.dispatchEvent(event));
  assert.equal(host.properties.titleHTML, '<h4 class="wps-line-24 wps-space-before"><strong><span class="wps-color-brand wps-size-28">One</span></strong><br><u>Two</u></h4>');
  assert.equal(host.properties.title, 'One Two');
  assert.equal(editor.getJSON().content.length, 1);
  h.act(() => editor.commands.undo()); assert.equal(editor.getHTML(), '<h2>initial</h2>');
});

test('engine refuses transactions with invalid protocol attributes without writes', async t => {
  const h = createHarness(); t.after(() => h.close());
  const host = h.instance('invalid'); host.displayMode = 2;
  const fixture = h.mount(host, 'Child'); await h.settle(); const editor = h.editor(fixture);
  h.act(() => editor.commands.updateAttributes('heading', { level: 5, lineSpacing: 1.5 }));
  assert.equal(editor.getHTML(), '<h2>initial</h2>'); assert.equal(host.properties.titleHTML, undefined);
});

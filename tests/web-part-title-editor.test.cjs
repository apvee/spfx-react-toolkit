const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./fixtures/web-part-shell/title-editor-harness.cjs');
const base = 'controls/editable-text/';

test('adapter round-trips finite block attributes, breaks, marks and rejects invalid engine input', t => {
  const h = createHarness(); t.after(() => h.close());
  const adapter = h.load(base + 'editable-text-editor-adapter.internal.ts');
  const doc = { type: 'heading', level: 4, lineSpacing: 2.4, spaceBefore: true, spaceAfter: true,
    content: [{ type: 'text', text: 'Title', marks: { bold: true, italic: true, underline: true, size: 96, color: 'brand' } }, { type: 'hardBreak' }] };
  assert.deepEqual(adapter.editorContentToEditableTextDocument(adapter.editableTextDocumentToEditorContent(doc)), doc);
  for (const content of [{ type: 'doc', content: [] }, { type: 'doc', content: [{ type: 'heading', attrs: { level: 5 } }] },
    { type: 'doc', content: [{ type: 'paragraph', attrs: { lineSpacing: 1.5 } }] },
    { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'X', marks: [{ type: 'titleStyle', attrs: { color: '#fff' } }] }] }] }]) {
    assert.throws(() => adapter.editorContentToEditableTextDocument(content), /Invalid/);
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

test('UI culture changes update editor labels and empty placeholders without replacing view state', async t => {
  const h = createHarness(); t.after(() => h.close());
  const host = h.instance('locale-retention'); host.displayMode = 2;
  host.context.pageContext.cultureInfo.currentUICultureName = 'it-IT';
  host.context.pageContext.cultureInfo.currentCultureName = 'de-DE';
  const fixture = h.mount(host, 'Child'); await h.settle();
  const editor = h.editor(fixture); const view = editor.view;
  const { resolveWebPartShellStrings: strings } = h.load('controls/web-part-shell/localization/web-part-shell-locales.internal.ts');
  assert.equal(view.dom.getAttribute('aria-label'), strings('it-IT').titleEditorLabel);
  h.act(() => { editor.commands.setTextSelection({ from: 2, to: 5 }); editor.commands.toggleBold(); });
  const saved = host.properties.titleHTML;
  host.context.pageContext.cultureInfo.currentUICultureName = 'ar-SA';
  fixture.render(host); await h.settle();
  assert.equal(h.editor(fixture), editor); assert.equal(editor.view, view);
  assert.equal(view.dom.getAttribute('aria-label'), strings('ar-SA').titleEditorLabel);
  assert.equal(editor.state.selection.from, 2); assert.equal(editor.state.selection.to, 5);
  assert.equal(host.properties.titleHTML, saved);
  h.act(() => editor.commands.undo()); assert.equal(editor.getHTML(), '<h2>initial</h2>');
  h.act(() => editor.commands.clearContent());
  assert.equal(view.dom.querySelector('[data-placeholder]').getAttribute('data-placeholder'), strings('ar-SA').titlePlaceholder);
  const emptySaved = host.properties.titleHTML;
  host.context.pageContext.cultureInfo.currentUICultureName = 'de-DE';
  fixture.render(host); await h.settle();
  assert.equal(view.dom.querySelector('[data-placeholder]').getAttribute('data-placeholder'), strings('de-DE').titlePlaceholder);
  assert.equal(host.properties.titleHTML, emptySaved);
});

test('two shell instances resolve UI language independently and preserve author text', async t => {
  const h = createHarness(); t.after(() => h.close());
  const first = h.instance('locale-it'); first.displayMode = 2;
  const second = h.instance('locale-he'); second.displayMode = 2;
  first.context.pageContext.cultureInfo.currentUICultureName = 'it-IT';
  second.context.pageContext.cultureInfo.currentUICultureName = 'he-IL';
  const a = h.mount(first, 'A'); const b = h.mount(second, 'B'); await h.settle();
  const { resolveWebPartShellStrings: strings } = h.load('controls/web-part-shell/localization/web-part-shell-locales.internal.ts');
  assert.equal(h.editor(a).view.dom.getAttribute('aria-label'), strings('it-IT').titleEditorLabel);
  assert.equal(h.editor(b).view.dom.getAttribute('aria-label'), strings('he-IL').titleEditorLabel);
  assert.equal(h.editor(a).getText(), 'initial'); assert.equal(h.editor(b).getText(), 'initial');
  assert.equal(first.properties.titleHTML, undefined); assert.equal(second.properties.titleHTML, undefined);
  first.context.pageContext.cultureInfo.currentUICultureName = 'unknown'; a.render(first); await h.settle();
  assert.equal(h.editor(a).view.dom.getAttribute('aria-label'), strings('en-US').titleEditorLabel);
  assert.equal(h.editor(b).view.dom.getAttribute('aria-label'), strings('he-IL').titleEditorLabel);
});

test('lazy editor failure and Retry follow current UI culture without touching title content', async t => {
  const React = require('react'); const originalLazy = React.lazy;
  // Simulate a failed engine chunk at its existing lazy boundary. Retry restores
  // the real lazy loader, so the test also proves recovery through that boundary.
  React.lazy = () => function FailedEditorChunk() { throw new Error('Fixture chunk unavailable'); };
  let h;
  try { h = createHarness(); } finally { React.lazy = originalLazy; }
  t.after(() => h.close());
  const host = h.instance('locale-lazy-failure'); host.displayMode = 2;
  host.context.pageContext.cultureInfo.currentUICultureName = 'it-IT';
  const logError = console.error; console.error = () => {};
  let fixture;
  try { fixture = h.mount(host, 'Child'); await h.settle(); } finally { console.error = logError; }
  const { resolveWebPartShellStrings: strings } = h.load('controls/web-part-shell/localization/web-part-shell-locales.internal.ts');
  const { resolveEditableTextStrings: commonStrings } = h.load('controls/editable-text/localization/editable-text-locales.internal.ts');
  assert.ok(fixture.container.textContent.includes(commonStrings('it-IT').editorLoadError));
  host.context.pageContext.cultureInfo.currentUICultureName = 'he-IL'; fixture.render(host); await h.settle();
  assert.ok(fixture.container.textContent.includes(commonStrings('he-IL').editorLoadError));
  const retry = [...fixture.container.querySelectorAll('button')].find(node => node.textContent === commonStrings('he-IL').retry);
  assert.ok(retry); h.act(() => retry.click()); await h.settle();
  assert.ok(h.editor(fixture));
  assert.equal(h.editor(fixture).view.dom.getAttribute('aria-label'), strings('he-IL').titleEditorLabel);
  assert.equal(host.properties.titleHTML, undefined); assert.equal(host.properties.title, 'initial');
});

test('native mark shortcuts share whole-title caret and selected-range semantics in both letter cases', async t => {
  const h = createHarness(); t.after(() => h.close());
  const host = h.instance('whole-title-keyboard'); host.displayMode = 2;
  const fixture = h.mount(host, 'Child'); await h.settle(); const editor = h.editor(fixture);
  // The JSDOM platform is non-Mac; native browser acceptance covers Cmd on Mac.
  for (const [lower, upper, tag] of [['b', 'B', 'strong'], ['i', 'I', 'em'], ['u', 'U', 'u']]) {
    h.act(() => { editor.commands.setTextSelection(4); });
    const press = key => h.act(() => { editor.view.dom.dispatchEvent(new h.window.KeyboardEvent('keydown', { key, ctrlKey: true, bubbles: true, cancelable: true })); });
    press(lower); assert.equal(host.properties.titleHTML, `<h2><${tag}>initial</${tag}></h2>`);
    assert.equal(editor.state.selection.from, 4); assert.equal(editor.state.selection.to, 4);
    press(upper); assert.equal(host.properties.titleHTML, '<h2>initial</h2>');
    h.act(() => { editor.commands.undo(); }); assert.equal(host.properties.titleHTML, `<h2><${tag}>initial</${tag}></h2>`);
    h.act(() => { editor.commands.undo(); }); assert.equal(host.properties.titleHTML, '<h2>initial</h2>');
  }
  h.act(() => { editor.commands.setTextSelection({ from: 2, to: 5 }); });
  h.act(() => { editor.view.dom.dispatchEvent(new h.window.KeyboardEvent('keydown', { key: 'b', ctrlKey: true, bubbles: true, cancelable: true })); });
  assert.equal(host.properties.titleHTML, '<h2>i<strong>nit</strong>ial</h2>');
  assert.equal(editor.state.selection.from, 2); assert.equal(editor.state.selection.to, 5);
});

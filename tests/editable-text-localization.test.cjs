const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./fixtures/editable-text/runtime-harness.cjs');

test('native resources and partial dictionaries resolve per instance, including nested color fallback', async t => {
  const h = createHarness({ nativeCulture: 'it-IT' }); t.after(() => h.close());
  const left = h.mount({ mode: 'edit', value: '', strings: { bold: 'Strong', colorNames: { red: 'Danger' } } });
  const right = h.mount({ mode: 'edit', value: '', strings: { bold: 'Heavy' } }); await h.settle();
  const { resolveEditableTextStrings } = h.load('controls/editable-text/localization/editable-text-strings.internal.ts');
  const native = resolveEditableTextStrings(); const custom = resolveEditableTextStrings({ colorNames: { red: 'Danger' }, italic: undefined });
  assert.equal(custom.colorNames.red, 'Danger'); assert.equal(custom.colorNames.blue, native.colorNames.blue);
  assert.equal(resolveEditableTextStrings({ headingLevel: level => `Section ${level}` }).headingLevel(2), 'Section 2');
  assert.equal(custom.italic, native.italic); assert.equal(native.placeholder, 'Aggiungi un testo');
  assert.equal(left.editor.view.dom.querySelector('[data-placeholder]').getAttribute('data-placeholder'), native.placeholder);
  h.act(() => left.editor.view.dom.dispatchEvent(new h.window.FocusEvent('focusin', { bubbles: true }))); await h.settle();
  assert.ok(document.querySelector('[aria-label="Strong"]'));
  h.act(() => right.editor.view.dom.dispatchEvent(new h.window.FocusEvent('focusin', { bubbles: true }))); await h.settle();
  assert.ok(document.querySelector('[aria-label="Heavy"]'));
  assert.notEqual(left.editor, right.editor); assert.deepEqual(left.calls, []); assert.deepEqual(right.calls, []);
  assert.deepEqual([...new Set(h.nativeResources.requestedIds)], ['ApveeSpfxReactToolkitEditableTextStrings']);
});

test('dictionary replacement refreshes live toolbar/color labels without losing draft, selection, marks or undo', async t => {
  const h = createHarness(); t.after(() => h.close());
  const f = h.mount({ mode: 'edit', value: '<p>Initial</p>', strings: { formattingControls: 'First controls' } }); await h.settle();
  const e = f.editor, view = e.view;
  h.act(() => e.commands.insertContent('Draft ')); h.act(() => e.commands.setTextSelection({ from: 1, to: 5 }));
  h.act(() => e.commands.setMark('bold')); const state = { html: e.getHTML(), from: e.state.selection.from, to: e.state.selection.to, marks: e.state.storedMarks };
  const count = f.calls.length;
  h.act(() => view.dom.dispatchEvent(new h.window.FocusEvent('focusin', { bubbles: true }))); await h.settle();
  f.render({ strings: { formattingControls: 'Latest controls', bold: 'Strong', themeColor: 'Ink', colorNames: { red: 'Danger' } } }); await h.settle();
  assert.ok(document.querySelector('[aria-label="Latest controls"]')); assert.ok(document.querySelector('[aria-label="Strong"]'));
  const color = document.querySelector('[aria-label="Ink"]'); assert.ok(color); h.act(() => color.click()); await h.settle();
  assert.ok(document.querySelector('[data-title-color="red"][aria-label="Danger"]'));
  assert.equal(f.editor, e); assert.equal(e.view, view); assert.equal(e.getHTML(), state.html);
  assert.equal(e.state.selection.from, state.from); assert.equal(e.state.selection.to, state.to); assert.deepEqual(e.state.storedMarks, state.marks);
  assert.equal(f.calls.length, count); h.act(() => e.commands.undo()); assert.notEqual(e.getHTML(), state.html);
});

test('dictionary hints update while explicit empty placeholder and direct accessible names retain precedence', async t => {
  const h = createHarness(); t.after(() => h.close());
  const f = h.mount({ mode: 'edit', value: '', strings: { placeholder: 'First hint' }, 'aria-label': 'Explicit name', locale: { ignored: true } }); await h.settle();
  const e = f.editor; assert.equal(e.view.dom.querySelector('[data-placeholder]').getAttribute('data-placeholder'), 'First hint');
  f.render({ strings: { placeholder: 'New hint' } }); assert.equal(e.view.dom.querySelector('[data-placeholder]').getAttribute('data-placeholder'), 'New hint');
  f.render({ placeholder: '', strings: { placeholder: 'Ignored hint' } });
  assert.equal(e.view.dom.querySelector('[data-placeholder]')?.getAttribute('data-placeholder') ?? '', ''); assert.equal(e.view.dom.getAttribute('aria-label'), 'Explicit name');
  assert.equal(f.root.hasAttribute('locale'), false); assert.equal(f.root.hasAttribute('strings'), false); assert.deepEqual(f.calls, []);
});

test('loading, invalid-input, oversize and chunk recovery labels follow dictionary replacement', async t => {
  const h = createHarness({ failEngineLoads: 1 }); t.after(() => h.close());
  const f = h.mount({ mode: 'edit', strings: { loadingEditor: 'Preparing editor', editorLoadError: 'Unavailable', retry: 'Again' } });
  assert.ok(f.root.textContent.includes('Preparing editor'));
  const previous = console.error; console.error = () => {}; try { await h.settle(); } finally { console.error = previous; }
  assert.ok(f.root.textContent.includes('Unavailable')); f.render({ strings: { editorLoadError: 'Latest failure', retry: 'Try now' } });
  assert.ok(f.root.textContent.includes('Latest failure')); const retry = [...f.root.querySelectorAll('button')].find(b => b.textContent === 'Try now'); assert.ok(retry);
  h.act(() => retry.click()); await h.settle();
  f.render({ value: '<div>'.repeat(33) + 'bad' + '</div>'.repeat(33), strings: { savedFormattingWarning: 'Recovered content', tooLargeError: 'Size limit' } });
  assert.ok(f.root.textContent.includes('Recovered content'));
  h.act(() => f.editor.commands.setContent({ type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: '&'.repeat(14000) }] }] }));
  assert.ok(f.root.textContent.includes('Size limit')); f.render({ strings: { tooLargeError: 'Latest size error' } });
  assert.ok(f.root.textContent.includes('Latest size error')); assert.deepEqual(f.calls, []);
});

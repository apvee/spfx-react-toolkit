const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./fixtures/web-part-shell/title-toolbar-harness.cjs');
const base = 'controls/editable-text/';
async function setup(t, html) {
  const h = createHarness(); t.after(() => h.close());
  const host = h.instance('toolbar'); host.displayMode = 2;
  if (html) host.properties.titleHTML = html;
  const fixture = h.mount(host, 'Child'); await h.settle();
  const editor = h.editor(fixture);
  return { h, host, fixture, editor };
}
function button(label) { return label === 'Brand' ? document.querySelector('[data-title-color="brand"]') : document.querySelector(`[aria-label="${label}"]`); }

test('Fluent toolbar edits the selected range through the existing publisher and selection alone does not write', async t => {
  const { h, host, editor } = await setup(t);
  h.act(() => { editor.commands.setTextSelection({ from: 2, to: 5 }); editor.view.dom.dispatchEvent(new h.window.FocusEvent('focusin', { bubbles: true })); });
  assert.ok(document.querySelector('[role="toolbar"][aria-label="Text formatting"]'), 'Edit focus reveals the Fluent title toolbar');
  assert.equal(host.properties.titleHTML, undefined);
  const bold = button('Bold'); assert.ok(bold);
  h.act(() => bold.click());
  assert.equal(host.properties.titleHTML, '<h2>i<strong>nit</strong>ial</h2>');
  assert.equal(editor.state.selection.from, 2); assert.equal(editor.state.selection.to, 5);
  h.act(() => editor.commands.undo()); assert.equal(editor.getHTML(), '<h2>initial</h2>');
});

test('format state reports actual mixed selection and collapsed stored marks', async t => {
  const { h, editor } = await setup(t, '<h3><strong><span class="wps-color-brand wps-size-28">AB</span></strong>C</h3>');
  const p = h.load(base + 'editable-text-formatting.internal.ts');
  h.act(() => editor.commands.setTextSelection({ from: 1, to: 4 }));
  const mixed = p.readEditableTextSelectionFormat(editor);
  assert.equal(mixed.bold, 'mixed'); assert.equal(mixed.size, 'mixed'); assert.equal(mixed.color, 'mixed');
  assert.equal(mixed.block.type, 'heading'); assert.equal(mixed.block.level, 3);
  h.act(() => { editor.commands.setTextSelection(4); editor.commands.setMark('titleStyle', { size: 37, color: 'subtle' }); editor.commands.toggleItalic(); });
  const caret = p.readEditableTextSelectionFormat(editor);
  assert.equal(caret.size, 37); assert.equal(caret.color, 'subtle'); assert.equal(caret.italic, true);
});



test('clear removes only selection overrides and block spacing while preserving heading and text', async t => {
  const { h, editor } = await setup(t, '<h3 class="wps-line-20 wps-space-before wps-space-after"><strong><em><u><span class="wps-color-brand wps-size-28">ABCD</span></u></em></strong></h3>');
  const p = h.load(base + 'editable-text-formatting.internal.ts');
  h.act(() => { editor.commands.setTextSelection({ from: 2, to: 4 }); p.clearEditableTextFormat(editor); });
  assert.equal(editor.getHTML(), '<h3><strong><em><u><span class="wps-color-brand wps-size-28">A</span></u></em></strong>BC<strong><em><u><span class="wps-color-brand wps-size-28">D</span></u></em></strong></h3>');
  h.act(() => editor.commands.undo()); assert.match(editor.getHTML(), /wps-line-20 wps-space-before wps-space-after/);
});

test('integer size validation rejects invalid and fractional values without changing format', async t => {
  const { h, editor, host } = await setup(t);
  const p = h.load(base + 'editable-text-formatting.internal.ts');
  h.act(() => editor.commands.setTextSelection({ from: 1, to: 8 }));
  for (const value of ['', '7', '97', '20.5', '1e1', 'no']) {
    h.act(() => assert.equal(p.setEditableTextFontSize(editor, value), false));
    assert.equal(host.properties.titleHTML, undefined);
  }
  h.act(() => assert.equal(p.setEditableTextFontSize(editor, '37'), true));
  assert.equal(host.properties.titleHTML, '<h2><span class="wps-size-37">initial</span></h2>');
});

function reveal(h, editor) { h.act(() => editor.view.dom.dispatchEvent(new h.window.FocusEvent('focusin', { bubbles: true }))); }
function menuItem(label) { return [...document.querySelectorAll('[role^="menuitem"], [role="option"]')].find(node => node.textContent.trim() === label); }
function click(h, node) { assert.ok(node); h.act(() => node.click()); }

test('Heading, Normal, custom size, B/I/U and theme palette controls retain the selected range', async t => {
  const { h, editor, host, fixture } = await setup(t);
  h.act(() => editor.commands.setTextSelection({ from: 2, to: 5 })); reveal(h, editor);
  click(h, button('Heading')); click(h, menuItem('Heading 4'));
  assert.equal(host.properties.titleHTML, '<h4>initial</h4>');
  click(h, button('Heading')); click(h, menuItem('Normal'));
  assert.equal(host.properties.titleHTML, '<p>initial</p>');
  const { Simulate } = require('react-dom/test-utils');
  const input = button('Font size');
  h.act(() => Simulate.change(input, { target: { value: '37' } }));
  h.act(() => Simulate.keyDown(input, { key: 'Enter' }));
  assert.equal(host.properties.titleHTML, '<p>i<span class="wps-size-37">nit</span>ial</p>');
  for (const label of ['Bold', 'Italic', 'Underline']) click(h, button(label));
  assert.equal(host.properties.titleHTML, '<p>i<strong><em><u><span class="wps-size-37">nit</span></u></em></strong>ial</p>');
  click(h, button('Theme color'));
  assert.ok(document.querySelector('[data-title-color="automatic"]'));
  assert.ok(document.querySelector('[data-title-color="brand"]'));
  click(h, button('Brand'));
  assert.equal(host.properties.titleHTML, '<p>i<strong><em><u><span class="wps-color-brand wps-size-37">nit</span></u></em></strong>ial</p>');
  assert.equal(editor.state.selection.from, 2); assert.equal(editor.state.selection.to, 5);
  assert.ok(fixture.focusRequests >= 6);
});

test('spacing menu has preset default, six finite overrides and before/after toggles; clear preserves the heading', async t => {
  const { h, editor, host } = await setup(t);
  reveal(h, editor); click(h, button('Spacing'));
  assert.equal(document.querySelectorAll('[name="lineSpacing"]').length, 7);
  assert.equal(menuItem('Default line spacing').getAttribute('aria-checked'), 'true');
  click(h, menuItem('Line spacing 2.4')); assert.equal(host.properties.titleHTML, '<h2 class="wps-line-24">initial</h2>');
  // The fixture counts view.focus requests without native Selection/Range.
  // Reveal explicitly again; actual restored focus continuity is Task9.
  reveal(h, editor);
  click(h, button('Spacing')); assert.ok(menuItem('Space before'), 'Reopened spacing exposes before'); click(h, menuItem('Space before'));
  reveal(h, editor); click(h, button('Spacing')); click(h, menuItem('Space after'));
  assert.equal(host.properties.titleHTML, '<h2 class="wps-line-24 wps-space-after wps-space-before">initial</h2>');
  reveal(h, editor); click(h, button('Clear formatting')); assert.equal(host.properties.titleHTML, '<h2>initial</h2>');
});

test('font picker follows Fluent presets and live theme tokens without writing inherited sizes', async t => {
  const { h, host, fixture, editor } = await setup(t);
  reveal(h, editor);
  for (const [choice, size] of [['Heading 1', '32'], ['Heading 2', '28'], ['Heading 3', '24'], ['Heading 4', '20'], ['Normal', '14']]) {
    click(h, button('Heading')); click(h, menuItem(choice));
    assert.equal(button('Font size').value, size);
    assert.doesNotMatch(host.properties.titleHTML, /wps-size-/);
  }
  click(h, button('Heading')); click(h, menuItem('Heading 2'));
  const saved = host.properties.titleHTML;
  fixture.render(host, { theme: { fontSizeHero700: '40px' } }); await h.settle();
  assert.equal(button('Font size').value, '40');
  assert.equal(host.properties.titleHTML, saved);
  h.act(() => editor.commands.setMark('titleStyle', { size: 37 }));
  fixture.render(host, { theme: { fontSizeHero700: '44px' } }); await h.settle();
  assert.equal(button('Font size').value, '44', 'toolbar reads inherited whole-title size, not a native stored-caret override');
  assert.equal(h.load(base + 'editable-text-formatting.internal.ts').readEditableTextSelectionFormat(editor).size, 37);
});

test('default line spacing removes only its override and undo restores it', async t => {
  const { h, host, editor } = await setup(t, '<h2 class="wps-line-20 wps-space-before">initial</h2>');
  reveal(h, editor); click(h, button('Spacing')); click(h, menuItem('Default line spacing'));
  assert.equal(host.properties.titleHTML, '<h2 class="wps-space-before">initial</h2>');
  h.act(() => editor.commands.undo());
  assert.equal(host.properties.titleHTML, '<h2 class="wps-line-20 wps-space-before">initial</h2>');
});



test('collapsed clear removes all title marks while preserving text, heading and its undo boundary', async t => {
  const { h, editor } = await setup(t, '<h4><strong>AB</strong><em>CD</em></h4>');
  const p = h.load(base + 'editable-text-formatting.internal.ts');
  h.act(() => editor.commands.setTextSelection(5));
  assert.equal(p.readEditableTextSelectionFormat(editor).italic, true);
  h.act(() => p.clearEditableTextFormat(editor));
  assert.equal(p.readEditableTextSelectionFormat(editor).italic, false);
  assert.deepEqual(editor.state.storedMarks, []);
  assert.equal(editor.getHTML(), '<h4>ABCD</h4>');
  h.act(() => editor.commands.undo());
  assert.equal(editor.getHTML(), '<h4><strong>AB</strong><em>CD</em></h4>');
});

test('mixed controls expose empty size and mixed pressed state, Tab remains uncancelled', async t => {
  const { h, editor, host } = await setup(t, '<h2><strong><span class="wps-size-37">AB</span></strong>CD</h2>');
  h.act(() => editor.commands.setTextSelection({ from: 1, to: 5 })); reveal(h, editor);
  assert.equal(button('Bold').getAttribute('aria-pressed'), 'mixed');
  assert.equal(button('Font size').value, ''); assert.equal(button('Font size').getAttribute('placeholder'), 'Mixed');
  const event = new h.window.KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true });
  h.act(() => button('Font size').dispatchEvent(event)); assert.equal(event.defaultPrevented, false);
  assert.equal(host.properties.titleHTML, '<h2><strong><span class="wps-size-37">AB</span></strong>CD</h2>');
});

test('toolbar keyboard entry focuses Heading and leaves Tab available', async t => {
  const { h, editor } = await setup(t);
  reveal(h, editor);
  h.act(() => editor.view.dom.dispatchEvent(new h.window.KeyboardEvent('keydown', { key: 'F10', altKey: true, bubbles: true, cancelable: true })));
  await h.settle(); assert.equal(document.activeElement.getAttribute('aria-label'), 'Heading');
  const tab = new h.window.KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true });
  h.act(() => document.activeElement.dispatchEvent(tab)); assert.equal(tab.defaultPrevented, false);
});



test('UI caret toggles turn whole-title B/I/U on and off while preserving the caret', async t => {
  const { h, editor, host } = await setup(t);
  h.act(() => editor.commands.setTextSelection(3)); reveal(h, editor);
  for (const label of ['Bold', 'Italic', 'Underline']) {
    click(h, button(label)); assert.equal(button(label).getAttribute('aria-pressed'), 'true');
    click(h, button(label)); assert.equal(button(label).getAttribute('aria-pressed'), 'false', `${label} turns off on the second caret click`);
  }
  assert.equal(host.properties.titleHTML, '<h2>initial</h2>');
  assert.equal(editor.state.selection.from, 3); assert.equal(editor.state.selection.to, 3);
});

test('UI caret formats accumulate and future typing retains stored size color and marks', async t => {
  const { h, editor, host } = await setup(t);
  h.act(() => editor.commands.setTextSelection(3)); reveal(h, editor);
  const { Simulate } = require('react-dom/test-utils');
  h.act(() => Simulate.change(button('Font size'), { target: { value: '37' } }));
  h.act(() => Simulate.keyDown(button('Font size'), { key: 'Enter' }));
  click(h, button('Theme color')); click(h, button('Brand')); reveal(h, editor); click(h, button('Italic'));
  assert.equal(button('Font size').value, '37');
  assert.equal(button('Italic').getAttribute('aria-pressed'), 'true');
  assert.equal(editor.state.storedMarks.find(mark => mark.type.name === 'titleStyle').attrs.color, 'brand');
  assert.equal(host.properties.titleHTML, '<h2><em><span class="wps-color-brand wps-size-37">initial</span></em></h2>');
  h.act(() => editor.commands.insertContent('X'));
  assert.equal(host.properties.titleHTML, '<h2><em><span class="wps-color-brand wps-size-37">inXitial</span></em></h2>');
  h.act(() => editor.commands.undo()); assert.equal(editor.getHTML(), '<h2><em><span class="wps-color-brand wps-size-37">initial</span></em></h2>');
});

test('UI caret controls preserve explicitly empty stored marks instead of surrounding bold', async t => {
  const { h, editor, host } = await setup(t, '<h2><strong>ABCD</strong></h2>');
  h.act(() => editor.commands.setTextSelection(2)); reveal(h, editor);
  click(h, button('Bold')); assert.deepEqual(editor.state.storedMarks, []);
  click(h, button('Italic'));
  h.act(() => editor.commands.insertContent('X'));
  assert.equal(host.properties.titleHTML, '<h2><em>AXBCD</em></h2>');
});

test('editor click keeps the nonmodal toolbar open while real outside interaction dismisses it', async t => {
  const { h, editor, host } = await setup(t);
  reveal(h, editor);
  const clickEvent = new h.window.MouseEvent('click', { bubbles: true, cancelable: true });
  h.act(() => editor.view.dom.dispatchEvent(clickEvent));
  assert.ok(document.querySelector('[data-title-toolbar]'), 'editor remains inside the toolbar interaction area');
  assert.equal(clickEvent.defaultPrevented, false);
  assert.equal(host.properties.titleHTML, undefined);
  const outside = document.createElement('button'); document.body.appendChild(outside);
  h.act(() => outside.click()); assert.equal(document.querySelector('[data-title-toolbar]'), null);
  outside.remove();
});

test('editor touchstart retains toolbar and outside touch dismisses without formatting', async t => {
  const { h, editor, host } = await setup(t, '<h2><strong>AB</strong>CD</h2>');
  reveal(h, editor);
  h.act(() => editor.view.dom.dispatchEvent(new h.window.Event('touchstart', { bubbles: true, cancelable: true })));
  assert.ok(document.querySelector('[data-title-toolbar]'));
  h.act(() => document.body.dispatchEvent(new h.window.Event('touchstart', { bubbles: true, cancelable: true })));
  assert.equal(Boolean(document.querySelector('[data-title-toolbar]')), false);
  assert.equal(host.properties.titleHTML, '<h2><strong>AB</strong>CD</h2>');
});

for (const popup of ['palette', 'spacing']) test(`owned ${popup} popup touch and focus retain the parent toolbar and allow its command`, async t => {
  const { h, editor, host } = await setup(t);
  h.act(() => editor.commands.setTextSelection({ from: 2, to: 5 })); reveal(h, editor);
  click(h, button(popup === 'palette' ? 'Theme color' : 'Spacing'));
  const target = popup === 'palette' ? button('Brand') : menuItem('Line spacing 2');
  assert.ok(target);
  const touch = new h.window.Event('touchstart', { bubbles: true, cancelable: true });
  h.act(() => target.dispatchEvent(touch));
  assert.ok(document.querySelector('[data-title-toolbar]'), 'touch on this editor popup retains the outer toolbar');
  assert.equal(touch.defaultPrevented, false);
  h.act(() => target.dispatchEvent(new h.window.FocusEvent('focusin', { bubbles: true })));
  assert.ok(document.querySelector('[data-title-toolbar]'), 'focus on this editor popup retains the outer toolbar');
  assert.equal(host.properties.titleHTML, undefined);
  click(h, target);
  assert.equal(host.properties.titleHTML, popup === 'palette' ? '<h2>i<span class="wps-color-brand">nit</span>ial</h2>' : '<h2 class="wps-line-20">initial</h2>');
  assert.equal(editor.state.selection.from, 2); assert.equal(editor.state.selection.to, 5);
  await h.settle();
});

test('owned popup correction keeps another editor outside and preserves toolbar Escape dismissal', async t => {
  const { h, editor, host } = await setup(t);
  const second = h.instance('other-popup-editor'); second.displayMode = 2;
  const secondFixture = h.mount(second, 'Other'); await h.settle(); const other = h.editor(secondFixture);
  reveal(h, editor); click(h, button('Theme color'));
  h.act(() => other.view.dom.dispatchEvent(new h.window.Event('touchstart', { bubbles: true, cancelable: true })));
  assert.equal(document.querySelector('[data-title-toolbar]'), null);
  assert.equal(host.properties.titleHTML, undefined); assert.equal(second.properties.titleHTML, undefined);
  reveal(h, editor);
  h.act(() => document.querySelector('[data-title-toolbar]').dispatchEvent(new h.window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })));
  assert.equal(document.querySelector('[data-title-toolbar]'), null);
  assert.equal(host.properties.titleHTML, undefined);
});

test('default popup mounts retain the iframe owner document and its popup interaction area', async t => {
  const h = createHarness(); t.after(() => h.close());
  const frame = document.createElement('iframe'); document.body.appendChild(frame);
  const doc = frame.contentDocument; doc.getSelection = () => null;
  doc.defaultView.ResizeObserver = h.window.ResizeObserver;
  Object.defineProperty(doc.defaultView, 'crypto', { configurable: true, value: h.window.crypto });
  const host = h.instance('popup-frame'); host.displayMode = 2; host.context.domElement = doc.createElement('div');
  const fixture = h.mount(host, 'Child', {}, { document: doc }); await h.settle(); const editor = h.editor(fixture);
  h.act(() => editor.commands.setTextSelection({ from: 1, to: 8 }));
  h.act(() => editor.view.dom.dispatchEvent(new doc.defaultView.FocusEvent('focusin', { bubbles: true })));
  const toolbar = doc.querySelector('[data-title-toolbar]'); assert.ok(toolbar); assert.equal(document.querySelector('[data-title-toolbar]'), null);
  await h.settle();
  const color = doc.querySelector('[aria-label="Theme color"]'); click(h, color);
  await h.settle();
  const brand = doc.querySelector('[data-title-color="brand"]'); assert.ok(brand); assert.equal(brand.ownerDocument, doc);
  h.act(() => brand.dispatchEvent(new doc.defaultView.Event('touchstart', { bubbles: true, cancelable: true })));
  assert.ok(doc.querySelector('[data-title-toolbar]'));
  click(h, brand); assert.equal(host.properties.titleHTML, '<h2><span class="wps-color-brand">initial</span></h2>');
  await h.settle();
  fixture.unmount(); frame.remove();
});

test('Heading and size options restore editor focus after Fluent listbox click handling', async t => {
  const { h, editor, fixture, host } = await setup(t);
  h.act(() => editor.commands.setTextSelection({ from: 2, to: 5 })); reveal(h, editor);
  click(h, button('Heading'));
  const beforeHeading = fixture.focusRequests;
  click(h, menuItem('Heading 3'));
  const synchronousHeading = fixture.focusRequests;
  await h.settle();
  assert.ok(synchronousHeading > beforeHeading);
  assert.ok(fixture.focusRequests > synchronousHeading, 'restore after Fluent listbox bubble handler reclaims trigger focus');
  assert.equal(editor.state.selection.from, 2); assert.equal(editor.state.selection.to, 5);
  reveal(h, editor); click(h, button('Font size'));
  const beforeSize = fixture.focusRequests;
  click(h, menuItem('28'));
  const synchronousSize = fixture.focusRequests;
  await h.settle();
  assert.ok(synchronousSize > beforeSize);
  assert.ok(fixture.focusRequests > synchronousSize, 'size option shares the same Fluent listbox focus contract');
  assert.equal(host.properties.titleHTML, '<h3>i<span class="wps-size-28">nit</span>ial</h3>');
  assert.equal(editor.state.selection.from, 2); assert.equal(editor.state.selection.to, 5);
});

test('overflow measures and hides whole Fluent composite controls instead of their trigger slots', async t => {
  const { h, editor } = await setup(t); reveal(h, editor);
  const heading = button('Heading').closest('.fui-Dropdown');
  const size = button('Font size').closest('.fui-Combobox');
  assert.equal(heading.parentElement.hasAttribute('data-overflow-item'), true);
  assert.equal(size.parentElement.hasAttribute('data-overflow-item'), true);
  assert.equal(button('Heading').hasAttribute('data-overflow-item'), false);
  assert.equal(button('Font size').hasAttribute('data-overflow-item'), false);
  assert.equal([...document.querySelectorAll('[role="toolbar"] [role="separator"]')].every(node => node.hasAttribute('data-overflow-item')), true, 'separator footprints participate in Fluent overflow measurement');
});



test('V1 toolbar omits format painter and selection gestures never copy formatting', async t => {
  const { h, editor, host } = await setup(t, '<h2><strong>AB</strong>CD</h2>');
  reveal(h, editor);
  assert.equal(Boolean(button('Format painter')), false);
  assert.equal(document.querySelector('[data-overflow-item="painter"]'), null);
  h.act(() => {
    editor.commands.setTextSelection({ from: 3, to: 5 });
    editor.view.dom.dispatchEvent(new h.window.Event('pointerup', { bubbles: true }));
    editor.view.dom.dispatchEvent(new h.window.KeyboardEvent('keyup', { key: 'Shift', bubbles: true }));
  });
  await h.settle();
  assert.equal(host.properties.titleHTML, '<h2><strong>AB</strong>CD</h2>');
});

test('Automatic removes only selected color, preserves size and other marks, and is undoable', async t => {
  const { h, editor, host } = await setup(t, '<h2><strong><span class="wps-color-brand wps-size-37">ABCD</span></strong></h2>');
  h.act(() => editor.commands.setTextSelection({ from: 2, to: 4 })); reveal(h, editor);
  click(h, button('Theme color')); click(h, document.querySelector('[data-title-color="automatic"]'));
  assert.equal(host.properties.titleHTML, '<h2><strong><span class="wps-color-brand wps-size-37">A</span></strong><strong><span class="wps-size-37">BC</span></strong><strong><span class="wps-color-brand wps-size-37">D</span></strong></h2>');
  assert.equal(editor.state.selection.from, 2); assert.equal(editor.state.selection.to, 4);
  h.act(() => editor.commands.undo()); assert.match(editor.getHTML(), /wps-color-brand wps-size-37/);
});

test('picker deduplicates host tokens while retaining Automatic and the active semantic color', async t => {
  const { h, editor, fixture, host } = await setup(t, '<h2><span class="wps-color-brand">initial</span></h2>');
  fixture.render(host, { theme: { colorNeutralForeground1: '#123456', colorNeutralForeground2: '#123456', colorBrandForeground1: '#123456' } });
  await h.settle(); reveal(h, editor); click(h, button('Theme color'));
  assert.ok(document.querySelector('[data-title-color="brand"]'));
  assert.ok(document.querySelector('[data-title-color="automatic"]'));
  assert.equal(document.querySelector('[data-title-color="primary"]'), null);
  assert.equal(document.querySelector('[data-title-color="subtle"]'), null);
  assert.equal(host.properties.titleHTML, '<h2><span class="wps-color-brand">initial</span></h2>');
});

test('localized toolbar commands and palette labels update with the same editor selection', async t => {
  const { h, host, editor, fixture } = await setup(t);
  const strings = culture => require('./fixtures/localization/editor-dictionaries.cjs').editorDictionary(h, culture);
  h.act(() => editor.commands.setTextSelection({ from: 2, to: 5 })); reveal(h, editor);
  fixture.render(host, { titleEditorStrings: strings('it-IT') }); await h.settle();
  click(h, button(strings('it-IT').themeColor));
  for (const swatch of document.querySelectorAll('[data-title-color]')) {
    const id = swatch.getAttribute('data-title-color');
    assert.equal(swatch.getAttribute('aria-label'), id === 'automatic' ? strings('it-IT').automatic : strings('it-IT').colorNames[id]);
  }
  click(h, document.querySelector('[data-title-color="brand"]'));
  assert.equal(host.properties.titleHTML, '<h2>i<span class="wps-color-brand">nit</span>ial</h2>');
  assert.equal(editor.state.selection.from, 2); assert.equal(editor.state.selection.to, 5);
  assert.ok(button(strings('it-IT').bold));
});

test('compact heading presentation retains its complete localized value across command updates', async t => {
  const { h, host, editor, fixture } = await setup(t);
  const strings = culture => require('./fixtures/localization/editor-dictionaries.cjs').editorDictionary(h, culture);
  fixture.render(host, { titleEditorStrings: strings('de-DE') }); await h.settle();
  h.act(() => editor.commands.setTextSelection({ from: 2, to: 5 })); reveal(h, editor);
  const heading = button(strings('de-DE').heading);
  assert.equal(heading.getAttribute('value'), strings('de-DE').headingLevel(2));
  assert.equal(heading.querySelector('span[title]').textContent, strings('de-DE').headingLevel(2));
  assert.equal(heading.querySelector('span[title]').getAttribute('title'), strings('de-DE').headingLevel(2));
  click(h, heading); click(h, menuItem(strings('de-DE').headingLevel(4)));
  assert.equal(button(strings('de-DE').heading).querySelector('span[title]').getAttribute('title'), strings('de-DE').headingLevel(4));
  assert.equal(editor.state.selection.from, 2); assert.equal(editor.state.selection.to, 5);
  assert.equal(host.properties.titleHTML, '<h4>initial</h4>');
});

test('palette sections preserve ordered accessible groups and each swatch belongs to its semantic row', async t => {
  const { h, editor } = await setup(t); reveal(h, editor); click(h, button('Theme color'));
  const sections = [...document.querySelectorAll('[data-title-color-section]')];
  assert.deepEqual(sections.map(section => section.getAttribute('data-title-color-section')), ['automatic', 'theme', 'palette']);
  assert.deepEqual(sections.map(section => section.getAttribute('aria-label')), ['Automatic', 'Theme colors', 'Standard colors']);
  assert.ok(sections.every(section => section.getAttribute('role') === 'group'));
  assert.equal(document.querySelector('[data-title-color="automatic"]').closest('[data-title-color-row]').getAttribute('data-title-color-row'), 'automatic');
  assert.equal(document.querySelector('[data-title-color="brand"]').closest('[data-title-color-row]').getAttribute('data-title-color-row'), 'theme');
  assert.equal(document.querySelector('[data-title-color="red"]').closest('[data-title-color-row]').getAttribute('data-title-color-row'), 'palette');
});

test('font-size expand action uses the current localized name without Fluent English prefixes', async t => {
  const { h, host, editor, fixture } = await setup(t);
  const strings = culture => require('./fixtures/localization/editor-dictionaries.cjs').editorDictionary(h, culture);
  reveal(h, editor);
  for (const culture of ['it-IT', 'de-DE', 'ar-SA']) {
    fixture.render(host, { titleEditorStrings: strings(culture) }); await h.settle();
    const expand = document.querySelector('.fui-Combobox__expandIcon');
    assert.equal(expand.getAttribute('aria-label'), strings(culture).fontSize);
    assert.equal(button(strings(culture).fontSize).getAttribute('role'), 'combobox');
    assert.equal(host.properties.titleHTML, undefined);
  }
  const expand = document.querySelector('.fui-Combobox__expandIcon');
  // Fluent's editable combobox expands on mousedown so it can focus its input.
  h.act(() => expand.dispatchEvent(new h.window.MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })));
  assert.ok(menuItem('28'), 'the localized expand action still opens the font-size choices');
});

test('collapsed toolbar Bold applies to the entire title, preserves the caret and publishes one undoable transaction', async t => {
  const { h, host, editor } = await setup(t);
  h.act(() => editor.commands.setTextSelection(3)); reveal(h, editor);
  let documentTransactions = 0;
  editor.on('transaction', ({ transaction }) => { if (transaction.docChanged) documentTransactions++; });
  click(h, button('Bold'));
  assert.equal(host.properties.titleHTML, '<h2><strong>initial</strong></h2>');
  assert.equal(editor.state.selection.from, 3); assert.equal(editor.state.selection.to, 3);
  assert.equal(documentTransactions, 1);
  assert.equal(button('Bold').getAttribute('aria-pressed'), 'true');
  click(h, button('Bold'));
  assert.equal(host.properties.titleHTML, '<h2>initial</h2>');
  assert.equal(documentTransactions, 2);
  h.act(() => editor.commands.undo()); assert.equal(host.properties.titleHTML, '<h2><strong>initial</strong></h2>');
  assert.equal(editor.state.selection.from, 3);
  h.act(() => editor.commands.undo()); assert.equal(host.properties.titleHTML, '<h2>initial</h2>');
});

test('collapsed toolbar state reads mixed whole-title marks, applies all and then removes all', async t => {
  const { h, host, editor } = await setup(t, '<h2><strong>AB</strong>CD</h2>');
  h.act(() => editor.commands.setTextSelection(2)); reveal(h, editor);
  assert.equal(button('Bold').getAttribute('aria-pressed'), 'mixed');
  click(h, button('Bold'));
  assert.equal(host.properties.titleHTML, '<h2><strong>ABCD</strong></h2>');
  assert.equal(button('Bold').getAttribute('aria-pressed'), 'true');
  click(h, button('Bold'));
  assert.equal(host.properties.titleHTML, '<h2>ABCD</h2>');
  assert.equal(button('Bold').getAttribute('aria-pressed'), 'false');
  assert.equal(editor.state.selection.from, 2);
});

test('empty-title toolbar styles prepare future typing without publishing and composing clicks never format', async t => {
  const { h, host, editor } = await setup(t, '<h2></h2>');
  const before = host.properties.titleHTML; reveal(h, editor);
  const { Simulate } = require('react-dom/test-utils');
  h.act(() => { Simulate.change(button('Font size'), { target: { value: '37' } }); });
  h.act(() => { Simulate.keyDown(button('Font size'), { key: 'Enter' }); });
  click(h, button('Theme color')); click(h, button('Brand')); click(h, button('Bold'));
  assert.equal(host.properties.titleHTML, before);
  assert.equal(button('Font size').value, '37'); assert.equal(button('Bold').getAttribute('aria-pressed'), 'true');
  h.act(() => { editor.commands.insertContent('Title'); });
  assert.equal(host.properties.titleHTML, '<h2><strong><span class="wps-color-brand wps-size-37">Title</span></strong></h2>');
  const saved = host.properties.titleHTML;
  const descriptor = Object.getOwnPropertyDescriptor(editor.view, 'composing');
  Object.defineProperty(editor.view, 'composing', { configurable: true, value: true });
  click(h, button('Italic')); click(h, button('Clear formatting'));
  assert.equal(host.properties.titleHTML, saved);
  if (descriptor) Object.defineProperty(editor.view, 'composing', descriptor); else delete editor.view.composing;
  h.act(() => { editor.commands.undo(); }); assert.equal(host.properties.titleHTML, '');
});

test('rapid heading and spacing toolbar gestures remain separate undo events on the native transaction', async t => {
  const { h, host, editor } = await setup(t); reveal(h, editor);
  h.act(() => { editor.commands.setTextSelection(3); });
  click(h, button('Heading')); click(h, menuItem('Heading 4'));
  click(h, button('Heading')); click(h, menuItem('Normal'));
  assert.equal(host.properties.titleHTML, '<p>initial</p>');
  h.act(() => { editor.commands.undo(); }); assert.equal(host.properties.titleHTML, '<h4>initial</h4>');
  h.act(() => { editor.commands.undo(); }); assert.equal(host.properties.titleHTML, '<h2>initial</h2>');
  reveal(h, editor); click(h, button('Spacing')); click(h, menuItem('Line spacing 2'));
  reveal(h, editor); click(h, button('Spacing')); click(h, menuItem('Space before'));
  reveal(h, editor); click(h, button('Spacing')); click(h, menuItem('Space after'));
  h.act(() => { editor.commands.undo(); }); assert.equal(host.properties.titleHTML, '<h2 class="wps-line-20 wps-space-before">initial</h2>');
  h.act(() => { editor.commands.undo(); }); assert.equal(host.properties.titleHTML, '<h2 class="wps-line-20">initial</h2>');
  h.act(() => { editor.commands.undo(); }); assert.equal(host.properties.titleHTML, '<h2>initial</h2>');
  assert.equal(editor.state.selection.from, 3); assert.equal(editor.state.selection.to, 3);
});

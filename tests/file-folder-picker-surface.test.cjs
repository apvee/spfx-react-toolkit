const test = require('node:test');
const assert = require('node:assert/strict');
const { createPickerHarness } = require('./fixtures/picker-source-harness.cjs');
const value = { kind: 'file', tenantId: '11111111-1111-4111-8111-111111111111', endpoint: 'https://tenant.sharepoint.com/_api/v2.0', driveId: 'drive', itemId: 'one', name: 'Report.pdf' };
function setup() {
  const actions = [];
  const runtime = { values: [], label: v => v.name, extension: () => '.pdf', remove: v => actions.push(v.itemId), browse: () => actions.push('browse'), browseButton: { current: null }, busy: false, ready: true, cancel() {}, attachFrame() {} };
  const h = createPickerHarness({ './useFileFolderPicker.internal': { useFileFolderPicker: () => runtime } });
  const { FileFolderPicker } = h.load('controls/file-folder-picker/file-folder-picker.tsx');
  return { h, runtime, actions, mount: props => h.mountComponent(() => h.React.createElement(FileFolderPicker, props)) };
}
test('one Fluent input surface exposes the icon action and empty-field action without a text input', () => {
  const s = setup(); const m = s.mount({ itemTypes: 'all' });
  try {
    const surface = m.container.querySelector('.fui-Input');
    assert.ok(surface, 'reuse native Input border classes');
    assert.equal(surface.querySelector('input'), null, 'file selection is not editable text');
    const action = surface.querySelector('button[aria-haspopup="dialog"]');
    assert.ok(action); assert.match(action.getAttribute('aria-label'), /file e cartelle/i);
    s.h.act(() => require('react-dom/test-utils').Simulate.click(action));
    s.h.act(() => require('react-dom/test-utils').Simulate.click(surface.querySelector('button:not([aria-haspopup])')));
    assert.deepEqual(s.actions, ['browse', 'browse']);
    assert.doesNotMatch(surface.textContent, /Rimuovi tutti/);
  } finally { m.unmount(); s.h.close(); }
});
test('selected tags remove individually; browsing icon and labels reflect each item mode', () => {
  const icons = [];
  for (const itemTypes of ['files', 'folders', 'all']) {
    const s = setup(); s.runtime.values = [value]; const m = s.mount({ itemTypes, size: 'large', label: 'Attachments' });
    try {
      assert.equal(m.container.querySelector('button[aria-haspopup="dialog"]').textContent, '');
      assert.doesNotMatch(m.container.textContent, /Rimuovi tutti/);
      const action = m.container.querySelector('button[aria-haspopup="dialog"]');
      icons.push(action.querySelector('svg').innerHTML);
      assert.ok(action.getAttribute('aria-describedby'), 'opening action has a Fluent tooltip description');
      s.h.act(() => require('react-dom/test-utils').Simulate.click(m.container.querySelector('.fui-Tag')));
      assert.deepEqual(s.actions, ['one']);
    } finally { m.unmount(); s.h.close(); }
  }
  assert.equal(new Set(icons).size, 3, 'file, folder and mixed sources use distinct icons');
});
test('busy and disabled surfaces preserve tags and cannot start another opening', () => {
  const s = setup(); s.runtime.values = [value]; s.runtime.busy = true;
  const m = s.mount({ disabled: true, validationState: 'error', validationMessage: 'Invalid source' });
  try {
    const surface = m.container.querySelector('.fui-Input');
    assert.equal(surface.getAttribute('aria-busy'), 'true');
    assert.equal(surface.querySelector('button[aria-haspopup="dialog"]').disabled, true);
    assert.match(surface.textContent, /Report.pdf/);
    assert.equal(surface.querySelector('.fui-Spinner'), null);
    assert.doesNotMatch(m.container.textContent, /Selettore aperto/);
    assert.match(m.container.textContent, /Invalid source/);
  } finally { m.unmount(); s.h.close(); }
});

test('empty first action preserves Field naming, hints and validation without duplicating the icon ID', () => {
  const s = setup(); const m = s.mount({ id: 'invoice-files', label: 'Invoice documents', hint: 'Choose an invoice', validationState: 'error', validationMessage: 'An invoice is required', required: true });
  try {
    const empty = m.container.querySelector('.fui-Input button:not([aria-haspopup])');
    const icon = m.container.querySelector('button[aria-haspopup="dialog"]');
    assert.ok(empty.getAttribute('aria-labelledby'));
    assert.equal(empty.getAttribute('aria-labelledby'), icon.getAttribute('aria-labelledby'));
    const descriptions = empty.getAttribute('aria-describedby');
    assert.ok(descriptions);
    assert.match(descriptions.split(' ').map(id => m.container.querySelector(`[id="${id}"]`)?.textContent || '').join(' '), /Choose an invoice/);
    assert.equal(empty.getAttribute('aria-invalid'), 'true');
    assert.equal(m.container.querySelectorAll('#invoice-files').length, 1);
  } finally { m.unmount(); s.h.close(); }
});

test('hosted dialog keeps an accessible name without visible header, close button or padding', () => {
  const s = setup(); s.runtime.dialog = { name: 'test-hosted', epoch: 1 };
  window.crypto = require('node:crypto').webcrypto;
  const m = s.mount({ browseLabel: 'Choose attachments' });
  try {
    const dialog = document.querySelector('[role="dialog"]');
    assert.ok(dialog);
    const title = document.getElementById(dialog.getAttribute('aria-labelledby'));
    assert.match(title.textContent, /Choose attachments/);
    const header = title.closest('[data-stage-header]');
    assert.equal(window.getComputedStyle(header).display, 'none');
    assert.equal(dialog.querySelector('button'), null, 'Microsoft owns the visible exit actions inside its iframe');
    assert.equal(window.getComputedStyle(dialog).width, '1080px');
    assert.equal(window.getComputedStyle(dialog).height, '720px');
    // JSDOM 15 does not implement physical/logical padding cascade; actual zero-padding geometry is verified in the browser.
    assert.equal(dialog.querySelector('iframe').getAttribute('title'), 'Choose attachments');
  } finally { m.unmount(); s.h.close(); }
});

test('consumer hint and source icon stay unchanged while opening without automatic status', () => {
  const s = setup(); s.runtime.busy = true;
  const m = s.mount({ hint: 'Descrizione del consumer', itemTypes: 'folders' });
  try {
    assert.equal(m.container.querySelector('.fui-Spinner'), null);
    assert.match(m.container.textContent, /Descrizione del consumer/);
    assert.doesNotMatch(m.container.textContent, /Selettore aperto/);
    assert.ok(m.container.querySelector('button[aria-haspopup] svg'));
    assert.equal(m.container.querySelector('[role="status"]'), null);
  } finally { m.unmount(); s.h.close(); }
});

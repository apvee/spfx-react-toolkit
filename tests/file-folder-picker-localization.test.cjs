const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');

test('file dictionaries merge all known groups, retain empty fields and insert display data literally', () => {
  const h = createHarness();
  try {
    const { resolveFileFolderPickerStrings } = h.load('controls/file-folder-picker/file-folder-picker-strings.internal.ts');
    const native = resolveFileFolderPickerStrings();
    assert.equal(native.browse.files, 'Browse files…');
    assert.equal(native.labels.removeItem('$& {name}'), 'Remove $& {name}');
    const customized = resolveFileFolderPickerStrings({ browse: { files: '', folders: undefined }, labels: { file: 'Document' }, errors: { opening: 'Open again' } });
    assert.equal(customized.browse.files, ''); assert.equal(customized.browse.folders, native.browse.folders);
    assert.equal(customized.labels.file, 'Document'); assert.equal(customized.labels.folder, native.labels.folder);
    assert.equal(customized.errors.opening, 'Open again'); assert.equal(customized.errors.communication, native.errors.communication);
    assert.deepEqual(resolveFileFolderPickerStrings(native), native, 'complete callback dictionaries are valid instance overrides');
  } finally { h.close(); }
});

test('safe file failures classify raw diagnostics before projecting closed stable instance codes', () => {
  const h = createHarness();
  try {
    const { resolveFileFolderPickerStrings } = h.load('controls/file-folder-picker/file-folder-picker-strings.internal.ts');
    const { FileFolderPickerError, fileFolderPickerFailure } = h.load('controls/file-folder-picker/file-folder-picker-errors.internal.ts');
    const native = resolveFileFolderPickerStrings();
    const strings = resolveFileFolderPickerStrings({ errors: Object.fromEntries(Object.keys(native.errors).map(code => [code, `Safe ${code}`])) });
    for (const code of Object.keys(native.errors)) {
      const failure = fileFolderPickerFailure(new FileFolderPickerError(code), 'opening', strings);
      assert.equal(failure.code, code); assert.equal(failure.message, `Safe ${code}`);
    }
    for (const [raw, code] of [['metadata identity mismatch PRIVATE TOKEN', 'selection-identity'], ['outside scope PRIVATE 403', 'selection-scope'], ['extension not allowed PRIVATE HTTP', 'selection-policy'], ['deleted PRIVATE 404', 'selection-unavailable'], ['PRIVATE CLAIMS', 'selection-invalid']]) {
      const failure = fileFolderPickerFailure(new Error(raw), 'selection-invalid', strings);
      assert.equal(failure.code, code); assert.equal(failure.message, `Safe ${code}`);
      assert.doesNotMatch(failure.message, /PRIVATE|HTTP|TOKEN|CLAIMS|403|404/);
    }
  } finally { h.close(); }
});

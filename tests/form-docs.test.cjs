const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { collectPublicSurface, assertDocumentedSurface } = require('../scripts/public-surface.helpers.cjs');

test('Form docs cover compound members and Form-owned alias props, including nested error state', () => {
  const surface = collectPublicSurface(path.resolve('packages/spfx-react-toolkit/src/controls/form/index.ts'), { includeTypeMembers: true });
  const document = fs.readFileSync(path.resolve('docs/api/controls/form.md'), 'utf8');
  assertDocumentedSurface(document, surface);
  for (const name of ['ScopedFormSuite.FileInput', 'ScopedFormSuite.ArraySection', 'FormRootProps.onSubmit', 'FormMethods.formState.errors']) {
    const missing = document.replaceAll(name, name + 'Extra');
    assert.throws(() => assertDocumentedSurface(missing, surface), new RegExp(name.replaceAll('.', '\\.')));
  }
});

test('opt-in alias coverage requires own nested member JSDoc and does not invent inherited upstream props', () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'form-docs-'));
  const file = path.join(directory, 'index.ts');
  try {
    fs.writeFileSync(file, '/** Options. */ export type Options = External<{ as?: "button" }> & {\n/** State. */\nstate: { value: string } };');
    const surface = collectPublicSurface(file, { includeTypeMembers: true });
    assert.deepEqual(surface[0].members.map(member => member.name), ['Options.state', 'Options.state.value']);
    assert.throws(() => assertDocumentedSurface('Options Options.state Options.state.value', surface), /JSDoc.*Options.state.value/);
    fs.writeFileSync(file, '/** Options. */ export type Options = External<{ as?: "button" }> & {\n/** State. */\nstate: {\n/** Value. */\nvalue: string } };');
    assertDocumentedSurface('Options Options.state Options.state.value', collectPublicSurface(file, { includeTypeMembers: true }));
  } finally { fs.rmSync(directory, { recursive: true, force: true }); }
});

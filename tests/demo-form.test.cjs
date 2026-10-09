const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
const panel = fs.readFileSync(path.join(base, 'panels/FormPanel.tsx'), 'utf8');
const registry = fs.readFileSync(path.join(base, 'demoRegistry.ts'), 'utf8');
const source = ts.createSourceFile('FormPanel.tsx', panel, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
const adapters = ['Input', 'Textarea', 'SearchBox', 'Checkbox', 'Switch', 'SpinButton', 'Slider', 'Select', 'RadioGroup', 'Dropdown', 'DropdownMulti', 'CheckboxGroup', 'Combobox', 'ComboboxMulti', 'TagPicker', 'Rating', 'ColorPicker', 'SwatchPicker', 'DatePicker', 'TimePicker', 'DateTime', 'FileInput'];
function walk(node, visit) { visit(node); ts.forEachChild(node, child => walk(child, visit)); }
function jsxNames() {
  const names = [];
  walk(source, node => { if (ts.isJsxOpeningElement(node) || ts.isJsxSelfClosingElement(node)) names.push(node.tagName.getText(source)); });
  return names;
}
test('Form registry owns one lazy entry and each public value exactly once', () => {
  for (const value of ["key: 'form'", "symbol: 'Form'", "symbol: 'createForm'", "webpackChunkName: 'spfx-demo-form'"]) {
    assert.equal(registry.split(value).length - 1, 1, value);
  }
  assert.match(registry, /import\(.*\.\/panels\/FormPanel/);
  assert.match(panel, /from '@apvee\/spfx-react-toolkit'/);
  assert.doesNotMatch(panel, /packages\/.*src|controls\/Form|style=|\.style\./);
});
test('Complete family demo renders real adapter and composition JSX rather than a symbol list', () => {
  const names = jsxNames();
  for (const member of [...adapters, 'Root', 'Submit', 'Reset', 'Error', 'Object', 'Array', 'ObjectSection', 'ArraySection']) {
    assert.ok(names.includes(`Order.${member}`), member);
  }
  for (const member of adapters) assert.ok(panel.includes(`data-form-control="${member}"`), member);
  for (const member of ['Root', 'Input', 'Submit', 'Reset']) assert.ok(names.includes(`Form.${member}`), `dynamic ${member}`);
  assert.match(panel, /Order\.useForm\(/);
  assert.match(panel, /Order\.useField<string>\(/);
  assert.match(panel, /Numeric\.useForm\(/);
  assert.match(panel, /Numeric\.Input name="quantity"/);
  assert.match(panel, /quantity: string.*quantity: number/);
});
test('RHF owns live values, errors, dirty/touched state, resets and native submissions', () => {
  for (const marker of ['panel', 'frame', 'theme', 'direction', 'resize', 'disabled', 'mode', 'save-delay', 'save-fail', 'save-count', 'status', 'values', 'dirty', 'touched', 'errors', 'submit', 'reset', 'new-record', 'dirty-refresh', 'clear-file', 'request-submit', 'schema-values', 'schema-output']) {
    assert.ok(panel.includes(`data-form-demo="${marker}"`), marker);
  }
  assert.match(panel, /const values = methods\.watch\(\)/);
  assert.match(panel, /dirtyFields, touchedFields, errors, isDirty, isSubmitting/);
  assert.match(panel, /methods\.reset\(createRecord\('Order B'\)\)/);
  assert.match(panel, /keepDirtyValues: true/);
  assert.match(panel, /methods\.resetField\('files'\)/);
  assert.equal(panel.split('formRef.current?.requestSubmit()').length - 1, 2);
  assert.match(panel, /await new Promise<void>.*setTimeout\(resolve, 900\)/);
  assert.match(panel, /throw new Error\('Simulated save failure'\)/);
  assert.match(panel, /throw new Error\('Simulated validator failure'\)/);
  assert.doesNotMatch(panel, /useState<DemoRecord>\(createRecord|\.files\s*=|new DataTransfer|setValue\('files'/);
  let inComponent = false;
  walk(source, node => {
    if (ts.isCallExpression(node) && node.expression.getText(source) === 'createForm') {
      let parent = node.parent;
      while (parent) { if (ts.isFunctionLike(parent)) inComponent = true; parent = parent.parent; }
    }
  });
  assert.equal(inComponent, false, 'suite identity must be stable outside render');
});
test('Headless operations and localized section actions retain typed complete records', () => {
  for (const operation of ['append', 'prepend', 'insert', 'remove', 'update', 'replace', 'move', 'swap']) {
    assert.match(panel, new RegExp(`array\\.${operation}\\(`), operation);
    assert.ok(panel.includes(`data-form-array-action="${operation}"`), operation);
  }
  assert.match(panel, /const createRow = \(\): DemoRow => \(\{ code: 'New row', note: '' \}\)/);
  assert.match(panel, /createItem=\{createRow\} minItems=\{1\} maxItems=\{3\} reorderable itemLabel=/);
  for (const member of ['remove', 'moveUp', 'moveDown', 'added', 'removed']) assert.match(panel, new RegExp(`${member}: label =>`));
  assert.match(panel, /moved: \(label, position, total\) =>/);
  assert.match(panel, /key=\{item\.key\}/);
  assert.match(panel, /readOnly/);
  assert.match(panel, /required rules=\{\{ required: 'Reference is required'/);
});

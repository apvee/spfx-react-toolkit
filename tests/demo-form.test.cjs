const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
const panel = fs.readFileSync(path.resolve(__dirname, 'fixtures/demo-control-isolation/form-regressions.tsx'), 'utf8');
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
  for (const marker of ['panel', 'frame', 'direction', 'resize', 'disabled', 'mode', 'save-delay', 'save-fail', 'save-count', 'status', 'values', 'dirty', 'touched', 'errors', 'submit', 'reset', 'new-record', 'dirty-refresh', 'clear-file', 'request-submit', 'schema-values', 'schema-output']) {
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
test('Settings rows opt in around existing RHF fields without competing labels or a separate form', () => {
  const rows = [];
  walk(source, node => {
    if (ts.isJsxElement(node) && node.openingElement.tagName.getText(source) === 'SettingsGroup.Item') rows.push(node);
  });
  assert.equal(rows.length, 2, 'reference and enabled need working settings rows');
  const expected = ['reference', 'enabled'];
  rows.forEach((row, index) => {
    const attrs = row.openingElement.attributes.properties.map(prop => prop.name?.getText(source));
    for (const attr of ['field', 'title', 'description']) assert.ok(attrs.includes(attr), attr);
    const controls = row.children.filter(ts.isJsxSelfClosingElement);
    assert.equal(controls.length, 1, 'one logical Form control per row');
    const control = controls[0];
    assert.equal(control.tagName.getText(source), index === 0 ? 'Order.Input' : 'Order.Switch');
    const name = control.attributes.properties.find(prop => prop.name?.getText(source) === 'name');
    assert.equal(name.initializer.text, expected[index], 'settings use the existing record model');
    for (const attr of ['label', 'hint', 'aria-label', 'aria-labelledby']) {
      assert.ok(!control.attributes.properties.some(prop => prop.name?.getText(source) === attr), `row owns ${attr}`);
    }
    let ancestor = row.parent;
    while (ancestor && !(ts.isJsxElement(ancestor) && ancestor.openingElement.tagName.getText(source) === 'Order.Root')) ancestor = ancestor.parent;
    assert.ok(ancestor, 'settings share existing submit/reset ownership');
  });
  for (const marker of ['settings-group', 'settings-reference', 'settings-enabled']) assert.ok(panel.includes(`data-form-demo="${marker}"`), marker);
  assert.equal(registry.split("symbol: 'SettingsGroup'").length - 1, 1, 'retain unique SettingsGroup owner');
});

test('Form presentation sample observes inner refs events and callback content', () => {
  for (const marker of ['inspect-slots', 'slot-observation', 'slot-label', 'slot-hint', 'slot-section-title', 'slot-section-description', 'slot-no-options']) assert.ok(panel.includes(`'${marker}'`) || panel.includes(`"${marker}"`), marker);
  assert.match(panel, /slotLabelRef\.current\?\.tagName/);
  assert.match(panel, /slotHintRef\.current\?\.tagName/);
  assert.match(panel, /onClick: \(\) => setSlotClicks/);
  assert.match(panel, /children: \(Component, props\) => <Component \{\.\.\.props\}>Callback hint/);
});

test('Form configuration uses Fluent 9 and retained diagnostics', () => {
  const source = ts.createSourceFile('panel.tsx', panel, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  const tags = [];
  function visit(node) {
    if (ts.isJsxOpeningElement(node) || ts.isJsxSelfClosingElement(node)) tags.push(node.tagName.getText(source));
    ts.forEachChild(node, visit);
  }
  visit(source);
  for (const tag of ['input', 'select', 'textarea', 'button', 'details']) assert.ok(!tags.includes(tag), tag);
  for (const tag of ['DemoCard', 'DemoAccordionPanel', 'Checkbox', 'Select', 'Field']) assert.ok(tags.includes(tag), tag);
  assert.match(panel, /controls=/); assert.match(panel, /diagnostics=/);
});

test('ColorArea names its two native input slots separately', () => {
  assert.match(panel, /inputX=\{\{ 'aria-label': 'Accent saturation' \}\}/);
  assert.match(panel, /inputY=\{\{ 'aria-label': 'Accent value' \}\}/);
});


test('Form scenario separates configuration commands, outcomes and advanced diagnostics', () => {
  const panel = fs.readFileSync('tests/fixtures/demo-control-isolation/form-regressions.tsx', 'utf8');
  const ast = require('typescript');
  const parsed = ast.createSourceFile('panel.tsx', panel, ast.ScriptTarget.Latest, true, ast.ScriptKind.TSX);
  const cards = [];
  function visit(node) {
    if (ast.isJsxOpeningElement(node) && node.tagName.getText(parsed) === 'DemoCard') cards.push(node);
    ast.forEachChild(node, visit);
  }
  visit(parsed);
  assert.equal(parsed.parseDiagnostics.length, 0);
  assert.equal(cards.length, 1, 'one configuration shell owns the scenario');
  const attrs = new Map(cards[0].attributes.properties.filter(ast.isJsxAttribute).map(attr => [attr.name.getText(parsed), attr.getText(parsed)]));
  for (const name of ['controls', 'commands', 'status', 'diagnostics']) assert.ok(attrs.has(name), name);
  assert.ok(!attrs.has('observations'));
  assert.match(attrs.get('commands'), /new-record/);
  assert.match(attrs.get('diagnostics'), /schema-output/);
  assert.doesNotMatch(attrs.get('commands'), /inspect-slots/);
});

test('schema mode hides record submission commands and saved status while keeping both forms mounted', () => {
  const gatedMarkers = new Set(['new-record', 'dirty-refresh', 'clear-file', 'request-submit', 'save-count', 'status']);
  const checked = new Set();
  walk(source, node => {
    if (!ts.isJsxElement(node) && !ts.isJsxSelfClosingElement(node)) return;
    const opening = ts.isJsxElement(node) ? node.openingElement : node;
    const marker = opening.attributes.properties.find(attr => ts.isJsxAttribute(attr) && attr.name.getText(source) === 'data-form-demo');
    if (!marker || !ts.isStringLiteral(marker.initializer) || !gatedMarkers.has(marker.initializer.text)) return;
    let gated = false;
    for (let ancestor = node.parent; ancestor; ancestor = ancestor.parent) {
      if (ts.isJsxElement(ancestor)) {
        const hidden = ancestor.openingElement.attributes.properties.find(attr => ts.isJsxAttribute(attr) && attr.name.getText(source) === 'hidden');
        if (hidden && /mode !== 'rules'/.test(hidden.getText(source))) gated = true;
      }
    }
    assert.ok(gated, `${marker.initializer.text} is unavailable when the record form is hidden`);
    checked.add(marker.initializer.text);
  });
  assert.equal(checked.size, gatedMarkers.size);
  assert.match(panel, /<div hidden=\{mode !== 'rules'\}>\s*<Order.Root/);
  assert.match(panel, /<div hidden=\{mode !== 'schema'\}><SchemaExample/);
});


test('Product scenario owns one configuration and selectable variants without comparison galleries', () => {
  const product = fs.readFileSync(path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/FormPanel.tsx'), 'utf8');
  const compiler = require('typescript');
  const parsed = compiler.createSourceFile('product.tsx', product, compiler.ScriptTarget.Latest, true, compiler.ScriptKind.TSX);
  const cards = []; let createInRender = false;
  function inspect(node) {
    if (compiler.isJsxOpeningElement(node) && node.tagName.getText(parsed) === 'DemoCard') cards.push(node);
    if (compiler.isCallExpression(node) && ['createForm', 'createDataGrid'].includes(node.expression.getText(parsed))) {
      for (let parent = node.parent; parent; parent = parent.parent) if (compiler.isFunctionLike(parent)) createInRender = true;
    }
    compiler.forEachChild(node, inspect);
  }
  inspect(parsed);
  assert.equal(parsed.parseDiagnostics.length, 0);
  assert.equal(cards.length, 1);
  assert.equal(createInRender, false);
  assert.match(product, /data-demo-variant="FormAdapter"/);
  assert.doesNotMatch(product, /Independent |comparison|hidden=\{(?:mode|family) !==/);
  const browserChecks = fs.readFileSync(path.resolve(__dirname, 'fixtures/demo-control-isolation/task3-product.browser.cjs'), 'utf8');
  assert.match(browserChecks, /formProduct/);
});

test('actual product schema and dynamic configuration owns disabled, dirty state and model loading', async () => {
  process.env.NODE_ENV = 'test';
  const hostHarness = require('./fixtures/form-selection-harness.cjs');
  const sourceHarness = require('./fixtures/formatted-text-harness.cjs');
  await hostHarness.setup();
  document.createRange = () => ({ selectNodeContents() {}, getClientRects() { return []; } });
  let fixture;
  try {
    const library = await sourceHarness.loadModule('tests/fixtures/sx-browser/form-toolkit-boundary.ts', { reactHookForm: true });
    const shared = await sourceHarness.loadModule('apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/index.ts');
    const compiled = ts.transpileModule(fs.readFileSync(path.join(base, 'panels/FormPanel.tsx'), 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React } });
    const actual = { exports: {} };
    new Function('require', 'module', 'exports', compiled.outputText)(name => name === '@apvee/spfx-react-toolkit' ? library : name === '../shared' ? shared : require(name), actual, actual.exports);
    fixture = hostHarness.mount(hostHarness.React.createElement(actual.exports.default));
    const find = marker => fixture.host.querySelector(`[data-form-demo="${marker}"]`);
    assert.equal(find('slot-hint').textContent, 'Summarize team results, decisions and next steps.');
    for (const family of ['schema', 'dynamic']) {
      hostHarness.change(find('mode'), family);
      assert.equal(fixture.host.querySelectorAll('form').length, 1);
      const input = fixture.host.querySelector(`input[name="${family === 'schema' ? 'quantity' : 'dynamic'}"]`);
      assert.ok(input, 'the current model supplies the actual input');
      assert.match(find('status').textContent, /Dirty: false/);
      hostHarness.act(() => find('disabled').click());
      assert.equal(input.disabled, true, `${family} field is disabled by current configuration`);
      hostHarness.act(() => find('disabled').click());
      await hostHarness.act(async () => { hostHarness.change(input, family === 'schema' ? '23' : 'Edited dynamic'); await Promise.resolve(); });
      assert.match(find('status').textContent, /Dirty: true/);
      hostHarness.click(find('new-record'));
      assert.equal(input.value, family === 'schema' ? '12' : 'Annual report', `${family} loads the current model`);
      assert.match(find('status').textContent, /Dirty: false/);
    }
  } finally { fixture?.close(); await hostHarness.cleanup(); }
});

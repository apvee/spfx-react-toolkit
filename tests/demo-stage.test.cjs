const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
function readPanel() { return fs.readFileSync(path.resolve(__dirname, 'fixtures/demo-control-isolation/stage-regressions.tsx'), 'utf8'); }
test('Stage has one lazy registry entry with concrete surface and form coverage', () => {
  const registry = fs.readFileSync(path.join(base, 'demoRegistry.ts'), 'utf8');
  for (const text of ["key: 'stage'", "symbol: 'Stage'", "webpackChunkName: 'spfx-demo-stage'"]) assert.equal(registry.split(text).length - 1, 1, text);
  assert.match(registry, /import\(.*\.\/panels\/StagePanel/);
});
test('Stage demo exposes executable acceptance scenarios through stable markers', () => {
  const panel = readPanel();
  for (const marker of ['panel', 'instance', 'trigger', 'mode', 'size', 'header', 'body', 'footer', 'request', 'counter', 'mount-count', 'dirty', 'save', 'reset', 'pending', 'complete-save', 'discard', 'keep-editing', 'nested-menu', 'menu-inline', 'combobox', 'date-picker', 'arrow-toggle']) {
    assert.ok(panel.includes(`data-stage-demo="${marker}"`), marker);
  }
  assert.match(panel, /from '@apvee\/spfx-react-toolkit'/);
  assert.doesNotMatch(panel, /style=|\.style\.|handleSubmit\(/);
  const source = ts.createSourceFile('StagePanel.tsx', panel, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  const names = [];
  function walk(node) {
    if (ts.isJsxOpeningElement(node) || ts.isJsxSelfClosingElement(node)) names.push(node.tagName.getText(source));
    ts.forEachChild(node, walk);
  }
  walk(source);
  assert.equal(names.filter(name => name === 'Stage').length, 1, 'all modes share one content tree');
  for (const name of ['Stage', 'Stage.Header', 'Stage.Body', 'Stage.Footer', 'Editor.Root', 'Editor.Input', 'Editor.Submit', 'Editor.Reset', 'Editor.Combobox', 'Editor.DatePicker', 'Menu']) assert.ok(names.includes(name), name);
  assert.match(panel, /const \{ isDirty, isSubmitting \} = methods\.formState/);
  assert.match(panel, /disabled=\{open\}/);
  assert.match(panel, /presentation === 'popover'/);
  assert.doesNotMatch(panel, /Independent RTL instance/);
  const fixture = fs.readFileSync(path.resolve(__dirname, 'fixtures/demo-control-isolation/stage.tsx'), 'utf8');
  assert.match(fixture, /dir="rtl"/);
  for (const marker of ['card', 'card-selection']) assert.ok(fixture.includes(`data-stage-demo="${marker}"`));
  assert.match(fixture, /StageExample instance="B"/);
});

function assertFooterSubmissionPlacement(panel) {
  const source = ts.createSourceFile('StagePanel.tsx', panel, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  let checked = false;
  function walk(node) {
    const tag = ts.isJsxElement(node) ? node.openingElement.tagName
      : ts.isJsxSelfClosingElement(node) ? node.tagName : undefined;
    if (tag && tag.getText(source) === 'Editor.Submit') {
      const ancestors = [];
      for (let parent = node.parent; parent; parent = parent.parent) {
        if (ts.isJsxElement(parent)) ancestors.push(parent.openingElement.tagName.getText(source));
      }
      assert.ok(ancestors.includes('Stage.Footer'), 'submit is in Footer');
      assert.ok(ancestors.includes('Editor.Root'), 'submit is associated with the real form');
      assert.ok(ancestors.includes('Stage'), 'form is inside surface portal');
      checked = true;
    }
    ts.forEachChild(node, walk);
  }
  walk(source);
  assert.ok(checked, 'real footer submission was checked');
}

test('Footer submission belongs to the form inside the portaled Stage surface', () => {
  assertFooterSubmissionPlacement(readPanel());
});

test('Placement coverage treats paired and self-closing Submit nodes identically', () => {
  for (const submit of ['<Editor.Submit>Save</Editor.Submit>', '<Editor.Submit />']) {
    assert.doesNotThrow(() => assertFooterSubmissionPlacement(
      `<Stage><Editor.Root><Stage.Footer>${submit}</Stage.Footer></Editor.Root></Stage>`
    ));
    assert.throws(() => assertFooterSubmissionPlacement(
      `<Stage><Stage.Footer>${submit}</Stage.Footer></Stage>`
    ), /submit is associated with the real form/);
    assert.throws(() => assertFooterSubmissionPlacement(
      `<Editor.Root><Stage.Footer>${submit}</Stage.Footer></Editor.Root>`
    ), /form is inside surface portal/);
    assert.throws(() => assertFooterSubmissionPlacement(
      `<Stage><Editor.Root>${submit}</Editor.Root></Stage>`
    ), /submit is in Footer/);
  }
});

test('Narrow Stage fields constrain native controls as well as Field wrappers', () => {
  const panel = readPanel();
  const source = ts.createSourceFile('StagePanel.tsx', panel, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  const controls = new Map();
  function walk(node) {
    if (ts.isJsxSelfClosingElement(node) && ['Editor.Input', 'Editor.Combobox', 'Editor.DatePicker'].includes(node.tagName.getText(source))) {
      controls.set(node.tagName.getText(source), new Map(node.attributes.properties.filter(ts.isJsxAttribute).map(attribute => [attribute.name.getText(source), attribute.getText(source)])));
    }
    ts.forEachChild(node, walk);
  }
  walk(source);
  for (const name of ['Editor.Input', 'Editor.Combobox', 'Editor.DatePicker']) {
    assert.match(controls.get(name)?.get('className') ?? '', /narrowField/, `${name} constrains Field`);
  }
  for (const name of ['Editor.Input', 'Editor.Combobox']) {
    assert.match(controls.get(name)?.get('root') ?? '', /controlWidth/, `${name} constrains its native root`);
    assert.match(controls.get(name)?.get('input') ?? '', /minWidth.zero/, `${name} allows native input to shrink without clipping adornments`);
  }
  assert.match(panel, /\[`& > \.\$\{datePickerClassNames.root\}`\]/, 'DatePicker uses the public class within its own Field only');
  assert.match(panel, /const controlWidth = sx\(minWidth.zero, width.full\)/);
});

test('Stage slot demo exposes callback actions and inspectable inner title ref', () => {
  const panel = readPanel();
  for (const marker of ['title-slot', 'action-slot', 'inspect-slot', 'slot-observation', 'callback-icon', 'long-text']) assert.ok(panel.includes(marker));
  assert.match(panel, /titleSlotRef.current/); assert.match(panel, /children: \(Tag, props\)/);
});

test('Stage configuration uses Fluent 9 and retained diagnostics', () => {
  const panel = readPanel();
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


test('Stage scenario separates configuration commands, outcomes and advanced diagnostics', () => {
  const panel = fs.readFileSync('tests/fixtures/demo-control-isolation/stage-regressions.tsx', 'utf8');
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
  assert.match(attrs.get('contentPlacement'), /configuration/);
  assert.doesNotMatch(panel, /Independent RTL instance/);
});


test('Product scenario owns one configuration and selectable variants without comparison galleries', () => {
  const product = fs.readFileSync(path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/StagePanel.tsx'), 'utf8');
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
  assert.match(product, /data-demo-variant="Presentation"/);
  assert.doesNotMatch(product, /Independent |comparison|hidden=\{(?:mode|family) !==/);
  const browserChecks = fs.readFileSync(path.resolve(__dirname, 'fixtures/demo-control-isolation/task3-product.browser.cjs'), 'utf8');
  assert.match(browserChecks, /stageProduct/);
});

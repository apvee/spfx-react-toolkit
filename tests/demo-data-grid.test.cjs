const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const read = file => fs.readFileSync(path.resolve(__dirname, '..', file), 'utf8');
const base = 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/';

test('DataGrid sample covers the complete family with consumer-owned observable states', () => {
  const registry = read(base + 'demoRegistry.ts');
  const panel = read('tests/fixtures/demo-control-isolation/data-grid-regressions.tsx');
  assert.match(registry, /key: 'data-grid'/);
  assert.match(registry, /webpackChunkName: 'spfx-demo-data-grid'/);
  for (const symbol of ['DataGrid', 'createDataGrid']) {
    assert.equal([...registry.matchAll(new RegExp(`symbol: '${symbol}'`, 'g'))].length, 1);
  }
  for (const column of ['Text', 'Number', 'Boolean', 'Date', 'DateTime', 'Badge', 'Icon', 'Link', 'Person', 'People', 'Tags', 'Progress', 'Rating', 'Sparkline', 'Actions']) {
    assert.ok(panel.includes(`<Grid.${column}Column `), column);
  }
  assert.match(panel, /<Grid\.Column /);
  assert.match(panel, /data-grid-slot-callback="header"/);
  assert.match(panel, /headerSlotRef/);
  assert.match(panel, /suppressSlots \? false :/);
  assert.match(panel, /Grid\.defineColumns/);
  assert.match(panel, /<DataGrid /);
  for (const id of ['dg-page', 'dg-hidden', 'dg-loading', 'dg-error', 'dg-empty', 'dg-selection', 'dg-sort', 'dg-action', 'dg-navigation', 'dg-activation-output', 'dg-ref', 'dg-update-sparkline', 'dg-color', 'dg-resizable', 'dg-slot-suppress', 'dg-inspect-slot-refs', 'dg-slot-refs', 'dg-slot-clicks']) assert.ok(panel.includes(`id="${id}"`), id);
  assert.match(panel, /new Set\(\['off-page'\]\)/);
  assert.match(panel, /new Set\(data\.selectedItems\)/);
  assert.match(panel, /previous\.map\(row => \(\{ \.\.\.row, trend:/);
  assert.match(panel, /from '@apvee\/spfx-react-toolkit'/);
  assert.doesNotMatch(panel, /style=|DataGrid\.View|\/lib\/|createDataGrid<DemoRow>\(\).*=>/);
});

 test('demo configuration, preview and retained diagnostics use shared v9 regions', () => {
  const panel = fs.readFileSync('tests/fixtures/demo-control-isolation/data-grid-regressions.tsx', 'utf8');
  assert.match(panel, /<DemoCard/);
  assert.match(panel, /controls=\{/);
  assert.match(panel, /diagnostics=\{/);
  assert.match(panel, /<DemoAccordionPanel/);
  assert.doesNotMatch(panel, /<(input|select|textarea|button|details)\b/);
 });


test('DataGrid scenario separates configuration commands, outcomes and advanced diagnostics', () => {
  const panel = fs.readFileSync('tests/fixtures/demo-control-isolation/data-grid-regressions.tsx', 'utf8');
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
  assert.match(attrs.get('commands'), /dg-apply-sort/);
  assert.doesNotMatch(attrs.get('commands'), /dg-inspect/);
  assert.match(attrs.get('diagnostics'), /dg-selection/);
  assert.doesNotMatch(attrs.get('status'), /JSON.stringify/);
});


test('Product scenario owns one configuration and selectable variants without comparison galleries', () => {
  const product = fs.readFileSync(path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/DataGridPanel.tsx'), 'utf8');
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
  assert.match(product, /data-demo-variant="API"/);
  assert.doesNotMatch(product, /Independent |comparison|hidden=\{(?:mode|family) !==/);
  const browserChecks = fs.readFileSync(path.resolve(__dirname, 'fixtures/demo-control-isolation/task3-product.browser.cjs'), 'utf8');
  assert.match(browserChecks, /dataGridProduct/);
});


test('actual product renders selectable defaults and mutually exclusive activation for every API', async () => {
  process.env.NODE_ENV = 'test';
  const hostHarness = require('./fixtures/form-selection-harness.cjs');
  const sourceHarness = require('./fixtures/formatted-text-harness.cjs');
  await hostHarness.setup();
  document.createRange = () => ({ selectNodeContents() {}, getClientRects() { return []; } });
  let fixture;
  try {
    const library = await sourceHarness.loadModule('tests/fixtures/sx-browser/data-grid-toolkit-boundary.ts');
    const shared = await sourceHarness.loadModule('apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/index.ts');
    const compiler = require('typescript');
    const source = read(base + 'panels/DataGridPanel.tsx');
    const compiled = compiler.transpileModule(source, { compilerOptions: { module: compiler.ModuleKind.CommonJS, target: compiler.ScriptTarget.ES2020, jsx: compiler.JsxEmit.React } });
    const panel = { exports: {} };
    new Function('require', 'module', 'exports', compiled.outputText)(name => name === '@apvee/spfx-react-toolkit' ? library : name === '../shared' ? shared : require(name), panel, panel.exports);
    fixture = hostHarness.mount(hostHarness.React.createElement(panel.exports.default));
    const change = (id, value) => hostHarness.change(fixture.host.querySelector(`#${id}`), value);
    for (const api of ['typed', 'descriptors', 'dynamic']) {
      change('dg-api', api); change('dg-selection-mode', 'multiselect');
      assert.equal(fixture.host.querySelectorAll('[role=grid]').length, 1, `${api} default renders one actual grid`);
      assert.ok(fixture.host.querySelectorAll('#data-grid-main input[type=checkbox]').length > 0, `${api} exposes native selection`);
      const firstRow = fixture.host.querySelector('#data-grid-main [data-grid-row-id]');
      const rowId = api === 'dynamic' ? 'valid' : 'a';
      const wasSelected = JSON.parse(fixture.host.querySelector('#dg-selection').textContent).includes(rowId);
      hostHarness.click(firstRow);
      const selected = JSON.parse(fixture.host.querySelector('#dg-selection').textContent);
      assert.equal(selected.includes(rowId), !wasSelected, `${api} selection toggles the retained model without activation`);
      change('dg-selection-mode', 'none');
      const activationRow = fixture.host.querySelector('#data-grid-main [data-grid-row-id]');
      const activationCell = activationRow.querySelector('[role=gridcell]');
      activationCell.focus();
      assert.equal(document.activeElement, activationCell, `${api} keyboard starts at the native gridcell`);
      hostHarness.key(activationCell, 'Enter');
      assert.deepEqual(JSON.parse(fixture.host.querySelector('#dg-activation-output').textContent), { rowId: api === 'dynamic' ? 'valid' : 'a' });
      assert.equal(fixture.host.querySelectorAll('#data-grid-main input[type=checkbox]').length, 0, `${api} activation has no conflicting selection`);
    }
  } finally { fixture?.close(); await hostHarness.cleanup(); }
});

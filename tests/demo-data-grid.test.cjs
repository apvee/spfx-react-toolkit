const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const read = file => fs.readFileSync(path.resolve(__dirname, '..', file), 'utf8');
const base = 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/';

test('DataGrid sample covers the complete family with consumer-owned observable states', () => {
  const catalogue = read(base + 'generated/jsdoc-examples/catalogue.ts');
  const panel = read('tests/fixtures/demo-control-isolation/data-grid-regressions.tsx');
  assert.doesNotMatch(catalogue, /key: 'data-grid'|spfx-demo-data-grid|panels\/DataGridPanel/);
  for (const symbol of ['DataGrid', 'createDataGrid']) {
    const catalogue = read(base + 'generated/jsdoc-examples/catalogue.ts');
    assert.ok(catalogue.includes(`"featureTitle": ${JSON.stringify(symbol)}`), symbol);
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


test('canonical DataGrid recipes own one active subject without a copied API controller', () => {
  const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
  for (const title of ['Explore typed document columns and controlled interaction', 'Supply descriptor columns with explicit sort and row order', 'Classify dynamic row values and preserve controlled selection']) {
    const { source } = scenario(title);
    const compiler = require('typescript');
    const parsed = compiler.createSourceFile('grid-example.tsx', source, compiler.ScriptTarget.Latest, true, compiler.ScriptKind.TSX);
    let grids = 0, createInRender = false;
    function inspect(node) {
      if ((compiler.isJsxOpeningElement(node) || compiler.isJsxSelfClosingElement(node)) && ['Grid','DataGrid'].includes(node.tagName.getText(parsed))) grids++;
      if (compiler.isCallExpression(node) && node.expression.getText(parsed) === 'createDataGrid') {
        for (let parent=node.parent; parent; parent=parent.parent) if (compiler.isFunctionLike(parent)) createInRender = true;
      }
      compiler.forEachChild(node, inspect);
    }
    inspect(parsed); assert.equal(parsed.parseDiagnostics.length,0); assert.equal(grids,1); assert.equal(createInRender,false);
    for (const region of ['controls:', 'commands:', 'status:', 'diagnostics:', 'children:']) assert.ok(source.includes(region),region);
    assert.doesNotMatch(source,/dg-api|private|panels\//);
  }
  const browserChecks = read('tests/fixtures/demo-control-isolation/task3-product.browser.cjs');
  assert.match(browserChecks,/dataGridProduct/); assert.match(browserChecks,/grid-scenario/);
});

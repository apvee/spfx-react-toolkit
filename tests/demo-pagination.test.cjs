const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
const readPanel = () => fs.readFileSync(path.resolve(__dirname, 'fixtures/demo-control-isolation/pagination-regressions.tsx'), 'utf8');

test('Pagination registry owns one lazy executable scenario', () => {
  const registry = fs.readFileSync(path.join(base, 'demoRegistry.ts'), 'utf8');
  for (const token of ["key: 'pagination'", "symbol: 'Pagination'", "webpackChunkName: 'spfx-demo-pagination'"]) {
    assert.equal(registry.split(token).length - 1, 1, token);
  }
  assert.match(registry, /import\(.*\.\/panels\/PaginationPanel/);
});

test('Pagination demo exposes result navigation and consumer lifecycle acceptance markers', () => {
  const panel = readPanel();
  for (const marker of ['panel', 'instance', 'mode', 'page-size', 'filter', 'reduce-total', 'manual-loading', 'hidden', 'rtl', 'localization', 'size', 'appearance', 'narrow', 'results', 'summary', 'form', 'navigation', 'complete', 'inspect-ref', 'submit', 'request', 'ref', 'submissions']) {
    assert.ok(panel.includes(`data-pagination-demo="${marker}"`), marker);
  }
  assert.match(panel, /from '@apvee\/spfx-react-toolkit'/);
  assert.doesNotMatch(panel, /style=|\.style\.|setTimeout|setInterval/);
  assert.match(panel, /const Grid = createDataGrid<Order>\(\)/);
  assert.match(panel, /available\.slice\(start, end\)/);
  assert.match(panel, /role="status" aria-live="polite" aria-atomic="true"/);
  assert.match(panel, /setPageIndex\(0\); setPending\(undefined\)/);
  assert.match(panel, /loading=\{pending !== undefined\}/);
  for (const mode of ['known', 'unknown', 'cursor', 'loadMore']) assert.ok(panel.includes(`value="${mode}"`), mode);
  const source = ts.createSourceFile('PaginationPanel.tsx', panel, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  const names = [];
  let suiteOutsideRender = false;
  function walk(node) {
    if (ts.isJsxOpeningElement(node) || ts.isJsxSelfClosingElement(node)) names.push(node.tagName.getText(source));
    if (ts.isVariableStatement(node) && node.parent === source && node.getText(source).startsWith('const Grid =')) suiteOutsideRender = true;
    ts.forEachChild(node, walk);
  }
  walk(source);
  assert.ok(suiteOutsideRender, 'typed suite belongs to module scope');
  for (const name of ['Pagination', 'Grid', 'Grid.TextColumn', 'FluentProvider', 'Select', 'Checkbox', 'form']) assert.ok(names.includes(name), name);
  assert.equal(names.filter(name => name === 'PaginationExample').length, 1, 'visible demo has one live instance');
  assert.match(panel, /export const PaginationExample/);
  assert.match(panel, /initialRtl = false/);
});

 test('demo configuration, preview and retained diagnostics use shared v9 regions', () => {
  const panel = fs.readFileSync('tests/fixtures/demo-control-isolation/pagination-regressions.tsx', 'utf8');
  assert.match(panel, /<DemoCard/);
  assert.match(panel, /controls=\{/);
  assert.match(panel, /diagnostics=\{/);
  assert.doesNotMatch(panel, /Independent RTL instance/);
  assert.doesNotMatch(panel, /<(input|select|textarea|button|details)\b/);
 });


test('Product scenario owns one configuration and selectable variants without comparison galleries', () => {
  const product = fs.readFileSync(path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/PaginationPanel.tsx'), 'utf8');
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
  assert.match(product, /data-demo-variant="Mode"/);
  assert.doesNotMatch(product, /Independent |comparison|hidden=\{(?:mode|family) !==/);
  const browserChecks = fs.readFileSync(path.resolve(__dirname, 'fixtures/demo-control-isolation/task3-product.browser.cjs'), 'utf8');
  assert.match(browserChecks, /paginationProduct/);
});

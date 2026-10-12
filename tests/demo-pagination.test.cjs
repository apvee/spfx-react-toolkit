const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
const readPanel = () => fs.readFileSync(path.join(base, 'panels/PaginationPanel.tsx'), 'utf8');

test('Pagination registry owns one lazy executable scenario', () => {
  const registry = fs.readFileSync(path.join(base, 'demoRegistry.ts'), 'utf8');
  for (const token of ["key: 'pagination'", "symbol: 'Pagination'", "webpackChunkName: 'spfx-demo-pagination'"]) {
    assert.equal(registry.split(token).length - 1, 1, token);
  }
  assert.match(registry, /import\(.*\.\/panels\/PaginationPanel/);
});

test('Pagination demo exposes result navigation and consumer lifecycle acceptance markers', () => {
  const panel = readPanel();
  for (const marker of ['panel', 'instance', 'mode', 'page-size', 'filter', 'reduce-total', 'manual-loading', 'hidden', 'theme', 'rtl', 'localization', 'size', 'appearance', 'narrow', 'results', 'summary', 'form', 'navigation', 'complete', 'inspect-ref', 'submit', 'request', 'ref', 'submissions']) {
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
  assert.equal(names.filter(name => name === 'PaginationExample').length, 2, 'two instances own independent state');
});

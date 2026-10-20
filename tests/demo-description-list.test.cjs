const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');

test('Description List registry loads one executable compound-control scenario', () => {
  const registry = fs.readFileSync(path.join(base, 'demoRegistry.ts'), 'utf8');
  for (const token of ["key: 'description-list'", "symbol: 'DescriptionList'", "webpackChunkName: 'spfx-demo-description-list'"]) {
    assert.equal(registry.split(token).length - 1, 1, token);
  }
  assert.match(registry, /import\(.*\.\/panels\/DescriptionListPanel/);
});

test('Description List demo retains the acceptance scenarios used for rendered host verification', () => {
  const panel = fs.readFileSync(path.resolve(__dirname, 'fixtures/demo-control-isolation/description-list-regressions.tsx'), 'utf8');
  for (const marker of ['panel', 'instance', 'layout', 'width', 'composition', 'reverse', 'hidden', 'item-hidden', 'empty', 'rtl', 'container', 'nested-container', 'list', 'pair', 'counter', 'inspect-ref', 'ref', 'clicks', 'conditional-slot', 'slot-clicks']) {
    assert.ok(panel.includes(`data-description-list-demo="${marker}"`), marker);
  }
  assert.match(panel, /from '@apvee\/spfx-react-toolkit'/);
  assert.doesNotMatch(panel, /style=|\.style\.|setTimeout|setInterval/);
  assert.match(panel, /\[320, 479, 480, 639, 640, 900\]/);
  assert.match(panel, /container\.inlineSize, width\.px\(pixels\)/);
  assert.match(panel, /container\.inlineSize, width\.px\(320\)/);
  assert.match(panel, /key=\{entry\.id\}/);
  assert.match(panel, /value: 0/);
  assert.match(panel, /value: ''/);
  assert.match(panel, /value: undefined/);
  assert.match(panel, /children: \(Component, props\)/);
  assert.match(panel, /value: showConditional &&/);
  assert.match(panel, /ref: slotRef/);
  assert.match(panel, /'data-description-list-demo': 'empty-slot-container'/);
  assert.match(panel, /createRef<HTMLDListElement>/);
  assert.match(panel, /createRef<HTMLDivElement>/);
  const source = ts.createSourceFile('DescriptionListPanel.tsx', panel, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  assert.equal(source.parseDiagnostics.length, 0, 'panel parses as TSX');
  const tags = [];
  function walk(node) {
    if (ts.isJsxOpeningElement(node) || ts.isJsxSelfClosingElement(node)) tags.push(node.tagName.getText(source));
    ts.forEachChild(node, walk);
  }
  walk(source);
  for (const name of ['DescriptionList', 'DescriptionList.Item', 'Counter', 'FluentProvider', 'Badge', 'Link']) assert.ok(tags.includes(name), name);
  assert.equal(tags.filter(name => name === 'DescriptionListExample').length, 1, 'visible demo has one live instance');
  assert.match(panel, /export const DescriptionListExample/);
  assert.match(panel, /initialRtl = false/);
});

 test('demo configuration, preview and retained diagnostics use shared v9 regions', () => {
  const panel = fs.readFileSync('tests/fixtures/demo-control-isolation/description-list-regressions.tsx', 'utf8');
  assert.match(panel, /<DemoCard/);
  assert.match(panel, /controls=\{/);
  assert.match(panel, /diagnostics=\{/);
  assert.doesNotMatch(panel, /Independent RTL instance/);
  assert.doesNotMatch(panel, /<(input|select|textarea|button|details)\b/);
 });


test('Product scenario owns one configuration and selectable variants without comparison galleries', () => {
  const product = fs.readFileSync(path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/DescriptionListPanel.tsx'), 'utf8');
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
  assert.match(browserChecks, /descriptionListProduct/);
});

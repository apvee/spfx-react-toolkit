const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');

test('DescriptionList is available through its generated owner scenario', () => {
  const { entry } = require('./fixtures/executable-examples/cards-description-list.cjs').scenario('Inspect responsive metadata with data and compound APIs');
  assert.equal(entry.featureTitle, 'DescriptionList');
  assert.equal(entry.metadata.host, 'none');
  assert.ok(entry.metadata.observe.length > 0);
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


test('source metadata scenario owns regions and its data/compound variant', () => {
  const { source } = require('./fixtures/executable-examples/cards-description-list.cjs').scenario('Inspect responsive metadata with data and compound APIs');
  const parsed = ts.createSourceFile('source.tsx', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  assert.equal(parsed.parseDiagnostics.length, 0);
  for (const token of ['controls:', 'status:', 'diagnostics:', 'children:', 'data-demo-variant="API"', 'key={entry.id}']) assert.ok(source.includes(token), token);
  assert.doesNotMatch(source, /Independent |comparison|components\/|panels\/|style=/);
  const browserChecks = fs.readFileSync(path.resolve(__dirname, 'fixtures/demo-control-isolation/task3-product.browser.cjs'), 'utf8');
  assert.match(browserChecks, /descriptionListProduct/);
});

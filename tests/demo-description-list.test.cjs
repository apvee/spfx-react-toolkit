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
  const panel = fs.readFileSync(path.join(base, 'panels/DescriptionListPanel.tsx'), 'utf8');
  for (const marker of ['panel', 'instance', 'layout', 'width', 'composition', 'reverse', 'hidden', 'item-hidden', 'empty', 'theme', 'rtl', 'container', 'nested-container', 'list', 'pair', 'counter', 'inspect-ref', 'ref', 'clicks']) {
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
  assert.equal(tags.filter(name => name === 'DescriptionListExample').length, 2, 'instances own independent state');
});

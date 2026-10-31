const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const read = file => fs.readFileSync(path.resolve(__dirname, '..', file), 'utf8');
const base = 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/';

test('groups have lazy interactive coverage without duplicate hook/provider inventory', () => {
  const registry = read(base + 'demoRegistry.ts');
  const panel = read('tests/fixtures/demo-control-isolation/groups-regressions.tsx');
  assert.match(registry, /key: 'groups'/);
  assert.match(registry, /webpackChunkName: 'spfx-demo-groups'/);
  assert.match(registry, /import\(.*GroupsPanel/);
  for (const symbol of ['SettingsGroup', 'AccordionGroup']) {
    assert.equal([...registry.matchAll(new RegExp(`symbol: '${symbol}'`, 'g'))].length, 1);
    assert.ok(panel.includes(`${symbol}.Item`));
  }
  for (const marker of ['direction', 'resize', 'multiple', 'unmount', 'hidden', 'disabled', 'settings', 'accordion', 'email', 'note', 'close-all', 'inspect', 'observations', 'title', 'description', 'action-label', 'accordion-title', 'accordion-description', 'open-items']) {
    assert.ok((panel.includes(`data-groups-demo="${marker}"`) || panel.includes(`'data-groups-demo': '${marker}'`)), marker);
  }
  for (const marker of ['Input aria-label=', 'Switch aria-label=', 'DemoIcon name=', 'onToggle=', 'openItems=', 'setOpenItems([])', 'unmountOnClose=', 'defaultValue="Retained note"', 'title={<span>', 'action=', 'ref=']) {
    assert.ok(panel.includes(marker), marker);
  }
  assert.match(panel, /\[320, 500, 639, 640, 900\]/);
  assert.match(panel, /from '@apvee\/spfx-react-toolkit'/);
  assert.doesNotMatch(panel, /style=|Layout\.|\/lib\//);
});

test('group guides explain labels, trigger restrictions and separate host execution', () => {
  const settings = read('docs/api/controls/settings-group.md');
  const accordion = read('docs/api/controls/accordion-group.md');
  assert.match(settings, /not implicit labels/);
  assert.match(settings, /wrap/);
  assert.match(accordion, /Never place interactive JSX/);
  assert.match(accordion, /unmountOnClose/);
  assert.match(accordion, /remain mounted/);
  assert.match(accordion, /native inert/);
  assert.match(read('docs/development/sharepoint-validation/checklists/groups.md'), /^# SettingsGroup and AccordionGroup/m);
});

test('group slot demo exposes inner refs and callback actions', () => {
  const panel = read('tests/fixtures/demo-control-isolation/groups-regressions.tsx');
  assert.match(panel, /title-slot/); assert.match(panel, /action-slot/);
  assert.match(panel, /object-icon/); assert.match(panel, /callback-icon/); assert.match(panel, /width=\{80\} height=\{80\}/);
  assert.match(panel, /titleSlotRef.current/); assert.match(panel, /children: \(Tag, props\)/);
});

// Configuration controls belong to Fluent and remain separate from the demonstrated control.
test('Groups configuration uses Fluent native slots and retained diagnostic regions', () => {
  const panel = fs.readFileSync(path.resolve(__dirname, 'fixtures/demo-control-isolation/groups-regressions.tsx'), 'utf8');
  assert.doesNotMatch(panel, /<(?:input|select|textarea|button|details)\b|style=|from '@fluentui\/react'/);
  assert.match(panel, /controls=\{/);
  assert.match(panel, /diagnostics=\{/);
  assert.doesNotMatch(panel, /defaultOpenItems/);
  assert.match(panel, /input=\{\{/);
  assert.match(panel, /onChange=\{\(_event, data\)/);
});


test('Product scenario owns one configuration and selectable variants without comparison galleries', () => {
  const product = fs.readFileSync(path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/GroupsPanel.tsx'), 'utf8');
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
  assert.match(product, /data-demo-variant="GroupType"/);
  assert.doesNotMatch(product, /Independent |comparison|hidden=\{(?:mode|family) !==/);
  const browserChecks = fs.readFileSync(path.resolve(__dirname, 'fixtures/demo-control-isolation/task3-product.browser.cjs'), 'utf8');
  assert.match(browserChecks, /groupsProduct/);
});

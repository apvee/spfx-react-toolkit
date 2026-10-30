const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const read = name => fs.readFileSync(path.resolve(__dirname,'..',name),'utf8');
test('StateMessage has a lazy interactive demo and real browser fixture coverage', () => {
  const base='apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/';
  const registry=read(base+'demoRegistry.ts'), panel=read(base+'panels/StateMessagePanel.tsx');
  assert.match(registry,/key: 'state-message'/);assert.match(registry,/symbol: 'StateMessage'/);
  assert.match(registry,/import\(.*StateMessagePanel/);
  for(const marker of ['preset','direction','resize','long','role','visual','heading','preview','inspect','observations', 'slots', 'title', 'description'])assert.ok((panel.includes(`data-state-message-demo="${marker}"`) || panel.includes(`'data-state-message-demo': '${marker}'`)),marker);
  assert.match(panel,/primaryAction=/);assert.match(panel,/secondaryAction=/);assert.match(panel,/tertiaryAction=/);
  assert.doesNotMatch(panel,/style=|Layout\./);
  assert.match(read('tests/fixtures/sx-browser/index.tsx'),/<StateMessageFixture/);
  assert.match(read('tests/fixtures/styles-panel-browser/index.tsx'),/<StateMessagePanel/);
});

// Configuration controls belong to Fluent and remain separate from the demonstrated control.
test('StateMessage configuration uses Fluent native slots and retained diagnostic regions', () => {
  const panel = fs.readFileSync(path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/StateMessagePanel.tsx'), 'utf8');
  assert.doesNotMatch(panel, /<(?:input|select|textarea|button|details)\b|style=|from '@fluentui\/react'/);
  assert.match(panel, /controls=\{/);
  assert.match(panel, /diagnostics=\{/);
  assert.doesNotMatch(panel, /defaultOpenItems/);
  assert.match(panel, /input=\{\{/);
  assert.match(panel, /onChange=\{\(_event, data\)/);
});

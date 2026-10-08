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
  for(const marker of ['preset','theme','direction','resize','long','role','visual','heading','preview','inspect','observations'])assert.ok(panel.includes(`data-state-message-demo="${marker}"`),marker);
  assert.match(panel,/primaryAction=/);assert.match(panel,/secondaryAction=/);assert.match(panel,/tertiaryAction=/);
  assert.doesNotMatch(panel,/style=|Layout\./);
  assert.match(read('tests/fixtures/sx-browser/index.tsx'),/<StateMessageFixture/);
  assert.match(read('tests/fixtures/styles-panel-browser/index.tsx'),/<StateMessagePanel/);
});

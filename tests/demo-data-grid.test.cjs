const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const read = file => fs.readFileSync(path.resolve(__dirname, '..', file), 'utf8');
const base = 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/';

test('DataGrid sample covers the complete family with consumer-owned observable states', () => {
  const registry = read(base + 'demoRegistry.ts');
  const panel = read(base + 'panels/DataGridPanel.tsx');
  assert.match(registry, /key: 'data-grid'/);
  assert.match(registry, /webpackChunkName: 'spfx-demo-data-grid'/);
  for (const symbol of ['DataGrid', 'createDataGrid']) {
    assert.equal([...registry.matchAll(new RegExp(`symbol: '${symbol}'`, 'g'))].length, 1);
  }
  for (const column of ['Text', 'Number', 'Boolean', 'Date', 'DateTime', 'Badge', 'Icon', 'Link', 'Person', 'People', 'Tags', 'Progress', 'Rating', 'Sparkline', 'Actions']) {
    assert.ok(panel.includes(`<Grid.${column}Column `), column);
  }
  assert.match(panel, /<Grid\.Column /);
  assert.match(panel, /Grid\.defineColumns/);
  assert.match(panel, /<DataGrid /);
  for (const id of ['dg-page', 'dg-hidden', 'dg-loading', 'dg-error', 'dg-empty', 'dg-selection', 'dg-sort', 'dg-action', 'dg-navigation', 'dg-activation-output', 'dg-ref', 'dg-update-sparkline', 'dg-color', 'dg-resizable']) assert.ok(panel.includes(`id="${id}"`), id);
  assert.match(panel, /new Set\(\['off-page'\]\)/);
  assert.match(panel, /new Set\(data\.selectedItems\)/);
  assert.match(panel, /previous\.map\(row => \(\{ \.\.\.row, trend:/);
  assert.match(panel, /from '@apvee\/spfx-react-toolkit'/);
  assert.doesNotMatch(panel, /style=|DataGrid\.View|\/lib\/|createDataGrid<DemoRow>\(\).*=>/);
});

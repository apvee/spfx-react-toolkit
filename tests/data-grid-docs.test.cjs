const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { collectPublicSurface, assertDocumentedSurface } = require('../scripts/public-surface.helpers.cjs');

test('DataGrid source JSDoc and public guide cover all exposed members and row/state contracts', () => {
  const surface = collectPublicSurface(path.resolve('packages/spfx-react-toolkit/src/controls/data-grid/index.ts'), { includeTypeMembers: true });
  const document = fs.readFileSync(path.resolve('docs/api/controls/data-grid.md'), 'utf8');
  assertDocumentedSurface(document, surface);
  for (const name of ['DataGridProps.getRowId', 'DataGridProps.selectedItems', 'DataGridSuite.defineColumns', 'DataGridSparklineSummary.changePercent', 'DataGridStrings.sparklineSummary']) {
    const missing = document.replaceAll(name, name + 'Extra');
    assert.throws(() => assertDocumentedSurface(missing, surface), new RegExp(name.replaceAll('.', '\\.')));
  }
  for (const contract of ['off-page', 'consumer-supplied order', 'missing', 'invalid', 'NaN', '0000', '9.3.27', 'showValue={false}', 'Number.MAX_VALUE', 'HTMLDivElement', 'ReadonlyMap', 'no-visible-columns']) assert.ok(document.toLowerCase().includes(contract.toLowerCase()), contract);
  assert.match(fs.readFileSync('docs/api/INDEX.md', 'utf8'), /\.\/controls\/data-grid\.md/);
  assert.match(fs.readFileSync('docs/development/sharepoint-validation/checklists/data-grid.md', 'utf8'), /^# DataGrid/m);
});

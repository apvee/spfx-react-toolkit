const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

test('Stage-only builders reject conflicting selections before producing browser assets', () => {
  for (const [script, stageFlag, conflictFlag] of [
    ['build-sx-browser-fixture.cjs', 'SX_BROWSER_STAGE_ONLY', 'SX_BROWSER_SHELL_ONLY'],
    ['build-sx-browser-fixture.cjs', 'SX_BROWSER_STAGE_ONLY', 'SX_BROWSER_GROUPS_ONLY'],
    ['build-styles-panel-browser.cjs', 'STYLES_PANEL_STAGE_ONLY', 'STYLES_PANEL_DATA_GRID_ONLY'],
    ['build-styles-panel-browser.cjs', 'STYLES_PANEL_STAGE_ONLY', 'STYLES_PANEL_FORM_ONLY'],
  ]) {
    const result = spawnSync(process.execPath, [path.resolve(__dirname, '../scripts', script)], {
      env: { ...process.env, [stageFlag]: '1', [conflictFlag]: '1' }, encoding: 'utf8', timeout: 30000,
    });
    assert.ifError(result.error);
    assert.equal(result.status, 1, `${script} accepted contradictory fixture selection`);
    assert.match(result.stderr, /Stage-only fixture cannot be combined with another Only flag/);
  }
});

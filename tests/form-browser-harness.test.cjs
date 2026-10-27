const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function selectFixture(env) {
  const script = path.resolve(__dirname, '../scripts/build-styles-panel-browser.cjs');
  let selected;
  const webpack = config => { selected = config; };
  webpack.NormalModuleReplacementPlugin = class {};
  const scriptRequire = require('node:module').createRequire(script);
  const buildRequire = name => name === 'webpack' ? webpack : scriptRequire(name);
  buildRequire.resolve = require.resolve;
  vm.runInNewContext(fs.readFileSync(script, 'utf8'), {
    require: buildRequire, __dirname: path.dirname(script), process: { env }, console,
  }, { filename: script });
  return selected;
}

test('Form-only compiled builder selects the narrow actual sample entry and boundary', () => {
  const config = selectFixture({ STYLES_PANEL_FORM_ONLY: '1' });
  assert.equal(path.basename(config.entry), 'form.tsx');
  assert.equal(path.basename(config.resolve.alias['@apvee/spfx-react-toolkit$']), 'form-toolkit-boundary.ts');
  assert.equal(config.plugins.length, 0, 'Form fixture needs no SPFx module substitution');
});

test('Existing compiled fixture selections retain their entry and boundary', () => {
  for (const [env, entry, boundary] of [
    [{}, 'index.tsx', 'toolkit-boundary.ts'],
    [{ STYLES_PANEL_STAGE_ONLY: '1' }, 'stage.tsx', 'stage-toolkit-boundary.ts'],
    [{ STYLES_PANEL_SHELL_ONLY: '1' }, 'web-part-shell.tsx', 'web-part-shell-toolkit-boundary.ts'],
    [{ STYLES_PANEL_DATA_GRID_ONLY: '1' }, 'data-grid.tsx', 'data-grid-toolkit-boundary.ts'],
  ]) {
    const config = selectFixture(env);
    assert.equal(path.basename(config.entry), entry);
    assert.equal(path.basename(config.resolve.alias['@apvee/spfx-react-toolkit$']), boundary);
  }
});

const test = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path');
const root = path.resolve(__dirname, '..'), lib = path.join(root, 'packages/spfx-react-toolkit/lib');
const names = ['SitePicker', 'ListPicker', 'FieldPicker', 'ViewPicker', 'ContentTypePicker', 'ListItemPicker', 'PrincipalPicker'];
test('clean picker entrypoints expose seven React and native APIs without new legacy aliases', () => {
  const manifest = require('../packages/spfx-react-toolkit/package.json');
  const { assertEntrypointContract, collectPublishedFiles } = require('../scripts/package-entrypoints.cjs');
  assertEntrypointContract(manifest, collectPublishedFiles());
  assert.equal(manifest.exports['./controls/pickers'].default, './lib/controls/pickers/index.js');
  assert.equal(manifest.exports['./property-pane/pickers'].default, './lib/property-pane/pickers/index.js');
  assert.equal(manifest.exports['./lib/controls/pickers'], undefined);
  assert.equal(manifest.peerDependencies['@microsoft/sp-property-pane'], '>=1.18.0 <2.0.0');
  for (const name of names) {
    assert.match(fs.readFileSync(path.join(lib, 'controls/pickers/index.d.ts'), 'utf8'), new RegExp(`export \\{ ${name} \\}`));
    assert.match(fs.readFileSync(path.join(lib, 'property-pane/pickers/index.d.ts'), 'utf8'), new RegExp(`export \\{ PropertyPane${name} \\}`));
  }
});
test('generic root import graph never reaches native property-pane runtime', () => {
  const visited = new Set();
  function visit(file) {
    if (visited.has(file)) return; visited.add(file);
    assert.ok(!file.includes('/property-pane/'), `Root reaches native runtime ${file}`);
    const source = fs.readFileSync(file, 'utf8');
    for (const match of source.matchAll(/(?:require\(|import\(|from\s+)['"](\.[^'"]+)['"]/g)) {
      const base = path.resolve(path.dirname(file), match[1]);
      const resolved = [base + '.js', path.join(base, 'index.js')].find(candidate => fs.existsSync(candidate));
      if (resolved) visit(resolved);
    }
  }
  visit(path.join(lib, 'index.js'));
  assert.ok([...visited].some(file => file.endsWith('/controls/pickers/site-picker.js')));
});

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const root = path.resolve(__dirname, '..');
const importsPanel = path.join(root, 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/ImportsPanel.tsx');
const importsPanelSpecifiers = new Set([
  '@apvee/spfx-react-toolkit/hooks',
  '@apvee/spfx-react-toolkit/styles',
  '@apvee/spfx-react-toolkit/lib/hooks/useStableCallback',
  '@apvee/spfx-react-toolkit/lib/hooks/useSx',
  '@apvee/spfx-react-toolkit/lib/helpers/styles/width',
  '@apvee/spfx-react-toolkit/lib/helpers/styles/typography',
  '@apvee/spfx-react-toolkit/lib/helpers/styles/padding-inline-start',
]);
test('library and SPFx app have independent package boundaries', () => {
  const workspace = require('../package.json');
  assert.equal(workspace.private, true);
  assert.deepEqual(workspace.workspaces, ['packages/spfx-react-toolkit', 'apps/spfx-react-toolkit-test']);
  const library = require('../packages/spfx-react-toolkit/package.json');
  const app = require('../apps/spfx-react-toolkit-test/package.json');
  assert.equal(library.name, '@apvee/spfx-react-toolkit');
  assert.equal(library.main, 'lib/index.js');
  assert.equal(library.types, 'lib/index.d.ts');
  assert.equal(app.private, true);
  assert.equal(app.dependencies[library.name], library.version);
  const walk = dir => fs.readdirSync(dir, {withFileTypes:true}).flatMap(e => e.isDirectory() ? walk(path.join(dir,e.name)) : [path.join(dir,e.name)]);
  for (const f of walk(path.join(root, 'apps/spfx-react-toolkit-test/src')).filter(f => /\.tsx?$/.test(f))) {
    const source = ts.createSourceFile(f, fs.readFileSync(f, 'utf8'), ts.ScriptTarget.Latest, true);
    const imports = source.statements.filter(statement => ts.isImportDeclaration(statement) || ts.isExportDeclaration(statement))
      .map(statement => statement.moduleSpecifier?.text).filter(Boolean);
    for (const specifier of imports) {
      const approvedImportsScenario = f === importsPanel && importsPanelSpecifiers.has(specifier);
      if (specifier.startsWith(library.name)) assert.ok(specifier === library.name || approvedImportsScenario, `${f}: app must consume the public root entry or an exact Imports scenario alias`);
      assert.ok(!specifier.includes('packages/') && (!specifier.includes('/lib/') || approvedImportsScenario), `${f}: accidental internal import ${specifier}`);
      if (specifier.startsWith('.')) assert.ok(path.resolve(path.dirname(f),specifier).startsWith(path.join(root,'apps/spfx-react-toolkit-test/src') + path.sep), `${f}: import escapes app source`);
    }
  }
  for (const f of walk(path.join(root, 'packages/spfx-react-toolkit/src')).filter(f => /\.tsx?$/.test(f))) {
    const source = ts.createSourceFile(f, fs.readFileSync(f, 'utf8'), ts.ScriptTarget.Latest, true);
    for (const statement of source.statements) {
      if (!ts.isImportDeclaration(statement) && !ts.isExportDeclaration(statement)) continue;
      const specifier = statement.moduleSpecifier?.text;
      if (!specifier) continue;
      assert.ok(!specifier.includes('spfx-react-toolkit-test'), `${f}: library must not depend on app`);
      if (specifier.startsWith('.')) assert.ok(path.resolve(path.dirname(f),specifier).startsWith(path.join(root,'packages/spfx-react-toolkit/src') + path.sep), `${f}: import escapes library source`);
    }
  }
});

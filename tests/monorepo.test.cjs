const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const root = path.resolve(__dirname, '..');
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
      if (specifier.startsWith(library.name)) assert.equal(specifier, library.name, `${f}: app must consume the public root entry`);
      assert.ok(!specifier.includes('packages/') && !specifier.includes('/lib/'), `${f}: accidental internal import ${specifier}`);
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

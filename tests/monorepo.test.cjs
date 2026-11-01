const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const root = path.resolve(__dirname, '..');
const { cleanAliases } = require('../scripts/package-entrypoints.cjs');
const { executableRoot, ownership } = require('../scripts/executable-examples-render.cjs');
const generatedBase = path.join(root, executableRoot);
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
const nativePickerBase = path.join(root, 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest');
const nativePickerSpecifiers = new Map([
  [path.join(nativePickerBase, 'SpFxReactToolkitTestWebPart.ts'), new Set([
    '@apvee/spfx-react-toolkit/property-pane/pickers',
    '@apvee/spfx-react-toolkit/property-pane/file-folder-picker',
    '@apvee/spfx-react-toolkit/controls/file-folder-picker',
  ])],
  [path.join(nativePickerBase, 'components/shared/fileFolderPickerNavigation.ts'), new Set([
    '@apvee/spfx-react-toolkit/controls/file-folder-picker',
  ])],
]);
function assertAppImport(file, specifier, libraryName, sourceText) {
  const approvedImportsScenario = file === importsPanel && importsPanelSpecifiers.has(specifier);
  const ownedExecutable = path.dirname(file) === generatedBase
    && /^(?:[a-f0-9]{64}\.tsx|catalogue\.ts)$/.test(path.basename(file))
    && sourceText.startsWith(ownership + '\n');
  const alias = specifier === libraryName ? '.' : '.' + specifier.slice(libraryName.length);
  const approvedExecutableEntry = ownedExecutable && Object.hasOwn(cleanAliases, alias);
  if (specifier.startsWith(libraryName)) assert.ok(specifier === libraryName || approvedImportsScenario || nativePickerSpecifiers.get(file)?.has(specifier) || approvedExecutableEntry, `${file}: app must consume the public root entry, an exact approved scenario alias or a source-owned executable clean entry`);
  assert.ok(!specifier.includes('packages/') && (!specifier.includes('/lib/') || approvedImportsScenario), `${file}: accidental internal import ${specifier}`);
  if (specifier.startsWith('.')) assert.ok(path.resolve(path.dirname(file), specifier).startsWith(path.join(root, 'apps/spfx-react-toolkit-test/src') + path.sep), `${file}: import escapes app source`);
}

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
    const sourceText = fs.readFileSync(f, 'utf8');
    const source = ts.createSourceFile(f, sourceText, ts.ScriptTarget.Latest, true);
    const imports = source.statements.filter(statement => ts.isImportDeclaration(statement) || ts.isExportDeclaration(statement))
      .map(statement => statement.moduleSpecifier?.text).filter(Boolean);
    for (const specifier of imports) {
      assertAppImport(f, specifier, library.name, sourceText);
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

test('only owned generated executables may consume exact clean aliases without weakening import boundaries', () => {
  const libraryName = require('../packages/spfx-react-toolkit/package.json').name;
  const generated = path.join(generatedBase, 'a'.repeat(64) + '.tsx');
  const owned = ownership + '\nexport {};\n';
  const pagination = libraryName + '/controls/pagination';
  assert.doesNotThrow(() => assertAppImport(generated, pagination, libraryName, owned));
  assert.doesNotThrow(() => assertAppImport(path.join(generatedBase, 'catalogue.ts'), libraryName + '/controls/web-part-shell', libraryName, owned));
  assert.doesNotThrow(() => assertAppImport(importsPanel, libraryName + '/lib/hooks/useStableCallback', libraryName, ''));
  for (const [file, specifier, text] of [
    [generated, libraryName + '/controls/not-a-public-entry', owned],
    [generated, libraryName + '/controls/pagination/Pagination', owned],
    [generated, libraryName + '/hooks/useAsyncInvoke.internal', owned],
    [generated, libraryName + '/lib/hooks/useStableCallback', owned],
    [generated, libraryName + '/src/index', owned],
    [generated, pagination, '// Manual file\n'],
    [path.join(generatedBase, 'manual.tsx'), pagination, owned],
    [path.join(generatedBase, '..', 'elsewhere.tsx'), pagination, owned],
    [generated, '../../../../../../../../packages/spfx-react-toolkit/src/index', owned],
    [generated, '../../../../../../../../outside.ts', owned],
    [importsPanel, libraryName + '/lib/hooks/not-approved', ''],
  ]) assert.throws(() => assertAppImport(file, specifier, libraryName, text), /app must consume|accidental internal import|import escapes app source/, specifier);
});

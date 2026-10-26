const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { execFileSync } = require('node:child_process');
const { createRequire } = require('node:module');
const root = path.resolve(__dirname, '..');
const library = path.join(root, 'packages/spfx-react-toolkit');
const manifest = require('../packages/spfx-react-toolkit/package.json');

test('EditableText has one explicit canonical alias and exactly four public exports', () => {
  assert.deepEqual(manifest.exports['./controls/editable-text'], { types: './lib/controls/editable-text/index.d.ts', default: './lib/controls/editable-text/index.js' });
  assert.deepEqual(manifest.typesVersions['*']['controls/editable-text'], ['lib/controls/editable-text/index.d.ts']);
  for (const key of Object.keys(manifest.exports)) {
    assert.ok(!key.includes('*'));
    assert.ok(!key.includes('editable-text/') || key === './controls/editable-text');
  }
});

test('packed declarations preserve all four canonical exports and exclude internal types', () => {
  fs.mkdirSync(path.join(root, 'temp'), { recursive: true });
  const directory = fs.mkdtempSync(path.join(root, 'temp/editable-text-package-'));
  try {
    const packed = JSON.parse(execFileSync('npm', ['pack', '--ignore-scripts', '--json', '--pack-destination', directory, '--cache', '/private/tmp/editable-text-npm-cache'], { cwd: library, encoding: 'utf8' }))[0];
    const installed = path.join(directory, 'node_modules/@apvee/spfx-react-toolkit');
    fs.mkdirSync(installed, { recursive: true });
    execFileSync('tar', ['-xzf', path.join(directory, packed.filename), '-C', installed, '--strip-components=1']);
    const from = createRequire(path.join(directory, 'package.json'));
    for (const suffix of ['controls/editable-text/editable-text-editor.internal', 'controls/editable-text/editable-text-codec.internal', 'controls/editable-text/editable-text-document.types.internal', 'lib/controls/editable-text/index', 'controls/web-part-shell/title/web-part-title-editor.internal', 'controls/web-part-shell/title/web-part-title-color-picker.internal', 'controls/web-part-shell/title/web-part-title-toolbar.internal']) {
      assert.throws(() => from.resolve('@apvee/spfx-react-toolkit/' + suffix), { code: 'ERR_PACKAGE_PATH_NOT_EXPORTED' });
    }
    assert.equal(from.resolve('@apvee/spfx-react-toolkit/controls/editable-text'), path.join(installed, 'lib/controls/editable-text/index.js'));
    const fixture = path.join(directory, 'consumer.tsx');
    fs.copyFileSync(path.join(root, 'tests/fixtures/editable-text-package-consumer.tsx'), fixture);
    for (const [moduleResolution, module] of [[ts.ModuleResolutionKind.Node10, ts.ModuleKind.ESNext], [ts.ModuleResolutionKind.Bundler, ts.ModuleKind.ESNext], [ts.ModuleResolutionKind.Node16, ts.ModuleKind.Node16]]) {
      const program = ts.createProgram([fixture], { noEmit: true, strict: true, skipLibCheck: true, moduleResolution, module, jsx: ts.JsxEmit.React, target: ts.ScriptTarget.ES2020, types: [] });
      const diagnostics = ts.getPreEmitDiagnostics(program);
      assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, { getCanonicalFileName: value => value, getCurrentDirectory: () => directory, getNewLine: () => '\n' }));
      assert.equal(program.getSourceFiles().filter(source => /[/\\](?:@tiptap|prosemirror-[^/\\]+)[/\\]/.test(source.fileName)).length, 0, 'Public declarations must not resolve engine types');
      assert.equal(program.getSourceFiles().filter(source => /[/\\]editable-text[/\\].*\.internal\.d\.ts$/.test(source.fileName)).length, 0, 'Public declarations must not resolve private AST or engine declarations');
      const checker = program.getTypeChecker(), source = program.getSourceFile(fixture), definitions = {};
      const index = program.getSourceFile(path.join(installed, 'lib/controls/editable-text/index.d.ts'));
      assert.deepEqual(checker.getExportsOfModule(checker.getSymbolAtLocation(index)).map(symbol => symbol.name).sort(), ['EditableText', 'EditableTextBlock', 'EditableTextChangeData', 'EditableTextProps']);
      for (const node of source.statements.filter(ts.isImportDeclaration)) for (const element of node.importClause?.namedBindings?.elements ?? []) {
        const symbol = checker.getSymbolAtLocation(element.name);
        definitions[element.name.text] = checker.getAliasedSymbol(symbol).declarations;
      }
      for (const family of [['rootText', 'domainText', 'controlText'], ['RootProps', 'DomainProps', 'ControlProps'], ['RootBlock', 'DomainBlock', 'ControlBlock'], ['RootChange', 'DomainChange', 'ControlChange']]) {
        assert.deepEqual(definitions[family[0]], definitions[family[1]]);
        assert.deepEqual(definitions[family[0]], definitions[family[2]]);
      }
    }
  } finally { fs.rmSync(directory, { recursive: true, force: true }); }
});


test('actual packed standalone Read/Edit loads only its engine and isolates two regions', async () => {
  const { compileProbe, executeProbe } = require('../scripts/verify-production-runtime.cjs');
  fs.mkdirSync(path.join(root, 'temp'), { recursive: true });
  const directory = fs.mkdtempSync(path.join(root, 'temp/editable-text-runtime-'));
  try {
    const packed = JSON.parse(execFileSync('npm', ['pack', '--ignore-scripts', '--json', '--pack-destination', directory, '--cache', '/private/tmp/editable-text-npm-cache'], { cwd: library, encoding: 'utf8' }))[0];
    const installed = path.join(directory, 'node_modules/@apvee/spfx-react-toolkit');
    fs.mkdirSync(installed, { recursive: true });
    execFileSync('tar', ['-xzf', path.join(directory, packed.filename), '-C', installed, '--strip-components=1']);
    const bundle = await compileProbe({ consumerRoot: directory, outputDir: path.join(directory, 'assets'), scenario: 'editable-text', variant: 'control' });
    assert.ok(bundle.emittedModules.some(item => item.resource.includes('@tiptap/')));
    assert.ok(bundle.emittedModules.every(item => !item.resource.includes('/web-part-shell/') && !item.resource.startsWith('@microsoft/sp-')));
    const stats = JSON.parse(fs.readFileSync(path.join(directory, 'assets/stats.json'), 'utf8'));
    const contract = require('./fixtures/tree-shaking/contract.json');
    const { checkFixture } = require('../scripts/bundle-contract.internal.cjs');
    const initialChunkIds = [...new Set(Object.values(stats.entrypoints).flatMap(entry => entry.chunks))];
    assert.deepEqual(checkFixture({ id: 'editable-text/control', initialChunkIds, emittedModules: bundle.emittedModules, retentionObservations: [] }, contract.fixtures.find(item => item.id === 'editable-text'), contract), []);
    const observation = executeProbe(bundle.bundlePath, 'editable-text', directory);
    assert.deepEqual(observation.observation.destroyed, [true, true]);
    assert.ok(observation.observation.assetRequests.length > 0);
  } finally { fs.rmSync(directory, { recursive: true, force: true }); }
});

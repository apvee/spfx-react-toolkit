const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { execFileSync } = require('node:child_process');
const { createRequire } = require('node:module');
const { checkFixture } = require('../scripts/bundle-contract.internal.cjs');
const root = path.resolve(__dirname, '..');
const library = path.join(root, 'packages/spfx-react-toolkit');
const manifest = require('../packages/spfx-react-toolkit/package.json');

test('Shell has one explicit canonical control alias with private internals', () => {
  assert.deepEqual(manifest.exports['./controls/web-part-shell'], { types: './lib/controls/web-part-shell/index.d.ts', default: './lib/controls/web-part-shell/index.js' });
  assert.deepEqual(manifest.typesVersions['*']['controls/web-part-shell'], ['lib/controls/web-part-shell/index.d.ts']);
  for (const key of Object.keys(manifest.exports)) {
    assert.ok(!key.includes('*'));
    assert.ok(!key.includes('web-part-shell/') || key === './controls/web-part-shell');
  }
});

test('packed declarations preserve app properties, canonical aliases and exclude engine types', () => {
  fs.mkdirSync(path.join(root, 'temp'), { recursive: true });
  const directory = fs.mkdtempSync(path.join(root, 'temp/web-part-shell-package-'));
  try {
    const packed = JSON.parse(execFileSync('npm', ['pack', '--ignore-scripts', '--json', '--pack-destination', directory, '--cache', path.join(directory, '.npm-cache')], { cwd: library, encoding: 'utf8' }))[0];
    const installed = path.join(directory, 'node_modules/@apvee/spfx-react-toolkit');
    fs.mkdirSync(installed, { recursive: true });
    execFileSync('tar', ['-xzf', path.join(directory, packed.filename), '-C', installed, '--strip-components=1']);
    const from = createRequire(path.join(directory, 'package.json'));
    assert.equal(from.resolve('@apvee/spfx-react-toolkit/controls/web-part-shell'), path.join(installed, 'lib/controls/web-part-shell/index.js'));
    const fixture = path.join(directory, 'consumer.tsx');
    fs.copyFileSync(path.join(root, 'tests/fixtures/web-part-shell-package-consumer.tsx'), fixture);
    for (const [moduleResolution, module] of [[ts.ModuleResolutionKind.Node10, ts.ModuleKind.ESNext], [ts.ModuleResolutionKind.Bundler, ts.ModuleKind.ESNext], [ts.ModuleResolutionKind.Node16, ts.ModuleKind.Node16]]) {
      const program = ts.createProgram([fixture], { noEmit: true, strict: true, skipLibCheck: true, moduleResolution, module, jsx: ts.JsxEmit.React, target: ts.ScriptTarget.ES2020, types: [] });
      const diagnostics = ts.getPreEmitDiagnostics(program);
      assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, { getCanonicalFileName: value => value, getCurrentDirectory: () => directory, getNewLine: () => '\n' }));
      assert.equal(program.getSourceFiles().filter(source => /[/\\](?:@tiptap|prosemirror-[^/\\]+)[/\\]/.test(source.fileName)).length, 0, 'Public declarations must not resolve engine types');
      const checker = program.getTypeChecker(), source = program.getSourceFile(fixture), definitions = {};
      for (const node of source.statements.filter(ts.isImportDeclaration)) for (const element of node.importClause?.namedBindings?.elements ?? []) {
        const symbol = checker.getSymbolAtLocation(element.name);
        definitions[element.name.text] = checker.getAliasedSymbol(symbol).declarations;
      }
      for (const family of [['rootShell', 'domainShell', 'controlShell'], ['rootMetadata', 'domainMetadata', 'controlMetadata']]) {
        assert.deepEqual(definitions[family[0]], definitions[family[1]]);
        assert.deepEqual(definitions[family[0]], definitions[family[2]]);
      }
    }
  } finally { fs.rmSync(directory, { recursive: true, force: true }); }
});

test('initial module assertions reject eager engine while allowing the async union', () => {
  const fixture = { forbiddenModuleGroups: [], requiredModuleGroups: ['shell', 'tiptap'], forbiddenInitialModuleGroups: ['tiptap', 'prosemirror'] };
  const contract = { moduleGroups: { shell: ['shell'], tiptap: ['@tiptap/'], prosemirror: ['prosemirror-'] } };
  const item = { id: 'shell/control', initialChunkIds: [1], emittedModules: [{ resource: 'shell', chunks: [1] }, { resource: '@tiptap/core', chunks: [2] }], retentionObservations: [] };
  assert.deepEqual(checkFixture(item, fixture, contract), []);
  item.emittedModules[1].chunks = [1, 2];
  assert.match(checkFixture(item, fixture, contract).join('\n'), /forbidden initial module group tiptap/);
});

// The public contract must retain all existing dependency checks while adding the
// exact implementation-only engine inventory (including optional menu pins).
test('package compatibility gate permits only the exact private engine dependency inventory', () => {
  const { assertCompatiblePackageContract } = require('../scripts/api-compatibility.helpers.cjs');
  const baseline = require('./fixtures/package-baseline.json');
  assert.doesNotThrow(() => assertCompatiblePackageContract(manifest, baseline));
  const changed = structuredClone(manifest); changed.dependencies['@tiptap/core'] = '^3.30.3';
  assert.throws(() => assertCompatiblePackageContract(changed, baseline), /Runtime dependencies changed/);
});

test('actual packed runtime requests no editor chunk in Read and loads and destroys engine in Edit', async () => {
  const { compileProbe, executeProbe } = require('../scripts/verify-production-runtime.cjs');
  fs.mkdirSync(path.join(root, 'temp'), { recursive: true });
  const directory = fs.mkdtempSync(path.join(root, 'temp/web-part-shell-runtime-'));
  try {
    const packed = JSON.parse(execFileSync('npm', ['pack', '--ignore-scripts', '--json', '--pack-destination', directory, '--cache', path.join(directory, '.npm-cache')], { cwd: library, encoding: 'utf8' }))[0];
    const installed = path.join(directory, 'node_modules/@apvee/spfx-react-toolkit');
    fs.mkdirSync(installed, { recursive: true });
    execFileSync('tar', ['-xzf', path.join(directory, packed.filename), '-C', installed, '--strip-components=1']);
    const bundle = await compileProbe({ consumerRoot: directory, outputDir: path.join(directory, 'assets'), scenario: 'web-part-shell', variant: 'control' });
    const observation = executeProbe(bundle.bundlePath, 'web-part-shell', directory);
    assert.equal(observation.observation.editorDestroyed, true);
    assert.ok(observation.observation.assetRequests.length > 0);
    assert.equal(observation.observation.childMounts, 1);
  } finally { fs.rmSync(directory, { recursive: true, force: true }); }
});

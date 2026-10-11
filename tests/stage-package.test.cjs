const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { execFileSync } = require('node:child_process');
const { createRequire } = require('node:module');

const root = path.resolve(__dirname, '..');
const library = path.join(root, 'packages/spfx-react-toolkit');

test('Stage has one explicit canonical alias without public implementation paths', () => {
  const manifest = JSON.parse(fs.readFileSync(path.join(library, 'package.json'), 'utf8'));
  assert.deepEqual(manifest.exports['./controls/stage'], {
    types: './lib/controls/stage/index.d.ts', default: './lib/controls/stage/index.js',
  });
  assert.deepEqual(manifest.typesVersions['*']['controls/stage'], ['lib/controls/stage/index.d.ts']);
  for (const key of Object.keys(manifest.exports)) {
    assert.ok(!key.includes('*'), 'Stage must not introduce wildcard exports');
    assert.ok(!key.includes('/stage/') && !key.startsWith('./lib/controls/stage'), `Private Stage path exported: ${key}`);
  }
});

test('packed Stage declarations resolve canonical compound exports and precise presentation props', () => {
  const temporaryRoot = path.join(root, 'temp');
  fs.mkdirSync(temporaryRoot, { recursive: true });
  const directory = fs.mkdtempSync(path.join(temporaryRoot, 'stage-package-'));
  try {
    const packed = JSON.parse(execFileSync('npm', [
      'pack', '--ignore-scripts', '--json', '--pack-destination', directory,
      '--cache', '/private/tmp/stage-npm-cache',
    ], { cwd: library, encoding: 'utf8' }))[0];
    const installed = path.join(directory, 'node_modules/@apvee/spfx-react-toolkit');
    fs.mkdirSync(installed, { recursive: true });
    execFileSync('tar', ['-xzf', path.join(directory, packed.filename), '-C', installed, '--strip-components=1']);
    const from = createRequire(path.join(directory, 'package.json'));
    assert.equal(from.resolve('@apvee/spfx-react-toolkit/controls/stage'), path.join(installed, 'lib/controls/stage/index.js'));
    for (const suffix of ['/controls/stage/stage', '/controls/stage/stage-context.internal', '/lib/controls/stage']) {
      assert.throws(() => from.resolve('@apvee/spfx-react-toolkit' + suffix), /not.*exported/i);
    }
    const fixture = path.join(directory, 'consumer.tsx');
    fs.copyFileSync(path.join(root, 'tests/fixtures/stage-package-consumer.tsx'), fixture);
    for (const [moduleResolution, module] of [
      [ts.ModuleResolutionKind.Node10, ts.ModuleKind.ESNext],
      [ts.ModuleResolutionKind.Bundler, ts.ModuleKind.ESNext],
      [ts.ModuleResolutionKind.Node16, ts.ModuleKind.Node16],
    ]) {
      const program = ts.createProgram([fixture], {
        noEmit: true, strict: true, skipLibCheck: true, moduleResolution, module,
        jsx: ts.JsxEmit.React, target: ts.ScriptTarget.ES2020, types: [],
      });
      const diagnostics = ts.getPreEmitDiagnostics(program);
      assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, {
        getCanonicalFileName: value => value, getCurrentDirectory: () => directory, getNewLine: () => '\n',
      }));
      const checker = program.getTypeChecker();
      const source = program.getSourceFile(fixture);
      const definitions = {};
      for (const node of source.statements.filter(ts.isImportDeclaration)) {
        for (const element of node.importClause?.namedBindings?.elements ?? []) {
          const symbol = checker.getSymbolAtLocation(element.name);
          definitions[element.name.text] = checker.getAliasedSymbol(symbol).declarations;
        }
      }
      for (const family of [['rootStage', 'domainStage', 'controlStage'], ['RootProps', 'DomainProps', 'ControlProps']]) {
        assert.deepEqual(definitions[family[0]], definitions[family[1]]);
        assert.deepEqual(definitions[family[0]], definitions[family[2]]);
      }
      assert.equal(program.getSourceFiles().filter(file => /[/\\](?:@tiptap|prosemirror-[^/\\]+)[/\\]/.test(file.fileName)).length, 0);
    }
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

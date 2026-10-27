const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { execFileSync } = require('node:child_process');
const { createRequire } = require('node:module');

const root = path.resolve(__dirname, '..');
const library = path.join(root, 'packages/spfx-react-toolkit');

test('DescriptionList import aliases share zero-growth budgets and unrelated consumers exclude the control', () => {
  const contract = require('./fixtures/tree-shaking/contract.json');
  const { checkFixture } = require('../scripts/bundle-contract.internal.cjs');
  const fixture = contract.fixtures.find(item => item.id === 'description-list');
  assert.ok(fixture);
  assert.deepEqual(Object.keys(fixture.variants), ['root', 'domain', 'control']);
  for (const alias of ['root', 'domain']) {
    const comparison = contract.comparisons.find(item => item.candidate === `description-list/${alias}`);
    assert.equal(comparison.reference, 'description-list/control');
    assert.equal(comparison.maxIncreaseBytes, 0);
  }
  for (const group of ['pnp', 'form', 'data-grid', 'shell', 'stage', 'pagination', 'markdown']) {
    assert.ok(fixture.forbiddenModuleGroups.includes(group), group);
  }
  for (const other of contract.fixtures.filter(item => item !== fixture)) {
    assert.ok(other.forbiddenModuleGroups.includes('description-list'), other.id);
    assert.ok(checkFixture({ id: `${other.id}/root`, initialChunkIds: [1], emittedModules: [{
      resource: '@apvee/spfx-react-toolkit/lib/controls/description-list/description-list.js',
      chunks: [1],
    }], retentionObservations: [] }, other, contract).some(message => message.includes('forbidden module group')));
  }
});

test('DescriptionList has one explicit canonical alias without public implementation paths', () => {
  const manifest = JSON.parse(fs.readFileSync(path.join(library, 'package.json'), 'utf8'));
  assert.deepEqual(manifest.exports['./controls/description-list'], {
    types: './lib/controls/description-list/index.d.ts', default: './lib/controls/description-list/index.js',
  });
  assert.deepEqual(manifest.typesVersions['*']['controls/description-list'], ['lib/controls/description-list/index.d.ts']);
  for (const key of Object.keys(manifest.exports)) {
    assert.ok(!key.includes('*'), 'DescriptionList must not introduce wildcard exports');
    assert.ok(!key.includes('/description-list/') && !key.startsWith('./lib/controls/description-list'), `Private DescriptionList path exported: ${key}`);
  }
});

test('packed DescriptionList declarations resolve canonical exports and mutually exclusive content props', () => {
  const temporaryRoot = path.join(root, 'temp');
  fs.mkdirSync(temporaryRoot, { recursive: true });
  const directory = fs.mkdtempSync(path.join(temporaryRoot, 'description-list-package-'));
  try {
    const packed = JSON.parse(execFileSync('npm', [
      'pack', '--ignore-scripts', '--json', '--pack-destination', directory,
      '--cache', path.join(directory, '.npm-cache'),
    ], { cwd: library, encoding: 'utf8' }))[0];
    const installed = path.join(directory, 'node_modules/@apvee/spfx-react-toolkit');
    fs.mkdirSync(installed, { recursive: true });
    execFileSync('tar', ['-xzf', path.join(directory, packed.filename), '-C', installed, '--strip-components=1']);
    const from = createRequire(path.join(directory, 'package.json'));
    assert.equal(from.resolve('@apvee/spfx-react-toolkit/controls/description-list'), path.join(installed, 'lib/controls/description-list/index.js'));
    for (const suffix of ['/controls/description-list/description-list', '/controls/description-list/description-list-context.internal', '/lib/controls/description-list']) {
      assert.throws(() => from.resolve('@apvee/spfx-react-toolkit' + suffix), /not.*exported/i);
    }
    const fixture = path.join(directory, 'consumer.tsx');
    fs.copyFileSync(path.join(root, 'tests/fixtures/description-list-package-consumer.tsx'), fixture);
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
      for (const family of [['rootDescriptionList', 'domainDescriptionList', 'controlDescriptionList'], ['RootProps', 'DomainProps', 'ControlProps']]) {
        assert.deepEqual(definitions[family[0]], definitions[family[1]]);
        assert.deepEqual(definitions[family[0]], definitions[family[2]]);
      }
      assert.equal(program.getSourceFiles().filter(file => /[/\\](?:@tiptap|prosemirror-[^/\\]+)[/\\]/.test(file.fileName)).length, 0);
    }
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

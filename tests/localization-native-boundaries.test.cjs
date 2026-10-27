'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const packageRoot = path.resolve(__dirname, '../packages/spfx-react-toolkit');
const sourceRoot = path.join(packageRoot, 'src');
const families = JSON.parse(fs.readFileSync(path.join(packageRoot, 'localization/resources.config.json'), 'utf8')).families;
function createLoader(nativeResources) {
  const cache = new Map();
  return function load(relative) {
    const file = path.resolve(sourceRoot, relative);
    if (cache.has(file)) return cache.get(file).exports;
    const module = { exports: {} }; cache.set(file, module);
    const output = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: {
      target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.CommonJS
    } }).outputText;
    const sourceRequire = id => {
      if (!id.startsWith('.')) return nativeResources.require(id);
      return load(path.resolve(path.dirname(file), id + '.ts'));
    };
    vm.runInThisContext('(function(require,module,exports){' + output + '\n})', { filename: file })(sourceRequire, module, module.exports);
    return module.exports;
  };
}
function fixture(options = { culture: 'en-US' }) {
  return require('./fixtures/localization/native-resources.cjs').createNativeResources(options);
}
test('native fixture rejects unknown IDs and requires an explicit supported resource culture', () => {
  const resources = fixture();
  assert.equal(resources.has('ApveeSpfxReactToolkitStageStrings'), true);
  assert.equal(resources.has('ApveeSpfxReactToolkitUnknownStrings'), false);
  assert.throws(() => resources.require('ApveeSpfxReactToolkitUnknownStrings'), /unknown native resource/i);
  assert.throws(() => fixture({}), /culture/i);
  assert.throws(() => fixture({ culture: 'xx-XX' }), /culture/i);
  assert.throws(() => fixture({ culture: 'en-US', overrides: { UnexpectedStrings: {} } }), /unknown native resource/i);
});
test('native fixture rejects a missing emitted AMD asset instead of synthesizing source data', t => {
  const assetRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'toolkit-missing-native-assets-'));
  t.after(() => fs.rmSync(assetRoot, { recursive: true, force: true }));
  const resources = fixture({ culture: 'en-US', assetRoot });
  assert.throws(() => resources.require('ApveeSpfxReactToolkitPaginationStrings'), error => {
    assert.match(error.message, /missing native resource AMD asset/i);
    assert.ok(error.message.includes(path.join(assetRoot, 'lib/controls/pagination/loc/en-us.js')));
    assert.match(error.message, /npm run build:library/);
    return true;
  });
  assert.deepEqual(fs.readdirSync(assetRoot), []);
});
for (const family of families) test(`${family.family} imports only its own native resource and preserves culture data`, () => {
  const nativeResources = fixture({ culture: 'it-IT' });
  const load = createLoader(nativeResources);
  const name = family.resourceType.export.replace(/Resource$/, '');
  const native = load(`controls/${family.family}/localization/${family.family}-native.internal.ts`)[`native${name}Resource`];
  assert.equal(native.__locale, 'it-IT');
  assert.deepEqual(nativeResources.requestedIds, [family.id]);
  const source = JSON.parse(fs.readFileSync(path.join(packageRoot, family.sourceDirectory, 'it-IT.json'), 'utf8'));
  assert.deepEqual(JSON.parse(JSON.stringify(native)), { ...source, __locale: 'it-IT' });
});
test('template interpolation is literal, nonrecursive and rejects missing own values', () => {
  const { formatResourceTemplate } = createLoader(fixture())('localization/resource-formatters.internal.ts');
  const literal = "$& $' $` {name} {other}";
  assert.equal(formatResourceTemplate('{name}: {other} / {name}', { name: literal, other: 12, extra: 'ignored' }), `${literal}: 12 / ${literal}`);
  assert.equal(formatResourceTemplate('{0}', { 0: 'Calendar' }), 'Calendar');
  assert.equal(formatResourceTemplate('{name}', { name: '' }), '');
  assert.throws(() => formatResourceTemplate('{missing}', {}), /missing.*missing/i);
  assert.throws(() => formatResourceTemplate('{toString}', {}), /missing.*toString/i);
});
test('plural templates use native grammar, other fallback and caller formatted values', () => {
  const { formatResourcePlural } = createLoader(fixture())('localization/resource-formatters.internal.ts');
  const forms = { zero: 'zero {count}', one: 'one {count}', two: 'two {count}', few: 'few {count}', many: 'many {count}', other: 'other {count}' };
  assert.equal(formatResourcePlural(forms, 'ar-SA', 0), 'zero 0');
  assert.equal(formatResourcePlural(forms, 'ar-SA', 2), 'two 2');
  assert.equal(formatResourcePlural(forms, 'ar-SA', 3), 'few 3');
  assert.equal(formatResourcePlural(forms, 'en-US', 3), 'other 3');
  assert.equal(formatResourcePlural({ other: '{count}' }, 'ar-SA', 3, { count: '٣' }), '٣');
  assert.equal(formatResourcePlural({ one: '{formattedCount} rating', other: '{formattedCount} ratings' }, 'en-US', 1, { formattedCount: '1,0' }), '1,0 rating');
});
test('native fixture overrides are isolated and evaluated actual AMD data is cached per fixture', () => {
  const id = 'ApveeSpfxReactToolkitPaginationStrings';
  const first = fixture({ culture: 'en-US', overrides: { [id]: { nextPage: 'Next override' } } });
  assert.equal(first.require(id).nextPage, 'Next override');
  assert.equal(first.require(id), first.require(id));
  assert.equal(fixture().require(id).nextPage, 'Next page');
  assert.equal(fixture({ culture: 'it-IT' }).require(id).nextPage, 'Pagina successiva');
  assert.deepEqual(first.requestedIds, [id, id, id]);
});
test('native declarations compile from direct private data contracts and emit no ambient import dependency', () => {
  const output = new Map();
  const entries = families.map(family => path.join(sourceRoot, 'controls', family.family, 'localization', `${family.family}-native.internal.ts`));
  entries.push(path.join(sourceRoot, 'localization/native-resources.internal.d.ts'), path.join(sourceRoot, 'localization/resource-formatters.internal.ts'));
  const options = { strict: true, skipLibCheck: true, target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ESNext,
    moduleResolution: ts.ModuleResolutionKind.Node10, types: [], declaration: true, emitDeclarationOnly: true };
  const program = ts.createProgram(entries, options);
  const diagnostics = ts.getPreEmitDiagnostics(program);
  assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: file => file, getCurrentDirectory: () => process.cwd(), getNewLine: () => '\n'
  }));
  const result = program.emit(undefined, (file, bytes) => output.set(file, bytes));
  assert.equal(result.emitSkipped, false);
  for (const [file, bytes] of output) {
    assert.doesNotMatch(bytes, /ApveeSpfxReactToolkit\w+Strings|editable-text-document\.types\.internal/, file);
  }
});
test('public declaration graph never requires native ambient resources or private editor document types', () => {
  const output = new Map();
  const entry = path.join(sourceRoot, 'index.ts');
  const program = ts.createProgram([entry, path.join(sourceRoot, 'localization/native-resources.internal.d.ts')], {
    skipLibCheck: true, target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ESNext,
    moduleResolution: ts.ModuleResolutionKind.Node10, jsx: ts.JsxEmit.React,
    declaration: true, emitDeclarationOnly: true, rootDir: sourceRoot, outDir: '/virtual-localization-declarations'
  });
  program.emit(undefined, (file, bytes) => output.set(file, bytes));
  const visited = new Set();
  function visit(file) {
    if (visited.has(file)) return;
    visited.add(file);
    const bytes = output.get(file);
    assert.ok(bytes, `Expected emitted declaration ${file}`);
    assert.doesNotMatch(bytes, /ApveeSpfxReactToolkit\w+Strings|native-resources\.internal|editable-text-document\.types\.internal/, file);
    const ast = ts.createSourceFile(file, bytes, ts.ScriptTarget.ES2020, true);
    function inspect(node) {
      if (ts.isStringLiteral(node) && node.text.startsWith('.')) {
        const target = path.resolve(path.dirname(file), node.text);
        const dependency = [target + '.d.ts', path.join(target, 'index.d.ts')].find(candidate => output.has(candidate));
        if (dependency) visit(dependency);
      }
      ts.forEachChild(node, inspect);
    }
    inspect(ast);
  }
  visit('/virtual-localization-declarations/index.d.ts');
});

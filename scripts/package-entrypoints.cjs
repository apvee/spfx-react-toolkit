// Versioned additive entrypoints. Generation is explicit; checks never write.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createRequire } = require('node:module');
const { ResolverFactory } = require('enhanced-resolve');
const ts = require('typescript');
const root = path.resolve(__dirname, '..');
const library = path.join(root, 'packages/spfx-react-toolkit');
const inventory = require('../tests/fixtures/package-entrypoints.json');
const cleanAliases = Object.freeze({
  '.': 'lib/index', './core': 'lib/core/index', './hooks': 'lib/hooks/index',
  './helpers': 'lib/helpers/index', './services': 'lib/services/index',
  './styles': 'lib/styles/index', './styles/useSx': 'lib/hooks/useSx',
  './hooks/useStableCallback': 'lib/hooks/useStableCallback',
});

/** Derive explicit exports without changing the frozen historical inventory. */
function createEntrypointContract({ publishedFiles, cleanAliases: aliases = cleanAliases }) {
  assert.deepEqual(aliases, cleanAliases, 'Unapproved clean alias contract');
  const files = new Set(publishedFiles);
  assert.equal(files.size, publishedFiles.length, 'Duplicate published file');
  for (const file of inventory.files) assert.ok(files.has(file), `Missing historical published file: ${file}`);
  const exports = {};
  function add(key, target) {
    assert.equal(exports[key], undefined, `Entrypoint collision: ${key}`);
    exports[key] = target;
  }
  function moduleTarget(base) {
    for (const extension of ['.d.ts', '.js']) assert.ok(files.has(base + extension), `Missing entrypoint target: ${base + extension}`);
    return { types: './' + base + '.d.ts', default: './' + base + '.js' };
  }
  const mappings = {};
  for (const [alias, base] of Object.entries(cleanAliases)) {
    add(alias, moduleTarget(base));
    if (alias !== '.') mappings[alias.slice(2)] = [base + '.d.ts'];
  }
  // Only historical physical paths become legacy aliases. New modules require
  // an explicit clean entrypoint decision rather than a permissive wildcard.
  for (const file of [...inventory.files].sort()) {
    assert.ok(!file.startsWith('/') && !file.split('/').includes('..'), `Invalid packed path: ${file}`);
    if (file.endsWith('.js')) {
      const base = file.slice(0, -3);
      const target = moduleTarget(base);
      add('./' + base, target);
      add('./' + file, target);
      if (base.endsWith('/index')) add('./' + base.slice(0, -6), target);
    } else add('./' + file, './' + file);
  }
  return { exports, typesVersions: { '*': mappings } };
}

function collectPublishedFiles(directory = library) {
  const files = ['LICENSE', 'README.md', 'package.json'];
  function visit(folder) {
    for (const entry of fs.readdirSync(folder, { withFileTypes: true })) {
      const file = path.join(folder, entry.name);
      if (entry.isDirectory()) visit(file);
      else files.push(path.relative(directory, file).split(path.sep).join('/'));
    }
  }
  visit(path.join(directory, 'lib'));
  for (const file of files) {
    assert.ok(fs.existsSync(path.join(directory, file)), `Missing published file: ${file}`);
    assert.ok(/^(package\.json|README\.md|LICENSE|lib\/(index\.[^/]+|(core|hooks|services|helpers|utils|styles)\/.+))$/.test(file), `Unexpected published file: ${file}`);
    if (!file.startsWith('lib/')) continue;
    const source = file.replace(/^lib\//, 'src/').replace(/\.(d\.ts|js)(\.map)?$/, '');
    assert.ok(fs.existsSync(path.join(directory, source + '.ts')) || fs.existsSync(path.join(directory, source + '.tsx')), `Published file lacks canonical source: ${file}`);
  }
  return files.sort();
}

function assertEntrypointContract(manifest, publishedFiles = [...inventory.files, 'lib/styles/index.js', 'lib/styles/index.d.ts']) {
  const contract = createEntrypointContract({ publishedFiles, cleanAliases });
  assert.deepEqual(manifest.exports, contract.exports, 'Package exports differ from the exact additive contract');
  // Object equality alone does not check condition priority.
  for (const [alias, target] of Object.entries(manifest.exports)) {
    if (typeof target === 'object') assert.deepEqual(Object.keys(target), ['types', 'default'], `Entrypoint condition order changed: ${alias}`);
  }
  assert.deepEqual(manifest.typesVersions, contract.typesVersions, 'Package typesVersions differ from exact clean aliases');
  assert.equal(manifest.type, undefined, 'Native Node module mode is not approved');
  assert.equal(manifest.module, undefined, 'An alternate runtime entrypoint is not approved');
}

/** Exercise actual consumer resolution; no paths aliases or resolver mocks. */
function verifyConsumerResolution(consumer, evidenceDirectory) {
  const fromHost = createRequire(path.join(consumer, 'package.json'));
  const packageFile = fromHost.resolve('@apvee/spfx-react-toolkit/package.json');
  const packageDirectory = path.dirname(packageFile);
  const resolver = ResolverFactory.createResolver({
    fileSystem: fs, useSyncFileSystemCalls: true, extensions: ['.js'], conditionNames: ['default'],
    exportsFields: ['exports'], mainFields: ['main'], symlinks: false,
  });
  const requests = {
    '@apvee/spfx-react-toolkit': 'lib/index.js',
    ...Object.fromEntries(Object.entries(cleanAliases).filter(([key]) => key !== '.').map(([key, base]) => ['@apvee/spfx-react-toolkit' + key.slice(1), base + '.js'])),
    '@apvee/spfx-react-toolkit/lib': 'lib/index.js',
    '@apvee/spfx-react-toolkit/lib/helpers/styles': 'lib/helpers/styles/index.js',
    '@apvee/spfx-react-toolkit/lib/hooks/useSx': 'lib/hooks/useSx.js',
    '@apvee/spfx-react-toolkit/lib/hooks/useSx.js': 'lib/hooks/useSx.js',
    '@apvee/spfx-react-toolkit/lib/hooks/useSx.d.ts': 'lib/hooks/useSx.d.ts',
    '@apvee/spfx-react-toolkit/lib/hooks/useSx.js.map': 'lib/hooks/useSx.js.map',
    '@apvee/spfx-react-toolkit/lib/hooks/useSx.d.ts.map': 'lib/hooks/useSx.d.ts.map',
    '@apvee/spfx-react-toolkit/lib/hooks/useAsyncInvoke.internal': 'lib/hooks/useAsyncInvoke.internal.js',
    '@apvee/spfx-react-toolkit/package.json': 'package.json',
    '@apvee/spfx-react-toolkit/README.md': 'README.md',
    '@apvee/spfx-react-toolkit/LICENSE': 'LICENSE',
  };
  const runtime = [];
  for (const [request, expected] of Object.entries(requests)) {
    const resolved = resolver.resolveSync({}, consumer, request);
    assert.equal(fs.realpathSync(resolved), fs.realpathSync(path.join(packageDirectory, expected)), `Runtime resolution changed: ${request}`);
    runtime.push({ request, file: resolved });
  }
  for (const suffix of ['/hooks/useAsyncInvoke.internal', '/styles/descriptor.internal', '/styles/types']) {
    assert.throws(() => resolver.resolveSync({}, consumer, '@apvee/spfx-react-toolkit' + suffix), /not exported/);
  }
  const fixture = path.join(consumer, 'compatibility/entrypoint-resolution.ts');
  fs.mkdirSync(path.dirname(fixture), { recursive: true });
  fs.copyFileSync(path.join(root, 'tests/fixtures/tree-shaking/resolution/consumer.ts'), fixture);
  const modes = [];
  const declarations = {};
  const importedNames = ['rootWidth', 'helperWidth', 'width', 'directoryWidth', 'rootSx', 'useSx', 'leafSx', 'legacySx', 'rootCallback', 'useStableCallback', 'legacyCallback'];
  for (const [mode, moduleResolution, module] of [
    ['node', ts.ModuleResolutionKind.Node10, ts.ModuleKind.ESNext],
    ['bundler', ts.ModuleResolutionKind.Bundler, ts.ModuleKind.ESNext],
    ['node16', ts.ModuleResolutionKind.Node16, ts.ModuleKind.Node16],
  ]) {
    const options = { strict: true, skipLibCheck: true, noEmit: true, target: ts.ScriptTarget.ES2020, module, moduleResolution, types: [] };
    const program = ts.createProgram([fixture], options);
    const diagnostics = ts.getPreEmitDiagnostics(program);
    assert.equal(diagnostics.length, 0, `${mode} resolution failed:\n${ts.formatDiagnosticsWithColorAndContext(diagnostics, { getCanonicalFileName: file => file, getCurrentDirectory: () => consumer, getNewLine: () => '\n' })}`);
    const resolvedRoot = ts.resolveModuleName('@apvee/spfx-react-toolkit', fixture, options, ts.sys).resolvedModule;
    assert.equal(fs.realpathSync(resolvedRoot.resolvedFileName), fs.realpathSync(path.join(packageDirectory, 'lib/index.d.ts')), 'Root must not be rewritten to lib/lib');
    const symbols = new Map();
    const checker = program.getTypeChecker();
    for (const statement of program.getSourceFile(fixture).statements) {
      if (!ts.isImportDeclaration(statement) || !statement.importClause?.namedBindings || !ts.isNamedImports(statement.importClause.namedBindings)) continue;
      for (const element of statement.importClause.namedBindings.elements) {
        if (importedNames.includes(element.name.text)) symbols.set(element.name.text, checker.getAliasedSymbol(checker.getSymbolAtLocation(element.name)));
      }
    }
    for (const group of [['rootWidth', 'helperWidth', 'width', 'directoryWidth'], ['rootSx', 'useSx', 'leafSx', 'legacySx'], ['rootCallback', 'useStableCallback', 'legacyCallback']]) {
      for (const name of group) assert.equal(symbols.get(name), symbols.get(group[0]), `Canonical symbol identity changed: ${name}`);
    }
    declarations[mode] = resolvedRoot.resolvedFileName;
    modes.push(mode);
  }
  const result = { typescript: ts.version, modes, runtime, declarations };
  if (evidenceDirectory) fs.writeFileSync(path.join(evidenceDirectory, 'entrypoint-resolution.json'), JSON.stringify(result, null, 2));
  return result;
}

function main(args = process.argv.slice(2)) {
  assert.ok(args.length === 1 && ['--check', '--write'].includes(args[0]), 'Use --check or --write');
  const file = path.join(library, 'package.json');
  const manifest = JSON.parse(fs.readFileSync(file, 'utf8'));
  const publishedFiles = collectPublishedFiles();
  if (args[0] === '--write') {
    Object.assign(manifest, createEntrypointContract({ publishedFiles, cleanAliases }));
    if (!manifest.files.includes('lib/styles/**/*')) manifest.files.push('lib/styles/**/*');
    fs.writeFileSync(file, JSON.stringify(manifest, null, 2) + '\n');
  }
  assertEntrypointContract(manifest, publishedFiles);
  console.log(`Package entrypoints passed: ${Object.keys(manifest.exports).length} exact exports, ${inventory.files.length} historical packed files.`);
}
if (require.main === module) main();
module.exports = { createEntrypointContract, assertEntrypointContract, collectPublishedFiles, verifyConsumerResolution, cleanAliases, main };

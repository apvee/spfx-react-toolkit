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
  './controls/file-folder-picker': 'lib/controls/file-folder-picker/index',
  './property-pane/file-folder-picker': 'lib/property-pane/file-folder-picker/index',
  './controls/pickers': 'lib/controls/pickers/index',
  './property-pane/pickers': 'lib/property-pane/pickers/index',
  './controls': 'lib/controls/index', './controls/formatted-text': 'lib/controls/formatted-text/index',
  './controls/relative-time': 'lib/controls/relative-time/index',
  './controls/auto-skeleton': 'lib/controls/auto-skeleton/index',
  './controls/state-message': 'lib/controls/state-message/index',
  './controls/settings-group': 'lib/controls/settings-group/index',
  './controls/accordion-group': 'lib/controls/accordion-group/index',
  './controls/form': 'lib/controls/form/index',
  './controls/data-grid': 'lib/controls/data-grid/index',
  './controls/editable-text': 'lib/controls/editable-text/index',
  './controls/web-part-shell': 'lib/controls/web-part-shell/index',
  './controls/stage': 'lib/controls/stage/index',
  './controls/pagination': 'lib/controls/pagination/index',
  './controls/description-list': 'lib/controls/description-list/index',
  './controls/content-card': 'lib/controls/content-card/index',
  './controls/preview-card': 'lib/controls/preview-card/index',
  './controls/sharepoint-permission-trim': 'lib/controls/sharepoint-permission-trim/index',
  './controls/sharepoint-group-trim': 'lib/controls/sharepoint-group-trim/index',
  './controls/entra-group-trim': 'lib/controls/entra-group-trim/index',
});

/** Exact build-only contract for native resource assets and the installed setup CLI. */
function getLocalizationPackageContract(directory = library) {
  const file = path.join(directory, 'localization/resources.json');
  if (!fs.existsSync(file)) return { extraFiles: [], nativeIds: new Set(), assetSources: new Map() };
  const data = JSON.parse(fs.readFileSync(file, 'utf8'));
  assert.equal(data.schemaVersion, 1, 'Unsupported localization inventory');
  assert.ok(Array.isArray(data.resources), 'Localization resources must be an array');
  const nativeIds = new Set();
  const assetSources = new Map();
  for (const resource of data.resources) {
    assert.match(resource.id, /^ApveeSpfxReactToolkit[A-Za-z0-9]+Strings$/);
    assert.ok(!nativeIds.has(resource.id), `Duplicate native resource ID: ${resource.id}`);
    nativeIds.add(resource.id);
    const match = /^lib\/controls\/([a-z]+(?:-[a-z]+)*)\/loc\/\{locale\}\.js$/.exec(resource.path);
    assert.ok(match, `Invalid native resource path: ${resource.path}`);
    assert.ok(Array.isArray(resource.locales) && resource.locales.includes('en-US'), 'Native resources need en-US');
    for (const culture of resource.locales) {
      assert.match(culture, /^[a-z]{2,3}(?:-[A-Z][a-z]{3})?-[A-Z]{2}$/);
      const asset = resource.path.replace('{locale}', culture.toLowerCase());
      assert.ok(!assetSources.has(asset), `Duplicate native resource asset: ${asset}`);
      assetSources.set(asset, `src/controls/${match[1]}/localization/resources/${culture}.json`);
    }
  }
  return { extraFiles: ['cli/spfx-react-toolkit.cjs', 'localization/resources.json'], nativeIds, assetSources };
}

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
  const localization = getLocalizationPackageContract(directory);
  const files = ['LICENSE', 'README.md', 'THIRD-PARTY-NOTICES.md', 'package.json', ...localization.extraFiles];
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
    assert.ok(localization.extraFiles.includes(file) || /^(package\.json|README\.md|LICENSE|THIRD-PARTY-NOTICES\.md|lib\/(index\.[^/]+|(core|hooks|services|helpers|utils|styles|controls|property-pane|localization)\/.+))$/.test(file), `Unexpected published file: ${file}`);
    if (!file.startsWith('lib/')) continue;
    if (localization.assetSources.has(file)) {
      assert.ok(fs.existsSync(path.join(directory, localization.assetSources.get(file))), `Native resource lacks canonical translator source: ${file}`);
      continue;
    }
    const source = file.replace(/^lib\//, 'src/').replace(/\.(d\.ts|js)(\.map)?$/, '');
    assert.ok(fs.existsSync(path.join(directory, source + '.ts')) || fs.existsSync(path.join(directory, source + '.tsx')), `Published file lacks canonical source: ${file}`);
  }
  return files.sort();
}

function assertEntrypointContract(manifest, publishedFiles = [...inventory.files, ...['lib/controls/file-folder-picker/index', 'lib/property-pane/file-folder-picker/index', 'lib/controls/pickers/index', 'lib/property-pane/pickers/index', 'lib/styles/index', 'lib/controls/index', 'lib/controls/formatted-text/index', 'lib/controls/relative-time/index', 'lib/controls/auto-skeleton/index', 'lib/controls/state-message/index', 'lib/controls/settings-group/index', 'lib/controls/accordion-group/index', 'lib/controls/form/index', 'lib/controls/data-grid/index', 'lib/controls/editable-text/index', 'lib/controls/web-part-shell/index', 'lib/controls/stage/index', 'lib/controls/pagination/index', 'lib/controls/description-list/index', 'lib/controls/content-card/index', 'lib/controls/preview-card/index', 'lib/controls/sharepoint-permission-trim/index', 'lib/controls/sharepoint-group-trim/index', 'lib/controls/entra-group-trim/index'].flatMap(base => ['.js', '.d.ts'].map(ext => base + ext))]) {
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
  for (const suffix of ['/hooks/useAsyncInvoke.internal', '/styles/descriptor.internal', '/styles/types', '/controls/form/form', '/controls/form/create-form', '/controls/form/controls/form-input', '/controls/form/plugins', '/controls/form/advanced', '/lib/controls/form', '/lib/controls/form/index', '/controls/data-grid/data-grid', '/controls/data-grid/create-data-grid', '/controls/data-grid/column-family.internal', '/controls/data-grid/cells/data-grid-cells.internal', '/controls/data-grid/view', '/lib/controls/data-grid', '/lib/controls/data-grid/index', '/controls/editable-text/editable-text', '/controls/editable-text/editable-text-editor.internal', '/controls/editable-text/editable-text-codec.internal', '/controls/editable-text/editable-text-document.types.internal', '/lib/controls/editable-text', '/lib/controls/editable-text/index', '/controls/web-part-shell/web-part-shell', '/controls/web-part-shell/title/web-part-title-editor.internal', '/controls/web-part-shell/title/web-part-title-color-picker.internal', '/controls/web-part-shell/title/web-part-title-toolbar.internal', '/lib/controls/web-part-shell/index', '/controls/stage/stage', '/controls/stage/stage-context.internal', '/lib/controls/stage', '/controls/content-slot.internal', '/controls/content-slot.types.internal', '/controls/cards/card-frame.internal', '/controls/content-card/content-card', '/controls/preview-card/preview-card', '/lib/controls/cards', '/lib/controls/content-card', '/lib/controls/preview-card', '/hooks/useAccessRequest.internal', '/services/spfx-scoped-permissions.internal', '/services/spfx-sharepoint-groups.internal', '/services/spfx-entra-groups.internal', '/controls/access-trim-children.internal', '/controls/sharepoint-permission-trim/sharepoint-permission-trim', '/controls/sharepoint-group-trim/sharepoint-group-trim', '/controls/entra-group-trim/entra-group-trim', '/lib/controls/sharepoint-permission-trim', '/lib/controls/sharepoint-group-trim', '/lib/controls/entra-group-trim', '/lib/services/spfx-permission-target.types']) {
    assert.throws(() => resolver.resolveSync({}, consumer, '@apvee/spfx-react-toolkit' + suffix), /not exported/);
  }
  const fixture = path.join(consumer, 'compatibility/entrypoint-resolution.ts');
  fs.mkdirSync(path.dirname(fixture), { recursive: true });
  fs.copyFileSync(path.join(root, 'tests/fixtures/tree-shaking/resolution/consumer.ts'), fixture);
  const modes = [];
  const declarations = {};
  const importedNames = ['rootEditable','domainEditable','controlEditable','RootEditableProps','DomainEditableProps','ControlEditableProps','RootEditableBlock','DomainEditableBlock','ControlEditableBlock','RootEditableChange','DomainEditableChange','ControlEditableChange','rootSitePicker','controlSitePicker','domainSitePicker','rootListPicker','controlListPicker','domainListPicker','rootFieldPicker','controlFieldPicker','domainFieldPicker','rootViewPicker','controlViewPicker','domainViewPicker','rootItemPicker','controlItemPicker','domainItemPicker','rootPrincipalPicker','controlPrincipalPicker','domainPrincipalPicker','RootPrincipalValue','ControlPrincipalValue','DomainPrincipalValue','NativePrincipalValue','rootWidth', 'helperWidth', 'width', 'directoryWidth', 'rootSx', 'useSx', 'leafSx', 'legacySx', 'rootCallback', 'useStableCallback', 'legacyCallback', 'rootText', 'domainText', 'controlText', 'rootRelative', 'domainRelative', 'controlRelative', 'rootSkeleton', 'domainSkeleton', 'controlSkeleton', 'rootState', 'domainState', 'controlState', 'rootForm', 'domainForm', 'controlForm', 'rootFactory', 'domainFactory', 'controlFactory', 'RootMethods', 'DomainMethods', 'ControlMethods', 'rootGrid', 'domainGrid', 'controlGrid', 'rootGridFactory', 'domainGridFactory', 'controlGridFactory', 'RootGridProps', 'DomainGridProps', 'ControlGridProps', 'rootShell', 'domainShell', 'controlShell', 'rootMetadata', 'domainMetadata', 'controlMetadata', 'RootShellProps', 'DomainShellProps', 'ControlShellProps', 'rootStage', 'domainStage', 'controlStage', 'RootStageProps', 'DomainStageProps', 'ControlStageProps', 'rootContentCard', 'domainContentCard', 'controlContentCard', 'RootCardProps', 'DomainCardProps', 'ControlCardProps', 'rootPreviewCard', 'domainPreviewCard', 'controlPreviewCard', 'RootPreviewProps', 'DomainPreviewProps', 'ControlPreviewProps', 'rootPermissionTrim', 'domainPermissionTrim', 'controlPermissionTrim', 'RootPermissionTrimProps', 'DomainPermissionTrimProps', 'ControlPermissionTrimProps', 'rootSPGroupTrim', 'domainSPGroupTrim', 'controlSPGroupTrim', 'RootSPGroupTrimProps', 'DomainSPGroupTrimProps', 'ControlSPGroupTrimProps', 'rootEntraTrim', 'domainEntraTrim', 'controlEntraTrim', 'RootEntraTrimProps', 'DomainEntraTrimProps', 'ControlEntraTrimProps', 'rootScopedHook', 'hooksScopedHook', 'legacyScopedHook', 'rootScopedInfo', 'hooksScopedInfo', 'legacyScopedInfo', 'rootScopedOptions', 'hooksScopedOptions', 'legacyScopedOptions', 'RootSharePointPermissionTrimBranchProps', 'DomainSharePointPermissionTrimBranchProps', 'ControlSharePointPermissionTrimBranchProps', 'RootSharePointGroupSelector', 'DomainSharePointGroupSelector', 'ControlSharePointGroupSelector', 'RootSharePointGroupTrimBranchProps', 'DomainSharePointGroupTrimBranchProps', 'ControlSharePointGroupTrimBranchProps', 'RootEntraGroupSelector', 'DomainEntraGroupSelector', 'ControlEntraGroupSelector', 'RootEntraGroupType', 'DomainEntraGroupType', 'ControlEntraGroupType', 'RootEntraGroupRelationship', 'DomainEntraGroupRelationship', 'ControlEntraGroupRelationship', 'RootEntraGroupTrimBranchProps', 'DomainEntraGroupTrimBranchProps', 'ControlEntraGroupTrimBranchProps'];
  for (const [mode, moduleResolution, module] of [
    ['node', ts.ModuleResolutionKind.Node10, ts.ModuleKind.ESNext],
    ['bundler', ts.ModuleResolutionKind.Bundler, ts.ModuleKind.ESNext],
    ['node16', ts.ModuleResolutionKind.Node16, ts.ModuleKind.Node16],
  ]) {
    const options = { strict: true, skipLibCheck: true, noEmit: true, target: ts.ScriptTarget.ES2020, module, moduleResolution, types: [] };
    const program = ts.createProgram([fixture], options);
    const diagnostics = ts.getPreEmitDiagnostics(program);
    assert.equal(diagnostics.length, 0, `${mode} resolution failed:\n${ts.formatDiagnosticsWithColorAndContext(diagnostics, { getCanonicalFileName: file => file, getCurrentDirectory: () => consumer, getNewLine: () => '\n' })}`);
    if (mode !== 'node') {
      // Node10 intentionally ignores exports maps and can discover packed
      // physical modules. Only modern resolution enforces this boundary.
      for (const suffix of ['/lib/controls/form', '/lib/controls/form/index', '/lib/controls/data-grid', '/lib/controls/data-grid/index', '/controls/editable-text/editable-text', '/controls/editable-text/editable-text-editor.internal', '/controls/editable-text/editable-text-codec.internal', '/controls/editable-text/editable-text-document.types.internal', '/lib/controls/editable-text', '/lib/controls/editable-text/index', '/controls/web-part-shell/web-part-shell', '/controls/web-part-shell/title/web-part-title-editor.internal', '/controls/web-part-shell/title/web-part-title-color-picker.internal', '/controls/web-part-shell/title/web-part-title-toolbar.internal', '/lib/controls/web-part-shell/index', '/controls/stage/stage', '/controls/stage/stage-context.internal', '/lib/controls/stage', '/controls/content-slot.internal', '/controls/content-slot.types.internal', '/controls/cards/card-frame.internal', '/controls/content-card/content-card', '/controls/preview-card/preview-card', '/lib/controls/cards', '/lib/controls/content-card', '/lib/controls/preview-card', '/hooks/useAccessRequest.internal', '/services/spfx-scoped-permissions.internal', '/services/spfx-sharepoint-groups.internal', '/services/spfx-entra-groups.internal', '/controls/access-trim-children.internal', '/controls/sharepoint-permission-trim/sharepoint-permission-trim', '/controls/sharepoint-group-trim/sharepoint-group-trim', '/controls/entra-group-trim/entra-group-trim', '/lib/controls/sharepoint-permission-trim', '/lib/controls/sharepoint-group-trim', '/lib/controls/entra-group-trim', '/lib/services/spfx-permission-target.types']) assert.equal(ts.resolveModuleName('@apvee/spfx-react-toolkit' + suffix, fixture, options, ts.sys).resolvedModule, undefined, `${mode} must reject unexported control physical paths`);
    }
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
    for (const group of [['rootEditable', 'domainEditable', 'controlEditable'],['RootEditableProps', 'DomainEditableProps', 'ControlEditableProps'],['RootEditableBlock', 'DomainEditableBlock', 'ControlEditableBlock'],['RootEditableChange', 'DomainEditableChange', 'ControlEditableChange'], ['rootSitePicker', 'controlSitePicker', 'domainSitePicker'],['rootListPicker', 'controlListPicker', 'domainListPicker'],['rootFieldPicker', 'controlFieldPicker', 'domainFieldPicker'],['rootViewPicker', 'controlViewPicker', 'domainViewPicker'],['rootItemPicker', 'controlItemPicker', 'domainItemPicker'],['rootPrincipalPicker', 'controlPrincipalPicker', 'domainPrincipalPicker'],['RootPrincipalValue', 'ControlPrincipalValue', 'DomainPrincipalValue', 'NativePrincipalValue'], ['rootWidth', 'helperWidth', 'width', 'directoryWidth'], ['rootSx', 'useSx', 'leafSx', 'legacySx'], ['rootCallback', 'useStableCallback', 'legacyCallback'], ['rootText', 'domainText', 'controlText'], ['rootRelative', 'domainRelative', 'controlRelative'], ['rootSkeleton', 'domainSkeleton', 'controlSkeleton'], ['rootState', 'domainState', 'controlState'], ['rootForm', 'domainForm', 'controlForm'], ['rootFactory', 'domainFactory', 'controlFactory'], ['RootMethods', 'DomainMethods', 'ControlMethods'], ['rootGrid', 'domainGrid', 'controlGrid'], ['rootGridFactory', 'domainGridFactory', 'controlGridFactory'], ['RootGridProps', 'DomainGridProps', 'ControlGridProps'], ['rootShell', 'domainShell', 'controlShell'], ['rootMetadata', 'domainMetadata', 'controlMetadata'], ['RootShellProps', 'DomainShellProps', 'ControlShellProps'], ['rootStage', 'domainStage', 'controlStage'], ['RootStageProps', 'DomainStageProps', 'ControlStageProps'], ['rootContentCard', 'domainContentCard', 'controlContentCard'], ['RootCardProps', 'DomainCardProps', 'ControlCardProps'], ['rootPreviewCard', 'domainPreviewCard', 'controlPreviewCard'], ['RootPreviewProps', 'DomainPreviewProps', 'ControlPreviewProps'], ['rootPermissionTrim', 'domainPermissionTrim', 'controlPermissionTrim'], ['RootPermissionTrimProps', 'DomainPermissionTrimProps', 'ControlPermissionTrimProps'], ['rootSPGroupTrim', 'domainSPGroupTrim', 'controlSPGroupTrim'], ['RootSPGroupTrimProps', 'DomainSPGroupTrimProps', 'ControlSPGroupTrimProps'], ['rootEntraTrim', 'domainEntraTrim', 'controlEntraTrim'], ['RootEntraTrimProps', 'DomainEntraTrimProps', 'ControlEntraTrimProps'], ['rootScopedHook', 'hooksScopedHook', 'legacyScopedHook'], ['rootScopedInfo', 'hooksScopedInfo', 'legacyScopedInfo'], ['rootScopedOptions', 'hooksScopedOptions', 'legacyScopedOptions'], ['RootSharePointPermissionTrimBranchProps', 'DomainSharePointPermissionTrimBranchProps', 'ControlSharePointPermissionTrimBranchProps'], ['RootSharePointGroupSelector', 'DomainSharePointGroupSelector', 'ControlSharePointGroupSelector'], ['RootSharePointGroupTrimBranchProps', 'DomainSharePointGroupTrimBranchProps', 'ControlSharePointGroupTrimBranchProps'], ['RootEntraGroupSelector', 'DomainEntraGroupSelector', 'ControlEntraGroupSelector'], ['RootEntraGroupType', 'DomainEntraGroupType', 'ControlEntraGroupType'], ['RootEntraGroupRelationship', 'DomainEntraGroupRelationship', 'ControlEntraGroupRelationship'], ['RootEntraGroupTrimBranchProps', 'DomainEntraGroupTrimBranchProps', 'ControlEntraGroupTrimBranchProps']]) {
      for (const name of group) {
        assert.ok(symbols.get(name), `Missing canonical imported symbol: ${name}`);
        assert.equal(symbols.get(name), symbols.get(group[0]), `Canonical symbol identity changed: ${name}`);
      }
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
    for (const glob of ['lib/styles/**/*', 'lib/controls/**/*']) if (!manifest.files.includes(glob)) manifest.files.push(glob);
    fs.writeFileSync(file, JSON.stringify(manifest, null, 2) + '\n');
  }
  assertEntrypointContract(manifest, publishedFiles);
  console.log(`Package entrypoints passed: ${Object.keys(manifest.exports).length} exact exports, ${inventory.files.length} historical packed files.`);
}
if (require.main === module) main();
module.exports = { createEntrypointContract, assertEntrypointContract, collectPublishedFiles, verifyConsumerResolution, getLocalizationPackageContract, cleanAliases, main };

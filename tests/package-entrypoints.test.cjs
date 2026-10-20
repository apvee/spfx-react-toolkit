const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const manifest = require('../packages/spfx-react-toolkit/package.json');
const inventory = require('./fixtures/package-entrypoints.json');
const aliases = {
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
  './controls/web-part-shell': 'lib/controls/web-part-shell/index',
  './controls/stage': 'lib/controls/stage/index',
  './controls/pagination': 'lib/controls/pagination/index',
  './controls/description-list': 'lib/controls/description-list/index',
  './controls/content-card': 'lib/controls/content-card/index',
  './controls/preview-card': 'lib/controls/preview-card/index',
  './controls/sharepoint-permission-trim': 'lib/controls/sharepoint-permission-trim/index',
  './controls/sharepoint-group-trim': 'lib/controls/sharepoint-group-trim/index',
  './controls/entra-group-trim': 'lib/controls/entra-group-trim/index',
};
const files = [...inventory.files, ...['lib/controls/file-folder-picker/index', 'lib/property-pane/file-folder-picker/index', 'lib/controls/pickers/index', 'lib/property-pane/pickers/index', 'lib/styles/index', 'lib/controls/index', 'lib/controls/formatted-text/index', 'lib/controls/relative-time/index', 'lib/controls/auto-skeleton/index', 'lib/controls/state-message/index', 'lib/controls/settings-group/index', 'lib/controls/accordion-group/index', 'lib/controls/form/index', 'lib/controls/data-grid/index', 'lib/controls/web-part-shell/index', 'lib/controls/stage/index', 'lib/controls/pagination/index', 'lib/controls/description-list/index', 'lib/controls/content-card/index', 'lib/controls/preview-card/index', 'lib/controls/sharepoint-permission-trim/index', 'lib/controls/sharepoint-group-trim/index', 'lib/controls/entra-group-trim/index'].flatMap(base => ['.js', '.d.ts', '.js.map', '.d.ts.map'].map(ext => base + ext))];
test('clean entrypoints select their canonical declarations before runtime', () => {
  for (const [alias, target] of Object.entries(aliases)) {
    assert.deepEqual(manifest.exports?.[alias], { types: './' + target + '.d.ts', default: './' + target + '.js' });
    assert.deepEqual(Object.keys(manifest.exports[alias]), ['types', 'default']);
  }
  assert.deepEqual(manifest.typesVersions, { '*': {
    core: ['lib/core/index.d.ts'], hooks: ['lib/hooks/index.d.ts'], helpers: ['lib/helpers/index.d.ts'],
    services: ['lib/services/index.d.ts'], styles: ['lib/styles/index.d.ts'],
    'controls/file-folder-picker': ['lib/controls/file-folder-picker/index.d.ts'], 'property-pane/file-folder-picker': ['lib/property-pane/file-folder-picker/index.d.ts'],
    'controls/pickers': ['lib/controls/pickers/index.d.ts'], 'property-pane/pickers': ['lib/property-pane/pickers/index.d.ts'],
    controls: ['lib/controls/index.d.ts'], 'controls/formatted-text': ['lib/controls/formatted-text/index.d.ts'],
    'controls/relative-time': ['lib/controls/relative-time/index.d.ts'],
    'controls/auto-skeleton': ['lib/controls/auto-skeleton/index.d.ts'],
    'controls/state-message': ['lib/controls/state-message/index.d.ts'],
    'controls/settings-group': ['lib/controls/settings-group/index.d.ts'],
    'controls/accordion-group': ['lib/controls/accordion-group/index.d.ts'],
    'controls/form': ['lib/controls/form/index.d.ts'],
    'controls/data-grid': ['lib/controls/data-grid/index.d.ts'],
    'controls/web-part-shell': ['lib/controls/web-part-shell/index.d.ts'],
    'controls/stage': ['lib/controls/stage/index.d.ts'],
    'controls/pagination': ['lib/controls/pagination/index.d.ts'],
    'controls/description-list': ['lib/controls/description-list/index.d.ts'],
    'controls/content-card': ['lib/controls/content-card/index.d.ts'],
    'controls/preview-card': ['lib/controls/preview-card/index.d.ts'],
    'controls/sharepoint-permission-trim': ['lib/controls/sharepoint-permission-trim/index.d.ts'],
    'controls/sharepoint-group-trim': ['lib/controls/sharepoint-group-trim/index.d.ts'],
    'controls/entra-group-trim': ['lib/controls/entra-group-trim/index.d.ts'],
    'styles/useSx': ['lib/hooks/useSx.d.ts'], 'hooks/useStableCallback': ['lib/hooks/useStableCallback.d.ts'],
  } });
  assert.equal(manifest.main, 'lib/index.js');
  assert.equal(manifest.types, 'lib/index.d.ts');
  assert.equal(manifest.type, undefined);
  assert.equal(manifest.module, undefined);
});
test('every frozen packed path remains accessible in its original form', () => {
  for (const file of inventory.files) {
    const actual = manifest.exports?.['./' + file];
    if (file.endsWith('.js')) {
      assert.deepEqual(actual, { types: './' + file.replace(/\.js$/, '.d.ts'), default: './' + file });
      assert.deepEqual(manifest.exports['./' + file.slice(0, -3)], actual);
      if (file.endsWith('/index.js')) assert.deepEqual(manifest.exports['./' + file.slice(0, -9)], actual);
    } else assert.equal(actual, './' + file);
  }
});
test('new clean private and wildcard paths are absent', () => {
  assert.ok(manifest.exports);
  for (const alias of ['./hooks/useAsyncInvoke.internal', './styles/descriptor.internal', './styles/types', './styles/*', './lib/*', './controls/formatted-text/formatted-text-content.internal', './controls/*', './controls/relative-time/relative-time.internal', './helpers/spfx-time-zone.internal', './controls/form/form', './controls/form/create-form', './controls/form/controls/form-input', './controls/form/plugins', './controls/form/advanced', './lib/controls/form', './lib/controls/form/index', './controls/data-grid/data-grid', './controls/data-grid/create-data-grid', './controls/data-grid/column-family.internal', './lib/controls/data-grid/index', './lib/controls/data-grid/data-grid.js', './controls/stage/stage', './controls/stage/stage-context.internal', './lib/controls/stage', './controls/content-slot.internal', './controls/content-slot.types.internal', './controls/cards/card-frame.internal', './controls/content-card/content-card', './controls/preview-card/preview-card', './lib/controls/cards', './lib/controls/content-card', './lib/controls/preview-card', './hooks/useAccessRequest.internal', './services/spfx-scoped-permissions.internal', './services/spfx-sharepoint-groups.internal', './services/spfx-entra-groups.internal', './controls/access-trim-children.internal', './controls/sharepoint-permission-trim/sharepoint-permission-trim', './controls/sharepoint-group-trim/sharepoint-group-trim', './controls/entra-group-trim/entra-group-trim', './lib/controls/sharepoint-permission-trim', './lib/controls/sharepoint-group-trim', './lib/controls/entra-group-trim', './lib/services/spfx-permission-target.types']) {
    assert.equal(manifest.exports[alias], undefined);
  }
});
test('contract generation is deterministic and rejects loss outside the declaration baseline', () => {
  const { createEntrypointContract } = require('../scripts/package-entrypoints.cjs');
  const forward = createEntrypointContract({ publishedFiles: files, cleanAliases: aliases });
  assert.deepEqual(forward, createEntrypointContract({ publishedFiles: [...files].reverse(), cleanAliases: aliases }));
  assert.deepEqual(forward.exports['./lib/helpers/styles/width'], {
    types: './lib/helpers/styles/width.d.ts', default: './lib/helpers/styles/width.js',
  });
  assert.throws(() => createEntrypointContract({ publishedFiles: files.filter(file => file !== 'lib/helpers/styles/width.js.map'), cleanAliases: aliases }), /Missing historical published file/);
  assert.throws(() => createEntrypointContract({ publishedFiles: files, cleanAliases: { ...aliases, './styles/types': 'lib/helpers/styles/types' } }), /clean alias/);
  assert.throws(() => createEntrypointContract({ publishedFiles: files, cleanAliases: { ...aliases, './styles': 'lib/helpers/styles/index' } }), /clean alias/);
});
test('package guard rejects catch-all root rewriting, wrong condition order and missing legacy paths', () => {
  const { assertCompatiblePackageContract } = require('../scripts/api-compatibility.helpers.cjs');
  const original = require('./fixtures/package-baseline.json');
  for (const mutate of [
    candidate => { candidate.typesVersions['*']['*'] = ['lib/*']; },
    candidate => { candidate.exports['./styles'] = { default: './lib/styles/index.js', types: './lib/styles/index.d.ts' }; },
    candidate => { delete candidate.exports['./lib/helpers/styles/width.js.map']; },
    candidate => { candidate.exports['./styles/types'] = './lib/helpers/styles/types.js'; },
    candidate => { candidate.type = 'module'; },
  ]) {
    const candidate = structuredClone(manifest);
    mutate(candidate);
    assert.throws(() => assertCompatiblePackageContract(candidate, original));
  }
});
test('explicit check verifies built targets and never writes the manifest', () => {
  const file = path.resolve(__dirname, '../packages/spfx-react-toolkit/package.json');
  const before = fs.readFileSync(file, 'utf8');
  const result = spawnSync(process.execPath, ['scripts/package-entrypoints.cjs', '--check'], { cwd: path.resolve(__dirname, '..'), encoding: 'utf8' });
  assert.equal(result.status, 0, result.stdout + result.stderr);
  assert.equal(fs.readFileSync(file, 'utf8'), before);
});
test('real TypeScript and Webpack resolvers honor clean and legacy entries', async () => {
  const { verifyConsumerResolution } = require('../scripts/package-entrypoints.cjs');
  const temporaryRoot = path.resolve(__dirname, '../temp');
  fs.mkdirSync(temporaryRoot, { recursive: true });
  const consumer = fs.mkdtempSync(path.join(temporaryRoot, 'entrypoint-test-'));
  try {
    fs.writeFileSync(path.join(consumer, 'package.json'), '{"private":true}');
    const installed = path.join(consumer, 'node_modules/@apvee/spfx-react-toolkit');
    fs.mkdirSync(installed, { recursive: true });
    fs.copyFileSync(path.resolve(__dirname, '../packages/spfx-react-toolkit/package.json'), path.join(installed, 'package.json'));
    fs.cpSync(path.resolve(__dirname, '../packages/spfx-react-toolkit/lib'), path.join(installed, 'lib'), { recursive: true });
    for (const file of ['README.md', 'LICENSE']) fs.copyFileSync(path.resolve(__dirname, '../packages/spfx-react-toolkit', file), path.join(installed, file));
    const evidence = await verifyConsumerResolution(consumer);
    assert.deepEqual(evidence.modes, ['node', 'bundler', 'node16']);
    assert.ok(evidence.runtime.length >= 15);
  } finally { fs.rmSync(consumer, { recursive: true, force: true }); }
});

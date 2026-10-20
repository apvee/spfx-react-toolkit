const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { resolveModule, collectPublicSurface, assertDocumentedSurface } = require('./public-surface.helpers.cjs');

const root = path.resolve(__dirname, '..');

function read(relativePath) {
  return fs.readFileSync(path.join(root, relativePath), 'utf8');
}

function resolveBarrelExports(barrelPath) {
  const source = read(barrelPath);
  const modules = [];
  const exportRegex = /export\s+\*\s+from\s+['"](.+)['"];?/g;
  let match;

  while ((match = exportRegex.exec(source)) !== null) {
    modules.push(path.relative(root, resolveModule(path.join(root, barrelPath), match[1])));
  }

  return modules;
}

function collectExportedNames(filePath, pattern) {
  const source = read(filePath);
  const names = [];
  let match;

  while ((match = pattern.exec(source)) !== null) {
    names.push(match[1]);
  }

  return names;
}

function collectPublicDocs() {
  const docs = [];
  const docsRoot = path.join(root, 'docs');

  function walk(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const fullPath = path.join(dir, entry.name);
      const relativePath = path.relative(root, fullPath);

      if (entry.isDirectory()) {
        walk(fullPath);
      } else if (entry.isFile() && entry.name.endsWith('.md')) {
        docs.push(relativePath);
      }
    }
  }

  walk(docsRoot);
  return docs;
}

function assertContainsAll(documentPath, names, label) {
  const document = read(documentPath);
  const missing = names.filter(name => !document.includes(name));
  assert.deepStrictEqual(missing, [], `${documentPath} missing ${label}: ${missing.join(', ')}`);
}

const helperModules = resolveBarrelExports('packages/spfx-react-toolkit/src/helpers/index.ts');
const helperFunctions = helperModules.flatMap(modulePath => (
  collectExportedNames(modulePath, /export\s+function\s+([A-Za-z0-9_]+)/g)
));

const serviceModules = resolveBarrelExports('packages/spfx-react-toolkit/src/services/index.ts');
const serviceFactories = serviceModules.flatMap(modulePath => (
  collectExportedNames(modulePath, /export\s+function\s+([A-Za-z0-9_]+)/g)
));
const serviceInterfaces = serviceModules.flatMap(modulePath => (
  collectExportedNames(modulePath, /export\s+interface\s+([A-Za-z0-9_]+)/g)
));

assertContainsAll('docs/api/helpers/INDEX.md', helperFunctions, 'helper functions');
assertContainsAll('docs/api/services/INDEX.md', serviceFactories, 'service factories');
assertContainsAll('docs/api/services/INDEX.md', serviceInterfaces, 'service interfaces');

const styleEntry = path.join(root, 'packages/spfx-react-toolkit/src/helpers/styles/index.ts');
const styleSurface = collectPublicSurface(styleEntry);
const helperSurface = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/helpers/index.ts'));
assert.deepStrictEqual(
  helperSurface.filter(item => item.filePath.startsWith(path.dirname(styleEntry) + path.sep)).map(item => item.name).sort(),
  styleSurface.map(item => item.name).sort(),
  'Style exports must remain reachable through the public helper barrel'
);
assertDocumentedSurface(read('docs/api/helpers/styles.md'), styleSurface);
const sxHook = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/hooks/useSx.ts'));
assertDocumentedSurface(read('docs/api/hooks/react.md'), sxHook);
assertContainsAll('docs/api/hooks/INDEX.md', ['useSx', 'useStableCallback'], 'React utility hooks');

const controlSurface = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/controls/index.ts'));
assertDocumentedSurface(read('docs/api/controls/formatted-text.md') + '\n' + read('docs/api/controls/relative-time.md') + '\n' + read('docs/api/controls/auto-skeleton.md') + '\n' + read('docs/api/controls/state-message.md') + '\n' + read('docs/api/controls/settings-group.md') + '\n' + read('docs/api/controls/accordion-group.md') + '\n' + read('docs/api/controls/form.md') + '\n' + read('docs/api/controls/data-grid.md') + '\n' + read('docs/api/controls/web-part-shell.md') + '\n' + read('docs/api/controls/stage.md') + '\n' + read('docs/api/controls/pagination.md') + '\n' + read('docs/api/controls/description-list.md') + '\n' + read('docs/api/controls/content-card.md') + '\n' + read('docs/api/controls/preview-card.md') + '\n' + ['sharepoint-permission-trim', 'sharepoint-group-trim', 'entra-group-trim', 'pickers', 'file-folder-picker'].map(name => read(`docs/api/controls/${name}.md`)).join('\n'), controlSurface);
const formSurface = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/controls/form/index.ts'), { includeTypeMembers: true });
assertDocumentedSurface(read('docs/api/controls/form.md'), formSurface);
const gridSurface = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/controls/data-grid/index.ts'), { includeTypeMembers: true });
assertDocumentedSurface(read('docs/api/controls/data-grid.md'), gridSurface);
const shellSurface = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/controls/web-part-shell/index.ts'), { includeTypeMembers: true });
assertDocumentedSurface(read('docs/api/controls/web-part-shell.md'), shellSurface);
const descriptionListSurface = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/controls/description-list/index.ts'), { includeTypeMembers: true });
assertDocumentedSurface(read('docs/api/controls/description-list.md'), descriptionListSurface);
assertContainsAll('docs/api/controls/description-list.md', [
  'DescriptionListProps.items', 'DescriptionListProps.children', 'DescriptionListProps.layout',
  'DescriptionListEntry.id', 'DescriptionListEntry.label', 'DescriptionListEntry.value',
  'DescriptionListItemProps.label', 'DescriptionListItemProps.children',
], 'DescriptionList content members');
for (const control of ['content-card', 'preview-card']) {
  const surface = collectPublicSurface(path.join(root, `packages/spfx-react-toolkit/src/controls/${control}/index.ts`), { includeTypeMembers: true });
  assertDocumentedSurface(read(`docs/api/controls/${control}.md`), surface);
}
// PreviewCard extends the toolkit content props; the collector checks its own members only.
assertContainsAll('docs/api/controls/preview-card.md', [
  'ContentCardProps', 'title', 'headingLevel', 'description', 'icon', 'headerActions', 'children', 'footer', 'actions',
], 'inherited PreviewCard content contract');
const paginationSurface = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/controls/pagination/index.ts'), { includeTypeMembers: true });
assertDocumentedSurface(read('docs/api/controls/pagination.md'), paginationSurface);
const stageSurface = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/controls/stage/index.ts'), { includeTypeMembers: true });
assertDocumentedSurface(read('docs/api/controls/stage.md'), stageSurface);
// The collector does not follow StageProps' referenced common/state types.
// Require those qualified members explicitly without broadening the shared collector.
assertContainsAll('docs/api/controls/stage.md', [
  'StageProps.children', 'StageProps.open', 'StageProps.defaultOpen', 'StageProps.onOpenChange',
  'StageProps.dismissMode', 'StageProps.unmountOnClose', 'StageProps.mountNode',
], 'referenced StageProps common/state members');
assertDocumentedSurface(read('docs/api/hooks/environment.md'), collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/hooks/useSPFxLocaleInfo.ts')));
assertContainsAll('docs/api/controls/INDEX.md', controlSurface.map(item => item.name), 'control exports');

const hooksIndex = read('docs/api/hooks/INDEX.md');
for (const staleHook of ['useSPFxPropertyPane', 'useSPFxPnPSP', 'useSPFxPnPGraph']) {
  assert.ok(!hooksIndex.includes(staleHook), `docs/api/hooks/INDEX.md contains stale hook ${staleHook}`);
}

const publicDocLinks = [
  './api/helpers/INDEX.md',
  './api/services/INDEX.md',
  './api/controls/INDEX.md',
];

for (const link of publicDocLinks) {
  assert.ok(read('docs/INDEX.md').includes(link), `docs/INDEX.md missing ${link}`);
}

const forbiddenPublicDocPatterns = [
  /\bJotai\b/i,
  /\batom\b/i,
  /\batoms\b/i,
];

const offenders = [];
for (const docPath of collectPublicDocs()) {
  const text = read(docPath);
  for (const pattern of forbiddenPublicDocPatterns) {
    if (pattern.test(text)) {
      offenders.push(`${docPath} contains ${pattern}`);
    }
  }
}

assert.deepStrictEqual(offenders, []);



// Validate repository-relative links in public docs as well as the export inventory.
const brokenLinks = [];
for (const documentPath of ['README.md', 'packages/spfx-react-toolkit/README.md', ...collectPublicDocs()]) {
  const source = read(documentPath);
  for (const match of source.matchAll(/\]\(([^)]+)\)/g)) {
    const target = match[1].split(/\s+["']/)[0];
    if (/^(?:[a-z][a-z0-9+.-]*:|#|\/)/i.test(target)) continue;
    const relative = decodeURIComponent(target.split('#')[0]);
    if (relative && !fs.existsSync(path.resolve(root, path.dirname(documentPath), relative))) {
      brokenLinks.push(`${documentPath}: ${target}`);
    }
  }
}
assert.deepStrictEqual(brokenLinks, [], `Broken public documentation links:\n${brokenLinks.join('\n')}`);
console.log(`public docs verification passed (${styleSurface.length} qualified style exports with source JSDoc, historical helper/service inventory, WebPartShell surface/JSDoc and relative links)`);

const facadeSurface = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/styles/index.ts'));
assert.deepStrictEqual(facadeSurface.map(item => item.name).sort(),
  [...styleSurface, ...sxHook].map(item => item.name).sort(),
  'Clean styles facade must expose the canonical descriptor surface and useSx');
assertContainsAll('docs/PACKAGE-IMPORTS.md', [
  '/core', '/hooks', '/helpers', '/services', '/styles', '/styles/useSx', '/hooks/useStableCallback',
  '/lib/', 'typesVersions', 'moduleResolution', 'SxFunction', 'SxOptions', 'SxInput',
  'CommonJS', 'sideEffects', '1,023', '512', 'PnP',
], 'facade imports and compatibility constraints');
assert.ok(read('docs/INDEX.md').includes('./PACKAGE-IMPORTS.md'), 'Package import guide must be indexed');

for (const family of ['sharepoint-permission-trim', 'sharepoint-group-trim', 'entra-group-trim', 'pickers', 'file-folder-picker']) {
  assertDocumentedSurface(read(`docs/api/controls/${family}.md`), collectPublicSurface(path.join(root, `packages/spfx-react-toolkit/src/controls/${family}/index.ts`), { includeTypeMembers: true }));
}
assertContainsAll('docs/api/controls/sharepoint-permission-trim.md', ['SPFxScopedPermissionsOptions', ...['scope', 'webUrl', 'listId', 'listTitle', 'listPath', 'itemId', 'folderPath'].map(member => `SharePointPermissionTrimProps.${member}`)], 'inherited scoped permission contract');
assertDocumentedSurface(read('docs/api/hooks/permissions.md'), collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/hooks/useSPFxCrossSitePermissions.ts'), { includeTypeMembers: true }));
console.log('Verified all three trim family surfaces and scoped permission overload.');

assertContainsAll('docs/api/hooks/permissions.md', ['scope', 'webUrl', 'listId', 'listTitle', 'listPath', 'itemId', 'folderPath'].map(member => `SPFxScopedPermissionsOptions.${member}`), 'scoped hook inherited selectors');

const nativePickerSurface = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/property-pane/pickers/index.ts'));
assertDocumentedSurface(read('docs/api/property-pane/pickers.md') + '\n' + read('docs/api/controls/pickers.md'), nativePickerSurface);

const nativeFileFolderSurface = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/property-pane/file-folder-picker/index.ts'));
assertDocumentedSurface(read('docs/api/property-pane/file-folder-picker.md') + '\n' + read('docs/api/controls/file-folder-picker.md'), nativeFileFolderSurface);

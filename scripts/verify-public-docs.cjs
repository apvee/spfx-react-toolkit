const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { resolveModule, collectPublicSurface, assertDocumentedSurface } = require('./public-surface.helpers.cjs');
const { generatedArtifacts, checkArtifacts, findBrokenLinks } = require('./documentation.cjs');

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
  const document = Array.isArray(documentPath) ? documentPath.map(read).join('\n') : read(documentPath);
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

const helperDocs = collectPublicDocs().filter(file => file.startsWith('docs/api/helpers/'));
const serviceDocs = collectPublicDocs().filter(file => file.startsWith('docs/api/services/'));
assertContainsAll(helperDocs, helperFunctions, 'helper functions');
assertContainsAll(serviceDocs, serviceFactories, 'service factories');
assertContainsAll(serviceDocs, serviceInterfaces, 'service interfaces');
const precheckHelperSurface = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/helpers/spfx-api-permission-precheck.helpers.ts'));
assertContainsAll('docs/api/helpers/api-permission-precheck.md', precheckHelperSurface.map(item => item.name), 'precheck helper functions, types and constants');

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
assertContainsAll('docs/api/INDEX.md', ['useSx', 'useStableCallback'], 'React utility hooks');

const apiIndex = read('docs/api/INDEX.md');
for (const staleHook of ['useSPFxPropertyPane', 'useSPFxPnPSP', 'useSPFxPnPGraph']) {
  assert.ok(!apiIndex.includes(staleHook), `docs/api/INDEX.md contains stale hook ${staleHook}`);
}

const publicDocLinks = [
  'api/INDEX.md#helpers-and-styles',
  'api/INDEX.md#non-react-services',
  'api/INDEX.md',
  'getting-started/installation.md',
  'getting-started/quick-start.md',
  'development/INDEX.md',
  'development/sharepoint-validation/INDEX.md',
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



// Validate current local paths and Markdown fragments, including agent guidance.
const brokenLinks = findBrokenLinks(root, ['AGENTS.md', 'README.md', 'packages/spfx-react-toolkit/README.md', ...collectPublicDocs()]);
assert.deepStrictEqual(brokenLinks, [], `Broken public documentation links:\n${brokenLinks.join('\n')}`);
assert.deepStrictEqual(checkArtifacts(root, generatedArtifacts(root)), [], 'Regenerate the API catalog and npm README with node scripts/documentation.cjs');

const facadeSurface = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/styles/index.ts'));
assert.deepStrictEqual(facadeSurface.map(item => item.name).sort(),
  [...styleSurface, ...sxHook].map(item => item.name).sort(),
  'Clean styles facade must expose the canonical descriptor surface and useSx');
assertContainsAll('docs/guides/package-imports.md', [
  '/core', '/hooks', '/helpers', '/services', '/styles', '/styles/useSx', '/hooks/useStableCallback',
  '/lib/', 'typesVersions', 'moduleResolution', 'SxFunction', 'SxOptions', 'SxInput',
  'CommonJS', 'sideEffects', '1,023', '512', 'PnP',
], 'facade imports and compatibility constraints');
assert.ok(read('docs/INDEX.md').includes('guides/package-imports.md'), 'Package import guide must be indexed');

console.log(`public docs verification passed (${styleSurface.length} qualified style exports/JSDoc, helper/service coverage, generated catalog/README and current local links/fragments)`);

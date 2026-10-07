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

const hooksIndex = read('docs/api/hooks/INDEX.md');
for (const staleHook of ['useSPFxPropertyPane', 'useSPFxPnPSP', 'useSPFxPnPGraph']) {
  assert.ok(!hooksIndex.includes(staleHook), `docs/api/hooks/INDEX.md contains stale hook ${staleHook}`);
}

const publicDocLinks = [
  './api/helpers/INDEX.md',
  './api/services/INDEX.md',
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
console.log(`public docs verification passed (${styleSurface.length} qualified style exports with source JSDoc, historical helper/service inventory and relative links)`);

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

const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');

function read(relativePath) {
  return fs.readFileSync(path.join(root, relativePath), 'utf8');
}

function resolveBarrelExports(barrelPath) {
  const source = read(barrelPath);
  const barrelDir = path.dirname(barrelPath);
  const modules = [];
  const exportRegex = /export\s+\*\s+from\s+['"](.+)['"];?/g;
  let match;

  while ((match = exportRegex.exec(source)) !== null) {
    modules.push(path.join(barrelDir, `${match[1]}.ts`));
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
        if (relativePath.startsWith(`docs${path.sep}superpowers`) || relativePath.startsWith(`docs${path.sep}maintenance`)) {
          continue;
        }
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
console.log('public docs verification passed (API inventory and relative links)');

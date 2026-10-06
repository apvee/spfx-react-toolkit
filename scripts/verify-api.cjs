const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { assertCompatibleDeclaration } = require('./api-compatibility.helpers.cjs');
const root = path.resolve(__dirname, '..');
const baseline = require('../tests/fixtures/api-baseline.json');
const requireApprovedAdditions = fs.existsSync(path.join(root, 'packages/spfx-react-toolkit/src/hooks/useSPFxSiteKeyValueStore.ts'));
for (const [name, snapshot] of Object.entries(baseline)) {
  const current = fs.readFileSync(path.join(root, 'packages/spfx-react-toolkit/lib', name), 'utf8');
  assertCompatibleDeclaration(name, current, snapshot.declaration, { requireApprovedAdditions });
}
const manifest = require('../packages/spfx-react-toolkit/package.json');
const originalPackage = require('../tests/fixtures/package-baseline.json');
// The user explicitly selected shared Fluent peers. Preserve all historical
// peers and use exactly the original Fluent version ranges for this change.
const sharedFluentPeers = {
  '@fluentui/react-migration-v8-v9': originalPackage.dependencies['@fluentui/react-migration-v8-v9'],
  '@fluentui/react-theme': originalPackage.dependencies['@fluentui/react-theme'],
};
assert.deepEqual(sharedFluentPeers, {
  '@fluentui/react-migration-v8-v9': '^9.9.12',
  '@fluentui/react-theme': '^9.2.0',
});
assert.deepEqual(manifest.peerDependencies, { ...originalPackage.peerDependencies, ...sharedFluentPeers }, 'Runtime peer contracts changed outside the approved shared Fluent peers');
// Approved dependency cleanup: this ES2020 library emits no tslib helpers.
// Keep the historical baseline unchanged; allow only this explicit removal.
const expectedRuntimeDependencies = { ...originalPackage.dependencies };
assert.equal(expectedRuntimeDependencies.tslib, '2.3.1', 'Unexpected historical tslib contract');
delete expectedRuntimeDependencies.tslib;
for (const name of Object.keys(sharedFluentPeers)) delete expectedRuntimeDependencies[name];
assert.deepEqual(manifest.dependencies ?? {}, expectedRuntimeDependencies, 'Runtime dependencies changed outside the approved tslib cleanup and shared Fluent peers');

// A dependency owned by the app/root or present transitively is not sufficient
// for standalone consumers of the published library, including its types.
const declaredPackages = new Set([
  ...Object.keys(manifest.dependencies ?? {}),
  ...Object.keys(manifest.peerDependencies),
]);
function verifyImport(specifier, file) {
  if (specifier.startsWith('.')) return;
  const packageName = specifier.startsWith('@')
    ? specifier.split('/').slice(0, 2).join('/')
    : specifier.split('/')[0];
  assert.notEqual(packageName, 'tslib', `Generated tslib import requires a direct runtime dependency: ${file}`);
  assert.ok(declaredPackages.has(packageName), `Undeclared package ${packageName} imported by ${file}`);
}
function verifyPublishedImports(directory) {
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    const file = path.join(directory, entry.name);
    if (entry.isDirectory()) { verifyPublishedImports(file); continue; }
    if (!entry.name.endsWith('.js') && !entry.name.endsWith('.d.ts')) continue;
    const source = ts.createSourceFile(file, fs.readFileSync(file, 'utf8'), ts.ScriptTarget.Latest, true);
    function visit(node) {
      if ((ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) && node.moduleSpecifier && ts.isStringLiteral(node.moduleSpecifier)) {
        verifyImport(node.moduleSpecifier.text, file);
      }
      if (ts.isCallExpression(node) && node.arguments.length > 0 && ts.isStringLiteral(node.arguments[0]) &&
          (node.expression.kind === ts.SyntaxKind.ImportKeyword || (ts.isIdentifier(node.expression) && node.expression.text === 'require'))) {
        verifyImport(node.arguments[0].text, file);
      }
      ts.forEachChild(node, visit);
    }
    visit(source);
  }
}
verifyPublishedImports(path.join(root, 'packages/spfx-react-toolkit/lib'));
assert.equal(manifest.private, false);
assert.deepEqual(manifest.files, originalPackage.files, 'Published path allowlist changed');
assert.equal(manifest.name, '@apvee/spfx-react-toolkit');
assert.equal(manifest.version, '2.1.0');
assert.equal(manifest.main, 'lib/index.js');
assert.equal(manifest.types, 'lib/index.d.ts');
assert.equal(manifest.exports, undefined, 'Do not restrict pre-existing deep imports');
console.log(`API compatibility passed: ${Object.keys(baseline).length} declaration modules, name/version/entry paths preserved.`);

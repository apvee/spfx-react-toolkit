const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { assertCompatibleDeclaration, assertCompatiblePackageContract } = require('./api-compatibility.helpers.cjs');
const { assertEntrypointContract, collectPublishedFiles } = require('./package-entrypoints.cjs');
const root = path.resolve(__dirname, '..');
const baseline = require('../tests/fixtures/api-baseline.json');
const requireApprovedAdditions = fs.existsSync(path.join(root, 'packages/spfx-react-toolkit/src/hooks/useSPFxSiteKeyValueStore.ts'));
const requirePnPListAdditions = fs.existsSync(path.join(root, 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPListById.ts'));
const requireSxAdditions = fs.existsSync(path.join(root, 'packages/spfx-react-toolkit/src/hooks/useSx.ts'));
for (const [name, snapshot] of Object.entries(baseline)) {
  const current = fs.readFileSync(path.join(root, 'packages/spfx-react-toolkit/lib', name), 'utf8');
  assertCompatibleDeclaration(name, current, snapshot.declaration, {
    requireApprovedAdditions, requirePnPListAdditions, requireSxAdditions,
    requirePnPListWebsRegistration: true,
  });
}
const manifest = require('../packages/spfx-react-toolkit/package.json');
const originalPackage = require('../tests/fixtures/package-baseline.json');
assertCompatiblePackageContract(manifest, originalPackage);
assertEntrypointContract(manifest, collectPublishedFiles());

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
      if (ts.isImportTypeNode(node) && ts.isLiteralTypeNode(node.argument) && ts.isStringLiteral(node.argument.literal)) {
        verifyImport(node.argument.literal.text, file);
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
console.log(`API compatibility passed: ${Object.keys(baseline).length} declaration modules, name/version/entry paths preserved.`);

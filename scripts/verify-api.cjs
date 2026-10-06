const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const root = path.resolve(__dirname, '..');
const baseline = require('../docs/maintenance/evidence/api-baseline.json');
const printer = ts.createPrinter({removeComments:true});
const canonical = text => printer.printFile(ts.createSourceFile('module.d.ts', text, ts.ScriptTarget.Latest, true));
for (const [name, snapshot] of Object.entries(baseline)) {
  const current = fs.readFileSync(path.join(root, 'packages/spfx-react-toolkit/lib', name), 'utf8');
  assert.equal(canonical(current), canonical(snapshot.declaration), `Declaration contract changed: ${name}`);
}
const manifest = require('../packages/spfx-react-toolkit/package.json');
const originalPackage = require('../docs/maintenance/evidence/package-baseline.json');
assert.deepEqual(manifest.peerDependencies, originalPackage.peerDependencies, 'Runtime peer contracts changed');
assert.deepEqual(manifest.dependencies, originalPackage.dependencies, 'Runtime dependencies changed');
assert.equal(manifest.private, false);
assert.deepEqual(manifest.files, originalPackage.files, 'Published path allowlist changed');
assert.equal(manifest.name, '@apvee/spfx-react-toolkit');
assert.equal(manifest.version, '2.1.0');
assert.equal(manifest.main, 'lib/index.js');
assert.equal(manifest.types, 'lib/index.d.ts');
assert.equal(manifest.exports, undefined, 'Do not restrict pre-existing deep imports');
console.log(`API compatibility passed: ${Object.keys(baseline).length} declaration modules, name/version/entry paths preserved.`);

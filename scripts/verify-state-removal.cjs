const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const sourceFiles = [];

function walk(dir) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      walk(fullPath);
    } else if (/\.(ts|tsx|js|json|md)$/.test(entry.name)) {
      sourceFiles.push(fullPath);
    }
  }
}

for (const dir of ['src', 'lib', 'docs']) {
  const absoluteDir = path.join(root, dir);
  if (fs.existsSync(absoluteDir)) {
    walk(absoluteDir);
  }
}

for (const file of [
  'package.json',
  'package-lock.json',
  'README.md'
]) {
  sourceFiles.push(path.join(root, file));
}

const filteredSourceFiles = sourceFiles.filter(file => {
  const relative = path.relative(root, file);
  return !relative.startsWith(`docs${path.sep}superpowers${path.sep}`);
});

const forbiddenPatterns = [
  /\bjotai\b/i,
  /\bjotay\b/i,
  /\batomWithStorage\b/i,
  /\buseAtomValue\b/i,
  /\buseSetAtom\b/i,
  /\buseAtom\(/i,
  /\bspfxAtoms\b/i,
  /\batoms\.internal\b/i
];

const offenders = [];
for (const file of filteredSourceFiles) {
  const text = fs.readFileSync(file, 'utf8');
  for (const pattern of forbiddenPatterns) {
    if (pattern.test(text)) {
      offenders.push(`${path.relative(root, file)} contains ${pattern}`);
    }
  }
}

assert.deepStrictEqual(offenders, []);

const packageJson = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
assert.ok(packageJson.files.includes('lib/services/**/*'), 'package files must include lib/services/**/*');
assert.ok(packageJson.files.includes('lib/helpers/**/*'), 'package files must include lib/helpers/**/*');

console.log('state removal verification passed');

const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
const root = fs.realpathSync(path.resolve(__dirname, '../../../..'));
const { createDependencyAdmission } = require(path.join(root, 'scripts/documentation-inputs.cjs'));
const dependencies = createDependencyAdmission(root);
const within = file => { const relative = path.relative(root, file); return relative !== '..' && !relative.startsWith('..' + path.sep) && !path.isAbsolute(relative); };
const reads = new Set();
function check(file) {
  if (typeof file !== 'string' && !Buffer.isBuffer(file) && !(file instanceof URL)) return;
  const absolute = path.resolve(file instanceof URL ? require('node:url').fileURLToPath(file) : String(file));
  let real;
  try { real = fs.realpathSync(absolute); } catch (error) { if (error.code === 'ENOENT' || error.code === 'ENOTDIR') return; throw error; }
  if (!(within(absolute) && within(real) && !absolute.split(path.sep).includes('node_modules')) && !dependencies.admitted(real)) throw Object.assign(new Error(`DOC_FORBIDDEN_INPUT: out-of-perimeter read ${absolute}`), { code: 'DOC_FORBIDDEN_INPUT' });
  reads.add(real);
}
for (const method of ['readFileSync', 'readFile']) {
  const original = fs[method];
  fs[method] = function (file, ...args) { check(file); return original.call(this, file, ...args); };
}
const originalOpen = fs.openSync;
fs.openSync = function (file, flags, ...args) { if (typeof flags === 'string' && flags.startsWith('r')) check(file); return originalOpen.call(this, file, flags, ...args); };
const originalLoad = Module._load;
Module._load = function (name, parent, main) {
  const resolved = Module._resolveFilename(name, parent, main);
  if (path.isAbsolute(resolved)) check(resolved);
  return originalLoad.apply(this, arguments);
};
for (const filename of Object.keys(require.cache)) check(filename);
globalThis.__documentationReadAudit = () => ({ reads: reads.size, files: [...reads].sort(), packages: dependencies.packages.length });

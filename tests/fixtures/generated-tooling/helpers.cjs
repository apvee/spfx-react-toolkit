const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const crypto = require('node:crypto');
const { spawnSync } = require('node:child_process');
const repository = path.resolve(__dirname, '../../..');
const config = require('../../../scripts/documentation.config.cjs');
const { executableRoot, ownership } = require('../../../scripts/executable-examples-render.cjs');
function mirror(t, executable = true) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'generated-tooling-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  fs.cpSync(path.join(repository, 'tests/fixtures/documentation/artifact-set'), root, { recursive: true });
  fs.cpSync(path.join(repository, 'scripts'), path.join(root, 'scripts'), { recursive: true });
  fs.copyFileSync(path.join(repository, 'package-lock.json'), path.join(root, 'package-lock.json'));
  fs.copyFileSync(path.join(repository, config.packageRoot, 'tsdoc.json'), path.join(root, config.packageRoot, 'tsdoc.json'));
  fs.symlinkSync(path.join(repository, 'node_modules'), path.join(root, 'node_modules'), 'dir');
  if (executable) {
    const doc = fs.readFileSync(path.join(repository, 'tests/fixtures/executable-examples/cases/artifact-count.jsdoc'), 'utf8');
    fs.writeFileSync(path.join(root, config.sourceRoot, 'index.ts'), doc + 'export function count(value: number): number { return value; }\n');
    const filename = path.join(root, config.tsconfig);
    const project = JSON.parse(fs.readFileSync(filename, 'utf8'));
    project.compilerOptions.jsx = 'react';
    fs.writeFileSync(filename, JSON.stringify(project));
    // The editorial selector must still identify a unique source example.
    const guide = path.join(root, 'docs/_sources/guides/count.md');
    fs.writeFileSync(guide, fs.readFileSync(guide, 'utf8').replace('section="example"', 'section="example" title="Z first scenario"'));
  }
  return root;
}
function snapshot(root) {
  const result = {};
  function visit(file, route) {
    const stat = fs.lstatSync(file);
    result[route] = { mtime: stat.mtimeMs, kind: stat.isDirectory() ? 'directory' : stat.isSymbolicLink() ? 'symlink' : 'file' };
    if (stat.isDirectory()) for (const name of fs.readdirSync(file).sort()) visit(path.join(file, name), route + '/' + name);
    else if (stat.isFile()) result[route].hash = crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
  }
  visit(root, '.');
  return result;
}
function run(root, args = [], env = {}, script = 'generated.cjs') {
  return spawnSync(process.execPath, [path.join(root, 'scripts', script), ...args], { cwd: os.tmpdir(), env: { ...process.env, ...env }, encoding: 'utf8', timeout: 120000 });
}
module.exports = { repository, config, executableRoot, ownership, mirror, snapshot, run };

const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const crypto = require('node:crypto');
const { spawnSync } = require('node:child_process');
const repository = path.resolve(__dirname, '../../..');
const config = require('../../../scripts/docs/documentation.config.cjs');
const { executableRoot, ownership } = require('../../../scripts/docs/executable-examples-render.cjs');
// Build mirrors need the engine plus explicit lifecycle scripts. Other linked
// tooling is copied one file at a time as documentation-validation data.
function copyEngineInputs(root, buildTools = []) {
  fs.cpSync(path.join(repository, 'scripts/docs'), path.join(root, 'scripts/docs'), { recursive: true });
  for (const route of buildTools) {
    fs.mkdirSync(path.dirname(path.join(root, route)), { recursive: true });
    fs.copyFileSync(path.join(repository, route), path.join(root, route));
  }
  const { markdownLinks, externalTarget } = require('../../../scripts/docs/documentation-markdown.cjs');
  const excluded = route => route === executableRoot || route.startsWith(executableRoot + '/') ||
    /(?:^|\/)(?:\.git|\.docs|node_modules|lib)(?:\/|$)/.test(route) || route === 'llms.txt' ||
    route === 'docs/INDEX.md' || /^docs\/(?:api|guides|getting-started|development)(?:\/|$)/.test(route) ||
    route === path.posix.join(config.packageRoot, 'README.md');
  const pending = ['README.md'];
  const seen = new Set();
  function editorial(directory) {
    if (!fs.existsSync(directory)) return;
    for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
      const file = path.join(directory, entry.name);
      if (entry.isDirectory()) editorial(file);
      else if (entry.isFile() && file.endsWith('.md')) pending.push(path.relative(root, file).split(path.sep).join('/'));
    }
  }
  editorial(path.join(root, 'docs/_sources'));
  while (pending.length) {
    const route = pending.shift();
    if (seen.has(route) || !fs.existsSync(path.join(root, route))) continue;
    seen.add(route);
    for (const href of markdownLinks(fs.readFileSync(path.join(root, route), 'utf8'))) {
      if (externalTarget(href)) continue;
      const target = decodeURIComponent(href.split('#')[0].split('?')[0]);
      if (!target) continue;
      const local = path.posix.normalize(target.startsWith('/') ? target.slice(1) : path.posix.join(path.posix.dirname(route), target));
      if (local.startsWith('../')) throw new Error('Mirror link escapes root: ' + href);
      if (excluded(local)) continue;
      const owner = path.join(repository, local), destination = path.join(root, local);
      if (!fs.existsSync(destination) && fs.existsSync(owner)) {
        if (fs.statSync(owner).isDirectory()) fs.mkdirSync(destination, { recursive: true });
        else {
          fs.mkdirSync(path.dirname(destination), { recursive: true });
          fs.copyFileSync(owner, destination);
        }
      }
      if (local.endsWith('.md')) pending.push(local);
    }
  }
}
function mirror(t, executable = true) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'generated-tooling-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  fs.cpSync(path.join(repository, 'tests/fixtures/documentation/artifact-set'), root, { recursive: true });
  copyEngineInputs(root);
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
  return spawnSync(process.execPath, [path.join(root, 'scripts/docs', script), ...args], { cwd: os.tmpdir(), env: { ...process.env, ...env }, encoding: 'utf8', timeout: 120000 });
}
module.exports = { copyEngineInputs, repository, config, executableRoot, ownership, mirror, snapshot, run };

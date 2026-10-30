const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const vm = require('node:vm');
const { execFileSync, spawnSync } = require('node:child_process');
const root = path.resolve(__dirname, '..');
const library = path.join(root, 'packages/spfx-react-toolkit');
const manifest = require('../packages/spfx-react-toolkit/package.json');
const inventory = require('../packages/spfx-react-toolkit/localization/resources.json');

test('the package exposes a runnable setup executable without importing the browser entry point', () => {
  assert.deepEqual(manifest.bin, { 'spfx-react-toolkit': 'cli/spfx-react-toolkit.cjs' });
  const executable = path.join(library, manifest.bin['spfx-react-toolkit']);
  assert.ok(fs.readFileSync(executable, 'utf8').startsWith('#!/usr/bin/env node\n'));
  const result = spawnSync(process.execPath, [executable, 'configure-localization', '--list'], { cwd: os.tmpdir(), encoding: 'utf8' });
  assert.equal(result.status, 0, result.stderr);
  for (const resource of inventory.resources) assert.ok(result.stdout.includes(resource.id), resource.id);
});

test('the real tarball contains every native asset and its CLI registers an isolated consumer', t => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'apvee-localization-package-'));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const packed = JSON.parse(execFileSync('npm', ['pack', '--ignore-scripts', '--json', '--pack-destination', directory, '--cache', path.join(directory, 'npm-cache')], { cwd: library, encoding: 'utf8' }))[0];
  const paths = new Set(packed.files.map(file => file.path));
  assert.ok(paths.has('cli/spfx-react-toolkit.cjs'));
  assert.ok(paths.has('localization/resources.json'));
  const assets = inventory.resources.flatMap(resource => resource.locales.map(culture => resource.path.replace('{locale}', culture.toLowerCase())));
  assert.equal(assets.length, 400);
  for (const asset of assets) assert.ok(paths.has(asset), `Missing packed native resource: ${asset}`);
  assert.equal([...paths].filter(file => /^lib\/controls\/[^/]+\/loc\/[^/]+\.js$/.test(file)).length, 400);
  const installed = path.join(directory, 'node_modules/@apvee/spfx-react-toolkit');
  fs.mkdirSync(installed, { recursive: true });
  execFileSync('tar', ['-xzf', path.join(directory, packed.filename), '-C', installed, '--strip-components=1']);
  fs.writeFileSync(path.join(directory, 'package.json'), JSON.stringify({ name: 'isolated-localization-consumer', private: true, dependencies: { '@apvee/spfx-react-toolkit': manifest.version } }));
  fs.mkdirSync(path.join(directory, 'config'));
  const configPath = path.join(directory, 'config/config.json');
  fs.writeFileSync(configPath, JSON.stringify({ version: '2.0', localizedResources: { ConsumerStrings: 'lib/loc/{locale}.js' } }, undefined, 2) + '\n');
  const executable = path.join(installed, 'cli/spfx-react-toolkit.cjs');
  const run = args => spawnSync(process.execPath, [executable, 'configure-localization', ...args], { cwd: directory, encoding: 'utf8' });
  const setup = run([]);
  assert.equal(setup.status, 0, setup.stderr);
  const configured = fs.readFileSync(configPath, 'utf8');
  const mappings = JSON.parse(configured).localizedResources;
  assert.equal(mappings.ConsumerStrings, 'lib/loc/{locale}.js');
  for (const resource of inventory.resources) {
    assert.equal(mappings[resource.id], `node_modules/@apvee/spfx-react-toolkit/${resource.path}`);
    for (const culture of resource.locales) {
      const file = path.join(directory, mappings[resource.id].replace('{locale}', culture.toLowerCase()));
      let dictionary;
      vm.runInNewContext(fs.readFileSync(file, 'utf8'), { define: (dependencies, factory) => { assert.equal(dependencies.length, 0); dictionary = factory(); } });
      assert.equal(dictionary.__locale, culture);
    }
  }
  assert.equal(run([]).status, 0);
  assert.equal(run(['--check']).status, 0);
  assert.equal(fs.readFileSync(configPath, 'utf8'), configured);
});

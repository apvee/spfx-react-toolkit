'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { packageName, fixture, write, makePackage, makeProject, run, npm, ok } = require('./fixtures/localization-consumers/fixture.cjs');
const pagination = 'ApveeSpfxReactToolkitPaginationStrings';

function temporary(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'toolkit-localization-lifecycle-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  return root;
}

function pack(root) {
  const packageRoot = makePackage(path.join(root, 'package'));
  const result = npm(packageRoot, ['pack', '--ignore-scripts', '--json']);
  ok(result);
  return path.join(packageRoot, JSON.parse(result.stdout)[0].filename);
}

function installedConsumer(root, tarball, postinstall = 'spfx-react-toolkit configure-localization') {
  return makeProject(path.join(root, 'consumer'), { manifest: {
    dependencies: { [packageName]: `file:${tarball}` },
    scripts: { postinstall, 'localization:check': 'spfx-react-toolkit configure-localization --check', 'localization:repair': 'spfx-react-toolkit configure-localization' }
  } });
}

test('offline npm tarball installation links the bin and runs the composed consumer postinstall without lib', t => {
  const root = temporary(t);
  const tarball = pack(root);
  const consumer = installedConsumer(root, tarball, 'node -e "require(\'fs\').writeFileSync(\'existing-hook.txt\', \'kept\')" && spfx-react-toolkit configure-localization');
  const misleading = makeProject(path.join(root, 'wrong-app'));
  const untouched = fs.readFileSync(path.join(misleading, 'config/config.json'));
  const result = npm(consumer, ['install', '--offline', '--foreground-scripts', '--ignore-scripts=false'], { INIT_CWD: misleading });
  ok(result);
  assert.match(result.stdout, /configure-localization/);
  assert.equal(fs.readFileSync(path.join(consumer, 'existing-hook.txt'), 'utf8'), 'kept');
  assert.equal(fs.existsSync(path.join(consumer, 'node_modules/.bin/spfx-react-toolkit')), true);
  assert.equal(fs.existsSync(path.join(consumer, 'node_modules', packageName, 'lib')), false);
  assert.equal(JSON.parse(fs.readFileSync(path.join(consumer, 'config/config.json'))).localizedResources[pagination], 'node_modules/@apvee/spfx-react-toolkit/lib/controls/pagination/loc/{locale}.js');
  assert.deepEqual(fs.readFileSync(path.join(misleading, 'config/config.json')), untouched);
  ok(npm(consumer, ['run', 'localization:check']));
});

test('scripts-disabled installation leaves config untouched and explicit npm script repairs it', t => {
  const root = temporary(t);
  const consumer = installedConsumer(root, pack(root));
  const config = path.join(consumer, 'config/config.json');
  const before = fs.readFileSync(config);
  ok(npm(consumer, ['install', '--offline', '--ignore-scripts']));
  assert.deepEqual(fs.readFileSync(config), before);
  const check = npm(consumer, ['run', 'localization:check']);
  assert.notEqual(check.status, 0);
  assert.match(check.stderr, /missing/i);
  ok(npm(consumer, ['run', 'localization:repair']));
  ok(npm(consumer, ['run', 'localization:check']));
});

test('actual npm workspaces expose a hoisted symlinked bin to one app postinstall before compilation', t => {
  const root = temporary(t);
  write(root, 'package.json', { name: 'consumer-workspace', version: '1.0.0', private: true, workspaces: ['packages/*', 'apps/*'] });
  makePackage(path.join(root, 'packages/toolkit'));
  const target = makeProject(path.join(root, 'apps/target'), { manifest: { name: 'target-app', dependencies: { [packageName]: '*' }, scripts: { postinstall: 'spfx-react-toolkit configure-localization', check: 'spfx-react-toolkit configure-localization --check' } } });
  const other = makeProject(path.join(root, 'apps/other'), { manifest: { name: 'other-app' } });
  const before = fs.readFileSync(path.join(other, 'config/config.json'));
  ok(npm(root, ['install', '--offline', '--foreground-scripts', '--ignore-scripts=false']));
  const installed = path.join(root, 'node_modules', packageName);
  assert.equal(fs.lstatSync(installed).isSymbolicLink(), true);
  const configured = JSON.parse(fs.readFileSync(path.join(target, 'config/config.json'))).localizedResources;
  assert.equal(configured[pagination], '../../node_modules/@apvee/spfx-react-toolkit/lib/controls/pagination/loc/{locale}.js');
  assert.equal(JSON.stringify(configured).includes(root), false);
  assert.equal(JSON.stringify(configured).includes('packages/toolkit'), false);
  assert.deepEqual(fs.readFileSync(path.join(other, 'config/config.json')), before);
  assert.equal(fs.existsSync(path.join(root, 'packages/toolkit/lib')), false);
  ok(npm(root, ['run', 'check', '--workspace=target-app']));
});

test('cwd wins over INIT_CWD and --project targets exactly one app', t => {
  const f = fixture(t);
  const other = makeProject(path.join(f.root, 'other'));
  const before = fs.readFileSync(path.join(other, 'config/config.json'));
  ok(run(f.cli, f.project, [], { INIT_CWD: other }));
  assert.deepEqual(fs.readFileSync(path.join(other, 'config/config.json')), before);
  fs.mkdirSync(path.join(other, 'node_modules/@apvee'), { recursive: true });
  fs.symlinkSync(f.packageRoot, path.join(other, 'node_modules', packageName), 'dir');
  const selectedBefore = fs.readFileSync(f.config);
  ok(run(f.cli, f.project, ['--project', '../other'], { INIT_CWD: f.project }));
  assert.deepEqual(fs.readFileSync(f.config), selectedBefore);
  assert.equal(JSON.parse(fs.readFileSync(path.join(other, 'config/config.json'))).localizedResources[pagination], 'node_modules/@apvee/spfx-react-toolkit/lib/controls/pagination/loc/{locale}.js');
});

test('hoisted workspace paths survive relocation without serializing a physical checkout path', t => {
  const root = temporary(t);
  const original = path.join(root, 'original');
  const packageRoot = makePackage(path.join(original, 'packages/toolkit'));
  const project = makeProject(path.join(original, 'apps/consumer'));
  fs.mkdirSync(path.join(original, 'node_modules/@apvee'), { recursive: true });
  fs.symlinkSync('../../packages/toolkit', path.join(original, 'node_modules', packageName), 'dir');
  ok(run(path.join(packageRoot, 'cli/spfx-react-toolkit.cjs'), project));
  const config = path.join(project, 'config/config.json');
  const before = fs.readFileSync(config);
  assert.equal(JSON.parse(before).localizedResources[pagination], '../../node_modules/@apvee/spfx-react-toolkit/lib/controls/pagination/loc/{locale}.js');
  const relocated = path.join(root, 'relocated');
  fs.renameSync(original, relocated);
  const cli = path.join(relocated, 'node_modules', packageName, 'cli/spfx-react-toolkit.cjs');
  const consumer = path.join(relocated, 'apps/consumer');
  ok(run(cli, consumer, ['--check']));
  ok(run(cli, consumer));
  assert.deepEqual(fs.readFileSync(path.join(consumer, 'config/config.json')), before);
});

for (const version of ['2.1.0', '9.0.0']) test(`rejects an executing CLI from another installation even when target version is ${version}`, t => {
  const f = fixture(t);
  const other = makeProject(path.join(f.root, 'other'));
  makePackage(path.join(other, 'node_modules', packageName), { version });
  const config = path.join(other, 'config/config.json');
  const before = fs.readFileSync(config);
  const result = run(f.cli, f.project, ['--project', other]);
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /executing|installation|installed.*match/i);
  assert.deepEqual(fs.readFileSync(config), before);
});

test('rejects an uninstalled target instead of deriving paths to the executing checkout', t => {
  const f = fixture(t);
  const target = makeProject(path.join(f.root, 'uninstalled'));
  const config = path.join(target, 'config/config.json');
  const before = fs.readFileSync(config);
  const result = run(f.cli, target);
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /installed|resolve|package/i);
  assert.deepEqual(fs.readFileSync(config), before);
});

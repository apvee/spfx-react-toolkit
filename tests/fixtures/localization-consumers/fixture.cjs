'use strict';
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const packageName = '@apvee/spfx-react-toolkit';
const cliSource = path.resolve(__dirname, '../../..', 'packages/spfx-react-toolkit/cli/spfx-react-toolkit.cjs');
const inventory = require('./resources.json');

function write(root, file, value) {
  const target = path.join(root, file);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, typeof value === 'string' ? value : JSON.stringify(value, undefined, 2) + '\n');
}

function makePackage(root, options = {}) {
  write(root, 'package.json', {
    name: packageName,
    version: options.version || '2.1.0',
    bin: { 'spfx-react-toolkit': 'cli/spfx-react-toolkit.cjs' },
    // Deliberately unresolved browser entry: configuring a clean install must not load it.
    main: 'lib/index.js',
    exports: { '.': './lib/index.js' },
    files: ['cli/spfx-react-toolkit.cjs', 'localization/resources.json']
  });
  write(root, 'localization/resources.json', options.inventory || inventory);
  if (fs.existsSync(cliSource)) {
    write(root, 'cli/spfx-react-toolkit.cjs', fs.readFileSync(cliSource, 'utf8'));
    fs.chmodSync(path.join(root, 'cli/spfx-react-toolkit.cjs'), 0o755);
  }
  return root;
}

function makeProject(root, options = {}) {
  write(root, 'package.json', { name: 'localization-consumer', version: '1.0.0', private: true, ...options.manifest });
  write(root, 'config/config.json', options.config === undefined ? { version: '2.0', localizedResources: { WebPartStrings: 'lib/webparts/example/loc/{locale}.js' } } : options.config);
  return root;
}

function fixture(t, options = {}) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'toolkit-localization-consumer-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const project = makeProject(path.join(root, 'consumer'), options);
  const packageRoot = makePackage(path.join(project, 'node_modules', packageName), options);
  return { root, project, packageRoot, config: path.join(project, 'config/config.json'), cli: path.join(packageRoot, 'cli/spfx-react-toolkit.cjs') };
}

function run(cli, project, args = [], env = {}) {
  return spawnSync(process.execPath, [cli, 'configure-localization', ...args], {
    cwd: project,
    encoding: 'utf8',
    env: { ...process.env, ...env }
  });
}

function npm(root, args, env = {}) {
  // Every dependency is a local tarball/workspace, with its own cache and config.
  // This never installs the toolkit's peer graph or changes the repository install.
  const userconfig = path.join(root, '.npmrc-test');
  if (!fs.existsSync(userconfig)) fs.writeFileSync(userconfig, '');
  return spawnSync(process.platform === 'win32' ? 'npm.cmd' : 'npm', args, {
    cwd: root,
    encoding: 'utf8',
    env: {
      ...process.env,
      npm_config_cache: path.join(root, '.npm-cache'),
      npm_config_userconfig: userconfig,
      npm_config_offline: 'true',
      npm_config_audit: 'false',
      npm_config_fund: 'false',
      npm_config_update_notifier: 'false',
      ...env
    },
    timeout: 60000
  });
}

function ok(result) {
  require('node:assert/strict').equal(result.status, 0, `${result.stdout}\n${result.stderr}\n${result.error || ''}`);
}

module.exports = { packageName, inventory, write, makePackage, makeProject, fixture, run, npm, ok };

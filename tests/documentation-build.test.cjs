const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const crypto = require('node:crypto');
const { spawnSync } = require('node:child_process');
const repository = path.resolve(__dirname, '..');
const fixture = path.join(__dirname, 'fixtures/documentation/artifact-set');
function mirror(t, real = false) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-build-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  if (real) {
    for (const route of ['README.md', 'LICENSE', 'AGENTS.md', 'package.json', 'docs', 'packages/spfx-react-toolkit', 'apps/spfx-react-toolkit-test/src', 'apps/spfx-react-toolkit-test/config', 'apps/spfx-react-toolkit-test/package.json', 'apps/spfx-react-toolkit-test/gulpfile.js', '.github/workflows/verify.yml', '.vscode/launch.json', '.vscode/settings.json', 'tests/fixtures']) {
      fs.mkdirSync(path.dirname(path.join(root, route)), { recursive: true });
      fs.cpSync(path.join(repository, route), path.join(root, route), { recursive: true, filter: source => !/(?:^|\/)(?:lib|node_modules|\.docs)(?:\/|$)/.test(source) });
    }
  } else fs.cpSync(fixture, root, { recursive: true });
  fs.copyFileSync(path.join(repository, 'packages/spfx-react-toolkit/tsdoc.json'), path.join(root, 'packages/spfx-react-toolkit/tsdoc.json'));
  fs.copyFileSync(path.join(repository, 'package-lock.json'), path.join(root, 'package-lock.json'));
  fs.cpSync(path.join(repository, 'scripts'), path.join(root, 'scripts'), { recursive: true });
  fs.symlinkSync(path.join(repository, 'node_modules'), path.join(root, 'node_modules'), 'dir');
  return root;
}
function run(root, args = [], mode, cwd = root) {
  const env = { ...process.env }; delete env.TOOLKIT_DOCS_MODE;
  if (mode !== undefined) env.TOOLKIT_DOCS_MODE = mode;
  return spawnSync(process.execPath, [path.join(root, 'scripts/documentation.cjs'), ...args], { cwd, env, encoding: 'utf8' });
}
function snapshot(root) {
  const result = {};
  function visit(route) {
    const file = path.join(root, route), stat = fs.statSync(file);
    result[route] = { mtime: stat.mtimeMs };
    if (stat.isDirectory()) for (const name of fs.readdirSync(file).sort()) visit(route + '/' + name);
    else result[route].hash = crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
  }
  visit('docs');
  for (const route of ['llms.txt', 'packages/spfx-react-toolkit/README.md']) if (fs.existsSync(path.join(root, route))) visit(route);
  return result;
}
test('CLI defaults to write from another cwd and environment check never repairs outputs', t => {
  const root = mirror(t);
  const written = run(root, [], undefined, os.tmpdir());
  assert.equal(written.status, 0, written.stderr);
  assert.ok(fs.existsSync(path.join(root, 'llms.txt')), 'default write must generate the complete set');
  fs.writeFileSync(path.join(root, 'llms.txt'), 'stale\n');
  const before = snapshot(root);
  const checked = run(root, [], 'check');
  assert.notEqual(checked.status, 0);
  assert.match(checked.stderr, /stale.*llms.txt/is);
  assert.deepEqual(snapshot(root), before);
  assert.equal(run(root, [], 'write').status, 0);
});
test('CLI --check overrides write, rejects unknown env and arguments, and validates before writes', t => {
  const root = mirror(t);
  const initial = snapshot(root);
  const check = run(root, ['--check'], 'write');
  assert.notEqual(check.status, 0);
  assert.deepEqual(snapshot(root), initial);
  for (const [args, mode] of [[[], 'repair'], [['--unknown'], undefined], [['--check'], 'repair']]) {
    const result = run(root, args, mode);
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /mode|usage/i);
    assert.deepEqual(snapshot(root), initial);
  }
  const source = path.join(root, 'packages/spfx-react-toolkit/src/index.ts');
  fs.appendFileSync(source, '\nexport const broken = ;\n');
  const invalid = run(root);
  assert.notEqual(invalid.status, 0);
  assert.match(invalid.stderr, /TS1109/);
  assert.deepEqual(snapshot(root), initial);
});
test('library build and CI use one canonical CLI with job-wide check mode', () => {
  const metadata = require('../packages/spfx-react-toolkit/package.json');
  assert.equal(metadata.scripts.build.split('node ../../scripts/documentation.cjs').length - 1, 1);
  assert.match(metadata.scripts.build, /tsc.*localization-resources.cjs --emit && node ..\/..\/scripts\/documentation.cjs$/);
  const root = require('../package.json');
  assert.equal(root.scripts['docs:generate'], 'node scripts/documentation.cjs');
  assert.equal(root.scripts['docs:check'], 'node scripts/documentation.cjs --check');
  const workflow = fs.readFileSync(path.join(repository, '.github/workflows/verify.yml'), 'utf8');
  assert.match(workflow, /    env:\n      TOOLKIT_DOCS_MODE: check\n    steps:/);
  for (const command of ['build', 'package:solution', 'verify:package']) assert.ok(workflow.includes('- run: npm run ' + command));
});
test('real copied library build refuses stale sources and missing development/llms without changing document bytes or mtimes', { timeout: 240000 }, t => {
  const root = mirror(t, true);
  const baseline = run(root);
  assert.equal(baseline.status, 0, baseline.stderr);
  const source = path.join(root, 'packages/spfx-react-toolkit/src/hooks/useStableCallback.ts');
  const original = fs.readFileSync(source, 'utf8');
  fs.writeFileSync(source, original.replace('/**', '/**\n * Build integration freshness probe.'));
  const developmentSource = path.join(root, 'docs/_sources/development/documentation.md');
  const originalDevelopment = fs.readFileSync(developmentSource, 'utf8');
  fs.appendFileSync(developmentSource, '\nBuild integration development freshness probe.\n');
  const developmentRoutes = Object.keys(snapshot(root)).filter(route => route.startsWith('docs/development/') && route.endsWith('.md'));
  assert.equal(developmentRoutes.length, 5);
  function build(label, expectedRoutes) {
    const before = snapshot(root);
    const result = spawnSync('npm', ['run', 'build:library'], { cwd: root, env: { ...process.env, TOOLKIT_DOCS_MODE: 'check' }, encoding: 'utf8', timeout: 110000 });
    assert.notEqual(result.status, 0, label);
    assert.match(result.stderr, /stale generated documentation/i, result.stdout + result.stderr);
    for (const route of expectedRoutes) assert.ok(result.stderr.includes(route), label + ': missing stale route ' + route);
    assert.match(result.stdout, /Emitted|localization|resources/i);
    assert.ok(fs.existsSync(path.join(root, 'packages/spfx-react-toolkit/lib/index.js')), 'real tsc must run');
    const inventory = JSON.parse(fs.readFileSync(path.join(root, 'packages/spfx-react-toolkit/localization/resources.json'), 'utf8'));
    for (const resource of inventory.resources) for (const locale of resource.locales) {
      const output = path.join(root, 'packages/spfx-react-toolkit', resource.path.replace('{locale}', locale.toLowerCase()));
      assert.match(fs.readFileSync(output, 'utf8'), /define\(\[\], function/, 'real localization emission must run');
    }
    assert.deepEqual(snapshot(root), before, label + ': docs bytes and mtimes must survive');
    t.diagnostic(JSON.stringify({ label, status: result.status, documents: Object.keys(before).length, unchanged: true }));
  }
  build('stale JSDoc and development source', ['docs/api/hooks/react.md', 'docs/development/documentation.md']);
  fs.writeFileSync(source, original);
  fs.writeFileSync(developmentSource, originalDevelopment);
  fs.unlinkSync(path.join(root, 'llms.txt'));
  fs.rmSync(path.join(root, 'docs/development'), { recursive: true });
  build('missing development directory and llms.txt', [...developmentRoutes, 'llms.txt']);
  assert.equal(fs.existsSync(path.join(root, 'docs/development')), false, 'check-mode build must not recreate the missing directory');
});

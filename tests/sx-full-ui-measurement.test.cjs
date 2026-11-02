const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const vm = require('node:vm');
const { createRequire } = require('node:module');

const script = path.resolve(__dirname, '../scripts/measure-sx-bundle.cjs');
const realRequire = createRequire(script);
// Block subprocesses and compilation at the boundary: the RED must never build,
// pack or ship the repository. Real before/after measurements cover the compiler.
async function run(env) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'sx-full-ui-'));
  const commands = [];
  const baselineReads = [];
  let compiled;
  const archive = Buffer.from('isolated archive boundary');
  const baselineHash = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures/sx-bundle/baseline.json'), 'utf8')).tarballSha256;
  // These boundary tests require no ignored archive or emitted host asset.
  // Actual archive bytes/hashes, resource inventory and compiler execution are
  // established separately by the real scoped before/after measurements.
  const filesystem = { ...fs, mkdirSync() {}, rmSync() {}, copyFileSync() {}, writeFileSync() {},
    readFileSync(file, ...args) { if (file.includes('artifacts/sx-baseline')) baselineReads.push(file); return file.endsWith('.tgz') ? archive : fs.readFileSync(file, ...args); },
    readdirSync(file, ...args) { if (file.endsWith('artifacts/sx-baseline/dist')) { baselineReads.push(file); return []; } return fs.readdirSync(file, ...args); },
  };
  const context = {
    __dirname: path.dirname(script), Buffer, console: { log() {}, error() {} },
    process: { env: { SX_MEASURE_FULL_UI_ONLY: '1', SX_MEASURE_OUTPUT: directory, ...env }, version: process.version,
      execPath: process.execPath, stderr: { write() {} } },
    require(name) {
      if (name === 'node:fs') return filesystem;
      if (name === 'node:crypto') return { ...realRequire(name), createHash(algorithm) {
        const hash = realRequire(name).createHash(algorithm);
        let archiveBoundary = false;
        return { update(value) { archiveBoundary = value === archive; hash.update(value); return this; },
          digest(encoding) { return archiveBoundary ? baselineHash : hash.digest(encoding); } };
      } };
      if (name === 'node:child_process') return { spawnSync(executable, args) {
        commands.push({ executable, args });
        if (args.includes('build:library') || args.includes('pack') || args.includes('bundle')) throw new Error('BROAD_COMMAND_BLOCKED');
        return { status: 0, stdout: 'test-npm', stderr: '' };
      } };
      if (name === 'webpack') return config => { compiled = config; throw new Error('COMPILER_BOUNDARY'); };
      if (name === 'typescript') return { ...realRequire(name), readConfigFile: () => ({ config: {} }),
        parseJsonConfigFileContent: () => ({ fileNames: [], options: {} }), createProgram: () => ({ getSourceFiles: () => [] }) };
      if (name === './localization-webpack.internal.cjs') return { nativeResourceAliases: ({ packageRoot }) => ({ 'ExampleNativeStrings$': path.join(packageRoot, 'lib/native.js') }) };
      return realRequire(name);
    },
  };
  context.require.resolve = realRequire.resolve;
  try {
    await vm.runInNewContext(fs.readFileSync(script, 'utf8'), context, { filename: script });
    await new Promise(resolve => setImmediate(resolve));
    return { commands, compiled, baselineReads, exitCode: context.process.exitCode };
  } finally { fs.rmSync(directory, { recursive: true, force: true }); }
}

test('full UI mode rejects an absent explicit candidate before any subprocess', async () => {
  const result = await run({});
  assert.deepEqual(result.commands, []);
  assert.equal(result.exitCode, 1);
  assert.equal(result.compiled, undefined);
});

test('full UI mode reaches only candidate extraction, transitive typecheck and the full entry compiler', async () => {
  const candidate = path.join(os.tmpdir(), 'prepared-sx-test-candidate.tgz');
  const result = await run({ SX_MEASURE_CANDIDATE_TARBALL: candidate });
  assert.ok(result.compiled, 'The full UI compiler must be reached without broad commands');
  assert.equal(result.compiled.entry, path.resolve(__dirname, 'fixtures/sx-bundle/full.tsx'));
  const typecheck = result.commands.find(item => item.args.includes('--noEmit'));
  assert.ok(typecheck, 'Actual full entry is transitively typechecked');
  assert.ok(typecheck.args.includes('--project'));
  assert.equal(result.commands.filter(item => item.executable === 'tar').length, 1);
  assert.match(result.compiled.resolve.alias['ExampleNativeStrings$'], /full-ui-candidate\/package\/lib\/native\.js$/);
  assert.equal(result.compiled.resolve.alias['react/jsx-runtime$'], require.resolve('react/jsx-runtime.js'));
  assert.deepEqual(result.baselineReads, [], 'An explicit candidate is sufficient; narrow mode requires no ignored baseline assets');
  assert.equal(result.commands.some(item => item.args.includes('build:library') || item.args.includes('pack') || item.args.includes('bundle')), false);
});

test('default mode retains the existing library-build pipeline rather than silently narrowing', async () => {
  const result = await run({ SX_MEASURE_FULL_UI_ONLY: '' });
  assert.ok(result.commands.some(item => item.args.includes('build:library')));
  assert.equal(result.compiled, undefined, 'The test blocks the default build before it can run');
});

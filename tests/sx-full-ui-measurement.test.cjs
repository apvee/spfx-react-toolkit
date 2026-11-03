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
async function run(env, options = {}) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'sx-full-ui-'));
  const commands = [];
  const baselineReads = [];
  const removals = [];
  const compilers = [];
  const errors = [];
  const staleExtractions = [];
  for (const role of ['baseline', 'candidate', 'full-ui-candidate']) {
    fs.mkdirSync(path.join(directory, role));
    fs.writeFileSync(path.join(directory, role, 'stale-only.js'), 'must not survive extraction');
  }
  const extractionDirectory = file => path.join(directory, path.basename(file));
  let compiled;
  const archive = Buffer.from('isolated archive boundary');
  const baselineHash = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures/sx-bundle/baseline.json'), 'utf8')).tarballSha256;
  // These boundary tests require no ignored archive or emitted host asset.
  // Actual archive bytes/hashes, resource inventory and compiler execution are
  // established separately by the real scoped before/after measurements.
  const filesystem = { ...fs, mkdirSync() {}, rmSync(file) { removals.push(file); fs.rmSync(extractionDirectory(file), { recursive: true, force: true }); }, copyFileSync() {}, writeFileSync() {},
    existsSync(file) { return String(file).endsWith('.tgz') ? options.missingPack !== true : String(file).endsWith('localization/resources.json') ? options.inventory !== false : fs.existsSync(file); },
    readFileSync(file, ...args) { if (file.includes('artifacts/sx-baseline')) baselineReads.push(file); return file.endsWith('.tgz') ? archive : fs.readFileSync(file, ...args); },
    readdirSync(file, ...args) { if (file.endsWith('candidate-pack')) return ['old-2.1.0.tgz', 'fresh-3.0.0.tgz']; if (file.endsWith('artifacts/sx-baseline/dist')) { baselineReads.push(file); return []; } return fs.readdirSync(file, ...args); },
  };
  const context = {
    __dirname: path.dirname(script), Buffer, console: { log() {}, error(value) { errors.push(String(value)); } },
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
        if (executable === 'tar') staleExtractions.push(fs.existsSync(path.join(extractionDirectory(args[args.indexOf('-C') + 1]), 'stale-only.js')));
        if (!options.default && (args.includes('build:library') || args.includes('pack') || args.includes('bundle'))) throw new Error('BROAD_COMMAND_BLOCKED');
        if (args.includes('pack')) return { status: 0, stdout: options.packJson ?? JSON.stringify([{ name: '@apvee/spfx-react-toolkit', version: '3.0.0', filename: 'fresh-3.0.0.tgz' }]), stderr: '' };
        return { status: 0, stdout: 'test-npm', stderr: '' };
      } };
      if (name === 'webpack') return config => { compiled = config; compilers.push(config); if (options.escape || options.default && config.output.path.endsWith('unchanged-baseline')) return { run(done) { done(undefined, { hasErrors: () => false, toJson: () => ({ chunks: [], modules: options.escape && !config.output.path.endsWith('unchanged-baseline') ? [{ name: './packages/spfx-react-toolkit/lib/index.js' }] : [], errors: [], warnings: [] }) }); }, close(done) { done(); } }; throw new Error('COMPILER_BOUNDARY'); };
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
    return { commands, compiled, compilers, removals, staleExtractions, baselineReads, errors, exitCode: context.process.exitCode };
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

test('default fresh pack selects only the successful JSON filename despite a stale archive', async () => {
  const result = await run({ SX_MEASURE_FULL_UI_ONLY: '0' }, { default: true });
  const command = result.commands.find(command => command.args[0] === 'pack');
  assert.ok(command.args.includes('--json'));
  const extraction = result.commands.find(command => command.args.includes('extract-candidate')) || result.commands.find(command => command.args[0] === '-xzf' && command.args[1].endsWith('fresh-3.0.0.tgz'));
  assert.ok(extraction, JSON.stringify(result.commands));
  assert.ok(result.removals.some(file => file.endsWith('/baseline')) && result.removals.some(file => file.endsWith('/candidate')));
  assert.deepEqual(result.staleExtractions, [false, false], 'Neither extraction can retain a stale file absent from the selected archive');
  assert.equal(result.compilers.length, 2);
  assert.equal(result.compilers[0].resolve.alias['ExampleNativeStrings$'], undefined, 'baseline receives no candidate resources');
  assert.ok(result.compilers[0].resolve.alias['react/jsx-runtime$']?.endsWith('/react/jsx-runtime.js'), 'same pinned host React adapter also applies to historical baseline');
  assert.ok(result.compilers[1].resolve.alias['ExampleNativeStrings$'].includes('/candidate/package/'));
  assert.ok(result.compilers[1].resolve.alias['react/jsx-runtime$'].endsWith('/react/jsx-runtime.js'));
});
for (const packJson of ['invalid', '[]', '[{"filename":"a.tgz"},{"filename":"b.tgz"}]', '[{"filename":"../escape.tgz"}]']) test('default rejects malformed/ambiguous/escaping pack result ' + packJson, async () => {
  const result = await run({ SX_MEASURE_FULL_UI_ONLY: '0' }, { default: true, packJson });
  assert.equal(result.exitCode, 1);
  assert.equal(result.commands.filter(command => command.args[0] === '-xzf').length, 0);
  assert.equal(result.compilers.length, 0);
});
test('default explicit archive never packs and uses its exact selected path', async () => {
  const candidate = path.join(os.tmpdir(), 'explicit-default.tgz');
  const result = await run({ SX_MEASURE_FULL_UI_ONLY: '0', SX_MEASURE_CANDIDATE_TARBALL: candidate }, { default: true });
  assert.equal(result.commands.some(command => command.args[0] === 'pack'), false);
  assert.ok(result.commands.some(command => command.args[0] === '-xzf' && command.args[1] === candidate));
  assert.ok(result.compiled.resolve.alias['ExampleNativeStrings$']);
});
test('default rejects a successful pack result whose archive is missing', async () => {
  const result = await run({ SX_MEASURE_FULL_UI_ONLY: '0' }, { default: true, missingPack: true });
  assert.equal(result.exitCode, 1);
  assert.equal(result.compilers.length, 0);
});

test('an archived candidate without an inventory retains its genuine legacy adapter boundary', async () => {
  const result = await run({ SX_MEASURE_FULL_UI_ONLY: '0', SX_MEASURE_CANDIDATE_TARBALL: path.join(os.tmpdir(), 'old-candidate.tgz') }, { default: true, inventory: false });
  assert.equal(result.compiled.resolve.alias['ExampleNativeStrings$'], undefined);
  assert.ok(result.compiled.resolve.alias['react/jsx-runtime$']);
});
test('candidate compiler stats reject workspace fallback independently of alias selection', async () => {
  const result = await run({ SX_MEASURE_FULL_UI_ONLY: '0', SX_MEASURE_CANDIDATE_TARBALL: path.join(os.tmpdir(), 'candidate.tgz') }, { default: true, escape: true });
  assert.equal(result.exitCode, 1);
  assert.ok(result.errors.some(error => error.includes('Candidate archive imports escaped to the workspace library')), result.errors.join('\n'));
});

test('auto pack rejects a different package identity despite a contained filename', async () => {
  const result = await run({ SX_MEASURE_FULL_UI_ONLY: '0' }, { default: true, packJson: JSON.stringify([{name:'other-package',version:'3.0.0',filename:'fresh-3.0.0.tgz'}]) });
  assert.equal(result.exitCode, 1);
  assert.equal(result.commands.filter(command => command.args[0] === '-xzf').length, 0);
});

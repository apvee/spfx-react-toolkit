const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');

test('package verification records long shared-peer commands without exceeding filename limits or losing arguments', () => {
  const script = path.resolve(__dirname, '../scripts/verify-package.cjs');
  const localRequire = require('node:module').createRequire(script);
  const evidence = fs.mkdtempSync(path.join(os.tmpdir(), 'form-package-evidence-'));
  const processBoundary = { env: { SPFX_PACKAGE_EVIDENCE: evidence }, stdout: { write() {} }, stderr: { write() {} } };
  const source = fs.readFileSync(script, 'utf8').replace(/main\(\)\.catch[\s\S]*$/, 'return { run, temp };');
  let temporary;
  try {
    const api = new Function('require', '__dirname', 'process', source)(key => key === 'node:child_process'
      ? { spawnSync: () => ({ status: 0, signal: undefined, stdout: 'peer tree', stderr: '' }) }
      : localRequire(key), path.dirname(script), processBoundary);
    temporary = api.temp;
    const common = ['ls', '--parseable', '--all', ...Array.from({ length: 20 }, (_, index) => `@fluentui/shared-peer-name-${index}`)];
    api.run([...common, 'first'], process.cwd());
    api.run([...common, 'second'], process.cwd());
    const files = fs.readdirSync(evidence).filter(file => file.endsWith('.stdout.log'));
    assert.equal(files.length, 2);
    assert.ok(files.every(file => Buffer.byteLength(file) < 255));
    assert.ok(files.every(file => fs.readFileSync(path.join(evidence, file), 'utf8') === 'peer tree'));
    const commands = JSON.parse(fs.readFileSync(path.join(evidence, 'commands.json'), 'utf8'));
    assert.deepEqual(commands.map(command => command.args), [[...common, 'first'], [...common, 'second']]);
  } finally {
    fs.rmSync(evidence, { recursive: true, force: true });
    if (temporary) fs.rmSync(temporary, { recursive: true, force: true });
  }
});

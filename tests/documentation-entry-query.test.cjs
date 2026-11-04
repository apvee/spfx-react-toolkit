const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { buildApiGraph } = require('../scripts/docs/documentation-model.cjs');
const config = require('../scripts/docs/documentation.config.cjs');
const fixture = path.join(__dirname, 'fixtures/documentation/artifact-set');
test('requested entries use the guarded checker without changing clean public bindings', t => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-query-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  fs.cpSync(fixture, root, { recursive: true });
  const entry = config.sourceRoot + '/internal.ts';
  fs.writeFileSync(path.join(root, entry), '/** Internal query. */ export const internal = 1;\n');
  const original = buildApiGraph(root, config);
  const queried = buildApiGraph(root, config, { requestedSourceEntries: [entry] });
  assert.deepEqual(queried.bindings, original.bindings);
  assert.deepEqual(queried.entryQueries[entry].map(binding => binding.exportedName), ['internal']);
  assert.ok(!original.entryQueries);
  assert.ok(!original.bindings.some(binding => binding.exportedName === 'internal'));
  const internal = queried.entryQueries[entry][0];
  assert.equal(internal.specifier, '');
  assert.equal(queried.entities.find(entity => entity.key === internal.entityKey).visibility, 'support');
  const outside = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-query-outside-'));
  t.after(() => fs.rmSync(outside, { recursive: true, force: true }));
  fs.writeFileSync(path.join(outside, 'input.ts'), 'export const input = 1;\n');
  for (const route of ['.docs/input.ts', 'packages/spfx-react-toolkit/lib/input.ts']) {
    fs.mkdirSync(path.dirname(path.join(root, route)), { recursive: true });
    fs.writeFileSync(path.join(root, route), 'export const input = 1;\n');
  }
  fs.symlinkSync(path.join(outside, 'input.ts'), path.join(root, config.sourceRoot, 'escape.ts'));
  fs.mkdirSync(path.join(root, 'node_modules'));
  fs.symlinkSync(path.join(root, config.sourceRoot), path.join(root, 'node_modules/workspace'));
  for (const route of ['../' + path.basename(outside) + '/input.ts', '.docs/input.ts', 'packages/spfx-react-toolkit/lib/input.ts', config.sourceRoot + '/escape.ts', 'node_modules/workspace/internal.ts']) {
    assert.throws(() => buildApiGraph(root, config, { requestedSourceEntries: [route] }), /forbidden|outside|source entry/i, route);
  }
});

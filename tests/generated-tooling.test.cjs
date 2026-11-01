const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const h = require('./fixtures/generated-tooling/helpers.cjs');
const unified = () => require('../scripts/generated.cjs');

test('one canonical context/full validation feeds both renderers and matches targeted bytes', t => {
  const root = h.mirror(t);
  const modules = [require('../scripts/documentation-model.cjs'), require('../scripts/documentation-contract.cjs'), require('../scripts/documentation-examples.cjs')];
  const names = ['buildApiGraph', 'validateDocumentation', 'checkDocumentationExamples'];
  const calls = [0, 0, 0];
  const originals = names.map((name, index) => modules[index][name]);
  for (const [index, name] of names.entries()) modules[index][name] = (...args) => { calls[index]++; return originals[index](...args); };
  for (const name of ['generated', 'documentation-context', 'documentation-artifacts', 'executable-examples']) delete require.cache[require.resolve('../scripts/' + name + '.cjs')];
  let actual;
  try { actual = unified().generateAllArtifacts(root, h.config); }
  finally { for (const [index, name] of names.entries()) modules[index][name] = originals[index]; }
  assert.deepEqual(calls, [1, 1, 1]);
  assert.deepEqual(actual.documentation, require('../scripts/documentation-artifacts.cjs').generateArtifacts(root, h.config));
  assert.deepEqual(actual.examples, require('../scripts/executable-examples.cjs').generateExecutableArtifacts(root, h.config));
  const source = path.join(root, h.config.sourceRoot, 'index.ts');
  fs.writeFileSync(source, fs.readFileSync(source, 'utf8').replace('count(3)', 'count(4)'));
  const changed = unified().generateAllArtifacts(root, h.config);
  for (const family of ['documentation', 'examples']) assert.notDeepEqual(changed[family], actual[family]);
});

test('explicit CLI ignores targeted environment modes, is idempotent and rejects every unknown flag read-only', t => {
  const root = h.mirror(t);
  const write = h.run(root, [], { TOOLKIT_DOCS_MODE: 'invalid', TOOLKIT_EXAMPLES_MODE: 'check' });
  assert.equal(write.status, 0, write.stderr);
  const before = h.snapshot(root);
  for (const script of ['documentation.cjs', 'executable-examples.cjs']) {
    const targeted = h.run(root, [], { TOOLKIT_DOCS_MODE: 'write', TOOLKIT_EXAMPLES_MODE: 'write' }, script);
    assert.equal(targeted.status, 0, targeted.stderr);
    assert.deepEqual(h.snapshot(root), before, 'targeted CLI must produce identical bytes without rewriting');
  }
  assert.equal(h.run(root).status, 0);
  assert.deepEqual(h.snapshot(root), before);
  assert.equal(h.run(root, ['--check'], { TOOLKIT_DOCS_MODE: 'write', TOOLKIT_EXAMPLES_MODE: 'invalid' }).status, 0);
  for (const args of [['--unknown'], ['--check', '--check'], ['--write']]) {
    assert.notEqual(h.run(root, args).status, 0);
    assert.deepEqual(h.snapshot(root), before);
  }
});

test('check preserves bytes, mtimes and directories for missing/changed/orphan in both families', t => {
  const root = h.mirror(t);
  assert.equal(h.run(root).status, 0);
  const expected = unified().generateAllArtifacts(root, h.config);
  for (const [family, artifacts] of Object.entries(expected)) {
    const routes = Object.keys(artifacts);
    fs.unlinkSync(path.join(root, routes[0]));
    fs.writeFileSync(path.join(root, routes[1]), (family === 'examples' ? h.ownership + '\n' : '') + 'stale\n');
    const orphan = family === 'examples' ? h.executableRoot + '/orphan.ts' : 'docs/api/orphan.md';
    fs.writeFileSync(path.join(root, orphan), artifacts[routes[0]]);
  }
  const before = h.snapshot(root);
  const checked = h.run(root, ['--check']);
  assert.notEqual(checked.status, 0);
  assert.match(checked.stderr, /documentation/);
  assert.match(checked.stderr, /examples/);
  assert.deepEqual(h.snapshot(root), before);
  assert.equal(h.run(root).status, 0);
  fs.rmSync(path.join(root, h.executableRoot), { recursive: true });
  const missingDirectory = h.snapshot(root);
  assert.notEqual(h.run(root, ['--check']).status, 0);
  assert.deepEqual(h.snapshot(root), missingDirectory);
});

test('invalid executable, invalid document links and manual executable conflicts prevent all writes', t => {
  for (const failure of ['executable', 'link', 'manual']) {
    const root = h.mirror(t);
    assert.equal(h.run(root).status, 0);
    fs.writeFileSync(path.join(root, 'llms.txt'), 'stale docs\n');
    if (failure === 'executable') {
      const source = path.join(root, h.config.sourceRoot, 'index.ts');
      fs.writeFileSync(source, fs.readFileSync(source, 'utf8').replace("kind: 'snapshot'", "kind: 'invalid'"));
    } else if (failure === 'link') fs.appendFileSync(path.join(root, 'docs/_sources/development/INDEX.md'), '\n[Missing](no-such-page.md)\n');
    else fs.writeFileSync(path.join(root, h.executableRoot, 'catalogue.ts'), '// Manual file\n');
    const before = h.snapshot(root);
    const result = h.run(root);
    assert.notEqual(result.status, 0, failure);
    assert.match(result.stderr, failure === 'executable' ? /EXECUTABLE_/ : failure === 'link' ? /DOC_LINK_TARGET/ : /manual artifact/);
    assert.deepEqual(h.snapshot(root), before, failure);
  }
});

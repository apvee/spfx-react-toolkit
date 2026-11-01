const assert = require('node:assert/strict');
const path = require('node:path');
const test = require('node:test');
const expected = require('./fixtures/documentation/semantic/expected.json');
const { buildApiGraph, resolvePublicReference } = require('../scripts/documentation-model.cjs');
const fixture = path.join(__dirname, 'fixtures/documentation/semantic');
let graph;
test.before(() => { graph = buildApiGraph(fixture, { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' }); });
const entity = reference => resolvePublicReference(graph, reference);
const binding = reference => graph.bindings.find(item => `${item.exportedName}${item.memberPath.length ? '.' + item.memberPath.join('.') : ''}` === reference);

test('discovers clean public bindings and preserves local alias and namespace documentation', () => {
  for (const reference of expected.references) assert.ok(entity(reference), reference);
  assert.match(entity('publicCallback').rawDoc, /Locally documented alias/);
  assert.equal(entity('publicCallback').signatures.length, 1);
  assert.match(binding('metrics').rawDoc, /Local namespace contract/);
  assert.deepEqual(binding('metrics.nested.px').memberPath, ['nested', 'px']);
  assert.equal(binding('metrics.nested.px').exportedName, 'metrics');
  assert.equal(binding('metrics.nested.px').specifier, '@fixture/semantic');
  assert.equal(binding('metrics.nested.px').ownerKey, entity('metrics.nested').key);
  assert.equal(entity('renamed').key, entity('original').key);
  assert.throws(() => entity('privateSupport'), /Unknown public reference/);
  assert.ok(!graph.bindings.some(item => item.specifier === '@fixture/semantic' && item.exportedName === 'makePanel'));
  assert.ok(graph.bindings.some(item => item.specifier === '@fixture/semantic/pane' && item.exportedName === 'makePanel'));
  assert.ok(!graph.bindings.some(item => item.specifier.includes('/lib')));
  assert.deepEqual(graph.diagnostics, []);
});
test('preserves callable consts, overloads, Object.assign mounts and mount comments', () => {
  assert.equal(entity('callable').signatures.length, 1);
  assert.equal(entity('convert').signatures.length, 2);
  assert.deepEqual(entity('Compound').members.map(item => item.name), ['Header']);
  assert.match(binding('Compound.Header').rawDoc, /Mounted header contract/);
  assert.equal(entity('Compound.Header').signatures.length, 1);
});
test('anonymous member types retain directly enclosing property JSDoc with exact provenance and explicit overrides', () => {
  const find = name => graph.entities.find(item => item.qualifiedName === `CallbackOwner.${name}`);
  for (const [name, text, line] of [
    ['accessor', 'Pure synchronous source returning original values.', 3],
    ['tooltip', 'Tooltip for the original value.', 5],
    ['explicit', 'Explicit anonymous signature contract.', 8],
    ['config', 'Nested configuration contract.', 9],
  ]) {
    const value = find(name);
    assert.ok(value, name);
    assert.ok(value.rawDoc.includes(text), `${name}: ${value.rawDoc}`);
    assert.deepEqual(value.docSource, { file: 'src/member-docs.ts', line });
    for (const signature of value.signatures) {
      assert.ok(signature.rawDoc.includes(text), `${name} signature: ${signature.rawDoc}`);
      assert.deepEqual(signature.docSource, { file: 'src/member-docs.ts', line });
    }
  }
  assert.equal(find('undocumented').rawDoc, '');
  assert.equal(find('undocumented').docSource, undefined);
});
test('explicit anonymous type-position comments win through parenthesis, union and intersection wrappers', () => {
  for (const [name, text, line] of [
    ['parenthesized', 'Explicit parenthesized type.', 13],
    ['union', 'Explicit union type.', 15],
    ['intersection', 'Explicit intersection type.', 17],
    ['nested', 'Explicit nearest wrapper type.', 20],
  ]) {
    const values = graph.entities.filter(item => item.qualifiedName === `CallbackOwner.${name}`);
    assert.ok(values.length, name);
    for (const value of values) {
      assert.ok(value.rawDoc.includes(text), `${name}: ${value.rawDoc}`);
      assert.deepEqual(value.docSource, { file: 'src/member-docs.ts', line });
      for (const signature of value.signatures) {
        assert.ok(signature.rawDoc.includes(text), `${name} signature: ${signature.rawDoc}`);
        assert.deepEqual(signature.docSource, { file: 'src/member-docs.ts', line });
      }
    }
  }
});
test('preserves discriminated branch requirements, Omit/intersection and never exclusions', () => {
  const variants = entity('SurfaceProps').variants;
  assert.equal(variants.length, 2);
  for (const rule of expected.surfaceRules) {
    const branch = variants.find(item => item.members.some(member => member.name === 'mode' && member.type.text === JSON.stringify(rule.mode)));
    assert.ok(branch, rule.mode);
    const trigger = branch.members.find(item => item.name === 'trigger');
    assert.equal(trigger.optional, rule.triggerOptional);
    assert.equal(branch.members.find(item => item.name === 'forbidden').forbidden, true);
    assert.ok(!branch.members.some(item => item.name === 'removed'));
    assert.ok(branch.members.some(item => item.name === 'kept'));
  }
});
test('preserves generic factory return suites and support types without invented imports', () => {
  const factory = entity('makeSuite');
  assert.match(factory.signatures[0].typeParameters[0].constraint, /Record/);
  const suite = factory.signatures[0].returns;
  assert.match(suite.text, /Suite<T>/);
  assert.deepEqual(suite.members.map(item => item.name), expected.suiteMembers);
  const support = graph.entities.find(item => item.qualifiedName === 'HiddenResult');
  assert.ok(support);
  assert.equal(support.visibility, 'support');
  assert.ok(!graph.bindings.some(item => item.entityKey === support.key));
  assert.ok(entity('callable').references.includes(support.key));
});
test('graph is serializable, deterministic and uses semantic keys and relative locations', () => {
  const again = buildApiGraph(fixture, { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' });
  assert.equal(JSON.stringify(graph), JSON.stringify(again));
  assert.ok(!JSON.stringify(graph).includes(fixture));
  for (const item of graph.entities) {
    assert.ok(!/:\d+$/.test(item.key), item.key);
    assert.ok(!path.isAbsolute(item.source.file));
  }
});

test('real pilot preserves Stage branches, compound mounts, PreviewCard inheritance, styles, aliases and Form suites', () => {
  const root = path.resolve(__dirname, '..');
  const actual = buildApiGraph(root, require('../scripts/documentation.config.cjs'));
  const find = reference => resolvePublicReference(actual, reference);
  assert.ok(find('Stage.Header').signatures.length);
  const variants = find('StageProps').variants;
  assert.ok(variants.length >= 3);
  for (const variant of variants) {
    const presentation = variant.members.find(item => item.name === 'presentation');
    const trigger = variant.members.find(item => item.name === 'trigger');
    assert.ok(presentation);
    assert.equal(trigger.optional, presentation.type.text !== '"popover"');
  }
  assert.ok(find('PreviewCardProps').members.some(item => item.name === 'title'));
  assert.ok(find('PreviewCardProps').members.some(item => item.name === 'preview'));
  assert.ok(find('width.px').signatures.length);
  const nested = actual.bindings.find(item => item.exportedName === 'overflow' && item.memberPath.join('.') === 'horizontal.hidden');
  assert.ok(nested);
  assert.ok(find('useSx').signatures.length);
  assert.match(find('useStableCallback').rawDoc, /Alias of Fluent UI/);
  assert.ok(find('useStableCallback').signatures.length);
  assert.ok(find('Form.Root').signatures.length);
  const factory = find('createForm');
  assert.ok(factory.signatures.length >= 3);
  assert.ok(factory.signatures.every(item => item.returns.members.some(member => member.name === 'Root')));
  assert.ok(!actual.bindings.some(item => item.specifier === '@apvee/spfx-react-toolkit' && item.exportedName.startsWith('PropertyPane')));
  assert.deepEqual(actual.diagnostics, []);
});

test('retains type expression, heritage and generic constraints while peer entities stay opaque', () => {
  assert.match(entity('SurfaceProps').typeExpression, /Omit<Base, 'removed'>/);
  const extension = entity('ExtendedPeer');
  assert.deepEqual(extension.typeParameters, [{ name: 'T', constraint: 'string', default: 'string' }]);
  assert.equal(extension.heritage[0].text, 'PeerBase');
  const peer = graph.entities.find(item => item.qualifiedName === 'PeerBase');
  assert.equal(peer.visibility, 'external');
  assert.ok(extension.references.includes(peer.key));
  assert.deepEqual(peer.members, []);
});
test('retains scalar value type spelling without requiring renderers to extract declarations', () => {
  const rendered = buildApiGraph(path.join(__dirname, 'fixtures/documentation/render'), { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' });
  const answer = resolvePublicReference(rendered, 'answer');
  assert.deepEqual(answer.type, { text: '42', references: [] });
  assert.equal(answer.typeExpression, undefined);
});
test('retains effective readonly properties through inheritance and mapped add/remove modifiers', () => {
  const rendered = buildApiGraph(path.join(__dirname, 'fixtures/documentation/render'), { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' });
  for (const [name, expected] of [['ReadonlyBase', [false, true]], ['ReadonlyChild', [false, false, true]], ['Frozen', [true, true]], ['Mutable', [false, false]]]) {
    assert.deepEqual(resolvePublicReference(rendered, name).members.map(item => Boolean(item.readonly)), expected, name);
  }
});
test('projects anonymous callback support types with their real signature instead of any', () => {
  const rendered = buildApiGraph(path.join(__dirname, 'fixtures/documentation/render'), { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' });
  const callback = rendered.entities.find(item => item.qualifiedName === 'FormRootProps.onInvalid');
  assert.equal(callback.kind, 'type');
  assert.equal(callback.type.text, '(reason: string) => void | Promise<void>');
  assert.equal(callback.signatures.length, 1);
  assert.equal(callback.signatures[0].returns.text, 'void | Promise<void>');
  assert.ok(!rendered.bindings.some(item => item.entityKey === callback.key));
});
test('reports source compiler failures instead of silently accepting an unresolved graph', () => {
  const fs = require('node:fs');
  const os = require('node:os');
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-model-'));
  try {
    fs.cpSync(fixture, temporary, { recursive: true });
    fs.appendFileSync(path.join(temporary, 'src/index.ts'), '\nexport const invalid: string = 42;\n');
    const invalid = buildApiGraph(temporary, { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' });
    assert.ok(invalid.diagnostics.some(item => item.code === 'TS2322' && item.file === 'src/index.ts' && item.line > 0));
    assert.ok(!JSON.stringify(invalid.diagnostics).includes(temporary));
  } finally {
    fs.rmSync(temporary, { recursive: true, force: true });
  }
});

test('semantic identity is stable when the same source fixture is relocated', () => {
  const fs = require('node:fs');
  const os = require('node:os');
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-model-relocated-'));
  try {
    fs.cpSync(fixture, temporary, { recursive: true });
    const relocated = buildApiGraph(temporary, { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' });
    const originalSignature = entity('returnNamespace').signatures[0];
    const relocatedSignature = resolvePublicReference(relocated, 'returnNamespace').signatures[0];
    assert.equal(originalSignature.text, '(): typeof import("./src/metrics")');
    assert.equal(relocatedSignature.text, originalSignature.text);
    for (const signature of [originalSignature, relocatedSignature]) {
      assert.ok(!signature.text.includes(fixture) && !signature.text.includes(temporary));
      assert.ok(!signature.key.includes(fixture) && !signature.key.includes(temporary));
    }
    assert.deepEqual(relocated, graph);
  } finally {
    fs.rmSync(temporary, { recursive: true, force: true });
  }
});

test('symbol-keyed brand members use source identities and are not invented dotted accesses', () => {
  assert.equal(entity('Token').members[0].name, '[brand]');
  assert.ok(!graph.bindings.some(item => item.exportedName === 'token' && item.memberPath.length));
  assert.ok(!JSON.stringify(graph).includes('__@'));
});


test('dependency-qualified signatures use the same canonical spelling in relocated graphs', () => {
  const fs = require('node:fs');
  const os = require('node:os');
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-model-dependency-'));
  const roots = ['a', 'b'].map(name => path.join(temporary, name, 'source'));
  try {
    const graphs = roots.map(root => {
      fs.cpSync(fixture, root, { recursive: true });
      const peerFile = path.join(root, '../.cache/node_modules/@fixture/namespace-peer/index.d.ts');
      fs.mkdirSync(path.dirname(peerFile), { recursive: true });
      fs.copyFileSync(path.join(fixture, 'peer/index.d.ts'), peerFile);
      fs.copyFileSync(path.join(fixture, 'peer-install/package.json'), path.join(path.dirname(peerFile), 'package.json'));
      fs.copyFileSync(path.join(fixture, 'peer-install/package-lock.json'), path.join(root, 'package-lock.json'));
      // The controlled install remains outside the source mirror, admitted
      // through its own node_modules rather than ambient ancestor lookup.
      fs.symlinkSync(path.resolve(root, '../.cache/node_modules'), path.join(root, 'node_modules'), 'dir');
      const configFile = path.join(root, 'tsconfig.json');
      const config = JSON.parse(fs.readFileSync(configFile, 'utf8'));
      config.compilerOptions.paths['@fixture/namespace-peer'] = ['node_modules/@fixture/namespace-peer/index.d.ts'];
      fs.writeFileSync(configFile, JSON.stringify(config));
      fs.appendFileSync(path.join(root, 'src/index.ts'), '\nimport * as peerNamespace from "@fixture/namespace-peer";\nexport const returnPeerNamespace = () => peerNamespace;\n');
      return buildApiGraph(root, { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' });
    });
    for (const current of graphs) {
      assert.deepEqual(current.diagnostics, []);
      const signature = resolvePublicReference(current, 'returnPeerNamespace').signatures[0];
      assert.equal(signature.text, '(): typeof import("./node_modules/@fixture/namespace-peer/index")');
      assert.equal(signature.returns.text, 'typeof import("./node_modules/@fixture/namespace-peer/index")');
      assert.ok(!signature.text.includes(temporary));
      assert.ok(!signature.key.includes(temporary));
    }
    assert.deepEqual(graphs[0], graphs[1]);
  } finally {
    fs.rmSync(temporary, { recursive: true, force: true });
  }
});


test('no-file compiler diagnostics retain the canonical config location and compiler message', t => {
  const fs = require('node:fs');
  const os = require('node:os');
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-diagnostic-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  fs.cpSync(fixture, root, { recursive: true });
  const file = path.join(root, 'tsconfig.json');
  const project = JSON.parse(fs.readFileSync(file, 'utf8'));
  project.compilerOptions.types = ['no-such-documentation-type'];
  fs.writeFileSync(file, JSON.stringify(project));
  const build = () => buildApiGraph(root, { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' });
  const diagnostic = build().diagnostics.find(item => item.code === 'TS2688');
  assert.ok(diagnostic);
  assert.equal(diagnostic.file, 'tsconfig.json');
  assert.equal(diagnostic.line, 1);
  assert.match(diagnostic.message, /Cannot find type definition file for 'no-such-documentation-type'/);
  assert.deepEqual(build().diagnostics, build().diagnostics);
});

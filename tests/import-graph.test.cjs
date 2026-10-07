const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { auditImportGraph } = require('../scripts/audit-import-graph.cjs');
const sourceRoot = path.join(__dirname, 'fixtures/import-graph');
const audit = () => auditImportGraph({ sourceRoot, entrypoints: ['index.ts'] });
test('type-only edges introduce no runtime dependency', () => {
  const report = audit();
  assert.ok(report.edges.filter(e => e.from === 'leaf.ts' && e.resolvedTarget === 'types.ts').every(e => e.kind === 'type-only'));
  assert.deepEqual(report.publicExports.find(e => e.name === 'alias').dependencies, []);
  assert.equal(report.publicExports.find(e => e.name === 'onlyType').kind, 'type');
});
test('named and namespace reexports resolve canonical leaves', () => {
  const report = audit();
  assert.equal(report.publicExports.find(e => e.name === 'alias').canonicalSource, 'leaf.ts');
  assert.equal(report.publicExports.find(e => e.name === 'namespace.value').canonicalSource, 'leaf.ts');
});
test('internal barrel import and unresolved module are reported', () => {
  const report = audit();
  assert.ok(report.issues.some(i => i.kind === 'internal-barrel-import' && i.from === 'problem.ts'));
  assert.ok(report.issues.some(i => i.kind === 'unresolved-module' && i.specifier === './missing'));
  assert.ok(report.modules.every(m => m.effects.rationale && Array.isArray(m.directConsumers)));
});
test('runtime cycles are reported without losing unrelated exports', () => {
  const report = audit();
  assert.ok(report.issues.some(i => i.kind === 'runtime-cycle' && i.modules.includes('cycle-a.ts')));
  assert.ok(report.publicExports.some(e => e.name === 'alias'));
});
test('unmarked type-only reexport is absent in emitted JS', () => {
  const report = audit();
  assert.equal(report.edges.find(e => e.from==='index.ts' && e.specifier==='./types' && e.kind!=='type-only'),undefined);
});
test('mixed runtime and erased reexports sharing a specifier retain separate edge kinds', () => {
  const report = audit();
  assert.deepEqual(report.edges.filter(e=>e.from==='index.ts'&&e.specifier==='./mixed').map(e=>e.kind),['reexport','type-only']);
  assert.equal(report.publicExports.find(e=>e.name==='erasedNamespace.live').kind,'type');
});
test('unresolved external packages are reported', () => {
  assert.ok(audit().issues.some(i=>i.kind==='unresolved-module' && i.specifier==='missing-external-registration'));
});
test('renamed imported namespaces retain runtime leaves and dependencies', () => {
  const report = auditImportGraph({sourceRoot,entrypoints:['namespace-alias.ts']});
  const value = report.publicExports.find(e=>e.name==='renamed.value');
  assert.equal(value.canonicalSource,'namespace-leaf.ts');
  assert.equal(value.kind,'runtime');
  assert.deepEqual(value.dependencies,['@fluentui/react-utilities']);
  const namedValue = report.publicExports.find(e=>e.name==='namedAlias.value');
  assert.equal(namedValue.kind,'runtime');
  assert.equal(namedValue.canonicalSource,'namespace-leaf.ts');
  assert.deepEqual(namedValue.dependencies,['@fluentui/react-utilities']);
  const shape = report.publicExports.find(e=>e.name==='renamed.NamespaceShape');
  assert.equal(shape.kind,'type');
  assert.deepEqual(shape.dependencies,[]);
});

test('source rules reject internal barrel imports, global renderers and unclassified registrations', () => {
  const { checkSourceRules } = require('../scripts/audit-import-graph.cjs');
  const report = { edges: [
    { from: 'hooks/leaf.ts', kind: 'type-only', specifier: '../core', resolvedTarget: 'core/index.ts' },
    { from: 'hooks/leaf.ts', kind: 'import', specifier: '@apvee/spfx-react-toolkit', resolvedTarget: '@apvee/spfx-react-toolkit' }
  ], modules: [
    { path: 'services/new.ts', externalImports: [{ kind: 'bare-import', specifier: '@pnp/sp/webs' }], topLevelInitializations: [] },
    { path: 'helpers/styles/new.ts', externalImports: [], topLevelInitializations: [{ expression: 'makeStyles({ root: {} })' }] }
  ] };
  assert.deepEqual(checkSourceRules(report, []).map(issue => issue.kind), ['internal-barrel-import', 'internal-package-import', 'unclassified-external-registration', 'package-global-renderer']);
});

test('repository source effects match the manifest and internal imports use direct modules', () => {
  const fs = require('node:fs');
  const { checkSourceRules } = require('../scripts/audit-import-graph.cjs');
  const report = auditImportGraph({ sourceRoot: path.resolve(__dirname, '../packages/spfx-react-toolkit/src'), entrypoints: ['index.ts'] });
  const manifest = JSON.parse(fs.readFileSync(path.resolve(__dirname, '../packages/spfx-react-toolkit/package.json')));
  assert.deepEqual(checkSourceRules(report, manifest.sideEffects), []);
});

test('source rules preserve type-only reexports at public barrel boundaries', () => {
  const {checkSourceRules} = require('../scripts/audit-import-graph.cjs');
  const report = {edges:[{from:'index.ts',specifier:'./core',resolvedTarget:'core/index.ts',kind:'type-only',declarationKind:'reexport'}],modules:[]};
  assert.deepEqual(checkSourceRules(report, []), []);
  const typeEdge = audit().edges.find(edge => edge.from === 'index.ts' && edge.specifier === './types');
  assert.equal(typeEdge.declarationKind, 'reexport');
});

test('namespace barrel reexports remain allowed while type-only internal barrel imports are rejected', () => {
  const fs = require('node:fs');
  const os = require('node:os');
  const {checkSourceRules} = require('../scripts/audit-import-graph.cjs');
  const directory = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), 'graph-barrel-rule-')));
  try {
    fs.mkdirSync(path.join(directory, 'core'));
    fs.writeFileSync(path.join(directory, 'core/index.ts'), 'export interface Shape { id: number }\nexport const value = 1;');
    fs.writeFileSync(path.join(directory, 'index.ts'), "export * as namespace from './core';\nexport type * as model from './core';");
    fs.writeFileSync(path.join(directory, 'leaf.ts'), "import type { Shape } from './core';\nexport type LocalShape = Shape;");
    const report = auditImportGraph({sourceRoot:directory,entrypoints:['index.ts']});
    assert.ok(report.edges.filter(edge => edge.from === 'index.ts').every(edge => edge.declarationKind === 'reexport'));
    assert.equal(report.edges.find(edge => edge.from === 'leaf.ts').kind, 'type-only');
    assert.deepEqual(checkSourceRules(report, []).map(issue => [issue.kind,issue.from]), [['internal-barrel-import','leaf.ts']]);
  } finally {fs.rmSync(directory,{recursive:true,force:true});}
});

test('renderer factories are deferred until called, including immediately invoked wrappers', () => {
  const {checkSourceRules}=require('../scripts/audit-import-graph.cjs');
  const report=expression=>({edges:[],modules:[{path:'factory.ts',externalImports:[],topLevelInitializations:[{expression}]}]});
  for(const expression of ['() => createDOMRenderer()', 'function () { return makeStyles({}); }', '(() => () => createDOMRenderer())()']) assert.deepEqual(checkSourceRules(report(expression),[]),[],expression);
  for(const expression of ['createDOMRenderer()', '(() => createDOMRenderer())()', '(function () { return createRegistry(); })()', 'factory(createDOMRenderer())', '(() => () => createDOMRenderer())()()', '(function () { return () => createRegistry(); })()()']) assert.equal(checkSourceRules(report(expression),[])[0]?.kind,'package-global-renderer',expression);
});

test('object methods and accessors defer renderer calls while computed names execute', () => {
  const {checkSourceRules}=require('../scripts/audit-import-graph.cjs');
  const report=expression=>({edges:[],modules:[{path:'factory.ts',externalImports:[],topLevelInitializations:[{expression}]}]});
  for(const expression of [
    '({ method() { return createDOMRenderer(); } })',
    '({ get renderer() { return createDOMRenderer(); } })',
    '({ set renderer(value) { registerRenderer(value); } })',
    '({ method(renderer = createDOMRenderer()) { return renderer; } })'
  ]) assert.deepEqual(checkSourceRules(report(expression),[]),[],expression);
  assert.equal(checkSourceRules(report('({ [createDOMRenderer()]() {} })'),[])[0]?.kind,'package-global-renderer');
});

test('invoked default parameters execute renderer calls only when the argument is absent', () => {
  const {checkSourceRules}=require('../scripts/audit-import-graph.cjs');
  const report=expression=>({edges:[],modules:[{path:'factory.ts',externalImports:[],topLevelInitializations:[{expression}]}]});
  for(const expression of [
    '((renderer = createDOMRenderer()) => renderer)()',
    '(function (renderer = createRegistry()) { return renderer; })()',
    '((renderer = createDOMRenderer()) => renderer)(undefined)'
  ]) assert.equal(checkSourceRules(report(expression),[])[0]?.kind,'package-global-renderer',expression);
  for(const expression of [
    '(renderer = createDOMRenderer()) => renderer',
    '((renderer = createDOMRenderer()) => renderer)({})'
  ]) assert.deepEqual(checkSourceRules(report(expression),[]),[],expression);
});

test('returned conditional factories remain deferred until the returned function is invoked', () => {
  const {checkSourceRules}=require('../scripts/audit-import-graph.cjs');
  const report=expression=>({edges:[],modules:[{path:'factory.ts',externalImports:[],topLevelInitializations:[{expression}]}]});
  for(const expression of [
    '(() => true ? () => createDOMRenderer() : () => createRegistry())()()',
    '(function () { return condition ? () => createDOMRenderer() : () => createRegistry(); })()()'
  ]) assert.equal(checkSourceRules(report(expression),[])[0]?.kind,'package-global-renderer',expression);
  assert.deepEqual(checkSourceRules(report('(() => true ? () => createDOMRenderer() : () => createRegistry())()'),[]),[]);
});

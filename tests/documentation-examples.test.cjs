const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const { buildApiGraph } = require('../scripts/documentation-model.cjs');
const { buildDocumentationModel, getDocumentationExample } = require('../scripts/documentation-contract.cjs');

const fixture = path.join(__dirname, 'fixtures/documentation/examples');
const config = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' };
let model;
test.before(() => { model = buildDocumentationModel(buildApiGraph(fixture, config), config); });
function check(root, input) {
  assert.ok(fs.existsSync(path.join(__dirname, '../scripts/documentation-examples.cjs')), 'source example checker must exist');
  return require('../scripts/documentation-examples.cjs').checkDocumentationExamples(root, input, config);
}
function negative(name, input = model) {
  const graph = structuredClone(input.graph);
  const entity = graph.entities.find(item => item.qualifiedName === 'Pagination');
  const file = `negative/${name}.jsdoc`;
  entity.rawDoc = fs.readFileSync(path.join(fixture, file), 'utf8');
  entity.docSource = { file, line: 1 };
  const result = buildDocumentationModel(graph, config);
  return { ...result, examples: result.examples.filter(item => item.source.file === file) };
}
function custom(examples) {
  return { ...model, examples: examples.map((code, index) => ({ ownerKey: 'isolated', language: 'ts', code,
    source: { file: `isolated-${index}.ts`, line: 20 } })) };
}
function files(root) {
  return fs.readdirSync(root, { withFileTypes: true }).flatMap(entry => {
    const filename = path.join(root, entry.name);
    if (entry.isSymbolicLink()) return [];
    return entry.isDirectory() ? files(filename) : [path.relative(root, filename).split(path.sep).join('/')];
  }).sort();
}

test('typechecks normalized props, generics, callbacks and type/value examples without lib', () => {
  assert.equal(fs.existsSync(path.join(fixture, 'lib')), false);
  assert.equal(model.examples.length, 4);
  assert.deepEqual(check(fixture, model), []);
  const complete = getDocumentationExample(model, 'Pagination', 'Complete state');
  const single = getDocumentationExample(model, 'transform');
  assert.deepEqual(check(fixture, { ...model, examples: [complete, single] }), []);
  assert.throws(() => getDocumentationExample(model, 'Pagination'), /unique example title/);
});

test('reports the missing pageIndex at its exact original comment line', () => {
  assert.deepEqual(check(fixture, negative('missing-state')), [{
    code: 'TS2304', file: 'negative/missing-state.jsdoc', line: 6,
    symbol: 'src/pagination.tsx#Pagination', message: "Cannot find name 'pageIndex'.",
  }]);
});

test('retains JSX props, callback, generic argument and type-only import errors', () => {
  const diagnostics = check(fixture, negative('invalid-contracts'));
  assert.ok(diagnostics.some(item => item.code === 'TS2322' && item.line === 7), JSON.stringify(diagnostics));
  assert.ok(diagnostics.some(item => item.code === 'TS2339' && /toUpperCase/.test(item.message)), JSON.stringify(diagnostics));
  assert.ok(diagnostics.some(item => item.code === 'TS2345' && item.line === 8), JSON.stringify(diagnostics));
  assert.ok(diagnostics.some(item => item.code === 'TS2693' && item.line === 9), JSON.stringify(diagnostics));
});

test('rejects root Property Pane, private, legacy lib and nonexistent imports', () => {
  const diagnostics = check(fixture, negative('illegal-imports'));
  assert.ok(diagnostics.some(item => item.code === 'TS2305' && item.line === 4 && /PropertyPaneListPicker/.test(item.message)), JSON.stringify(diagnostics));
  for (const line of [5, 6, 7, 8]) assert.ok(diagnostics.some(item => item.code === 'TS2307' && item.line === line), JSON.stringify(diagnostics));
  assert.deepEqual(check(fixture, custom(["import { PropertyPaneListPicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';\nconst pane = PropertyPaneListPicker('items');"])), []);
});

test('uses installed peer declarations rather than a permissive virtual peer', () => {
  const diagnostics = check(fixture, negative('real-peer'));
  assert.ok(diagnostics.some(item => item.code === 'TS2322' && item.line === 5), JSON.stringify(diagnostics));
});

test('rejects private relative, absolute and aliased example imports identically with cold or stale lib', () => {
  const mirror = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-example-imports-'));
  try {
    fs.cpSync(fixture, mirror, { recursive: true });
    fs.symlinkSync(path.join(__dirname, '../node_modules'), path.join(mirror, 'node_modules'), 'dir');
    const specifiers = ['./private', '../src/private', path.join(mirror, 'src/private'),
      '../lib/private', path.join(mirror, 'lib/private'), '@example-private', '@example-stale'];
    const examples = custom(specifiers.map(specifier => `import { privateValue } from ${JSON.stringify(specifier)};`));
    const cold = check(mirror, examples);
    fs.mkdirSync(path.join(mirror, 'lib'));
    fs.writeFileSync(path.join(mirror, 'lib/private.d.ts'), 'export declare const privateValue: string;\n');
    const stale = check(mirror, examples);
    assert.deepEqual(cold.map(item => ({ code: item.code, file: item.file, line: item.line })), [
      { code: 'TS2307', file: 'isolated-0.ts', line: 20 },
      { code: 'TS2307', file: 'isolated-1.ts', line: 20 },
      { code: 'TS2307', file: 'isolated-2.ts', line: 20 },
      { code: 'TS2307', file: 'isolated-3.ts', line: 20 },
      { code: 'TS2307', file: 'isolated-4.ts', line: 20 },
      { code: 'TS2307', file: 'isolated-5.ts', line: 20 },
      { code: 'TS2307', file: 'isolated-6.ts', line: 20 },
    ]);
    assert.deepEqual(stale, cold);
    assert.deepEqual(check(mirror, custom(["import { PropertyPaneListPicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';\nimport type { CSSProperties } from 'react';\nconst peer: CSSProperties = { display: 'block' };\nconst picker = PropertyPaneListPicker('items');"])), []);
  } finally { fs.rmSync(mirror, { recursive: true, force: true }); }
});

test('rejects example triple-slash local path references without gaining stale global declarations', () => {
  const mirror = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-example-references-'));
  try {
    fs.cpSync(fixture, mirror, { recursive: true });
    fs.symlinkSync(path.join(__dirname, '../node_modules'), path.join(mirror, 'node_modules'), 'dir');
    const references = ['../lib/private-global.d.ts', path.join(mirror, 'lib/private-global.d.ts')];
    const examples = custom(references.map(reference => `/// <reference path=${JSON.stringify(reference)} />\nconst example: string = privateGlobal;`));
    const cold = check(mirror, examples);
    fs.mkdirSync(path.join(mirror, 'lib'));
    fs.writeFileSync(path.join(mirror, 'lib/private-global.d.ts'), 'declare const privateGlobal: string;\n');
    const stale = check(mirror, examples);
    assert.deepEqual(cold.map(item => ({ code: item.code, file: item.file, line: item.line })), [
      { code: 'DOC_EXAMPLE_LOCAL_REFERENCE', file: 'isolated-0.ts', line: 20 },
      { code: 'TS2304', file: 'isolated-0.ts', line: 21 },
      { code: 'DOC_EXAMPLE_LOCAL_REFERENCE', file: 'isolated-1.ts', line: 20 },
      { code: 'TS2304', file: 'isolated-1.ts', line: 21 },
    ]);
    assert.deepEqual(stale, cold);
  } finally { fs.rmSync(mirror, { recursive: true, force: true }); }
});

test('rejects local triple-slash type references while installed peer type references stay real', () => {
  const mirror = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-example-type-references-'));
  try {
    fs.cpSync(fixture, mirror, { recursive: true });
    fs.symlinkSync(path.join(__dirname, '../node_modules'), path.join(mirror, 'node_modules'), 'dir');
    const references = ['../lib/private-global', path.join(mirror, 'lib/private-global')];
    const examples = custom(references.map(reference => `/// <reference types=${JSON.stringify(reference)} />\nconst example: string = privateGlobal;`));
    const cold = check(mirror, examples);
    fs.mkdirSync(path.join(mirror, 'lib'));
    fs.writeFileSync(path.join(mirror, 'lib/private-global.d.ts'), 'declare const privateGlobal: string;\n');
    const stale = check(mirror, examples);
    assert.deepEqual(cold.map(item => ({ code: item.code, file: item.file, line: item.line })), [
      { code: 'DOC_EXAMPLE_LOCAL_REFERENCE', file: 'isolated-0.ts', line: 20 },
      { code: 'TS2304', file: 'isolated-0.ts', line: 21 },
      { code: 'DOC_EXAMPLE_LOCAL_REFERENCE', file: 'isolated-1.ts', line: 20 },
      { code: 'TS2304', file: 'isolated-1.ts', line: 21 },
    ]);
    assert.deepEqual(stale, cold);
    assert.deepEqual(check(mirror, custom(['/// <reference types="react" />\nconst style: React.CSSProperties = { display: "block" };'])), []);
  } finally { fs.rmSync(mirror, { recursive: true, force: true }); }
});

test('isolates examples with only a harmless module boundary and never supplies missing inputs', () => {
  const diagnostics = check(fixture, custom(['const name = 1;', 'const name = 2;', 'name.toFixed();']));
  assert.deepEqual(diagnostics.map(item => ({ code: item.code, file: item.file, line: item.line })), [
    { code: 'TS2339', file: 'isolated-2.ts', line: 20 },
  ]);
  assert.match(diagnostics[0].message, /void/); // real DOM global name, never the other example's number
  assert.deepEqual(check(fixture, custom(['const pageIndex = 0;', 'pageIndex.toFixed();'])).map(item => item.code), ['TS2304']);
});

test('forbids suppression comments and preserves TypeScript unused-expect-error diagnostics', () => {
  const diagnostics = check(fixture, negative('suppression'));
  assert.deepEqual(diagnostics.filter(item => item.code === 'DOC_EXAMPLE_SUPPRESSION').map(item => item.line), [4, 6]);
  assert.ok(diagnostics.some(item => item.code === 'TS2578' && item.line === 4), JSON.stringify(diagnostics));
  assert.deepEqual(check(fixture, negative('no-check')).map(item => ({ code: item.code, line: item.line })), [
    { code: 'DOC_EXAMPLE_SUPPRESSION', line: 4 },
  ]);
  assert.deepEqual(check(fixture, custom(["const text = '@ts-ignore';"])), []);
});

test('does not confuse regular expressions or JSX text with error suppression comments', () => {
  const input = custom(['const pattern = /[/*] @ts-ignore/;', "import * as React from 'react'; const view = <div>// @ts-ignore</div>;"]);
  input.examples[1].language = 'tsx';
  assert.deepEqual(check(fixture, input), []);
});

test('checks the same source mirror before and after stale lib appears without emitting or executing examples', () => {
  const mirror = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-examples-'));
  try {
    fs.cpSync(fixture, mirror, { recursive: true });
    fs.symlinkSync(path.join(__dirname, '../node_modules'), path.join(mirror, 'node_modules'), 'dir');
    const mirroredModel = buildDocumentationModel(buildApiGraph(mirror, config), config);
    const unsafe = custom(["Object.defineProperty(globalThis, '__documentation_example_executed__', { value: true });\nthrow new Error('Example code must never execute');"]);
    const before = files(mirror);
    assert.deepEqual(check(mirror, mirroredModel), []);
    assert.deepEqual(check(mirror, unsafe), []);
    assert.equal(Object.hasOwn(globalThis, '__documentation_example_executed__'), false);
    assert.deepEqual(files(mirror), before);
    const expected = check(mirror, negative('illegal-imports', mirroredModel));
    fs.mkdirSync(path.join(mirror, 'lib'), { recursive: true });
    fs.writeFileSync(path.join(mirror, 'lib/index.d.ts'), 'export declare const PropertyPaneListPicker: number;\nexport declare const Pagination: never;\n');
    fs.writeFileSync(path.join(mirror, 'lib/pagination.d.ts'), 'export declare const Pagination: number;\n');
    fs.writeFileSync(path.join(mirror, 'lib/absent.d.ts'), 'export declare const absent: number;\n');
    const withStale = files(mirror);
    assert.deepEqual(check(mirror, mirroredModel), []);
    assert.deepEqual(check(mirror, negative('illegal-imports', mirroredModel)), expected);
    assert.deepEqual(files(mirror), withStale);
  } finally { fs.rmSync(mirror, { recursive: true, force: true }); }
});

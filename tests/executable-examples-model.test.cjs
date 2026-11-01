const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const { buildApiGraph } = require('../scripts/documentation-model.cjs');
const { buildDocumentationModel } = require('../scripts/documentation-contract.cjs');

const fixture = path.join(__dirname, 'fixtures/executable-examples');
const config = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' };
let graph;
let base;
test.before(() => {
  graph = buildApiGraph(fixture, config);
  base = buildDocumentationModel(graph, config);
});
function collect(model = base) {
  assert.ok(fs.existsSync(path.join(__dirname, '../scripts/executable-examples-model.cjs')),
    'executable example projection must exist');
  return require('../scripts/executable-examples-model.cjs').collectExecutableExamples(fixture, model, config);
}
function replaceDoc(rawDoc, name = 'Widget') {
  const copy = structuredClone(graph);
  changeDoc(copy, rawDoc, name);
  return buildDocumentationModel(copy, config);
}
function changeDoc(copy, rawDoc, name) {
  const entity = copy.entities.find(item => item.qualifiedName === name);
  const oldDoc = entity.rawDoc;
  entity.rawDoc = rawDoc;
  entity.docSource = { file: 'cases/input.jsdoc', line: 1 };
  for (const contract of [...entity.signatures, ...copy.bindings.filter(binding => binding.entityKey === entity.key)]) {
    if (contract.rawDoc === oldDoc) {
      contract.rawDoc = rawDoc;
      contract.docSource = entity.docSource;
    }
  }
}
function snippet(code, title = 'Scenario', language = 'tsx') {
  return `/** A scenario.\n * @example ${title}\n * \`\`\`${language}\n${code.split('\n').map(line => ` * ${line}`).join('\n')}\n * \`\`\`\n */`;
}
const component = "import * as React from 'react';\nexport function Demo(): React.ReactElement { return <div />; }";
const metadata = "export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'Visible content.' };";
function invalid(code, expected, language = 'tsx') {
  const result = collect(replaceDoc(snippet(code, 'Scenario', language)));
  assert.equal(result.examples.length, 0);
  assert.ok(result.diagnostics.some(item => item.code === expected), JSON.stringify(result.diagnostics));
  for (const diagnostic of result.diagnostics) {
    assert.equal(diagnostic.file, 'cases/input.jsdoc');
    assert.equal(diagnostic.symbol, 'src/widget.tsx#Widget');
    assert.ok(diagnostic.line >= 4 && diagnostic.line <= code.split('\n').length + 3, JSON.stringify(diagnostic));
  }
  return result;
}

test('projects only the named executable contract and retains physical provenance', () => {
  const result = collect();
  assert.deepEqual(result.diagnostics, []);
  assert.equal(result.examples.length, 1);
  const example = result.examples[0];
  assert.equal(example.ownerKey, 'src/widget.tsx#Widget');
  assert.equal(example.title, 'Basic view');
  assert.equal(example.groupRoute, 'docs/api/controls/widget.md');
  assert.equal(example.groupKey, 'controls');
  assert.equal(example.featureKey, 'src/widget.tsx#Widget');
  assert.equal(example.featureTitle, 'Widget');
  assert.deepEqual(example.publicNames, ['AliasWidget', 'Widget']);
  assert.deepEqual(example.source, { file: 'src/widget.tsx', line: 8 });
  assert.match(example.id, /^[a-f0-9]{64}$/);
  assert.deepEqual(example.metadata, { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'A widget is visible.' });
  assert.match(example.code, /MisleadingFirst/);
});

test('leaves ordinary, default-only and type usage examples outside the catalog', () => {
  const ordinary = base.examples.filter(example => example.title !== 'Basic view');
  assert.equal(collect({ ...base, examples: ordinary }).examples.length, 0);
  assert.deepEqual(collect({ ...base, examples: ordinary }).diagnostics, []);
  const defaultNamed = collect(replaceDoc(snippet(component.replace('export function Demo', 'export default function Demo'))));
  assert.deepEqual(defaultNamed, { examples: [], diagnostics: [] });
});

test('keeps scenario IDs stable across physical line changes and added public aliases', () => {
  const original = collect().examples[0];
  const changed = structuredClone(base);
  changed.examples.forEach(example => { example.source.line += 30; });
  changed.graph.bindings.push({ ...changed.graph.bindings.find(binding => binding.exportedName === 'Widget'), exportedName: 'NewAlias' });
  const result = collect(changed);
  assert.equal(result.examples.length, 1);
  assert.equal(result.examples[0].id, original.id);
  assert.deepEqual(result.examples[0].publicNames, ['AliasWidget', 'NewAlias', 'Widget']);
  assert.equal(result.examples[0].featureKey, original.featureKey);
});

test('uses the canonical public compound spelling rather than a source-qualified declaration name', () => {
  const changed = structuredClone(base);
  changed.graph.bindings.push({ ...changed.graph.bindings.find(binding => binding.exportedName === 'Widget'),
    exportedName: 'Toolkit', memberPath: ['Widget'] });
  const result = collect(changed);
  assert.equal(result.examples.length, 1);
  assert.equal(result.examples[0].featureTitle, 'Toolkit.Widget');
  assert.equal(result.examples[0].featureKey, 'src/widget.tsx#Widget');
});

test('groups two scenarios of one API together and uses the owner family rather than snippet functions', () => {
  const copy = structuredClone(graph);
  changeDoc(copy,
    snippet(component + '\n' + metadata, 'First').replace(' */', '') +
    snippet(component + '\n' + metadata, 'Second').replace('/** A scenario.\n', ''), 'Widget');
  changeDoc(copy, snippet(component + '\n' + metadata, 'Clock').replace('/** A scenario.', '/** A scenario.\n * @docPage hooks/time'), 'tick');
  const result = collect(buildDocumentationModel(copy, config));
  assert.deepEqual(result.diagnostics, []);
  assert.equal(result.examples.length, 3);
  const first = result.examples.find(example => example.title === 'First');
  const second = result.examples.find(example => example.title === 'Second');
  const clock = result.examples.find(example => example.title === 'Clock');
  assert.equal(first.featureKey, second.featureKey);
  assert.notEqual(first.id, second.id);
  assert.equal(clock.groupKey, 'hooks');
  assert.equal(clock.featureTitle, 'tick');
});

test('accepts structural optional region/reset and public Shell ports', () => {
  for (const name of ['regions', 'shell']) {
    const model = replaceDoc(fs.readFileSync(path.join(fixture, `cases/${name}.jsdoc`), 'utf8'));
    const result = collect(model);
    assert.deepEqual(result.diagnostics, [], name);
    assert.equal(result.examples.length, 1, name);
  }
  const aliased = `${component.replace('export function Demo', 'function Actual')}\nexport { Actual as Demo };\n${metadata}`;
  assert.deepEqual(collect(replaceDoc(snippet(aliased))).diagnostics, []);
});

test('rejects incomplete named contracts while ignoring names inside nested scopes', () => {
  invalid(component, 'EXECUTABLE_INCOMPLETE_CONTRACT');
  invalid(metadata, 'EXECUTABLE_INCOMPLETE_CONTRACT');
  const nested = `function ordinary() { const demo = {}; function Demo() { return 1; } }`;
  assert.deepEqual(collect(replaceDoc(snippet(nested))), { examples: [], diagnostics: [] });
});

for (const declaration of ['export type { Actual as Demo };', 'export { type Actual as Demo };']) {
  test(`rejects erased component export: ${declaration}`, () => {
    const code = `${component.replace('export function Demo', 'function Actual')}\n${declaration}\n${metadata}`;
    const result = invalid(code, 'EXECUTABLE_RUNTIME_EXPORT');
    assert.ok(result.diagnostics.some(item => item.code === 'EXECUTABLE_RUNTIME_EXPORT' && item.line === 6));
  });
}

for (const declaration of ['export type { data as demo };', 'export { type data as demo };']) {
  test(`rejects erased metadata export: ${declaration}`, () => {
    const code = `${component}\n${metadata.replace('export const demo', 'const data')}\n${declaration}`;
    const result = invalid(code, 'EXECUTABLE_RUNTIME_EXPORT');
    assert.ok(result.diagnostics.some(item => item.code === 'EXECUTABLE_RUNTIME_EXPORT' && item.line === 7));
  });
}

for (const declaration of [
  'export declare function Demo(): React.ReactElement;',
  'export declare const Demo: () => React.ReactElement;',
  'declare function Actual(): React.ReactElement;\nexport { Actual as Demo };',
]) {
  test(`rejects ambient component without a runtime implementation: ${declaration}`, () => {
    invalid(`import * as React from 'react';\n${declaration}\n${metadata}`, 'EXECUTABLE_RUNTIME_EXPORT');
  });
}

test('retains real runtime aliases for both contract exports and implemented overloads', () => {
  const code = `import * as React from 'react';
function Actual(): React.ReactElement;
function Actual(): React.ReactElement { return <div />; }
${metadata.replace('export const demo', 'const data')}
export { Actual as Demo, data as demo };`;
  const result = collect(replaceDoc(snippet(code)));
  assert.deepEqual(result.diagnostics, []);
  assert.equal(result.examples.length, 1);
});

test('requires TSX for executable components', () => {
  invalid(component.replace('return <div />', 'return React.createElement("div")') + '\n' + metadata, 'EXECUTABLE_LANGUAGE', 'ts');
});

test('rejects missing and duplicate local titles and type-only owners', () => {
  const code = component + '\n' + metadata;
  const missing = collect(replaceDoc(snippet(code, '')));
  assert.ok(missing.diagnostics.some(item => item.code === 'EXECUTABLE_TITLE'));
  const repeated = snippet(code, 'Same').replace(' */', '') + snippet(code, 'Same').replace('/** A scenario.\n', '');
  const duplicate = collect(replaceDoc(repeated));
  assert.equal(duplicate.examples.length, 0);
  assert.ok(duplicate.diagnostics.some(item => item.code === 'EXECUTABLE_DUPLICATE_TITLE'));
  const typeOnly = collect(replaceDoc(snippet(code), 'WidgetOptions'));
  assert.ok(typeOnly.diagnostics.some(item => item.code === 'EXECUTABLE_RUNTIME_OWNER'));
  assert.equal(typeOnly.examples.length, 1, 'the unrelated Widget scenario is retained');
});

test('requires an executable title to be unique among all examples of its owner', () => {
  const repeated = snippet(component + '\n' + metadata, 'Same').replace(' */', '') +
    snippet('export const ordinary = 1;', 'Same').replace('/** A scenario.\n', '');
  const result = collect(replaceDoc(repeated));
  assert.equal(result.examples.length, 0);
  assert.ok(result.diagnostics.some(item => item.code === 'EXECUTABLE_DUPLICATE_TITLE'));
});

test('rejects expressions, spreads, unknown metadata and missing or empty observations', () => {
  const values = [
    "const host = 'none'; export const demo = { host, kind: 'snapshot', contentPlacement: 'subject', observe: 'Visible.' };",
    "export const demo = { ...{}, host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'Visible.' };",
    metadata.replace("observe: 'Visible content.'", "observe: 'Visible content.', extra: 'unknown'"),
    metadata.replace("observe: 'Visible content.'", "observe: '   '"),
    metadata.replace(", observe: 'Visible content.'", ''),
    metadata.replace("kind: 'snapshot'", "kind: 'execute'"),
    metadata.replace("host: 'none'", "host: 'tenant'"),
    metadata.replace("contentPlacement: 'subject'", "contentPlacement: 'other'"),
    metadata.replace("host: 'none'", "host: 'none', host: 'webPart'"),
  ];
  for (const value of values) invalid(component + '\n' + value, 'EXECUTABLE_METADATA');
});

test('checks actual synchronous React result types and mountable optional props', () => {
  const declarations = [
    'export function Demo(): string { return "not a React element"; }',
    'export async function Demo(): Promise<React.ReactElement> { return <div />; }',
    'export const Demo = 42;',
    'export function Demo(props: { required: string }): React.ReactElement { return <div />; }',
    'export function Demo(props: { resetRevision: number }): React.ReactElement { return <div />; }',
    'export function Demo(props: { resetRevision?: string }): React.ReactElement { return <div />; }',
    'export function Demo(props: { render?: (regions: { children?: number }) => React.ReactElement }): React.ReactElement { return <div />; }',
    'export function Demo(props?: { render?: (regions: { children?: number }) => React.ReactElement }): React.ReactElement { return <div />; }',
    'export function Demo(): any { return 42; }',
  ];
  for (const declaration of declarations) invalid("import * as React from 'react';\n" + declaration + '\n' + metadata, 'EXECUTABLE_COMPONENT');
});

test('rejects incoherent Shell adapters and incompatible Shell ports', () => {
  for (const patch of [
    value => value.replace("host: 'webPart'", "host: 'none'"),
    value => value.replace("contentPlacement: 'configuration'", "contentPlacement: 'subject'"),
    value => value.replace("hostAdapter: 'webPartShell'", "hostAdapter: 'other'"),
  ]) {
    const result = collect(replaceDoc(patch(fs.readFileSync(path.join(fixture, 'cases/shell.jsdoc'), 'utf8'))));
    assert.equal(result.examples.length, 0);
    assert.ok(result.diagnostics.some(item => item.code === 'EXECUTABLE_METADATA'));
  }
  const wrong = fs.readFileSync(path.join(fixture, 'cases/shell.jsdoc'), 'utf8')
    .replace('shell?: DemoShellPort', 'shell?: { presentation: string }');
  const result = collect(replaceDoc(wrong));
  assert.equal(result.examples.length, 0);
  assert.ok(result.diagnostics.some(item => item.code === 'EXECUTABLE_COMPONENT'));
});

test('rejects required Shell props and narrowed presentation update parameters', () => {
  const shell = fs.readFileSync(path.join(fixture, 'cases/shell.jsdoc'), 'utf8');
  for (const rawDoc of [shell.replace('shell?: DemoShellPort', 'shell: DemoShellPort'),
    shell.replace("patch: Partial<DemoShellPort['presentation']>", "patch: { appearance: 'plain' }")]) {
    const result = collect(replaceDoc(rawDoc));
    assert.equal(result.examples.length, 0);
    assert.ok(result.diagnostics.some(item => item.code === 'EXECUTABLE_COMPONENT'));
  }
});

test('reports snippet compiler errors at physical source lines without evaluating metadata', () => {
  const result = invalid(component + '\nconst broken: number = "text";\n' + metadata, 'TS2322');
  assert.ok(result.diagnostics.some(item => item.code === 'TS2322' && item.line === 6));
  invalid(component + '\nexport const demo = (() => { throw new Error("must never run"); })();', 'EXECUTABLE_METADATA');
});

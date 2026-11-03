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

test('accepts old, optional subset and full bounded Shell presentation consumers', () => {
  const shell = fs.readFileSync(path.join(fixture, 'cases/shell.jsdoc'), 'utf8');
  for (const fields of [undefined, "'surface' | 'dir'", "'surface' | 'dir' | 'strings' | 'titleEditorStrings' | 'action'"]) {
    const raw = fields ? shell.replace("'padding'>>", "'padding'>> & Pick<WebPartShellProps<IWebPartShellProperties>, " + fields + ">") : shell;
    const result = collect(replaceDoc(raw));
    assert.deepEqual(result.diagnostics, [], fields);
    assert.equal(result.examples.length, 1);
  }
});

test('rejects unsafe, required, narrowed and indexed Shell extensions and patch signatures', () => {
  const shell = fs.readFileSync(path.join(fixture, 'cases/shell.jsdoc'), 'utf8');
  const full = shell.replace("'padding'>>", "'padding'>> & Pick<WebPartShellProps<IWebPartShellProperties>, 'surface' | 'dir' | 'strings' | 'titleEditorStrings' | 'action'>");
  const invalid = [
    shell.replace("'padding'>>", "'padding'>> & { extra?: string }"),
    shell.replace("'padding'>>", "'padding'>> & { surface: 'auto' | 'section' | 'application' }"),
    shell.replace("'padding'>>", "'padding'>> & { surface?: 'auto' }"),
    shell.replace("'padding'>>", "'padding'>> & { dir?: any }"),
    shell.replace("'padding'>>", "'padding'>> & { strings?: unknown }"),
    shell.replace("'padding'>>", "'padding'>> & { strings?: { titleEditorLabel?: any; titlePlaceholder?: string } }"),
    shell.replace("'padding'>>", "'padding'>> & { action?: never }"),
    shell.replace("'padding'>>", "'padding'>> & { titleEditorStrings?: { bold?: string; colorNames?: { red?: string }; formattingShortcut?: (label: any, shortcut: string) => string } }"),
    shell.replace("'padding'>>", "'padding'>> & { [key: string]: unknown }"),
    full.replace("patch: Partial<DemoShellPort['presentation']>", "patch: { surface?: 'auto' }"),
    full.replace("patch: Partial<DemoShellPort['presentation']>", "patch: Partial<DemoShellPort['presentation']> & { surface: 'auto' }"),
    full.replace("patch: Partial<DemoShellPort['presentation']>", "patch: any"),
    full.replace("patch: Partial<DemoShellPort['presentation']>", "patch: Partial<DemoShellPort['presentation']> & { extra?: string }"),
  ];
  for (const raw of invalid) {
    const result = collect(replaceDoc(raw));
    assert.equal(result.examples.length, 0, raw);
    assert.ok(result.diagnostics.some(item => item.code === 'EXECUTABLE_COMPONENT'), JSON.stringify(result.diagnostics));
  }
});

const reviewedUnsafeShellShapes = [
  ['empty naming dictionary', 'strings', '{}'],
  ['missing naming member', 'strings', '{ titleEditorLabel?: string }'],
  ['unsupported construct signature', 'strings', '{ titleEditorLabel?: string; titlePlaceholder?: string; new(): {} }'],
  ['unsafe naming index', 'strings', '{ titleEditorLabel?: string; titlePlaceholder?: string; [key: string]: any }'],
  ['unsafe nested color index', 'titleEditorStrings', '{ bold?: string; colorNames?: { red?: string; [key: string]: any }; formattingShortcut?: (label: string, shortcut: string) => string }'],
];
for (const [label, field, shape] of reviewedUnsafeShellShapes) for (const boundary of ['presentation', 'patch']) {
  test(`rejects reviewed ${label} at Shell ${boundary} boundary`, () => {
    const shell = fs.readFileSync(path.join(fixture, 'cases/shell.jsdoc'), 'utf8');
    const extension = boundary === 'presentation' ? `{ ${field}?: ${shape} }` : `Pick<WebPartShellProps<IWebPartShellProperties>, '${field}'>`;
    let raw = shell.replace("'padding'>>", "'padding'>> & " + extension);
    if (boundary === 'patch') raw = raw.replace("patch: Partial<DemoShellPort['presentation']>",
      `patch: Partial<Omit<DemoShellPort['presentation'], '${field}'>> & { ${field}?: ${shape} }`);
    const result = collect(replaceDoc(raw));
    assert.equal(result.examples.length, 0, JSON.stringify(result.diagnostics));
    assert.ok(result.diagnostics.some(item => item.code === 'EXECUTABLE_COMPONENT'));
  });
}

test('reports snippet compiler errors at physical source lines without evaluating metadata', () => {
  const result = invalid(component + '\nconst broken: number = "text";\n' + metadata, 'TS2322');
  assert.ok(result.diagnostics.some(item => item.code === 'TS2322' && item.line === 6));
  invalid(component + '\nexport const demo = (() => { throw new Error("must never run"); })();', 'EXECUTABLE_METADATA');
});

const nativeDoc = fs.readFileSync(path.join(fixture, 'cases/native.jsdoc'), 'utf8');
test('accepts the complete optional function-valued native port without private canonical inputs', () => {
  const result = collect(replaceDoc(nativeDoc));
  assert.deepEqual(result.diagnostics, []);
  assert.equal(result.examples.length, 1);
  assert.equal(result.examples[0].metadata.hostAdapter, 'propertyPane');
});

for (const [name, before, after] of [
  ['mutable environment function', 'readonly refresh: () => void;', 'refresh: () => void;'],
  ['required port', 'propertyPane?:', 'propertyPane:'],
  ['method bivariance', 'readonly refresh: () => void;', 'refresh(): void;'],
  ['extra environment', 'readonly refresh: () => void;', 'readonly refresh: () => void; readonly extra: string;'],
  ['extra create argument', '(host: NativePropertyPaneEnvironment<P>)', '(host: NativePropertyPaneEnvironment<P>, required: string)'],
  ['narrow event callback', 'oldValue: unknown', 'oldValue: string'],
  ['any callback', 'newValue: unknown', 'newValue: any'],
  ['unknown environment', 'readonly instance: BaseClientSideWebPart<P>', 'readonly instance: unknown'],
  ['wrong factory return', '() => IPropertyPaneConfiguration', '() => string'],
  ['wrong preparation discriminator', "status: 'prepared'", "status: 'ready'"],
  ['wrong lease return', 'readonly release: () => void', 'readonly release: () => number'],
  ['extra required member', 'readonly dispose: () => void;', 'readonly dispose: () => void; readonly start: () => void;'],
  ['any source', 'pickerSites?: string[]', 'pickerSites?: any[]'],
  ['unknown source', 'pickerSites?: string[]', 'pickerSites?: unknown[]'],
  ['source index signature', 'pickerSites?: string[]', '[key: string]: string[] | undefined'],
  ['unsafe cast', "return <div>{props.propertyPane ? 'Ready' : 'Host missing'}</div>;", "return <div>{String((props.propertyPane as unknown) !== undefined)}</div>;"],
]) test(`native structural contract rejects ${name}`, () => {
  assert.ok(nativeDoc.includes(before), name);
  const result = collect(replaceDoc(nativeDoc.replace(before, after)));
  assert.equal(result.examples.length, 0, name);
  assert.ok(result.diagnostics.some(item => item.code === 'EXECUTABLE_COMPONENT'), JSON.stringify(result.diagnostics));
});
test('native metadata cannot be used for non-WebPart or subject content and native port cannot hide in ordinary metadata', () => {
  for (const [before, after] of [["host: 'webPart'", "host: 'none'"], ["contentPlacement: 'configuration'", "contentPlacement: 'subject'"]]) {
    assert.ok(collect(replaceDoc(nativeDoc.replace(before, after))).diagnostics.some(item => item.code === 'EXECUTABLE_METADATA'));
  }
  assert.ok(collect(replaceDoc(nativeDoc.replace(", hostAdapter: 'propertyPane'", ''))).diagnostics.some(item => item.code === 'EXECUTABLE_COMPONENT'));
});

const topCode = `import * as React from 'react';
type Regions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
export function Demo(props: { render?: (regions: Regions) => React.ReactElement; resetRevision?: number; applicationCustomizerTop?: { readonly placeholder: 'Top' } } = {}): React.ReactElement { return <div>{props.applicationCustomizerTop?.placeholder ?? 'Native Top required'}</div>; }
export const demo = { host: 'applicationCustomizer', hostAdapter: 'applicationCustomizerTop', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Actual native Top provider.' };`;
test('accepts the exact optional readonly Top signal without exposing native capabilities', () => {
  const result = collect(replaceDoc(snippet(topCode)));
  assert.deepEqual(result.diagnostics, []);
  assert.equal(result.examples[0].metadata.hostAdapter, 'applicationCustomizerTop');
});
for (const [name, before, after, code] of [
  ['missing adapter', ", hostAdapter: 'applicationCustomizerTop'", '', 'EXECUTABLE_METADATA'],
  ['wrong host', "host: 'applicationCustomizer'", "host: 'webPart'", 'EXECUTABLE_METADATA'],
  ['cross pair', "hostAdapter: 'applicationCustomizerTop'", "hostAdapter: 'propertyPane'", 'EXECUTABLE_METADATA'],
  ['unknown adapter', "hostAdapter: 'applicationCustomizerTop'", "hostAdapter: 'nativeTop'", 'EXECUTABLE_METADATA'],
  ['subject', "contentPlacement: 'configuration'", "contentPlacement: 'subject'", 'EXECUTABLE_METADATA'],
  ['required signal', 'applicationCustomizerTop?:', 'applicationCustomizerTop:', 'EXECUTABLE_COMPONENT'],
  ['mutable signal', "readonly placeholder: 'Top'", "placeholder: 'Top'", 'EXECUTABLE_COMPONENT'],
  ['broad signal', "readonly placeholder: 'Top'", 'readonly placeholder: string', 'EXECUTABLE_COMPONENT'],
  ['extra instance', "readonly placeholder: 'Top'", "readonly placeholder: 'Top'; readonly instance?: object", 'EXECUTABLE_COMPONENT'],
  ['extra context', "readonly placeholder: 'Top'", "readonly placeholder: 'Top'; readonly context?: object", 'EXECUTABLE_COMPONENT'],
  ['extra DOM', "readonly placeholder: 'Top'", "readonly placeholder: 'Top'; readonly domElement?: HTMLElement", 'EXECUTABLE_COMPONENT'],
  ['extra dispose', "readonly placeholder: 'Top'", "readonly placeholder: 'Top'; dispose(): void", 'EXECUTABLE_COMPONENT'],
  ['unsafe signal', "{ readonly placeholder: 'Top' }", 'any', 'EXECUTABLE_COMPONENT'],
  ['unknown signal', "{ readonly placeholder: 'Top' }", 'unknown', 'EXECUTABLE_COMPONENT'],
  ['never signal', "{ readonly placeholder: 'Top' }", 'never', 'EXECUTABLE_COMPONENT'],
  ['missing signal', "; applicationCustomizerTop?: { readonly placeholder: 'Top' }", '', 'EXECUTABLE_COMPONENT'],
]) test(`Top contract rejects ${name}`, () => invalid(topCode.replace(before, after), code));
test('ordinary metadata cannot hide a Top signal', () => {
  invalid(topCode.replace("host: 'applicationCustomizer', hostAdapter: 'applicationCustomizerTop'", "host: 'webPart'"), 'EXECUTABLE_COMPONENT');
});

for (const extra of ['instance?: object', 'context?: object', 'domElement?: HTMLElement', 'dispose?: () => void']) test(`Top Demo rejects an extra outer ${extra}`, () => {
  invalid(topCode.replace('resetRevision?: number;', `resetRevision?: number; ${extra};`), 'EXECUTABLE_COMPONENT');
});
test('Top rejects erased and async entrypoints before common loader erasure', () => {
  invalid(topCode.replace('export function Demo', 'export async function Demo').replace('): React.ReactElement {', '): Promise<React.ReactElement> {'), 'EXECUTABLE_COMPONENT');
  invalid(topCode.replace('export function Demo', 'function Actual') + '\nexport type { Actual as Demo };', 'EXECUTABLE_RUNTIME_EXPORT');
});

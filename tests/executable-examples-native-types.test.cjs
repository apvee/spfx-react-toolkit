const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { renderExecutableExamples, executableRoot } = require('../scripts/executable-examples-render.cjs');
const root = path.resolve(__dirname, '..');
const template = fs.readFileSync(path.join(__dirname, 'fixtures/executable-examples/cases/native.jsdoc'), 'utf8')
  .split(' * ```tsx\n')[1].split(' * ```')[0].split('\n').map(line => line.replace(/^ \* ?/, '')).join('\n')
  .replace("import * as React from 'react';", "import * as React from 'react';\nimport type { SitePickerValue, ListPickerValue, FieldPickerValue, ViewPickerValue, ContentTypePickerValue, ListItemPickerValue, PrincipalPickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';\nimport type { FileFolderPickerValue } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';");
const id = 'a'.repeat(64);
function emitted(code) {
  return renderExecutableExamples([{ id, ownerKey: 'source#Native', title: 'Native', groupRoute: 'docs/api/property-pane/pickers.md', groupKey: 'property-pane', featureKey: 'source#Native', featureTitle: 'Native', publicNames: ['Native'], source: { file: 'source.ts', line: 1 }, code, metadata: { host: 'webPart', kind: 'read', contentPlacement: 'configuration', hostAdapter: 'propertyPane', observe: 'Native fixture.' } }]);
}
function compile(domain) {
  const artifacts = emitted(domain.startsWith('type SourceProperties') ? template.replace('interface SourceProperties { pickerSites?: string[]; }', domain) : template.replace('pickerSites?: string[];', domain));
  const virtual = new Map(Object.entries(artifacts).map(([file, text]) => [path.join(root, file), text]));
  const config = ts.readConfigFile(path.join(root, 'apps/spfx-react-toolkit-test/tsconfig.json'), ts.sys.readFile);
  const { options } = ts.parseJsonConfigFileContent(config.config, ts.sys, path.join(root, 'apps/spfx-react-toolkit-test'));
  const host = ts.createCompilerHost({ ...options, noEmit: true }); const read = host.readFile.bind(host), exists = host.fileExists.bind(host);
  host.readFile = file => virtual.get(file) ?? read(file); host.fileExists = file => virtual.has(file) || exists(file);
  const entry = path.join(root, executableRoot, 'catalogue.ts');
  const program = ts.createProgram([entry], { ...options, noEmit: true, noUnusedLocals: false }, host);
  return { artifacts, diagnostics: ts.getPreEmitDiagnostics(program).filter(d => d.file && virtual.has(d.file.fileName)).map(d => ({ code: d.code, message: ts.flattenDiagnosticMessageText(d.messageText, '\n') })) };
}
const integrated = 'pickerSites?: SitePickerValue[]; pickerLists?: ListPickerValue[]; pickerFields?: FieldPickerValue[]; pickerViews?: ViewPickerValue[]; pickerContentTypes?: ContentTypePickerValue[]; pickerItems?: ListItemPickerValue[]; pickerPrincipals?: PrincipalPickerValue[]; fileFolderValues?: FileFolderPickerValue[];';
for (const [name, domain] of [['single', 'pickerSites?: SitePickerValue[];'], ['full integrated domains with extra actual host keys', integrated]]) test(`emitted actual-app wrapper accepts ${name} before common erasure`, () => {
  const result = compile(domain); assert.deepEqual(result.diagnostics, []);
  const catalogue = result.artifacts[executableRoot + '/catalogue.ts'];
  assert.ok(catalogue.indexOf('const CheckedDemo: NativeCompatible') < catalogue.indexOf('return { Demo: (props: ExecutableExampleProps)'));
  assert.match(catalogue, /import type \{ ISpFxReactToolkitTestWebPartProps \}/); assert.match(catalogue, /CheckedDemo\(props\)/);
});
for (const [name, domain] of [
  ['untyped any array', 'pickerSites?: any[];'],
  ['untyped unknown array', 'pickerSites?: unknown[];'],
  ['union branch-only missing actual key', 'type SourceProperties = { pickerSites?: SitePickerValue[] } | { pickerSites?: SitePickerValue[]; nonexistent?: string };'],
  ['wrong array', 'pickerSites?: string[];'],
  ['narrow element', 'pickerSites?: (SitePickerValue & { text: string })[];'],
  ['mixed array', 'pickerSites?: (SitePickerValue | ListPickerValue)[];'],
  ['readonly array', 'pickerSites?: readonly SitePickerValue[];'],
  ['required source property', 'pickerSites: SitePickerValue[];'],
  ['required source even with explicit undefined', 'pickerSites: SitePickerValue[] | undefined;'],
  ['missing actual key', 'nonexistent?: SitePickerValue[];'],
  ['callback domain', 'description?: (value: string) => void;'],
]) test(`emitted actual-app wrapper rejects ${name}`, () => {
  const result = compile(domain);
  assert.ok(result.diagnostics.some(d => d.code === 2322 && /never/.test(d.message)), JSON.stringify(result.diagnostics));
});

test('canonical native validation accepts actual public full integrated array domains without importing the app proof', () => {
  const config = require('../scripts/documentation.config.cjs');
  const graph = require('../scripts/documentation-model.cjs').buildApiGraph(root, config);
  const model = require('../scripts/documentation-contract.cjs').buildDocumentationModel(graph, config);
  const original = model.examples.find(example => example.title === 'Select native site references'); assert.ok(original);
  const code = template.replace('pickerSites?: string[];', integrated);
  const result = require('../scripts/executable-examples-model.cjs').collectExecutableExamples(root, { ...model, examples: [{ ...original, code }] }, config);
  assert.deepEqual(result.diagnostics, []); assert.equal(result.examples.length, 1);
});

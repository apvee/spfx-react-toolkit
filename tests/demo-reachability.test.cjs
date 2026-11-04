const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { spawnSync } = require('node:child_process');
const files = require('./fixtures/demo-reachability/files.cjs');
const auditFile = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/scripts/verification/verify-demo-reachability.cjs');
const app = 'apps/spfx-react-toolkit-test/';
function fixture(changes = {}) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'demo-reachability-'));
  fs.symlinkSync(path.resolve(__dirname, '../node_modules'), path.join(root, 'node_modules'), 'dir');
  for (const [name, content] of Object.entries({ ...files, ...changes })) {
    if (content === undefined) continue;
    fs.mkdirSync(path.dirname(path.join(root, name)), { recursive: true });
    fs.writeFileSync(path.join(root, name), content);
  }
  return root;
}
function snapshot(root) {
  const result = {};
  function walk(folder) {
    for (const entry of fs.readdirSync(folder, { withFileTypes: true })) {
      if (entry.isSymbolicLink()) continue;
      const file = path.join(folder, entry.name);
      if (entry.isDirectory()) walk(file);
      else result[path.relative(root, file)] = [fs.readFileSync(file).toString('base64'), fs.statSync(file).mtimeMs];
    }
  }
  walk(root);
  return result;
}
function check(changes, category, match) {
  const root = fixture(changes);
  try {
    const before = snapshot(root);
    const report = require(auditFile).analyzeDemoReachability(root);
    assert.deepEqual(Object.keys(report).sort(), ['roots', 'unreachableFiles', 'unusedExports', 'sourceDiagnostics', 'unresolvedDynamicReferences', 'retiredRuntimeReferences'].sort());
    if (category) {
      assert.ok(report[category].some(match), JSON.stringify(report, undefined, 2));
      const cli = spawnSync(process.execPath, [auditFile, '--root', root], { encoding: 'utf8' });
      assert.equal(cli.status, 1, cli.stdout + cli.stderr);
      assert.match(cli.stdout + cli.stderr, /(?:apps\/spfx-react-toolkit-test|packages\/spfx-react-toolkit)/);
    } else {
      assert.equal(report.roots.length, 2);
      for (const key of Object.keys(report).filter(key => key !== 'roots')) assert.deepEqual(report[key], [], key + ': ' + JSON.stringify(report, undefined, 2));
    }
    assert.deepEqual(snapshot(root), before, 'audit and CLI preserve bytes and mtime');
    return report;
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
}
test('native roots and actual catalogue Promise namespace demand retain only Demo', () => check({}));
test('a faithful extra scenario export is unused at its export slot', () => check({ [app + 'src/scenario.ts']: files[app + 'src/scenario.ts'] + 'export const spare = 2;\n' }, 'unusedExports', d => d.file === app + 'src/scenario.ts' && d.symbol === 'spare' && d.line === 2));
test('an unused alias of the same demanded declaration is a separate dead slot', () => check({ [app + 'src/scenario.ts']: files[app + 'src/scenario.ts'] + 'export { Demo as spare };\n' }, 'unusedExports', d => d.symbol === 'spare' && d.line === 2));
test('unknown module namespace escape fails instead of consuming all exports', () => check({ [app + 'src/catalogue.ts']: "declare function leak(value: unknown): Promise<{ Demo: () => number }>;\nexport interface Entry { load(): Promise<{ Demo: () => number }>; }\nexport const entries: readonly Entry[] = [{ load: () => leak(import('./scenario')) }];\n" }, 'unresolvedDynamicReferences', d => d.file === app + 'src/catalogue.ts' && d.line === 3));
const diagnostic = (file, code, line) => d => d.file === app + 'src/' + file && d.code === code && d.line === line;
test('disconnected tsconfig source stays disconnected despite a test-only importer', () => check({ [app + 'src/orphan.ts']: 'export const orphan = 1;\n', 'tests/old.test.cjs': "require('../apps/spfx-react-toolkit-test/src/orphan');\n" }, 'unreachableFiles', d => d === app + 'src/orphan.ts'));
test('barrel demand follows transit spelling without keeping an unused alias', () => check({ [app + 'src/scenario.ts']: "export { Demo } from './implementation';\nexport { Demo as spare } from './implementation';\n", [app + 'src/implementation.ts']: files[app + 'src/scenario.ts'] }, 'unusedExports', d => d.file === app + 'src/scenario.ts' && d.symbol === 'spare' && d.line === 2));
test('orphan concrete class method fails in a live class', () => check({ [app + 'src/scenario.ts']: 'class Subject { live(): number { return 1; } orphan(): number { return 2; } }\nexport function Demo(): number { return new Subject().live(); }\n' }, 'sourceDiagnostics', diagnostic('scenario.ts', 'DEMO_UNUSED_MEMBER', 1)));
test('assigned-only concrete property fails through a local interface', () => check({ [app + 'src/scenario.ts']: 'interface Value { live: number; orphan: number; }\nexport function Demo(): number { const value: Value = { live: 1, orphan: 2 }; return value.live; }\n' }, 'sourceDiagnostics', diagnostic('scenario.ts', 'DEMO_UNUSED_MEMBER', 2)));
test('narrowed discriminated union demand does not keep sibling same-name fields', () => check({ [app + 'src/scenario.ts']: "type Value = { kind: 'ready'; request: number } | { kind: 'error'; request: number };\nfunction read(value: Value): number { return value.kind === 'error' ? value.request : 0; }\nexport function Demo(): number { return read({ kind: 'ready', request: 1 }); }\n" }, 'sourceDiagnostics', diagnostic('scenario.ts', 'DEMO_UNUSED_MEMBER', 3)));
test('non-narrowed union demand consumes synthesized member declarations in every arm', () => check({ [app + 'src/scenario.ts']: "type Value = { kind: 'ready'; request: number } | { kind: 'error'; request: number };\nfunction read(value: Value): number { return value.kind ? value.request : 0; }\nexport function Demo(): number { return read({ kind: 'ready', request: 1 }) + read({ kind: 'error', request: 2 }); }\n" }));
test('strict compiler catches unused locals and unreachable statements', () => check({ [app + 'src/scenario.ts']: 'export function Demo(): number { const unused = 2; return 1; console.log(3); }\n' }, 'sourceDiagnostics', d => d.file === app + 'src/scenario.ts' && d.line === 1 && d.code === 'TS7027'));
test('unrelated unused parameter is not excused by native file ownership', () => check({ [app + 'src/Web.ts']: files[app + 'src/Web.ts'] + '\nfunction unrelated(ignored: number): number { return 1; }\nunrelated(1);\n' }, 'sourceDiagnostics', diagnostic('Web.ts', 'TS6133', 7)));
test('stale lazy import fails visibly', () => check({ [app + 'src/catalogue.ts']: files[app + 'src/catalogue.ts'].replace("'./scenario'", "'./missing'") }, 'sourceDiagnostics', diagnostic('catalogue.ts', 'DEMO_MISSING_MODULE', 2)));
test('nonliteral lazy import fails visibly', () => check({ [app + 'src/catalogue.ts']: files[app + 'src/catalogue.ts'].replace("import('./scenario')", "import(String('./scenario'))") }, 'unresolvedDynamicReferences', d => d.file === app + 'src/catalogue.ts' && d.line === 2));
test('native feature identity is checked against actual configured manifest', () => check({ [app + 'sharepoint/assets/elements.xml']: '<Elements><CustomAction Location="ClientSideExtension.ApplicationCustomizer" ClientSideComponentId="33333333-3333-4333-8333-333333333333" /></Elements>' }, 'sourceDiagnostics', d => d.file === app + 'sharepoint/assets/elements.xml' && d.code === 'DEMO_NATIVE_CONFIG' && d.line === 1));
test('missing configured native asset fails without directory repair', () => check({ [app + 'sharepoint/assets/elements.xml']: undefined }, 'sourceDiagnostics', d => d.file === app + 'config/package-solution.json' && d.code === 'DEMO_MISSING_ASSET' && d.line === 1));
function sourceModule(code) { return '// Generated by scripts/executable-examples.cjs. Do not edit.\nexport const code: string = ' + JSON.stringify(code) + ';\n'; }
function sourceConsumer(host) { return host.replace('entry.load().then(', 'entry.loadSource().then(module => console.log(module.code)); entry.load().then('); }
function metadataFixture({ namespace = false, extra = '', alias = false } = {}) {
  const { createHash } = require('node:crypto');
  const ownerKey = 'packages/subject/src/owner.ts#' + (namespace ? '<module>' : 'owner');
  const title = 'Scenario';
  const id = createHash('sha256').update(JSON.stringify([ownerKey, title])).digest('hex');
  const metadata = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'A number' };
  const code = (alias ? 'function Component(): number { return 1; }\nexport { Component as Demo };\nconst data = ' : 'export function Demo(): number { return 1; }\nexport const demo = ') + JSON.stringify(metadata) + ';\n' + (alias ? 'export { data as demo };\n' : '') + extra;
  const comment = '/**\n' + (namespace ? ' * @packageDocumentation\n' : '') + ' * @example\n * Scenario\n * ```tsx\n' + code.split('\n').slice(0, -1).map(line => ' * ' + line).join('\n') + '\n * ```\n */\n';
  const source = { file: 'packages/subject/src/owner.ts', line: namespace ? 6 : 5 };
  return {
    [source.file]: comment + 'export function owner(): number { return 1; }\n',
    [app + 'src/catalogue.ts']: `export interface Entry { id: string; ownerKey: string; title: string; source: { file: string; line: number }; metadata: { host: string; kind: string; contentPlacement: string; observe: string }; load(): Promise<{ Demo: () => number }>; loadSource(): Promise<{ readonly code: string }>; }\nexport const entries: readonly Entry[] = [{ ${JSON.stringify({id, ownerKey, title, source, metadata}).slice(1, -1)}, load: () => import('./${id}'), loadSource: () => import(/* webpackChunkName: 'spfx-jsdoc-source-${id}' */ "./${id}.source") }];\n`,
    [app + 'src/scenario.ts']: undefined,
    [app + 'src/host.ts']: sourceConsumer(files[app + 'src/host.ts']),
    [app + 'src/' + id + '.source.ts']: sourceModule(code),
    [app + 'src/' + id + '.tsx']: '// Generated by scripts/executable-examples.cjs. Do not edit.\n' + code,
  };
}
test('exact physical named metadata replica has a distinct mandatory artifact role', () => check(metadataFixture()));
test('physical namespace header metadata and aliased Demo/demo exports retain only required slots', () => check(metadataFixture({ namespace: true, alias: true })));
test('faithfully source-and-output-added unrelated export still fails', () => check(metadataFixture({ extra: 'export const unrelated = 1;\n' }), 'unusedExports', d => d.symbol === 'unrelated' && d.line === 4));
test('faithfully source-and-output-added second metadata alias is a dead export slot', () => check(metadataFixture({ alias: true, extra: 'export { data as spareMetadata };\n' }), 'unusedExports', d => d.symbol === 'spareMetadata' && d.line === 6));
test('forged metadata replica bytes fail exact physical provenance', () => { const changes = metadataFixture(); const file = Object.keys(changes).find(name => name.endsWith('.tsx')); changes[file] += '\n'; check(changes, 'sourceDiagnostics', d => d.code === 'DEMO_METADATA_PROVENANCE' && d.file === app + 'src/catalogue.ts' && d.line === 2); });
test('genuine positional native override retains a required middle argument', () => check({ [app + 'src/Web.ts']: files[app + 'src/Web.ts'].replace('  public render()', '  // SDK owns the positional middle argument; path and next remain observed.\n  protected onPropertyPaneFieldChanged(propertyPath: string, _oldValue: unknown, newValue: unknown): void { console.log(propertyPath, newValue); }\n  public render()') }));
test('literal brackets, destructuring and a finite union keep the exact namespace slots', () => check({ [app + 'src/host.ts']: "import { entries, type Entry } from './catalogue';\nexport function mount(): void { entries.forEach((entry: Entry) => { entry.load().then(value => { const alias = value; const { Demo } = alias; const key: 'Demo' = 'Demo'; return Demo() + alias['Demo']() + alias[key](); }); }); }\n" }));
test('finite dynamic object keys and JSON serialization retain only genuinely observed properties', () => check({ [app + 'src/scenario.ts']: "export function Demo(): number { const values = { first: 1, second: 2 }; const key: 'first' | 'second' = Math.random() ? 'first' : 'second'; const serialized = { result: values[key] }; return JSON.stringify(serialized).length; }\n" }));
test('retired bridge reference is rejected independently of resolution', () => check({ [app + 'src/host.ts']: files[app + 'src/host.ts'] + "import './demoRegistry';\n", [app + 'src/demoRegistry.ts']: 'console.log("retired");\n' }, 'retiredRuntimeReferences', d => d.file === app + 'src/host.ts' && d.line === 3 && d.target === './demoRegistry'));
const resourceFiles = {
  [app + 'src/strings.d.ts']: "declare module 'AppStrings' { const strings: { Label: string }; export = strings; }\n",
  [app + 'src/loc/en-us.js']: 'define([], function() { return { Label: "Label" }; });\n',
  [app + 'src/scenario.ts']: "import * as strings from 'AppStrings';\nexport function Demo(): number { return strings.Label.length; }\n",
};
function resourceConfig() { const config = JSON.parse(files[app + 'config/config.json']); config.localizedResources.AppStrings = './lib/loc/{locale}.js'; return JSON.stringify(config); }
test('owned ambient plus mapped AMD dictionary is a native source contract', () => check({ ...resourceFiles, [app + 'config/config.json']: resourceConfig() }));
test('removed required native mapping fails from surviving ambient/import contract', () => check(resourceFiles, 'sourceDiagnostics', d => d.file === app + 'config/config.json' && d.code === 'DEMO_MISSING_RESOURCE' && d.line === 1));
test('removed dictionary fails without repairing its configured path', () => check({ ...resourceFiles, [app + 'src/loc/en-us.js']: undefined, [app + 'config/config.json']: resourceConfig() }, 'sourceDiagnostics', d => d.file === app + 'config/config.json' && d.code === 'DEMO_MISSING_RESOURCE' && d.line === 1));
test('arbitrary extra bitmap does not gain a role from Teams packaging globs', () => check({ [app + 'teams/orphan.png']: 'not a configured icon' }, 'sourceDiagnostics', d => d.file === app + 'teams/orphan.png' && d.code === 'DEMO_UNUSED_ASSET' && d.line === 1));
const registrationFiles = {
  [app + 'src/scenario.ts']: "import './registration';\nimport { registered } from './runtime';\nexport function Demo(): number { return registered.value(); }\n",
  [app + 'src/runtime.ts']: 'export interface Registration {}\nexport const registered = {} as Registration;\n',
  [app + 'src/registration.ts']: "import { registered } from './runtime';\ndeclare module './runtime' { interface Registration { value(): number; } }\nregistered.value = () => 1;\n",
};
test('necessary registration retains its real augmentation and evaluation edge', () => check(registrationFiles));
test('removing necessary registration cannot be excused by a test-only importer', () => check({ ...registrationFiles, [app + 'src/scenario.ts']: registrationFiles[app + 'src/scenario.ts'].replace("import './registration';\n", ''), [app + 'src/registration.ts']: undefined, 'tests/old.test.cjs': "require('../apps/spfx-react-toolkit-test/src/registration');" }, 'sourceDiagnostics', diagnostic('scenario.ts', 'TS2339', 2)));
test('CLI rejects unknown options visibly', () => { const cli = spawnSync(process.execPath, [auditFile, '--repair'], { encoding: 'utf8' }); assert.equal(cli.status, 1); assert.match(cli.stderr, /Usage:/); });
function catalogueFixture() {
  const changes = metadataFixture();
  let catalogue = changes[app + 'src/catalogue.ts'];
  catalogue = catalogue.replace('export interface Entry {', 'export interface Entry { groupRoute: string; groupKey: string; featureKey: string; featureTitle: string; publicNames: readonly string[];');
  catalogue = catalogue.replace('export const entries:', 'export const executableExamples:').replace('load: () => import(', 'groupRoute: "docs/api/subject.md", groupKey: "controls", featureKey: "subject", featureTitle: "Subject", publicNames: ["owner"], load: () => import(');
  changes[app + 'src/catalogue.ts'] = catalogue + 'export const executableGroups = [{ key: "controls", title: "Controls", features: [{ key: "subject", title: "Subject", scenarios: [executableExamples[0]] }] }];\n';
  changes[app + 'src/host.ts'] = sourceConsumer(files[app + 'src/host.ts']).replace("import { entries, type Entry }", "import { executableGroups, type Entry }").replace('export function mount()', 'const entries = executableGroups.flatMap(group => group.features).flatMap(feature => feature.scenarios);\nexport function mount()');
  return changes;
}

const sourceOutput = changes => Object.keys(changes).find(name => name.endsWith('.source.ts'));
const sourceProvenance = diagnostic('catalogue.ts', 'DEMO_METADATA_PROVENANCE', 2);
test('source loader actually consumes only the exact canonical code export', () => check(catalogueFixture()));
test('missing required source loader fails physical provenance', () => {
  const changes = catalogueFixture();
  changes[app + 'src/catalogue.ts'] = changes[app + 'src/catalogue.ts'].replace(/, loadSource:.*?(?= }];)/, '');
  check(changes, 'sourceDiagnostics', sourceProvenance);
});
test('missing source module fails physical provenance', () => {
  const changes = catalogueFixture(); changes[sourceOutput(changes)] = undefined;
  check(changes, 'sourceDiagnostics', sourceProvenance);
});
test('source loader cannot target another ID even with the exact canonical string', () => {
  const changes = catalogueFixture(), output = sourceOutput(changes);
  const id = path.basename(output, '.source.ts');
  changes[app + 'src/' + '0'.repeat(64) + '.source.ts'] = changes[output]; changes[output] = undefined;
  changes[app + 'src/catalogue.ts'] = changes[app + 'src/catalogue.ts'].replace('./' + id + '.source', './' + '0'.repeat(64) + '.source');
  check(changes, 'sourceDiagnostics', sourceProvenance);
});
test('source loader cannot borrow the same ID from a different folder', () => {
  const changes = catalogueFixture(), output = sourceOutput(changes), id = path.basename(output, '.source.ts');
  changes[app + 'src/other/' + id + '.source.ts'] = changes[output]; changes[output] = undefined;
  changes[app + 'src/catalogue.ts'] = changes[app + 'src/catalogue.ts'].replace('./' + id + '.source', './other/' + id + '.source');
  check(changes, 'sourceDiagnostics', sourceProvenance);
});
test('source loader requires its exact literal chunk identity', () => {
  const changes = catalogueFixture();
  changes[app + 'src/catalogue.ts'] = changes[app + 'src/catalogue.ts'].replace('spfx-jsdoc-source-', 'spfx-jsdoc-forged-');
  check(changes, 'sourceDiagnostics', sourceProvenance);
});
test('a line-comment imitation cannot supply the source webpack chunk identity', () => {
  const changes = catalogueFixture();
  changes[app + 'src/catalogue.ts'] = changes[app + 'src/catalogue.ts'].replace('import(/* webpackChunkName:', 'import(// /* webpackChunkName:').replace(/(spfx-jsdoc-source-[a-f0-9]+['] \*\/) /, '$1\n');
  check(changes, 'sourceDiagnostics', sourceProvenance);
});
test('a later conflicting webpack chunk directive cannot override the proven source identity', () => {
  const changes = catalogueFixture();
  changes[app + 'src/catalogue.ts'] = changes[app + 'src/catalogue.ts'].replace(/(spfx-jsdoc-source-[a-f0-9]+' \*\/)/, '$1 /* webpackChunkName: "forged" */');
  check(changes, 'sourceDiagnostics', sourceProvenance);
});
test('a conflicting webpack chunk directive after the source argument cannot bypass provenance', () => {
  const changes = catalogueFixture();
  changes[app + 'src/catalogue.ts'] = changes[app + 'src/catalogue.ts'].replace(/(\.source")\)/, '$1 /* webpackChunkName: "forged" */)');
  check(changes, 'sourceDiagnostics', sourceProvenance);
});
test('source loader cannot replace the imported code through a wrapper', () => {
  const changes = catalogueFixture();
  changes[app + 'src/catalogue.ts'] = changes[app + 'src/catalogue.ts'].replace(/(loadSource:.*?import\(.*?\))(?= }];)/, '$1.then(() => ({code: "forged"}))');
  check(changes, 'sourceDiagnostics', sourceProvenance);
});
for (const [name, mutate] of [
  ['inferred source literal type', text => text.replace('export const code: string = ', 'export const code = ')],
  ['forged source text', text => text.replace('export function Demo()', 'export function Forged()')],
  ['extra source statement', text => text + 'console.log("unrelated");\n'],
  ['extra source export', text => text + 'export const spare = 1;\n'],
  ['nonliteral source expression', text => text.replace(/;\n$/, ' + "";\n')],
]) test(name + ' fails exact source-module provenance', () => {
  const changes = catalogueFixture(), output = sourceOutput(changes); changes[output] = mutate(changes[output]);
  check(changes, 'sourceDiagnostics', sourceProvenance);
});
test('metadata proof cannot keep the source code export live after its consumer is removed', () => {
  const changes = catalogueFixture();
  changes[app + 'src/host.ts'] = changes[app + 'src/host.ts'].replace('entry.loadSource().then(module => console.log(module.code)); ', '');
  check(changes, 'unusedExports', d => d.file === sourceOutput(changes) && d.symbol === 'code');
});
test('invoking the source loader without reading code leaves its export unused', () => {
  const changes = catalogueFixture();
  changes[app + 'src/host.ts'] = changes[app + 'src/host.ts'].replace('module => console.log(module.code)', '() => console.log("ready")');
  check(changes, 'unusedExports', d => d.file === sourceOutput(changes) && d.symbol === 'code');
});
test('exact reachable catalogue classification and hierarchy have a distinct tooling role', () => check(catalogueFixture()));
test('extra catalogue field cannot inherit a provenance role', () => { const changes = catalogueFixture(); changes[app + 'src/catalogue.ts'] = changes[app + 'src/catalogue.ts'].replace('groupRoute: "docs/api/subject.md"', 'orphan: true, groupRoute: "docs/api/subject.md"'); check(changes, 'sourceDiagnostics', diagnostic('catalogue.ts', 'DEMO_CATALOGUE_CONTRACT', 2)); });
test('second catalogue array alias is not an accepted export slot', () => { const changes = catalogueFixture(); changes[app + 'src/catalogue.ts'] += 'export { executableExamples as spareCatalogue };\n'; check(changes, 'unusedExports', d => d.file === app + 'src/catalogue.ts' && d.symbol === 'spareCatalogue' && d.line === 4); });
test('unrelated catalogue export is not a tooling artifact', () => { const changes = catalogueFixture(); changes[app + 'src/catalogue.ts'] += 'export const spare = 1;\n'; check(changes, 'unusedExports', d => d.symbol === 'spare' && d.line === 4); });
test('forged hierarchy classification fails independently of exact scenario bytes', () => { const changes = catalogueFixture(); changes[app + 'src/catalogue.ts'] = changes[app + 'src/catalogue.ts'].replace('key: "subject", title:', 'key: "different", title:'); check(changes, 'sourceDiagnostics', diagnostic('catalogue.ts', 'DEMO_CATALOGUE_CONTRACT', 2)); });
test('disconnected exact catalogue does not become a runtime root', () => check({ ...catalogueFixture(), [app + 'src/host.ts']: 'export function mount(): void {}\n' }, 'unreachableFiles', d => d === app + 'src/catalogue.ts'));
test('a local Record annotation does not excuse an unused concrete field', () => check({ [app + 'src/scenario.ts']: 'export function Demo(): number { const value: Record<string, number> = { live: 1, orphan: 2 }; return value.live; }\n' }, 'sourceDiagnostics', d => d.file === app + 'src/scenario.ts' && d.code === 'DEMO_UNUSED_MEMBER' && /orphan/.test(d.message)));
test('static namespace broad escape fails even alongside a genuine finite read', () => check({ [app + 'src/host.ts']: "import * as scenario from './scenario';\ndeclare function leak(value: unknown): void;\nexport function mount(): void { scenario.Demo(); leak(scenario); }\n" }, 'unresolvedDynamicReferences', d => d.file === app + 'src/host.ts' && d.line === 3));
const bindingFixtures = require('./fixtures/demo-reachability/bindings.cjs');
const unusedScenarioMember = name => d => d.file === app + 'src/scenario.tsx' && d.line > 0 && d.code === 'DEMO_UNUSED_MEMBER' && d.message.includes(`member ${name} `);
test('real aliased Form family preserves only the bound default value path', () => check(bindingFixtures.form()));
test('extra Form default key is not consumed by an unrelated bound name', () => check(bindingFixtures.form(", orphan: 'unused'"), 'sourceDiagnostics', unusedScenarioMember('orphan')));
test('same-spelling local Form family grants no toolkit field contract', () => check(bindingFixtures.form('', true), 'sourceDiagnostics', unusedScenarioMember('live')));
test('real Grid JSX fields preserve selected raw data and exact finite presentations', () => check(bindingFixtures.grid()));
test('real Grid classifier retains the original selected invalid object payload', () => check(bindingFixtures.grid({ invalid: true })));
test('nonselected raw Grid row member is still unused', () => check(bindingFixtures.grid({ extra: ", orphan: 'unused'" }), 'sourceDiagnostics', unusedScenarioMember('orphan')));
test('extra finite Grid mapping key is still unused', () => check(bindingFixtures.grid({ mapExtra: ", spare: {label:'Unused'}" }), 'sourceDiagnostics', unusedScenarioMember('spare')));
test('extra Grid presentation member is not an external contract', () => check(bindingFixtures.grid({ presentationExtra: ', orphan: 1' }), 'sourceDiagnostics', unusedScenarioMember('orphan')));
test('real aliased Grid descriptor demand does not keep an extra typed row member', () => check(bindingFixtures.grid({ descriptor: true, extra: ", orphan: 'unused'" }), 'sourceDiagnostics', unusedScenarioMember('orphan')));
test('same-spelling local Grid marker grants no selected-data contract', () => check(bindingFixtures.grid({lookalike:true}), 'sourceDiagnostics', unusedScenarioMember('live')));
test('genuine React static error update links only read state members', () => check(bindingFixtures.react()));
test('extra React static returned state member is still unused', () => check(bindingFixtures.react(true), 'sourceDiagnostics', unusedScenarioMember('orphan')));
test('same-spelling non-React static hook has no framework state role', () => check(bindingFixtures.react(false, true), 'sourceDiagnostics', unusedScenarioMember('getDerivedStateFromError')));
function topFixture({ extra = false, index = false, wrongValue = false } = {}) {
  const changes = catalogueFixture();
  const output = Object.keys(changes).find(name => name.endsWith('.tsx'));
  const metadata = {host:'applicationCustomizer', kind:'snapshot', contentPlacement:'configuration', observe:'A number', hostAdapter:'applicationCustomizerTop'};
  const code = `import * as React from 'react';\nexport function Demo({applicationCustomizerTop}: { ${index ? '[capability: string]: unknown;' : ''} applicationCustomizerTop?: {readonly placeholder:'Top'; ${extra ? 'readonly extra: number;' : ''}} } = {}): React.ReactElement { return <p>{applicationCustomizerTop ? 'Ready' : 'Missing'}</p>; }\nexport const demo = ${JSON.stringify(metadata)};\n`;
  changes[output] = '// Generated by scripts/executable-examples.cjs. Do not edit.\n' + code;
  changes[sourceOutput(changes)] = sourceModule(code);
  changes['packages/subject/src/owner.ts'] = '/**\n * @example\n * Scenario\n * ```tsx\n' + code.trimEnd().split('\n').map(line => ' * ' + line).join('\n') + '\n * ```\n */\nexport function owner(): number { return 1; }\n';
  changes[app+'src/catalogue.ts'] = "import type * as React from 'react';\nexport interface ExecutableExampleProps { applicationCustomizerTop?: {readonly placeholder:'Top'}; }\n" + changes[app+'src/catalogue.ts'].replace('Demo: () => number','Demo: React.ComponentType<ExecutableExampleProps>').replace(JSON.stringify({host:'none',kind:'snapshot',contentPlacement:'subject',observe:'A number'}),JSON.stringify(metadata)).replace('observe: string }; load()', 'observe: string; hostAdapter?: string }; load()');
  changes[app+'src/host.ts'] = undefined;
  changes[app+'src/host.tsx'] = `import * as React from 'react';\nimport {executableGroups, type Entry, type ExecutableExampleProps} from './catalogue';\nfunction load(entry: Entry): React.LazyExoticComponent<React.ComponentType<ExecutableExampleProps>> { entry.loadSource().then(module => console.log(module.code)); return React.lazy(() => entry.load().then(module => ({default: module.Demo}))); }\nfunction Top({entry}:{entry:Entry}): React.ReactElement { const Demo = load(entry); return <Demo applicationCustomizerTop={{placeholder:'${wrongValue ? 'Bottom' : 'Top'}'}} />; }\nexport function mount(): void { console.log(React.createElement(Top,{entry: executableGroups[0].features[0].scenarios[0]})); }\n`;
  return changes;
}
test('accepted exact Top prerequisite has a type contract role, separate from a leaf runtime read', () => check(topFixture()));
test('extra source Top prerequisite member fails its exact validated channel', () => check(topFixture({extra:true}), 'sourceDiagnostics', d => d.file.startsWith(app+'src/') && d.code === 'DEMO_TOP_CONTRACT' && d.line > 0));
test('indexed outer Top props cannot acquire an exact prerequisite role', () => check(topFixture({index:true}), 'sourceDiagnostics', d => d.code === 'DEMO_TOP_CONTRACT' && d.line > 0));
test('wrong Top prerequisite literal fails despite accepted copied source metadata', () => check(topFixture({wrongValue:true}), 'sourceDiagnostics', d => d.file === app+'src/host.tsx' && d.code === 'DEMO_TOP_CONTRACT' && d.line === 4));
test('Top-like object on an unrelated component has no accepted prerequisite role', () => check({ [app+'src/scenario.ts']:undefined, [app+'src/scenario.tsx']: "import * as React from 'react';\nfunction Subject(_props:{applicationCustomizerTop?:{readonly placeholder:'Top'}}):React.ReactElement {return <p>Ready</p>;}\nexport function Demo():React.ReactElement {return <Subject applicationCustomizerTop={{placeholder:'Top'}} />;}\n", [app+'src/catalogue.ts']:files[app+'src/catalogue.ts'].replace('() => number','() => unknown') }, 'sourceDiagnostics', unusedScenarioMember('placeholder')));
test('an external presentation annotation alone does not consume a local concrete member', () => check({[app+'src/scenario.ts']: "import type {DataGridBadgeValue} from '@apvee/spfx-react-toolkit/controls/data-grid';\nexport function Demo(): number { const value: DataGridBadgeValue & {observed:number} = {label:'unused',observed:1}; return value.observed; }\n"}, 'sourceDiagnostics', d=>d.file===app+'src/scenario.ts' && d.line===2 && d.code==='DEMO_UNUSED_MEMBER' && /member label /.test(d.message)));
test('a same-type Form object outside the actual bound defaults gains no selected-key role', () => {
  const changes = bindingFixtures.form();
  changes[app+'src/scenario.tsx'] = changes[app+'src/scenario.tsx']
    .replace('interface Fields { live: string;', 'interface Fields { live: string; observed?:string;')
    .replace('const defaults:', "const unrelated: Fields = {live:'unbound',observed:'kept'};\nconst defaults:")
    .replace('return <Editor.Root', 'return <><p>{unrelated.observed}</p><Editor.Root')
    .replace('</Editor.Root>;', '</Editor.Root></>;');
  check(changes, 'sourceDiagnostics', unusedScenarioMember('live'));
});
test('finite Map presentation contracts retain genuine keys and fields', () => {
  const changes = bindingFixtures.grid();
  changes[app+'src/scenario.tsx'] = changes[app+'src/scenario.tsx']
    .replace('const presentations = {ready: {label:\'Ready\'}, pending: {label:\'Pending\'}};', "const presentations = new Map<'ready'|'pending', import('@apvee/spfx-react-toolkit/controls/data-grid').DataGridBadgeValue>([['ready',{label:'Ready'}],['pending',{label:'Pending'}]]);");
  check(changes);
});
test('bound Form array paths retain selected nested fields', () => {
  const changes = bindingFixtures.form();
  changes[app+'src/scenario.tsx'] = changes[app+'src/scenario.tsx']
    .replace('interface Fields { live: string;  }', 'interface Fields { rows: {live:string}[]; }')
    .replace("{ live: 'kept' }", "{ rows: [{live:'kept'}] }")
    .replace('<Editor.Input name="live" />', '{[0].map(index=><Editor.Input key={index} name={`rows.${index}.live`} />)}');
  check(changes);
});
test('bound Form array paths do not retain a nonselected nested member', () => {
  const changes = bindingFixtures.form();
  changes[app+'src/scenario.tsx'] = changes[app+'src/scenario.tsx']
    .replace('interface Fields { live: string;  }', 'interface Fields { rows: {live:string; orphan:string}[]; }')
    .replace("{ live: 'kept' }", "{ rows: [{live:'kept',orphan:'unused'}] }")
    .replace('<Editor.Input name="live" />', '{[0].map(index=><Editor.Input key={index} name={`rows.${index}.live`} />)}');
  check(changes,'sourceDiagnostics',unusedScenarioMember('orphan'));
});
test('a faithfully paired extra native string remains an unused dictionary member', () => check({...resourceFiles,[app+'config/config.json']:resourceConfig(),[app+'src/strings.d.ts']:resourceFiles[app+'src/strings.d.ts'].replace('Label: string','Label: string; Unused: string'),[app+'src/loc/en-us.js']:'define([], function() { return { Label: "Label", Unused: "Unused" }; });\n'},'sourceDiagnostics',d=>d.file===app+'src/loc/en-us.js'&&d.line===1&&d.code==='DEMO_UNUSED_RESOURCE_MEMBER'&&/Unused/.test(d.message)));
const toolkitResourceFiles = {
 'packages/spfx-react-toolkit/localization/resources.json': JSON.stringify({schemaVersion:1,resources:[{id:'ApveeSpfxReactToolkitPickersStrings',path:'lib/controls/pickers/loc/{locale}.js',locales:['en-US']}]}),
 'packages/spfx-react-toolkit/src/controls/pickers/localization/resources/en-US.json':'{}\n',
};
function toolkitResourceConfig(){const config=JSON.parse(files[app+'config/config.json']);config.localizedResources.ApveeSpfxReactToolkitPickersStrings='../../node_modules/@apvee/spfx-react-toolkit/lib/controls/pickers/loc/{locale}.js';return JSON.stringify(config);}
test('toolkit mapping derives its actual versioned native source owner',()=>check({...toolkitResourceFiles,[app+'config/config.json']:toolkitResourceConfig()}));
test('removed toolkit native mapping fails against surviving inventory ownership',()=>check(toolkitResourceFiles,'sourceDiagnostics',d=>d.file===app+'config/config.json'&&d.line===1&&d.code==='DEMO_MISSING_RESOURCE'));
test('removed toolkit native source asset cannot borrow a compiled/workspace dictionary',()=>check({...toolkitResourceFiles,[app+'config/config.json']:toolkitResourceConfig(),'packages/spfx-react-toolkit/src/controls/pickers/localization/resources/en-US.json':undefined},'sourceDiagnostics',d=>d.file==='packages/spfx-react-toolkit/localization/resources.json'&&d.line===1&&d.code==='DEMO_MISSING_RESOURCE'));
test('unused public constructor parameter properties are concrete assigned-only members',()=>check({[app+'src/scenario.ts']:'class Subject { constructor(public orphan = 1) {} live(): number {return 1;} }\nexport function Demo(): number {return new Subject().live();}\n'},'sourceDiagnostics',d=>d.file===app+'src/scenario.ts'&&d.line===1&&d.code==='DEMO_UNUSED_MEMBER'&&/member orphan /.test(d.message)));
test('a getter read does not consume an unused setter body',()=>check({[app+'src/scenario.ts']:'class Subject { get value(): number {return 1;} set value(next: number) {console.log(next);} }\nexport function Demo(): number {return new Subject().value;}\n'},'sourceDiagnostics',d=>d.file===app+'src/scenario.ts'&&d.line===1&&d.code==='DEMO_UNUSED_MEMBER'&&/setter value/.test(d.message)));
test('a genuine setter assignment consumes the setter without pretending to read a field',()=>check({[app+'src/scenario.ts']:'class Subject {set value(next: number) {console.log(next);} }\nexport function Demo(): number {const subject=new Subject();subject.value=1;return 1;}\n'}));
const typedSource = "import type {Payload} from './types';\nexport function Demo(): number {const value:Payload={live:1};return value.live;}\n";
test('a pure type-only module retains its real compile contract without becoming an evaluation root',()=>check({[app+'src/scenario.ts']:typedSource,[app+'src/types.ts']:'export interface Payload {live:number;}\n'}));
test('type-only reachability does not excuse unreachable top-level evaluation effects',()=>check({[app+'src/scenario.ts']:typedSource,[app+'src/types.ts']:'export interface Payload {live:number;}\nconsole.log("not evaluated by the app");\n'},'sourceDiagnostics',diagnostic('types.ts','DEMO_TYPE_ONLY_RUNTIME',2)));
test('an ordinary import erased to a type contract still does not evaluate its module',()=>check({[app+'src/scenario.ts']:typedSource.replace('import type','import'),[app+'src/types.ts']:'export interface Payload {live:number;}\nconsole.log("not evaluated by the app");\n'},'sourceDiagnostics',diagnostic('types.ts','DEMO_TYPE_ONLY_RUNTIME',2)));
test('literal require namespace aliases have finite export-slot demand',()=>check({[app+'src/host.ts']:files[app+'src/host.ts'].replace('export function mount()',"const namespace = require('./scenario');\nexport function mount()").replace('entries.forEach(',"namespace['Demo'](); entries.forEach(")}));
test('literal require namespace cannot escape broadly alongside a finite read',()=>check({[app+'src/host.ts']:files[app+'src/host.ts'].replace('export function mount()',"const namespace = require('./scenario');\ndeclare function leak(value: unknown): void;\nexport function mount()").replace('entries.forEach(',"namespace.Demo(); leak(namespace); entries.forEach(")},'unresolvedDynamicReferences',d=>d.file===app+'src/host.ts'&&d.line===4));
test('enumerating own keys does not consume a class prototype method',()=>check({[app+'src/scenario.ts']:'class Subject {field=1; orphan(): number {return 2;} }\nexport function Demo(): number {return Object.keys(new Subject()).length;}\n'},'sourceDiagnostics',d=>d.file===app+'src/scenario.ts'&&d.line===1&&d.code==='DEMO_UNUSED_MEMBER'&&/member orphan /.test(d.message)));
test('serializing an instance does not consume its prototype getter',()=>check({[app+'src/scenario.ts']:'class Subject {field=1; get orphan(): number {return 2;} }\nexport function Demo(): number {return JSON.stringify(new Subject()).length;}\n'},'sourceDiagnostics',d=>d.file===app+'src/scenario.ts'&&d.line===1&&d.code==='DEMO_UNUSED_MEMBER'&&/member orphan /.test(d.message)));
test('actual external delivery follows a genuine React memo origin',()=>check({[app+'src/scenario.ts']:"import * as React from 'react';\nimport {useSPFxApiPermissionPrecheck} from '@apvee/spfx-react-toolkit';\nexport function Demo(): number {const config=React.useMemo(()=>({customApis:[{name:'Custom',resource:'api',scopes:['read']}]}),[]);const result=useSPFxApiPermissionPrecheck(config);return result.results.length;}\n"}));
test('an undelivered memo object cannot borrow an external contract annotation',()=>check({[app+'src/scenario.ts']:"import * as React from 'react';\nimport type {SPFxApiPermissionPrecheckConfig} from '@apvee/spfx-react-toolkit';\nexport function Demo(): number {const config=React.useMemo<SPFxApiPermissionPrecheckConfig>(()=>({customApis:[{name:'Custom',resource:'api',scopes:['read']}]}),[]);return config.customApis?.length ?? 0;}\n"},'sourceDiagnostics',d=>d.code==='DEMO_UNUSED_MEMBER'&&/member resource /.test(d.message)));
test('a type query does not execute or read the concrete member it describes',()=>check({[app+'src/scenario.ts']:'const data={live:1,orphan:2};\ntype Shape=typeof data.orphan;\nexport function Demo(): Shape {return data.live;}\n'},'sourceDiagnostics',d=>d.file===app+'src/scenario.ts'&&d.line===1&&d.code==='DEMO_UNUSED_MEMBER'&&/member orphan /.test(d.message)));
const sortCode="import {SearchQueryBuilder} from '@pnp/sp/search';\nexport function Demo(): number {const builder=SearchQueryBuilder();builder.sortList({Property:'Title',Direction:1});return 1;}\n";
test('external rest arguments consume their real element contract',()=>check({[app+'src/scenario.ts']:sortCode}));
test('extra rest argument data cannot inherit the external element contract',()=>check({[app+'src/scenario.ts']:sortCode.replace("builder.sortList({Property:'Title',Direction:1})","const sort={Property:'Title',Direction:1,orphan:1};builder.sortList(sort)")},'sourceDiagnostics',d=>d.code==='DEMO_UNUSED_MEMBER'&&/member orphan /.test(d.message)));
const freezeCode="import * as React from 'react';\nimport {useSPFxApiPermissionPrecheck, SPFxApiPermissionPrecheckConfig} from '@apvee/spfx-react-toolkit';\nexport function Demo(): number {const [config,setConfig]=React.useState<SPFxApiPermissionPrecheckConfig>({});const result=useSPFxApiPermissionPrecheck(config);React.useEffect(()=>setConfig(Object.freeze({customApis:[{name:'Custom',resource:'api',scopes:['read']}]})),[]);return result.results.length;}\n";
test('external delivery follows genuine frozen React state origins',()=>check({[app+'src/scenario.ts']:freezeCode}));
function callbackFixture(extra=''){return "import {PropertyPaneDropdown,IPropertyPaneDropdownProps} from '@microsoft/sp-property-pane';\ninterface Controller {options():IPropertyPaneDropdownProps['options'];}\nconst factory=()=>[{key:'one',text:'One'"+extra+"}];\nconst controller:Controller={options:factory};\nexport function Demo():number {PropertyPaneDropdown('x',{label:'Choice',options:controller.options()});return 1;}\n";}
test('actual external delivery follows a local typed callback implementation',()=>check({[app+'src/scenario.ts']:callbackFixture()}));
test('a callback result extra member does not inherit its external output contract',()=>check({[app+'src/scenario.ts']:callbackFixture(',orphan:1')},'sourceDiagnostics',d=>d.code==='DEMO_UNUSED_MEMBER'&&/member orphan /.test(d.message)));
test('actual external delivery follows a const assertion without dropping its inner object',()=>check({[app+'src/scenario.ts']:"import * as React from 'react';\nimport {useSPFxApiPermissionPrecheck} from '@apvee/spfx-react-toolkit';\nexport function Demo():number {const options=React.useMemo(()=>({autoCheck:false,mode:'passive'} as const),[]);return useSPFxApiPermissionPrecheck({},options).results.length;}\n"}));
function jsxForwardFixture(extra=''){return "import * as React from 'react';\nimport {PropertyPaneDropdown,IPropertyPaneDropdownOption} from '@microsoft/sp-property-pane';\nfunction Subject(props:{options:IPropertyPaneDropdownOption[]}):React.ReactElement {PropertyPaneDropdown('x',{label:'Choice',options:props.options});return <p>Choice</p>;}\nexport function Demo():React.ReactElement {return <Subject options={[{key:'one',text:'One'"+extra+"}]} />;}\n";}
test('a real local JSX forwarding call retains actual external nested data',()=>check({[app+'src/scenario.ts']:undefined,[app+'src/scenario.tsx']:jsxForwardFixture(),[app+'src/catalogue.ts']:files[app+'src/catalogue.ts'].replace('() => number','() => unknown')}));
test('local JSX forwarding does not consume an extra external data member',()=>check({[app+'src/scenario.ts']:undefined,[app+'src/scenario.tsx']:jsxForwardFixture(',orphan:1'),[app+'src/catalogue.ts']:files[app+'src/catalogue.ts'].replace('() => number','() => unknown')},'sourceDiagnostics',unusedScenarioMember('orphan')));
function destructuredFixture(extra=''){return "import {PropertyPaneDropdown,IPropertyPaneDropdownOption} from '@microsoft/sp-property-pane';\nfunction settings():{options:IPropertyPaneDropdownOption[]} {return {options:[{key:'one',text:'One'"+extra+"}]};}\nexport function Demo():number {const {options:choices}=settings();PropertyPaneDropdown('x',{label:'Choice',options:choices});return 1;}\n";}
test('external data delivery follows an actual destructured function result',()=>check({[app+'src/scenario.ts']:destructuredFixture()}));
test('destructured delivery still rejects a noncontract nested member',()=>check({[app+'src/scenario.ts']:destructuredFixture(',orphan:1')},'sourceDiagnostics',d=>d.code==='DEMO_UNUSED_MEMBER'&&/member orphan /.test(d.message)));
test('an import-type qualifier retains its actual exported type slot',()=>check({[app+'src/scenario.ts']:'export function Demo():number {const value:import("./types").Payload={live:1};return value.live;}\n',[app+'src/types.ts']:'export interface Payload {live:number;}\n'}));
test('a namespace type qualifier retains its actual exported type slot',()=>check({[app+'src/scenario.ts']:'import type * as Types from "./types";\nexport function Demo():number {const value:Types.Payload={live:1};return value.live;}\n',[app+'src/types.ts']:'export interface Payload {live:number;}\n'}));
test('a typeof-only export use does not claim that its callable runtime body is consumed',()=>check({[app+'src/scenario.ts']:'import type {create} from "./types";\ntype Factory=typeof create;\nexport function Demo():ReturnType<Factory> {return 1;}\n',[app+'src/types.ts']:'export function create():number {return 2;}\n'},'sourceDiagnostics',diagnostic('types.ts','DEMO_TYPE_ONLY_EXPORT',1)));
test('actual nested external delivery follows a local object spread',()=>check({[app+'src/scenario.ts']:"import {PropertyPaneDropdown,IPropertyPaneDropdownOption} from '@microsoft/sp-property-pane';\nexport function Demo():number {const patch:{options:IPropertyPaneDropdownOption[]}={options:[{key:'one',text:'One'}]};const next={...patch};PropertyPaneDropdown('x',{label:'Choice',options:next.options});return 1;}\n"}));
function motionFixture({extra='',unbound=false,lookalike=false}={}) {return "import * as React from 'react';\nimport {isValidElement as valid,cloneElement as copy} from 'react';\nimport {AccordionPanel as Panel} from '@fluentui/react-components';\ntype NativeProps=React.HTMLAttributes<HTMLDivElement>&{inert?:''};\n"+(lookalike?"const LocalPanel=(_props:React.ComponentProps<typeof Panel>):React.ReactElement=><div/>;\n":'')+"export function Demo():React.ReactElement {"+(unbound?"const unrelated={inert:'' as const,observed:1};console.log(unrelated.observed);":'')+"return <"+(lookalike?'LocalPanel':'Panel')+" collapseMotion={{children:(Motion,props)=>{const panel=valid<NativeProps>(props.children)?copy(props.children,{inert:props.visible?undefined:''"+extra+"}):props.children;return <Motion {...props} unmountOnExit={false}>{panel}</Motion>;}}}>Retained content</"+(lookalike?'LocalPanel':'Panel')+">;}\n";}
const motionChanges=options=>({[app+'src/scenario.ts']:undefined,[app+'src/scenario.tsx']:motionFixture(options),[app+'src/catalogue.ts']:files[app+'src/catalogue.ts'].replace('() => number','() => unknown')});
test('actual aliased Fluent motion native child receives the bounded React17 inert delta',()=>check(motionChanges()));
test('an extra motion clone delta member is not a native attribute contract',()=>check(motionChanges({extra:',orphan:1'}),'sourceDiagnostics',unusedScenarioMember('orphan')));
test('an unbound inert delta remains unused beside the actual native motion handoff',()=>check(motionChanges({unbound:true}),'sourceDiagnostics',unusedScenarioMember('inert')));
test('a local lookalike motion component grants no native child handoff',()=>check(motionChanges({lookalike:true}),'sourceDiagnostics',unusedScenarioMember('inert')));
test('actual destructured delivery follows shorthand return value identity',()=>check({[app+'src/scenario.ts']:destructuredFixture().replace("return {options:[{key:'one',text:'One'}]};","const options=[{key:'one',text:'One'}];return {options};")}));
test('a local clone lookalike cannot consume the native motion delta',()=>{const changes=motionChanges();changes[app+'src/scenario.tsx']=changes[app+'src/scenario.tsx'].replace('isValidElement as valid,cloneElement as copy','isValidElement as valid').replace('type NativeProps=',"function copy<P>(element:React.ReactElement<P>,_props:Partial<P>):React.ReactElement<P>{return element;}\ntype NativeProps=");check(changes,'sourceDiagnostics',unusedScenarioMember('inert'));});
test('an unused clone result cannot gain the returned native motion handoff role',()=>{const changes=motionChanges();changes[app+'src/scenario.tsx']=changes[app+'src/scenario.tsx'].replace('{panel}</Motion>','{props.children}</Motion>');check(changes,'sourceDiagnostics',unusedScenarioMember('inert'));});
test('a broad key asserted to namespace keyof is unresolved rather than evidence for every export',()=>check({[app+'src/host.ts']:"import * as namespace from './scenario';\ndeclare const key:string;\nexport function mount():void {namespace[key as keyof typeof namespace]();}\n"},'unresolvedDynamicReferences',d=>d.file===app+'src/host.ts'&&d.line===3));
test('a broad asserted local key alias still has an unresolved namespace demand',()=>check({[app+'src/host.ts']:"import * as namespace from './scenario';\ndeclare const key:string;\nconst selected=key as keyof typeof namespace;\nexport function mount():void {namespace[selected]();}\n"},'unresolvedDynamicReferences',d=>d.file===app+'src/host.ts'&&d.line===4));
test('transparent assertion satisfies and non-null wrappers do not manufacture finite namespace demand',()=>check({[app+'src/host.ts']:"import * as namespace from './scenario';\ndeclare const key:string;\nexport function mount():void {namespace[((key as keyof typeof namespace)!) satisfies keyof typeof namespace]();}\n"},'unresolvedDynamicReferences',d=>d.file===app+'src/host.ts'&&d.line===3));
test('a genuinely declared finite namespace union and a literal local alias remain supported',()=>check({[app+'src/scenario.ts']:files[app+'src/scenario.ts']+'export const Other=()=>2;\n',[app+'src/host.ts']:"import * as namespace from './scenario';\ndeclare const key:'Demo'|'Other';\nconst literal='Demo' as const;const alias=literal;\n"+files[app+'src/host.ts'].replace('entries.forEach(','namespace[key]();namespace[alias]();entries.forEach(')}));

// FIX1 independent review counterexamples: the valid finite read stays present.
const reviewFix1 = require('./fixtures/demo-reachability/review-fix1.cjs');
const escapedAt = line => d => d.file === app + 'src/host.ts' && d.line === line;
test('FIX1 I1 parenthesized namespace escape remains unresolved beside a finite read', () => {
  check(reviewFix1.namespaceConsumer('namespace.Demo(); leak((namespace))'), 'unresolvedDynamicReferences', escapedAt(4));
});
test('FIX1 I1 asserted unknown namespace alias preserves its broad escape origin', () => {
  check(reviewFix1.namespaceConsumer('namespace.Demo(); leak(alias)', 'const alias = (namespace as unknown);'), 'unresolvedDynamicReferences', escapedAt(5));
});
test('FIX1 I1 satisfies and non-null wrappers preserve namespace escape origin', () => {
  check(reviewFix1.namespaceConsumer('namespace.Demo(); leak(((namespace satisfies unknown)!))'), 'unresolvedDynamicReferences', escapedAt(4));
});
test('FIX1 I1 wrapped namespace spread cannot hide behind a genuine read', () => {
  check(reviewFix1.namespaceConsumer('namespace.Demo(); const copy = { ...((namespace)) }; console.log(copy.Demo)'), 'unresolvedDynamicReferences', escapedAt(4));
});
test('FIX1 I1 asserted alias spread does not consume an unrelated export slot', () => {
  const changes = reviewFix1.namespaceConsumer('namespace.Demo(); const copy = { ...(alias as typeof namespace) }; console.log(copy.Demo)', 'const alias = namespace as unknown;');
  changes[app + 'src/scenario.ts'] = files[app + 'src/scenario.ts'] + 'export const spare = 2;\n';
  const report = check(changes, 'unresolvedDynamicReferences', escapedAt(5));
  assert.ok(report.unusedExports.some(d => d.file === app + 'src/scenario.ts' && d.symbol === 'spare' && d.line === 2));
});
test('FIX1 I1 wrapped immutable namespace alias retains actual finite member demand', () => {
  check(reviewFix1.namespaceConsumer('alias.Demo()', 'const alias = ((namespace as typeof namespace) satisfies typeof namespace)!;'));
});
test('FIX1 I1 wrapped alias can enter a genuine finite declared consumer', () => {
  check(reviewFix1.namespaceConsumer('namespace.Demo(); consume((alias))', 'const alias = (namespace as typeof namespace);\nfunction consume(value: { Demo(): number }): number { return value.Demo(); }'));
});
const nativeProblem = (file, line = 1) => d => d.file === app + file && d.code === 'DEMO_NATIVE_CONFIG' && d.line === line;
test('FIX1 I2 wrong active XML ID cannot borrow the correct registration from a comment', () => {
  const { extensionId, wrongId, location } = reviewFix1;
  check({ [app + 'sharepoint/assets/elements.xml']: `<Elements>\n<!-- ${extensionId} ${location} -->\n<CustomAction Location="${location}" ClientSideComponentId="${wrongId}" />\n</Elements>` }, 'sourceDiagnostics', nativeProblem('sharepoint/assets/elements.xml', 3));
});
test('FIX1 I2 wrong active XML location cannot borrow a comment location', () => {
  const { extensionId, location } = reviewFix1;
  check({ [app + 'sharepoint/assets/elements.xml']: `<Elements>\n<!-- ${location} -->\n<CustomAction Location="Wrong.Extension" ClientSideComponentId="${extensionId}" />\n</Elements>` }, 'sourceDiagnostics', nativeProblem('sharepoint/assets/elements.xml', 3));
});
test('FIX1 I2 deleted Top custom action fails its required serve association', () => {
  check(reviewFix1.serveWithoutAction(), 'sourceDiagnostics', nativeProblem('config/serve.json'));
});
test('FIX1 I2 deleted customActions property cannot bypass native serve validation', () => {
  check(reviewFix1.serveWithoutAction({ removeProperty: true }), 'sourceDiagnostics', nativeProblem('config/serve.json'));
});
test('FIX1 I2 removed feature association fails while the native XML source survives', () => {
  check(reviewFix1.solutionWithoutAssociation(), 'sourceDiagnostics', nativeProblem('config/package-solution.json'));
});
test('FIX1 I2 deleted feature assets cannot bypass required native registration', () => {
  check(reviewFix1.solutionWithoutAssociation({ removeAssets: true }), 'sourceDiagnostics', nativeProblem('config/package-solution.json'));
});
test('FIX1 I2 active prefixed single-quoted native registrations keep both physical associations', () => {
  check(reviewFix1.nativePair());
});
test('FIX1 I2 deleting one of two required XML associations is not rescued by the other', () => {
  const changes = reviewFix1.nativePair();
  const solution = JSON.parse(changes[app + 'config/package-solution.json']);
  solution.solution.features[0].assets.elementManifests = ['instance.xml'];
  changes[app + 'config/package-solution.json'] = JSON.stringify(solution);
  check(changes, 'sourceDiagnostics', nativeProblem('config/package-solution.json'));
});
test('FIX1 I2 active ClientSideComponentInstance ID is validated independently of comments', () => {
  const changes = reviewFix1.nativePair();
  changes[app + 'sharepoint/assets/instance.xml'] = changes[app + 'sharepoint/assets/instance.xml'].replace('ComponentId=\'' + reviewFix1.extensionId, 'ComponentId=\'' + reviewFix1.wrongId) + `\n<!-- ${reviewFix1.extensionId} -->`;
  check(changes, 'sourceDiagnostics', nativeProblem('sharepoint/assets/instance.xml', 3));
});
test('FIX1 I2 malformed qualified tag cannot masquerade as an active native registration', () => {
  const { extensionId, location } = reviewFix1;
  check({ [app + 'sharepoint/assets/elements.xml']: `<sp:Elements xmlns:sp="http://schemas.microsoft.com/sharepoint/">\n<!-- Qualified names have at most one prefix. -->\n<sp:CustomAction:Other Location="${location}" ClientSideComponentId="${extensionId}" />\n</sp:Elements>` }, 'sourceDiagnostics', d => d.file === app + 'sharepoint/assets/elements.xml' && d.code === 'DEMO_NATIVE_CONFIG' && d.line > 0);
});

// Exercise both real npm layers; npm must not consume the audit's --root flag.
test('root npm audit facade forwards explicit fixture roots and diagnostics from a neutral cwd read-only', () => {
  const repository = path.resolve(__dirname, '..');
  for (const invalid of [false, true]) {
    const root = fixture(invalid ? { [app + 'src/orphan.ts']: 'export const orphan = 1;\n' } : {});
    const cwd = fs.mkdtempSync(path.join(os.tmpdir(), 'demo-audit-cwd-'));
    try {
      const before = snapshot(root);
      const cli = spawnSync('npm', ['--prefix', repository, 'run', 'verify:demo-reachability', '--', '--root', root], { cwd, encoding: 'utf8' });
      assert.equal(cli.status, invalid ? 1 : 0, cli.stdout + cli.stderr);
      assert.match(cli.stdout, /Demo reachability: 2 native roots, /);
      if (invalid) assert.match(cli.stderr, /apps\/spfx-react-toolkit-test\/src\/orphan.ts:1 unreachableFiles/);
      else assert.match(cli.stdout, /2 native roots, 0 problems/);
      assert.deepEqual(snapshot(root), before, 'real npm facade preserves fixture bytes and mtime');
      assert.deepEqual(fs.readdirSync(cwd), [], 'neutral cwd remains untouched');
    } finally {
      fs.rmSync(root, { recursive: true, force: true });
      fs.rmSync(cwd, { recursive: true, force: true });
    }
  }
});
test('root npm audit facade preserves unsupported argv and failure exit', () => {
  const cwd = fs.mkdtempSync(path.join(os.tmpdir(), 'demo-audit-cwd-'));
  try {
    const cli = spawnSync('npm', ['--prefix', path.resolve(__dirname, '..'), 'run', 'verify:demo-reachability', '--', '--repair'], { cwd, encoding: 'utf8' });
    assert.equal(cli.status, 1, cli.stdout + cli.stderr);
    assert.match(cli.stderr, /Usage: verify-demo-reachability.cjs \[--root <repository>\]/);
    assert.deepEqual(fs.readdirSync(cwd), []);
  } finally { fs.rmSync(cwd, { recursive: true, force: true }); }
});

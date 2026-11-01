const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const test = require('node:test');
const MarkdownIt = require('markdown-it');
const { buildApiGraph, resolvePublicReference } = require('../scripts/documentation-model.cjs');
const { buildDocumentationModel, bindingContractKey, resolveDocumentationReference } = require('../scripts/documentation-contract.cjs');
const { markdownAnchors } = require('../scripts/documentation.cjs');
const fixture = path.join(__dirname, 'fixtures/documentation/render');
const config = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json', repositoryUrl: 'https://github.com/fixture/render', repositoryRef: 'stable' };
let model;
test.before(() => { model = buildDocumentationModel(buildApiGraph(fixture, config), config); });
function api() {
  assert.ok(fs.existsSync(path.join(__dirname, '../scripts/documentation-render.cjs')), 'pure renderer must exist');
  return require('../scripts/documentation-render.cjs');
}
test('source-owned inherited details and examples link to the retained owner without erasing effective rows', () => {
  const artifacts = api().renderApiPages(model, config);
  const page = artifacts['docs/api/readability.md'];
  // The independent author keeps an equal example; only inherited views share.
  assert.equal(page.split("const sourceValue = 'owned';").length - 1, 2);
  // The standalone callable support entity retains its separate entity contract.
  assert.equal(page.split("const sourceReader = 'owned';").length - 1, 2);
  assert.match(page, /Examples: see \[ExampleBase\.value\]\(readability\.md#examplesexamplebase\)/);
  assert.match(page, /Contract: see \[ExampleBase\.read\]\(readability\.md#examplesexamplebase\)/);
  assert.ok(page.includes('## examples.ExampleChild\n'));
  assert.ok(api().renderCatalog(model, config).includes('[`examples.ExampleChild`]'));
  for (const text of ['ExampleChild.value` | `string', 'Optional, readonly', 'ExampleMutable.value` | `number', 'Required', 'exampleFactory return.value', 'exampleFactory return.forbidden', 'Forbidden', "mode: 'disabled'", 'Local explanation.', "const localValue = 'local';", 'The source owner keeps this explanation.']) assert.ok(page.includes(text), text);
  const { collectPublicTargets, validatePublicLinks } = require('../scripts/documentation-guides.cjs');
  assert.deepEqual(validatePublicLinks(artifacts, collectPublicTargets(fixture, artifacts, config), config), []);
});
test('ownership sharing fails closed for different contracts, origins, missing owners and cyclic ownership', () => {
  const child = model.graph.entities.find(entity => entity.qualifiedName === 'ExampleChild');
  const key = child.key + '::member:read';
  const mutations = [
    contract => { contract.summary += ' Local summary.'; },
    contract => { contract.useWhen.push('Local selection.'); },
    contract => { contract.params[0].description += ' Local parameter.'; },
    contract => { contract.returns += ' Local return.'; },
    contract => { contract.defaultValue = 'Local default.'; },
    contract => { contract.remarks = [{ title: 'Local requirement', body: 'Local remarks.' }]; },
    contract => { contract.typeParams = [{ name: 'Local', description: 'Local type parameter.' }]; },
    contract => { contract.examples = [{ ...contract.examples[0], code: contract.examples[0].code + '\n// Local example' }]; },
    contract => { contract.examples = [{ ...contract.examples[0], title: 'Local title' }]; },
    contract => { contract.examples = [{ ...contract.examples[0], ownerKey: 'unrelated::member:read' }]; },
    contract => { contract.examples = [{ ...contract.examples[0], source: { ...contract.examples[0].source, line: 999 } }]; },
    contract => { contract.source.line += 1; },
    contract => { contract.inheritedFrom = 'missing::member:read'; },
  ];
  for (const mutate of mutations) {
    const copy = structuredClone(model);
    mutate(copy.contracts[key]);
    const page = api().renderApiPages(copy, config)['docs/api/readability.md'];
    const section = page.slice(page.indexOf('## examples.ExampleChild\n'), page.indexOf('\n## ', page.indexOf('## examples.ExampleChild\n') + 1));
    // Changed prose may still share an exact example, but never the whole contract.
    assert.ok(!section.includes('Contract: see'), section);
    assert.ok(section.includes('The input value.') || section.includes('Local parameter.'));
    const readDetail = section.slice(section.indexOf('#### examples.ExampleChild.read'), section.indexOf('#### examples.ExampleChild.value'));
    const example = copy.contracts[key].examples[0];
    const original = model.contracts[key].examples[0];
    if (example.code !== original.code || example.title !== original.title || example.ownerKey !== original.ownerKey || example.source.line !== original.source.line) {
      assert.ok(readDetail.includes(example.code.trimEnd()));
      assert.ok(!readDetail.includes('Examples: see'));
    }
    if (copy.contracts[key].source.line !== model.contracts[key].source.line || copy.contracts[key].inheritedFrom.startsWith('missing')) assert.ok(section.includes("const sourceReader = 'owned';"));
  }
  const copy = structuredClone(model);
  copy.contracts[copy.contracts[key].inheritedFrom].inheritedFrom = key;
  assert.equal(api().renderApiPages(copy, config)['docs/api/readability.md'].split("const sourceReader = 'owned';").length - 1, 6);
  const unrendered = structuredClone(model);
  const ownerKey = unrendered.contracts[key].inheritedFrom.split('::member:')[0];
  for (const page of unrendered.pages) page.entityKeys = page.entityKeys.filter(key => key !== ownerKey);
  const page = api().renderApiPages(unrendered, config)['docs/api/readability.md'];
  const childSection = page.slice(page.indexOf('## examples.ExampleChild\n'), page.indexOf('\n## ', page.indexOf('## examples.ExampleChild\n') + 1));
  assert.ok(childSection.includes("const sourceReader = 'owned';"));
  assert.ok(!childSection.includes('Contract: see'));
});
test('renderer repeats identical bytes after graph, binding and contract map enumeration changes', () => {
  const first = api().renderApiPages(model, config);
  assert.deepEqual(api().renderApiPages(model, config), first);
  const copy = structuredClone(model);
  copy.graph.entities.reverse();
  copy.graph.bindings.reverse();
  copy.contracts = Object.fromEntries(Object.entries(copy.contracts).reverse());
  copy.bindingContracts = Object.fromEntries(Object.entries(copy.bindingContracts).reverse());
  assert.deepEqual(api().renderApiPages(copy, config), first);
  assert.equal(api().renderCatalog(copy, config), api().renderCatalog(model, config));
});
test('module introductions deduplicate only complete resolved content in deterministic source order', () => {
  const copy = structuredClone(model);
  const controls = copy.contracts['src/controls.ts'];
  copy.graph.entities.find(entity => entity.qualifiedName === 'StageBase').source.file = 'src/duplicate.ts';
  copy.contracts['src/duplicate.ts'] = { ...structuredClone(controls), source: { file: 'src/duplicate.ts', line: 1 } };
  const render = () => api().renderApiPages(copy, config)['docs/api/controls/stage.md'];
  assert.equal(render().split('Source-owned stage introduction.').length - 1, 1);
  for (const [field, value] of [['remarks', [{ title: '', body: 'Distinct module remarks.' }]], ['params', [{ name: 'input', description: 'Distinct module parameter.' }]], ['returns', 'Distinct module return.'], ['defaultValue', 'Distinct default.']]) {
    copy.contracts['src/duplicate.ts'][field] = value;
    assert.equal(render().split('Source-owned stage introduction.').length - 1, 2, field);
    copy.contracts['src/duplicate.ts'] = { ...structuredClone(controls), source: { file: 'src/duplicate.ts', line: 1 } };
  }
  copy.contracts['src/controls.ts'].summary += ' [Source](./guide.md)';
  copy.contracts['src/duplicate.ts'].summary = copy.contracts['src/controls.ts'].summary;
  copy.contracts['src/duplicate.ts'].source.file = 'src/nested/units.ts';
  assert.equal(render().split('Source-owned stage introduction.').length - 1, 2);
});
test('renders branch constraints, inherited ownership, source module introductions and legal qualified imports', () => {
  const artifacts = api().renderApiPages(model, config);
  const stage = artifacts['docs/api/controls/stage.md'];
  for (const text of ['Source-owned stage introduction.', 'Unit introduction stays separate.', 'StageProps.open', '"popover"', 'trigger', 'never', 'Promise results are observed', 'rejected Promises reach onError', 'Popover requires a trigger.', 'The caller owns open state.', 'width.px']) assert.ok(stage.includes(text), text);
  assert.ok(stage.indexOf('Source-owned stage introduction.') < stage.indexOf('Unit introduction stays separate.'));
  assert.match(stage, /import \{ width \} from '@fixture\/render';/);
  assert.doesNotMatch(stage, /import \{ px \}/);
  assert.match(stage, /\[StageBase\.open\]\([^)]*#stagebase\)/);
  assert.match(artifacts['docs/api/property-pane/picker.md'], /import \{ PropertyPanePicker \} from '@fixture\/render\/property-pane';/);
  assert.doesNotMatch(artifacts['docs/api/property-pane/picker.md'], /from '@fixture\/render';/);
  assert.match(stage, /T extends string = string/);
  assert.match(stage, /answer: 42/);
  assert.match(stage, /\| `Frozen\.size` \| `number` \| Required, readonly/);
  assert.match(stage, /\| `Mutable\.size` \| `number`/);
  assert.doesNotMatch(stage, /@packageDocumentation|FormRootProps\.onInvalid: any/);
  assert.deepEqual(api().renderApiPages(structuredClone(model), config), artifacts);
});
test('small rendered references match source-owned golden files', () => {
  const artifacts = api().renderApiPages(model, config);
  assert.equal(artifacts['docs/api/property-pane/picker.md'], fs.readFileSync(path.join(fixture, 'picker.expected.md'), 'utf8'));
  const stage = artifacts['docs/api/controls/stage.md'];
  for (const [heading, name] of [['StageProps', 'stage'], ['Derived', 'inherited'], ['action', 'callable'], ['width.px', 'namespace']]) {
    const marker = '\n## ' + heading + '\n';
    const start = stage.indexOf(marker) + 1;
    assert.ok(start > 0, heading);
    const next = stage.indexOf('\n## ', start + 1);
    const section = stage.slice(start, next === -1 ? undefined : next).trimEnd() + '\n';
    assert.equal(section, fs.readFileSync(path.join(fixture, name + '.expected.md'), 'utf8'), name);
  }
});
test('local inherited overrides retain their own defaults and requirements and link upstream ownership', () => {
  const source = path.join(__dirname, 'fixtures/documentation/contracts');
  const inherited = buildDocumentationModel(buildApiGraph(source, config), config);
  const page = api().renderApiPages(inherited, config)['docs/api/inheritance.md'];
  const row = page.split('\n').find(line => line.startsWith('| `Customized.label`'));
  assert.ok(row.includes('Local label.') && row.includes('[Original.label](inheritance.md#original)'));
  assert.ok(page.includes('Requires the local label context.'));
});
test('R3a local overrides and borrowed member prose keep distinct public fragment origins', t => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-override-origin-'));
  fs.cpSync(path.join(__dirname, 'fixtures/documentation/contracts'), root, { recursive: true });
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const file = path.join(root, 'src/inheritance.ts');
  const source = fs.readFileSync(file, 'utf8')
    .replace('/** Original options.', '/** Original options.\n * @docPage original')
    .replace('/** Derived options.', '/** Derived options.\n * @docPage derived')
    .replace('/** Custom label options.', '/** Custom label options.\n * @docPage customized')
    .replace('Upstream label. If omitted, no label is shown.', 'Upstream label. See [base details](#original). If omitted, no label is shown.')
    .replace('Local label. When omitted, a local label is used.', 'Local label. See [custom details](#customized). When omitted, a local label is used.')
    .replace('@defaultValue Local label.', '@defaultValue Local label. See [default details](#customized).')
    .replace('Requires the local label context.', 'Requires the local label context; see [requirements](#customized).');
  fs.writeFileSync(file, source);
  const normalized = buildDocumentationModel(buildApiGraph(root, config), config);
  const artifacts = { ...api().renderApiPages(normalized, config), 'docs/api/INDEX.md': api().renderCatalog(normalized, config) };
  const original = artifacts['docs/api/original.md'];
  const inherited = artifacts['docs/api/derived.md'];
  const customized = artifacts['docs/api/customized.md'];
  assert.ok(original.includes('See [base details](#original).'));
  const inheritedRow = inherited.split('\n').find(line => line.startsWith('| `Derived.label`'));
  assert.ok(inheritedRow.includes('See [base details](original.md#original).'));
  const overrideRow = customized.split('\n').find(line => line.startsWith('| `Customized.label`'));
  assert.ok(overrideRow.includes('See [custom details](#customized).'));
  assert.ok(overrideRow.includes('Local label. See [default details](#customized).'));
  assert.ok(overrideRow.includes('Inherited from [Original.label](original.md#original).'));
  assert.ok(customized.includes('Requires the local label context; see [requirements](#customized).'));
  const { collectPublicTargets, validatePublicLinks } = require('../scripts/documentation-guides.cjs');
  assert.deepEqual(validatePublicLinks(artifacts, collectPublicTargets(root, artifacts, config), config), []);
});
test('branch-specific member requirements remain with the branch rather than being replaced by another branch contract', () => {
  const copy = structuredClone(model);
  const key = Object.keys(copy.contracts).find(key => key.includes('StageProps::member:trigger::variant:'));
  copy.contracts[key].remarks.push({ title: 'Requirements', body: 'This Popover trigger must stay focusable.' });
  const page = api().renderApiPages(copy, config)['docs/api/controls/stage.md'];
  const popover = page.slice(page.indexOf('#### StageProps branch 2'));
  assert.ok(popover.includes('This Popover trigger must stay focusable.'));
  assert.ok(!page.slice(page.indexOf('#### StageProps branch 1'), page.indexOf('#### StageProps branch 2')).includes('This Popover trigger must stay focusable.'));
});
test('every public alias keeps its own heading and catalog row on a shared route', () => {
  const artifacts = api().renderApiPages(model, config);
  for (const page of model.pages) {
    const anchors = markdownAnchors(artifacts[page.route]);
    for (const anchor of [...Object.values(page.anchors), ...Object.values(page.bindingAnchors)]) assert.ok(anchors.has(anchor), anchor);
  }
  const catalog = api().renderCatalog(model, config);
  for (const name of ['Stage', 'action', 'perform', 'width.px']) assert.ok(catalog.includes('[`' + name + '`]'), name);
  assert.equal(catalog.split('\n').filter(line => line.startsWith('| [`action`]')).length, 1);
  const row = catalog.split('\n').find(line => line.startsWith('| [`action`]'));
  assert.ok(row.includes('@fixture/render/controls') && row.includes("'@fixture/render'"));
  const alias = structuredClone(model);
  const action = alias.graph.bindings.find(item => item.exportedName === 'action');
  alias.graph.bindings.push({ ...action, exportedName: 'a' });
  const page = alias.pages.find(item => item.entityKeys.includes(action.entityKey));
  page.bindingAnchors[bindingContractKey(alias.graph.bindings.at(-1))] = 'a';
  const anchors = markdownAnchors(api().renderApiPages(alias, config)[page.route]);
  assert.ok(anchors.has('action') && anchors.has('perform') && anchors.has('a'));
});
test('ambiguous public names are qualified, and conflicting natural anchors fail explicitly', () => {
  const copy = structuredClone(model);
  const picker = copy.graph.bindings.find(item => item.exportedName === 'PropertyPanePicker');
  copy.graph.bindings.push({ ...picker, exportedName: 'Stage' });
  const pickerPage = copy.pages.find(item => item.entityKeys.includes(picker.entityKey));
  pickerPage.bindingAnchors[bindingContractKey(copy.graph.bindings.at(-1))] = 'stage';
  const catalog = api().renderCatalog(copy, config);
  assert.match(catalog, /@fixture\/render\/property-pane#Stage/);
  const collision = structuredClone(model);
  const stage = collision.graph.bindings.find(item => item.exportedName === 'Stage');
  collision.graph.bindings.push({ ...stage, entityKey: picker.entityKey, exportedName: 'STAGE' });
  collision.pages.find(item => item.entityKeys.includes(stage.entityKey)).entityKeys.push(picker.entityKey);
  assert.throws(() => api().renderApiPages(collision, config), /anchor|collision/i);
});
test('renders examples byte-for-byte inside a safe fence and carries normalized contract fields', () => {
  const code = "const nested = `\n```md\n~~~\n`;";
  const rendered = api().renderDocumentationExample({ language: 'ts', code, title: 'Nested' });
  assert.equal(rendered, '#### Nested\n\n````ts\n' + code + '\n````\n');
  const markdown = new MarkdownIt();
  for (const source of ['const value = 1;\n', 'const value = `\n```\n`;\n\n\n']) {
    const output = api().renderDocumentationExample({ language: 'ts', code: source });
    assert.equal(markdown.parse(output, {}).find(token => token.type === 'fence').content, source);
  }
  const copy = structuredClone(model);
  const action = copy.graph.bindings.find(item => item.exportedName === 'action');
  const contract = copy.contracts[action.entityKey];
  contract.throws = ['Rejects invalid input.'];
  contract.deprecated = 'Use another helper.';
  contract.typeParams = [{ name: 'T', description: 'Preserved literal type.' }];
  contract.see = ['[Authored link](https://example.test/real)'];
  contract.examples = [{ language: 'ts', code, title: 'Nested' }];
  const renderedPage = api().renderApiPages(copy, config)['docs/api/controls/stage.md'];
  for (const text of ['Rejects invalid input.', 'Use another helper.', 'Preserved literal type.', '[Authored link](https://example.test/real)', code]) assert.ok(renderedPage.includes(text), text);
});
test('llms uses raw URLs, deduplicates routes only, and limits Optional to the two development indices', () => {
  const pages = [
    { route: 'docs/api/shared.md', title: 'First', summary: 'First API.', area: 'api' },
    { route: 'docs/api/shared.md', title: 'Second', summary: 'Alias.', area: 'api' },
    { route: 'docs/guides/usage.md', title: 'Usage', summary: 'Use it.', area: 'guides' },
    { route: 'docs/development/setup.md', title: 'Setup', summary: '', area: 'development' },
    { route: 'docs/development/INDEX.md', title: 'Contributor orientation', summary: 'Source-owned contributor workflows.', area: 'development' },
    { route: 'docs/development/sharepoint-validation/INDEX.md', title: 'Host evidence', summary: 'Authenticated evidence stays separate.', area: 'development' },
    ...['docs/_sources/source.md', 'docs/fixture/secret.md', 'docs/guides/_sources/secret.md', 'docs/..\\secret.md', '.docs/private.md', 'AGENTS.md', 'tests/fixtures/secret.md', 'llms-full.txt'].map(route => ({ route, title: route, summary: '', area: 'other' })),
  ];
  const llms = api().renderLlms(pages, { name: '@fixture/render', description: 'Fixture package.', version: '1.2.3' }, config);
  assert.equal(llms.split('https://raw.githubusercontent.com/fixture/render/stable/docs/api/shared.md').length - 1, 1);
  assert.match(llms, /## Optional/);
  assert.ok(llms.includes('[Contributor orientation]') && llms.includes('Source-owned contributor workflows.'));
  assert.ok(llms.includes('[Host evidence]') && llms.includes('Authenticated evidence stays separate.'));
  for (const target of ['setup.md', '_sources', '.docs', 'AGENTS', '/docs/fixture/', 'fixtures', 'secret', 'llms-full', '/blob/']) assert.ok(!llms.includes(target), target);
  assert.ok(llms.includes('Version: 1.2.3') && llms.includes('## api') && llms.includes('## guides'));
  assert.ok(llms.includes('/development/INDEX.md') && llms.includes('/sharepoint-validation/INDEX.md'));
  const artifacts = api().renderApiPages(model, config);
  assert.ok(!Object.keys(artifacts).some(file => /llms-full/.test(file)));
});
test('README retains npm blob URLs, local LICENSE and fenced bytes with explicit config', () => {
  const text = api().renderPackageReadme('[Guide](docs/start.md) [License](LICENSE)\r\n\r\n~~~~md\r\n[Code](docs/code.md)\r\n~~~~\r\n', config);
  assert.ok(text.includes('[Guide](https://github.com/fixture/render/blob/stable/docs/start.md) [License](LICENSE)'));
  assert.ok(text.includes('~~~~md\n[Code](docs/code.md)\n~~~~'));
  assert.ok(!text.includes('\r') && text.endsWith('\n') && !text.endsWith('\n\n'));
});
test('real source pilot keeps Stage state/Popover constraints, Promise observations, aliases and generic factory suites truthful', () => {
  const actualConfig = require('../scripts/documentation.config.cjs');
  const root = path.resolve(__dirname, '..');
  const actual = buildDocumentationModel(buildApiGraph(root, actualConfig), actualConfig);
  const artifacts = api().renderApiPages(actual, actualConfig);
  const reference = name => {
    const entity = resolvePublicReference(actual.graph, name);
    return artifacts[actual.pages.find(page => page.entityKeys.includes(entity.key)).route];
  };
  const stage = reference('StageProps');
  assert.match(stage, /StageProps\.open/);
  assert.match(stage, /Authoritative controlled state/);
  assert.match(stage, /Omit open for uncontrolled state/);
  assert.match(stage, /Required integrated accessible trigger/);
  assert.match(stage, /Full is unsupported/);
  const form = reference('FormRootProps');
  assert.match(form, /FormRootProps\.onInvalid/);
  assert.match(form, /A returned promise is awaited despite the notification-style void signature; failures are contained/);
  const factory = reference('createForm');
  assert.match(factory, /createForm return\.Root/);
  assert.match(factory, /TInput extends FieldValues/);
  assert.match(factory, /\n## CreateForm \(type\)\n/);
  assert.match(factory, /\n## createForm\n/);
  assert.equal(resolveDocumentationReference(actual, 'CreateForm').anchor, 'createform-type');
  assert.equal(resolveDocumentationReference(actual, 'createForm').anchor, 'createform');
  assert.equal(resolveDocumentationReference(actual, 'CreateForm').page.route, 'docs/api/controls/form.md');
  assert.ok(markdownAnchors(factory).has('createform-type'));
  const alias = reference('useStableCallback');
  assert.match(alias, /Alias of Fluent UI 9/);
  assert.match(alias, /useEventCallback\.ts/);
  assert.doesNotMatch(alias, /from '@fluentui\/react-utilities';/);
  const catalog = api().renderCatalog(actual, actualConfig);
  assert.match(catalog, /\[`width\.px`\]/);
  assert.ok(!catalog.includes('[`StageProps.open`]'));
});

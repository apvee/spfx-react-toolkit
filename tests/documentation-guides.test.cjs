const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const MarkdownIt = require('markdown-it');
const { buildApiGraph } = require('../scripts/docs/documentation-model.cjs');
const { buildDocumentationModel } = require('../scripts/docs/documentation-contract.cjs');
const render = require('../scripts/docs/documentation-render.cjs');
const fixture = path.join(__dirname, 'fixtures/documentation/guides');
const config = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' };
let model;
test.before(() => {
  model = buildDocumentationModel(buildApiGraph(path.join(__dirname, 'fixtures/documentation/render'), config), config);
  const stage = model.graph.bindings.find(binding => binding.exportedName === 'Stage');
  const contract = model.contracts[stage.entityKey];
  contract.useWhen = ['Read [owner](./owner.md#state) and {@link StageProps}.'];
  contract.avoidWhen = ['Avoid unmanaged state.'];
  contract.examples = [{ title: 'Basic', language: 'tsx', code: 'const text = `{@link Missing} [literal](bad.md)`;\n\n', source: contract.source }];
});
function api() {
  assert.ok(fs.existsSync(path.join(__dirname, '../scripts/docs/documentation-guides.cjs')), 'editorial resolver must exist');
  return require('../scripts/docs/documentation-guides.cjs');
}
function temporary(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-guides-'));
  fs.cpSync(fixture, root, { recursive: true });
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  return root;
}
test('discovers only convention sources, derives metadata/routes and home navigation', () => {
  const sources = api().discoverEditorialSources(fixture, config);
  assert.deepEqual(sources.map(source => source.route), ['docs/INDEX.md', 'docs/development/INDEX.md', 'docs/development/sharepoint-validation/INDEX.md', 'docs/getting-started/start.md', 'docs/guides/usage.md']);
  assert.equal(sources[4].title, 'Usage');
  assert.equal(sources[4].summary, 'Compose source-owned APIs for a concrete task.');
  const pages = api().renderEditorialPages(sources, model, config);
  assert.match(pages['docs/INDEX.md'], /\[Start\]\(getting-started\/start.md\)/);
  assert.match(pages['docs/INDEX.md'], /\[Usage\]\(guides\/usage.md\)/);
  for (const target of ['api/INDEX.md', 'development/INDEX.md', 'development/sharepoint-validation/INDEX.md']) assert.ok(pages['docs/INDEX.md'].includes(target));
  assert.ok(!pages['docs/INDEX.md'].includes('_sources'));
});
test('rejects missing H1/introduction, duplicate destinations and API ownership collisions', t => {
  const root = temporary(t);
  const file = path.join(root, 'docs/_sources/guides/bad.md');
  fs.writeFileSync(file, 'No heading.\n');
  assert.throws(() => api().discoverEditorialSources(root, config), /title|H1/i);
  fs.writeFileSync(file, '# Empty\n\n## A section\n');
  assert.throws(() => api().discoverEditorialSources(root, config), /introduction/i);
  fs.rmSync(file);
  const sources = api().discoverEditorialSources(root, config);
  assert.throws(() => api().renderEditorialPages([...sources, sources[1]], model, config), /collision|duplicate/i);
  const copy = structuredClone(model);
  copy.pages[0].route = sources[1].route;
  assert.throws(() => api().renderEditorialPages(sources, copy, config), /collision/i);
});
test('resolves symbols and normalized inclusions once, protects code and leaves ordinary comments', () => {
  const pages = api().renderEditorialPages(api().discoverEditorialSources(fixture, config), model, config);
  const usage = pages['docs/guides/usage.md'];
  assert.match(usage, /\[Open a stage\]\(\.\.\/api\/controls\/stage.md#stage\)/);
  assert.match(usage, /\[owner\]\(\.\.\/\.\.\/src\/owner.md#state\)/);
  assert.match(usage, /\[StageProps\]\(\.\.\/api\/controls\/stage.md#stageprops\)/);
  assert.ok(usage.includes('<!-- An ordinary comment remains literal. -->'));
  assert.ok(usage.includes('`{@link Missing} [literal](bad.md)`'));
  assert.ok(usage.includes('![Asset](../assets/a\\(b\\).svg)'));
  assert.ok(usage.includes('[manual]: ../development/INDEX.md#manual-heading "Manual"'));
  const tokens = new MarkdownIt().parse(usage, {});
  assert.ok(tokens.some(token => token.type === 'fence' && token.content === 'const text = `{@link Missing} [literal](bad.md)`;\n\n'));
});
test('rejects unknown/ambiguous symbols, selectors, attributes, recursive includes and executable syntax', () => {
  const source = content => [{ source: 'docs/_sources/guides/test.md', route: 'docs/guides/test.md', title: 'Test', summary: 'Test.', area: 'guides', content }];
  for (const directive of [
    '{@link Missing}',
    '<!-- toolkit:include symbol="Missing" section="example" -->',
    '<!-- toolkit:include symbol="Stage" section="signature" -->',
    '<!-- toolkit:include symbol="Stage" section="selection" title="Basic" -->',
    '<!-- toolkit:include symbol="Stage" section="example" execute="yes" -->',
    '<!-- toolkit:include symbol=${process.exit()} section="example" -->',
    '<!-- toolkit:include inventory="localization.resources" symbol="Stage" -->',
    '<!-- toolkit:include inventory="arbitrary" -->',
  ]) assert.throws(() => api().renderEditorialPages(source(directive), model, config), /reference|symbol|section|title|attribute|syntax|inventory|Missing|Unknown/i, directive);
  const copy = structuredClone(model);
  const contract = copy.contracts[copy.graph.bindings.find(binding => binding.exportedName === 'Stage').entityKey];
  contract.examples.push({ ...contract.examples[0], title: 'Second' });
  assert.throws(() => api().renderEditorialPages(source('<!-- toolkit:include symbol="Stage" section="example" -->'), copy, config), /unique|title/i);
  contract.useWhen = ['<!-- toolkit:include symbol="Stage" section="selection" -->'];
  assert.throws(() => api().renderEditorialPages(source('<!-- toolkit:include symbol="Stage" section="selection" -->'), copy, config), /recursive|include/i);
});
test('collects virtual generated targets, validates fragments, duplicate headings and local assets without disk fallback', t => {
  const root = temporary(t);
  const sources = api().discoverEditorialSources(root, config);
  const artifacts = { ...api().renderEditorialPages(sources, model, config), ...render.renderApiPages(model, config), 'docs/api/INDEX.md': '# Catalog\n' };
  // This test targets the known guide links; included source-owned links have their own virtual target.
  artifacts['src/owner.md'] = '# Owner\n\n## State\n';
  artifacts['docs/development/sharepoint-validation/INDEX.md'] = '# Host\n';
  fs.mkdirSync(path.join(root, 'docs/development'), { recursive: true });
  fs.writeFileSync(path.join(root, 'docs/development/INDEX.md'), '# Poisoned output\n');
  fs.writeFileSync(path.join(root, 'docs/development/stale.md'), '# Stale output\n');
  artifacts['docs/guides/usage.md'] += '\n[Missing anchor](../development/INDEX.md#absent)\n';
  fs.mkdirSync(path.join(root, 'docs/api'), { recursive: true });
  fs.writeFileSync(path.join(root, 'docs/api/stale.md'), '# Old generated reference\n');
  const targets = api().collectPublicTargets(root, artifacts, config);
  assert.ok(targets.some(target => target.route === 'docs/development/INDEX.md' && target.content.includes('Manual heading') && target.anchors.includes('manual-heading')));
  assert.ok(!targets.some(target => target.route.includes('_sources') || target.route === 'docs/api/stale.md' || target.route === 'docs/development/stale.md'));
  assert.deepEqual(api().validatePublicLinks(artifacts, targets, config).map(item => item.code), ['DOC_LINK_ANCHOR']);
  artifacts['docs/guides/usage.md'] += '\n[Old output](../api/stale.md)\n';
  assert.deepEqual(api().validatePublicLinks(artifacts, targets, config).map(item => item.code).sort(), ['DOC_LINK_ANCHOR', 'DOC_LINK_TARGET']);
});
test('localization inventory adapter uses actual schema/id/path sorted by ID and rejects malformed schemas', t => {
  const root = temporary(t);
  fs.mkdirSync(path.join(root, 'localization'), { recursive: true });
  const inventoryPath = path.resolve(__dirname, '../packages/spfx-react-toolkit/localization/resources.json');
  fs.copyFileSync(inventoryPath, path.join(root, 'localization/resources.json'));
  fs.writeFileSync(path.join(root, 'docs/_sources/guides/localization.md'), '# Localization\n\nResources.\n\n<!-- toolkit:include inventory="localization.resources" -->\n');
  const sources = api().discoverEditorialSources(root, config);
  const page = api().renderEditorialPages(sources, model, config)['docs/guides/localization.md'];
  const inventory = JSON.parse(fs.readFileSync(inventoryPath, 'utf8'));
  let previous = -1;
  for (const resource of [...inventory.resources].sort((a, b) => a.id < b.id ? -1 : 1)) {
    const index = page.indexOf('`' + resource.id + '`');
    assert.ok(index > previous && page.includes('`' + resource.path + '`'));
    previous = index;
  }
  fs.writeFileSync(path.join(root, 'localization/resources.json'), '{"schemaVersion":2,"resources":[]}');
  assert.throws(() => api().discoverEditorialSources(root, config), /schema|inventory/i);
});
test('shared Markdown transform handles references, parentheses, escapes and code literals in npm README', () => {
  const source = '[A](docs/a(b).md#part) ![B](docs/a\\(b\\).svg) [C][c]\n\n[c]: <docs/a(b).md> "Title"\n\n`[literal](docs/no.md)`\n\n```md\n[code](docs/no.md)\n```\n';
  const output = render.renderPackageReadme(source);
  assert.ok(output.includes('[A](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/a\\(b\\).md#part)'));
  assert.ok(output.includes('[c]: <https://github.com/apvee/spfx-react-toolkit/blob/main/docs/a(b).md> "Title"'));
  assert.ok(output.includes('`[literal](docs/no.md)`'));
  assert.ok(output.includes('```md\n[code](docs/no.md)\n```'));
});
test('API module, member, signature, entity and catalog prose use each normalized source before sharing a destination', () => {
  const copy = structuredClone(model);
  const page = copy.pages.find(page => page.route === 'docs/api/controls/stage.md');
  const files = [...new Set(page.entityKeys.map(key => copy.graph.entities.find(entity => entity.key === key).source.file))];
  assert.ok(files.length > 1);
  for (const [i, file] of files.entries()) {
    copy.contracts[file].source.file = `src/folder${i}/module.ts`;
    copy.contracts[file].summary = `Module ${i} [relative](./read.md) {@link StageProps}.`;
  }
  for (const contract of Object.values(copy.contracts)) if (contract.source.file === 'src/controls.ts') {
    contract.summary += ' {@link StageProps} [extra](./detail.md).';
  }
  const output = render.renderApiPages(copy, config)[page.route];
  assert.match(output, /src\/folder0\/read.md/);
  assert.match(output, /src\/folder1\/read.md/);
  assert.match(output, /src\/detail.md/);
  for (const token of new MarkdownIt().parse(output, {})) if (token.children) {
    for (const child of token.children) if (child.type === 'text') assert.ok(!child.content.includes('{@link'));
  }
  const catalog = render.renderCatalog(copy, config);
  assert.match(catalog, /src\/detail.md/);
  assert.ok(!catalog.includes('{@link'));
});
test('shared parser preserves block/list/table layouts, reference definitions, multiline links and all literal code', () => {
  const { transformMarkdown } = require('../scripts/docs/documentation-markdown.cjs');
  const source = '> [quote](docs/a.md)\n>\n> [q]: docs/q.md\n\n- [list](docs/b.md)\n  continuation [next](docs/c.md)\n\n| Name | Link |\n| --- | --- |\n| A | [table](docs/d.md) |\n\n[wrapped](\n  docs/e.md\n  "title"\n)\n\n    [literal](docs/no.md)\n\n`` [literal](docs/no.md) ` ``\n';
  const output = transformMarkdown(source, { link: target => 'prefix/' + target });
  for (const name of ['a', 'b', 'c', 'd', 'e', 'q']) assert.ok(output.includes('prefix/docs/' + name + '.md'), name);
  assert.ok(output.includes('    [literal](docs/no.md)'));
  assert.ok(output.includes('`` [literal](docs/no.md) ` ``'));
});
test('source spans remain correct for blockquoted includes, escaped table cells and image-label references', () => {
  const { transformMarkdown } = require('../scripts/docs/documentation-markdown.cjs');
  const handlers = { link: target => 'prefix/' + target, include: () => 'Included', reference: () => 'Resolved' };
  assert.equal(transformMarkdown('> <!-- toolkit:include symbol="Stage" section="selection" -->\n', handlers), '> Included\n');
  const table = '| Name | Link |\n| --- | --- |\n| a\\|b | [x](a.md) |\n';
  assert.equal(transformMarkdown(table, handlers), table.replace('(a.md)', '(prefix/a.md)'));
  assert.equal(transformMarkdown('![{@link Stage}](a.md)\n', handlers), '![Resolved](prefix/a.md)\n');
});
test('canonical references reject ambiguous exports and select exact titled examples', () => {
  const sources = [{ source: 'docs/_sources/guides/test.md', route: 'docs/guides/test.md', title: 'Test', summary: 'Test.', area: 'guides', content: '{@link Stage}' }];
  const copy = structuredClone(model);
  const other = copy.graph.bindings.find(binding => binding.exportedName === 'StageProps');
  copy.graph.bindings.push({ ...other, exportedName: 'Stage', memberPath: [] });
  assert.throws(() => api().renderEditorialPages(sources, copy, config), /ambiguous/i);
  const contract = model.contracts[model.graph.bindings.find(binding => binding.exportedName === 'Stage').entityKey];
  const multiple = structuredClone(model);
  multiple.contracts[multiple.graph.bindings.find(binding => binding.exportedName === 'Stage').entityKey].examples.push({ ...contract.examples[0], title: 'Second', code: 'const selected = 2;\n' });
  sources[0].content = '<!-- toolkit:include symbol="Stage" section="example" title="Second" -->';
  const output = api().renderEditorialPages(sources, multiple, config)[sources[0].route];
  assert.ok(output.includes('const selected = 2;'));
  assert.ok(!output.includes('const text ='));
});
test('all API prose fields and authored headings resolve with contract provenance', () => {
  const copy = structuredClone(model);
  const key = copy.graph.bindings.find(binding => binding.exportedName === 'action').entityKey;
  const contract = copy.contracts[key];
  const text = 'Read {@link StageProps} and [details](./details.md).';
  contract.remarks = [{ title: 'Related {@link StageProps}', body: text }];
  contract.deprecated = text;
  contract.useWhen = [text];
  contract.avoidWhen = [text];
  contract.throws = [text];
  contract.see = [text];
  contract.defaultValue = text;
  contract.params = [{ name: 'value', description: text }];
  contract.typeParams = [{ name: 'T', description: text }];
  contract.returns = text;
  const output = render.renderApiPages(copy, config)['docs/api/controls/stage.md'];
  for (const token of new MarkdownIt().parse(output, {})) if (token.children) for (const child of token.children) {
    if (child.type === 'text') assert.ok(!child.content.includes('{@link'), child.content);
  }
  assert.ok(output.split('src/details.md').length > 8);
});
test('qualified hash references round-trip through actual TypeScript JSDoc and pinned strict TSDoc normalization', t => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-qualified-'));
  fs.cpSync(path.join(__dirname, 'fixtures/documentation/render'), root, { recursive: true });
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const file = path.join(root, 'src/controls.ts');
  fs.writeFileSync(file, fs.readFileSync(file, 'utf8').replace('Source-owned stage introduction.', 'Source-owned stage introduction. See {@link @fixture/render/controls#StageProps | Stage inputs}.'));
  const actual = buildDocumentationModel(buildApiGraph(root, config), config);
  assert.deepEqual(actual.diagnostics.filter(item => item.code.startsWith('tsdoc-')), []);
  assert.ok(actual.contracts['src/controls.ts'].summary.includes('{@link @fixture/render/controls#StageProps | Stage inputs}'));
  const output = render.renderApiPages(actual, config)['docs/api/controls/stage.md'];
  assert.ok(output.includes('[Stage inputs](stage.md#stageprops)'));
  assert.ok(!output.includes('{@link'));
});
test('repository directory links use contained targets and virtual README anchors without admitting private sources', t => {
  const root = temporary(t);
  fs.mkdirSync(path.join(root, 'packages/toolkit'), { recursive: true });
  fs.mkdirSync(path.join(root, 'apps/sample'), { recursive: true });
  fs.mkdirSync(path.join(root, '.docs'), { recursive: true });
  fs.writeFileSync(path.join(root, '.docs/private.md'), '# Private\n');
  const artifacts = {
    'packages/toolkit/README.md': '# Virtual package\n\n## Install\n',
    'docs/development/directories.md': '# Directories\n\n[Toolkit](../../packages/toolkit#install) [Sample](../../apps/sample) [Private](../../.docs/private.md) [Source](../_sources/guides/usage.md)\n',
  };
  const targets = api().collectPublicTargets(root, artifacts, config);
  assert.ok(!targets.some(target => target.route === '.docs/private.md'));
  const diagnostics = api().validatePublicLinks(artifacts, targets, config);
  assert.equal(diagnostics.length, 2);
  assert.ok(diagnostics.every(item => item.code === 'DOC_LINK_TARGET' && /Private|_sources|\.docs/.test(item.message)));
});
test('home navigation summaries resolve from each discovered page provenance', () => {
  const sources = api().discoverEditorialSources(fixture, config);
  sources.find(source => source.area === 'guides').summary = 'Use {@link StageProps} and [setup](../getting-started/start.md).';
  const home = api().renderEditorialPages(sources, model, config)['docs/INDEX.md'];
  assert.ok(home.includes('[StageProps](api/controls/stage.md#stageprops)'));
  assert.ok(home.includes('[setup](getting-started/start.md)'));
  assert.ok(!home.includes('{@link'));
});
test('round1 I1 heading anchors use rendered inline text across links, images, code and entities', t => {
  const root = temporary(t);
  const manual = '# [Read][ref]\n\n## [Read](a(b).md)\n\n## A &amp; B\n\n## `Code` ![Image](asset.png) &lt;tag&gt;\n\n[ref]: target.md\n';
  const artifacts = { 'docs/guides/links.md': '# Links\n\n[One](../development/headings.md#read) [Two](../development/headings.md#read-1) [Entity](../development/headings.md#a--b) [Literals](../development/headings.md#code-image-tag)\n' };
  const targets = api().collectPublicTargets(root, { ...artifacts, 'docs/development/headings.md': manual }, config);
  assert.deepEqual(targets.find(target => target.route === 'docs/development/headings.md').anchors, ['read', 'read-1', 'a--b', 'code-image-tag']);
  assert.deepEqual([...require('../scripts/docs/documentation.cjs').markdownAnchors(manual)], ['read', 'read-1', 'a--b', 'code-image-tag']);
  assert.deepEqual(api().validatePublicLinks(artifacts, targets, config), []);
  artifacts['docs/guides/links.md'] += '\n[False anchor](../development/headings.md#readref)\n';
  assert.deepEqual(api().validatePublicLinks(artifacts, targets, config).map(item => item.code), ['DOC_LINK_ANCHOR']);
});
test('round1 I2 stale disk llms.txt is excluded while a supplied virtual llms.txt is valid', t => {
  const root = temporary(t);
  fs.writeFileSync(path.join(root, 'llms.txt'), 'Old generated content.\n');
  const artifacts = { 'docs/INDEX.md': '# Home\n\n[LLM index](../llms.txt)\n' };
  const targets = api().collectPublicTargets(root, artifacts, config);
  assert.ok(!targets.some(target => target.route === 'llms.txt'));
  assert.deepEqual(api().validatePublicLinks(artifacts, targets, config).map(item => item.code), ['DOC_LINK_TARGET']);
  artifacts['llms.txt'] = '# Virtual LLM index\n';
  assert.deepEqual(api().validatePublicLinks(artifacts, api().collectPublicTargets(root, artifacts, config), config), []);
});
test('round1 R1 encoded URI delimiters remain path characters and fragments decode independently', t => {
  const root = temporary(t);
  fs.writeFileSync(path.join(root, 'docs/assets/image#one.png'), 'image');
  fs.writeFileSync(path.join(root, 'docs/assets/image?two.png'), 'image');
  const artifacts = { 'docs/guides/links.md': '# Links\n\n![Hash](../assets/image%23one.png?download=1) ![Query](../assets/image%3Ftwo.png) [Unicode](../development/unicode.md#citt%C3%A0) [Escaped](../development/unicode.md#escaped\\-heading)\n' };
  artifacts['docs/development/unicode.md'] = '# Città\n\n## Escaped heading\n';
  assert.deepEqual(api().validatePublicLinks(artifacts, api().collectPublicTargets(root, artifacts, config), config), []);
  artifacts['docs/guides/links.md'] += '\n[Malformed](../assets/%ZZ.png)\n';
  assert.deepEqual(api().validatePublicLinks(artifacts, api().collectPublicTargets(root, artifacts, config), config).map(item => item.code), ['DOC_LINK_TARGET']);
});
test('round1 R2 rejects outside and private canonical targets through ancestor or leaf symlinks before reads', t => {
  const root = temporary(t);
  const outside = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-outside-'));
  t.after(() => fs.rmSync(outside, { recursive: true, force: true }));
  fs.writeFileSync(path.join(outside, 'private.md'), '# Outside private data\n');
  fs.symlinkSync(outside, path.join(root, 'docs/assets/linked'));
  fs.symlinkSync(path.join(outside, 'private.md'), path.join(root, 'docs/assets/leaf.md'));
  fs.mkdirSync(path.join(root, '.docs'));
  fs.writeFileSync(path.join(root, '.docs/private.md'), '# Private evidence\n');
  fs.symlinkSync(path.join(root, '.docs'), path.join(root, 'docs/assets/private'));
  fs.symlinkSync(path.join(root, 'docs/_sources'), path.join(root, 'docs/assets/authoring'));
  for (const href of ['../assets/linked/private.md', '../assets/leaf.md', '../assets/private/private.md', '../assets/authoring/guides/usage.md']) {
    const artifacts = { 'docs/guides/links.md': '# Links\n\n[Forbidden](' + href + ')\n' };
    assert.throws(() => api().collectPublicTargets(root, artifacts, config), error => error.code === 'DOC_OUTSIDE_INPUT' && /outside|private|source/i.test(error.message), href);
  }
  // Canonicalize the root too: a symlink spelling of the authorized checkout
  // must not make ordinary contained files or directory targets fail.
  const alias = path.join(outside, 'authorized-root');
  fs.symlinkSync(root, alias);
  fs.mkdirSync(path.join(root, 'assets'));
  fs.writeFileSync(path.join(root, 'assets/ok.png'), 'image');
  fs.symlinkSync(path.join(root, 'assets'), path.join(root, 'docs/assets/contained'));
  const artifacts = { 'docs/guides/links.md': '# Links\n\n![Contained](../assets/contained/ok.png) [Directory](../../assets)\n' };
  assert.deepEqual(api().validatePublicLinks(artifacts, api().collectPublicTargets(alias, artifacts, config), config), []);
});
test('round1 I1 canonical legal underscore symbol headings retain model/render/checker agreement', () => {
  const graph = structuredClone(model.graph);
  const key = graph.bindings.find(binding => binding.exportedName === 'answer').entityKey;
  for (const binding of graph.bindings.filter(binding => binding.entityKey === key)) binding.exportedName = '_answer_';
  const actual = buildDocumentationModel(graph, config);
  const page = actual.pages.find(page => page.entityKeys.includes(key));
  const output = render.renderApiPages(actual, config)[page.route];
  const anchors = require('../scripts/docs/documentation-markdown.cjs').markdownAnchors(output);
  assert.ok(anchors.includes(page.anchors[key]), page.anchors[key]);
  assert.ok(require('../scripts/docs/documentation.cjs').markdownAnchors(output).has(page.anchors[key]));
});
test('round1 R2 discovery and localization inventory reads also reject outside canonical inputs', t => {
  const outside = fs.mkdtempSync(path.join(os.tmpdir(), 'documentation-outside-source-'));
  t.after(() => fs.rmSync(outside, { recursive: true, force: true }));
  const root = temporary(t);
  fs.cpSync(path.join(root, 'docs/_sources'), path.join(outside, 'sources'), { recursive: true });
  fs.rmSync(path.join(root, 'docs/_sources'), { recursive: true });
  fs.symlinkSync(path.join(outside, 'sources'), path.join(root, 'docs/_sources'));
  assert.throws(() => api().discoverEditorialSources(root, config), error => error.code === 'DOC_OUTSIDE_INPUT');
  const inventoryRoot = temporary(t);
  fs.mkdirSync(path.join(inventoryRoot, 'localization'));
  fs.copyFileSync(path.resolve(__dirname, '../packages/spfx-react-toolkit/localization/resources.json'), path.join(outside, 'resources.json'));
  fs.symlinkSync(path.join(outside, 'resources.json'), path.join(inventoryRoot, 'localization/resources.json'));
  fs.writeFileSync(path.join(inventoryRoot, 'docs/_sources/guides/localization.md'), '# Localization\n\nResources.\n\n<!-- toolkit:include inventory="localization.resources" -->\n');
  assert.throws(() => api().discoverEditorialSources(inventoryRoot, config), error => error.code === 'DOC_OUTSIDE_INPUT');
});
test('round2 R2 discovery permits authoring sources and root aliases but always rejects canonical evidence', t => {
  const root = temporary(t);
  const alias = path.join(root, 'same-checkout');
  fs.symlinkSync(root, alias);
  assert.deepEqual(api().discoverEditorialSources(alias, config).map(source => source.route), ['docs/INDEX.md', 'docs/development/INDEX.md', 'docs/development/sharepoint-validation/INDEX.md', 'docs/getting-started/start.md', 'docs/guides/usage.md']);
  fs.mkdirSync(path.join(root, '.docs/hidden'), { recursive: true });
  fs.writeFileSync(path.join(root, '.docs/hidden/INDEX.md'), '# Private evidence\n\nMust never become public documentation.\n');
  fs.rmSync(path.join(root, 'docs/_sources'), { recursive: true });
  fs.symlinkSync(path.join(root, '.docs/hidden'), path.join(root, 'docs/_sources'));
  for (const spelling of [root, alias]) assert.throws(() => api().discoverEditorialSources(spelling, config),
    error => error.code === 'DOC_OUTSIDE_INPUT' && error.file === 'docs/_sources' && /evidence/i.test(error.message));
});

test('recursive development metadata and nested link provenance come exclusively from source', () => {
  const sources = api().discoverEditorialSources(fixture, config);
  const source = sources.find(item => item.route === 'docs/development/sharepoint-validation/INDEX.md');
  assert.ok(source);
  assert.equal(source.area, 'development');
  assert.equal(source.title, 'Host checks');
  assert.match(source.summary, /Read \[development\]/);
  const page = api().renderEditorialPages(sources, model, config)[source.route];
  assert.ok(page.includes('[development](../INDEX.md#manual-heading)'));
  assert.ok(page.includes('[start](../../getting-started/start.md)'));
  assert.ok(page.includes('[repository](../../../package.json)'));
  assert.ok(page.includes('[details](#details)'));
  assert.ok(page.includes('[owner](../../../src/owner.md#state)'));
  assert.ok(page.includes('[StageProps](../../api/controls/stage.md#stageprops)'));
  assert.ok(!page.includes('_sources'));
});
test('stale development targets and canonical aliases cannot satisfy missing virtual routes', t => {
  const root = temporary(t);
  fs.mkdirSync(path.join(root, 'docs/development'), { recursive: true });
  fs.writeFileSync(path.join(root, 'docs/development/stale.md'), '# Poisoned\n');
  fs.symlinkSync(path.join(root, 'docs/development/stale.md'), path.join(root, 'stale-alias.md'));
  const artifacts = { 'docs/guides/links.md': '# Links\n\n[Old](../development/stale.md#poisoned) [Alias](../../stale-alias.md#poisoned)\n' };
  const targets = api().collectPublicTargets(root, artifacts, config);
  assert.ok(!targets.some(target => ['docs/development/stale.md', 'stale-alias.md'].includes(target.route)));
  assert.deepEqual(api().validatePublicLinks(artifacts, targets, config).map(item => item.code), ['DOC_LINK_TARGET', 'DOC_LINK_TARGET']);
});

test('renders scenario code only in guides that select its exact local title', () => {
  const copy = structuredClone(model);
  const contract = copy.contracts[copy.graph.bindings.find(binding => binding.exportedName === 'Stage').entityKey];
  const scenario = { title: 'Basic', language: 'tsx', code: 'const scenarioContent = <div />;\n', source: contract.source };
  contract.scenarios = [scenario];
  const source = { source: 'docs/_sources/guides/test.md', route: 'docs/guides/test.md', title: 'Test', summary: 'Test.', area: 'guides' };
  const include = (section, title) => api().renderEditorialPages([{ ...source, content:
    '<!-- toolkit:include symbol="Stage" section="' + section + '"' + (title === undefined ? '' : ' title="' + title + '"') + ' -->' }], copy, config)[source.route];
  assert.match(include('scenario', 'Basic'), /const scenarioContent/);
  assert.doesNotMatch(include('example', 'Basic'), /const scenarioContent/);
  assert.doesNotMatch(render.renderApiPages(copy, config)['docs/api/controls/stage.md'], /const scenarioContent/);
  assert.throws(() => include('scenario'), /exact scenario title/);
  assert.throws(() => include('scenario', '  '), /exact scenario title/);
  assert.throws(() => include('scenario', 'basic'), /Missing or ambiguous scenario/);
  contract.scenarios.push(scenario);
  assert.throws(() => include('scenario', 'Basic'), /Missing or ambiguous scenario/);
});

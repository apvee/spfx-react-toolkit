const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const { buildApiGraph, resolvePublicReference } = require('../scripts/docs/documentation-model.cjs');
const fixture = path.join(__dirname, 'fixtures/documentation/contracts');
const config = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' };
let graph;
test.before(() => { graph = buildApiGraph(fixture, config); });
function api() {
  assert.ok(fs.existsSync(path.join(__dirname, '../scripts/docs/documentation-contract.cjs')), 'normalized contract module must exist');
  return require('../scripts/docs/documentation-contract.cjs');
}
function build(input = graph) { return api().buildDocumentationModel(input, config); }
function get(model, reference) { return model.contracts[resolvePublicReference(model.graph, reference).key]; }
function mutate(reference, rawDoc) {
  const copy = structuredClone(graph);
  copy.entities.find(item => item.key === resolvePublicReference(copy, reference).key).rawDoc = rawDoc;
  return copy;
}
function codes(model) { return api().validateDocumentation(model).map(item => item.code); }

test('normalizes source contracts, custom tags and legacy default once for renderer and guide consumers', () => {
  const model = build();
  const contract = get(model, '@fixture/contracts#paginate');
  assert.equal(contract.summary, 'Produces a page number.');
  assert.deepEqual(contract.useWhen, ['You own results and need controlled pages.']);
  assert.match(contract.avoidWhen[0], /server fetching/);
  assert.deepEqual(contract.params, [{ name: 'value', description: 'Requested page number.' }]);
  assert.equal(contract.returns, 'The requested page.');
  assert.deepEqual(contract.remarks, [{ title: 'Requirements', body: 'No provider is required.' }]);
  const options = resolvePublicReference(graph, 'PagingOptions');
  assert.equal(model.contracts[api().memberContractKey(options.key, 'enabled')].defaultValue, 'true');
  const selected = api().getDocumentationSelection(model, '@fixture/contracts#paginate');
  assert.strictEqual(selected, contract);
  assert.strictEqual(api().getDocumentationExample(model, '@fixture/contracts#paginate', 'Basic'), contract.examples[0]);
  assert.deepEqual(api().validateDocumentation(model), []);
});

test('separates package headers from first declaration docs with exact source provenance', () => {
  const entity = resolvePublicReference(graph, '@fixture/contracts#paginate');
  assert.equal(entity.docSource.line, 6);
  assert.equal(entity.moduleDocSource.line, 1);
  assert.doesNotMatch(entity.rawDoc, /packageDocumentation/);
  const model = build();
  assert.equal(get(model, '@fixture/contracts#paginate').source.line, 6);
  assert.equal(model.contracts['src/pagination.ts'].docPage, 'controls/pagination');
  assert.equal(get(model, '@fixture/contracts/other#paginate').source.line, 1);
});
test('module modifiers never leak into prose, while literal modifier text stays in prose and examples', () => {
  const copy = structuredClone(graph);
  const entity = copy.entities.find(item => item.qualifiedName === 'paginate' && item.source.file === 'src/pagination.ts');
  entity.moduleDoc = '/** Source introduction.\n * @packageDocumentation\n * @remarks\n * Literal `@packageDocumentation` remains.\n * @example Literal\n * ```ts\n * const text = "@packageDocumentation";\n * ```\n */';
  const model = build(copy);
  assert.equal(model.contracts['src/pagination.ts'].summary, 'Source introduction.');
  assert.equal(model.contracts['src/pagination.ts'].remarks[0].body, 'Literal `@packageDocumentation` remains.');
  assert.ok(model.contracts['src/pagination.ts'].examples[0].code.includes('"@packageDocumentation"'));
});
test('keeps authored see references alongside throws instead of losing TSDoc dedicated see blocks', () => {
  const model = build(mutate('@fixture/contracts#paginate', '/** A helper.\n * @throws Rejects invalid input.\n * @see https://example.test/upstream\n * @see [Other contract](https://example.test/other)\n */'));
  const contract = get(model, '@fixture/contracts#paginate');
  assert.deepEqual(contract.throws, ['Rejects invalid input.']);
  assert.deepEqual(contract.see, ['https://example.test/upstream', '[Other contract](https://example.test/other)']);
});

test('groups implementation and types on docPage without moving reexports to an aggregate barrel', () => {
  const model = build();
  const page = model.pages.find(item => item.route === 'docs/api/controls/pagination.md');
  assert.ok(page.entityKeys.includes(resolvePublicReference(graph, 'PagingOptions').key));
  assert.ok(page.entityKeys.includes(resolvePublicReference(graph, '@fixture/contracts#paginate').key));
  assert.ok(!model.pages.some(item => item.route === 'docs/api/aggregate.md'));
  assert.ok(model.pages.some(item => item.route === 'docs/api/other.md'));
  assert.equal(new Set(Object.values(page.anchors)).size, page.entityKeys.length);
});

test('binding comments apply only to that alias and precise references resolve ambiguous short names', () => {
  const model = build();
  assert.equal(api().getDocumentationSelection(model, 'navigate').summary, 'Alias-specific navigation binding.');
  assert.match(api().getDocumentationSelection(model, 'pages').summary, /Produces a page/);
  assert.throws(() => api().getDocumentationSelection(model, 'paginate'), /Ambiguous/);
  assert.match(api().getDocumentationSelection(model, '@fixture/contracts/other#paginate').summary, /another entry/);
  assert.equal(Object.keys(model.bindingContracts).filter(key => key.includes('#navigate')).length, 1);
});

test('reports unknown tags and nonexistent parameters/generics without silencing malformed legacy contracts', () => {
  const model = build(mutate('@fixture/contracts#paginate', `/** A helper.
   * @useWhen Choose it.
   * @mystery Bad tag.
   * @param missing - Invalid name.
   * @template Absent - Invalid generic.
   * @returns A number.
   */`));
  assert.ok(codes(model).includes('tsdoc-undefined-tag'));
  assert.ok(codes(model).includes('DOC_UNKNOWN_PARAM'));
  assert.ok(codes(model).includes('DOC_UNKNOWN_TYPE_PARAM'));
  assert.ok(codes(model).includes('DOC_MISSING_EXAMPLE'));
  assert.ok(codes(model).includes('DOC_MISSING_PARAM'));
});

test('requires explicit optional omission while useWhen is only required on directly selected runtime APIs', () => {
  const copy = structuredClone(graph);
  const options = copy.entities.find(item => item.qualifiedName === 'PagingOptions');
  options.members[0].rawDoc = '/** Navigation flag. @defaultValue true */';
  let model = build(copy);
  assert.ok(codes(model).includes('DOC_MISSING_OMISSION'));
  assert.ok(!api().validateDocumentation(model).some(item => item.code === 'DOC_MISSING_USE_WHEN' && /PagingOptions|Part/.test(item.symbol)));
  model = build(mutate('@fixture/contracts#paginate', '/** A helper. */'));
  assert.ok(codes(model).includes('DOC_MISSING_USE_WHEN'));
  assert.ok(!codes(model).includes('DOC_MISSING_AVOID_WHEN'));
});

test('requires fenced examples and unique local titles while ignoring tags inside code fences', () => {
  const model = build(mutate('@fixture/contracts#paginate', `/** A helper.
   * @useWhen Choose it.
   * @example Duplicate
   * ~~~ts
   * const text = '@docPage ../escape';
   * ~~~
   * @example Duplicate
   * \`\`\`tsx
   * const text = '@unknown';
   * \`\`\`
   */`));
  assert.ok(codes(model).includes('DOC_DUPLICATE_EXAMPLE_TITLE'));
  assert.ok(!codes(model).includes('DOC_INVALID_PAGE'));
  assert.ok(!codes(model).includes('tsdoc-undefined-tag'));
  assert.equal(get(model, '@fixture/contracts#paginate').examples.length, 2);
  const invalid = build(mutate('@fixture/contracts#paginate', '/** Helper.\n * @example prose only\n */'));
  assert.ok(codes(invalid).includes('DOC_INVALID_EXAMPLE'));
});

test('rejects unsafe pages, conflicting alias assignments, and case-insensitive route collisions', () => {
  for (const route of ['../escape', '/absolute', 'a/../b', 'page.md', 'page#x', 'page?x', 'a\\b', 'a//b']) {
    assert.ok(codes(build(mutate('@fixture/contracts#paginate', `/** Helper.\n * @docPage ${route}\n */`))).includes('DOC_INVALID_PAGE'), route);
  }
  const copy = structuredClone(graph);
  copy.bindings.find(item => item.exportedName === 'navigate').rawDoc = '/** Alias.\n * @docPage different\n */';
  assert.ok(codes(build(copy)).includes('DOC_PAGE_CONFLICT'));
  const collision = mutate('@fixture/contracts/other#paginate', '/** Other.\n * @docPage Controls/Pagination\n */');
  assert.ok(codes(build(collision)).includes('DOC_PAGE_COLLISION'));
});

test('namespace and compound parts inherit owner routes and selection without requiring repeated examples', () => {
  const model = build();
  const parent = resolvePublicReference(graph, 'widget');
  const child = resolvePublicReference(graph, 'widget.Part');
  const page = model.pages.find(item => item.entityKeys.includes(parent.key));
  assert.ok(page.entityKeys.includes(child.key));
  assert.deepEqual(api().getDocumentationSelection(model, 'widget.Part').useWhen, get(model, 'widget').useWhen);
  assert.ok(!api().validateDocumentation(model).some(item => item.symbol === child.key && /MISSING_EXAMPLE|MISSING_USE_WHEN/.test(item.code)));
});

test('inherited members retain upstream source and link while local documentation stays local', () => {
  const model = build();
  const original = resolvePublicReference(graph, 'Original');
  const derived = resolvePublicReference(graph, 'Derived');
  const inherited = model.contracts[api().memberContractKey(derived.key, 'label')];
  assert.equal(inherited.inheritedFrom, api().memberContractKey(original.key, 'label'));
  assert.equal(inherited.source.file, 'src/inheritance.ts');
  assert.equal(inherited.defaultValue, undefined);
  assert.equal(inherited.remarks.length, 0);
  assert.equal(model.contracts[api().memberContractKey(derived.key, 'size')].remarks[0].title, 'Inherited behavior');
});

test('diagnostics and models are deterministic, serializable and use relative source coordinates', () => {
  const model = build(mutate('@fixture/contracts#paginate', '/** Broken.\n * @unknown Bad.\n */'));
  assert.equal(JSON.stringify(model), JSON.stringify(build(model.graph)));
  assert.ok(!JSON.stringify(model).includes(fixture));
  const diagnostics = api().validateDocumentation(model);
  const sorted = [...diagnostics].sort((a, b) => a.file.localeCompare(b.file, 'en') || a.line - b.line || a.code.localeCompare(b.code, 'en') || a.message.localeCompare(b.message, 'en'));
  assert.deepEqual(diagnostics, sorted);
});

test('normalizes generic callable signatures and returned suite members without requiring an example for each support member', () => {
  const semantic = buildApiGraph(path.join(__dirname, 'fixtures/documentation/semantic'), config);
  const model = build(semantic);
  const suite = resolvePublicReference(semantic, 'makeSuite').signatures[0];
  assert.ok(model.contracts[api().memberContractKey(suite.key + '::returns', 'Field')]);
  const genericModel = build();
  const generic = resolvePublicReference(graph, 'identity');
  assert.ok(!api().validateDocumentation(genericModel).some(item => item.symbol === generic.key && item.code === 'DOC_UNKNOWN_TYPE_PARAM'));
  assert.equal(get(genericModel, 'identity').typeParams[0].name, 'T');
});

test('reports duplicate or empty remarks sections and incomplete example titles', () => {
  const model = build(mutate('@fixture/contracts#paginate', '/** Helper.\n * @remarks\n * ### Requirements\n * One provider.\n * ### Requirements\n * @example\n * ```ts\n * const a = 1;\n * ```\n * @example Named\n * ```ts\n * const b = 2;\n * ```\n */'));
  assert.ok(codes(model).includes('DOC_DUPLICATE_REMARK'));
  assert.ok(codes(model).includes('DOC_EMPTY_REMARK'));
  assert.ok(codes(model).includes('DOC_MISSING_EXAMPLE_TITLE'));
});

test('a binding-local contract is validated rather than silently accepting missing selection or examples', () => {
  const copy = structuredClone(graph);
  copy.bindings.find(item => item.exportedName === 'navigate').rawDoc = '/** Local alias. */';
  const diagnostics = api().validateDocumentation(build(copy));
  assert.ok(diagnostics.some(item => item.symbol === '@fixture/contracts#navigate' && item.code === 'DOC_MISSING_USE_WHEN'));
  assert.ok(diagnostics.some(item => item.symbol === '@fixture/contracts#navigate' && item.code === 'DOC_MISSING_EXAMPLE'));
});

test('declaration page overrides file default and preserves source filename dots in fallback routes', () => {
  const model = build(mutate('@fixture/contracts#paginate', '/** Helper.\n * @docPage controls/specific\n */'));
  assert.ok(model.pages.some(page => page.route === 'docs/api/controls/specific.md' && page.entityKeys.includes(resolvePublicReference(graph, '@fixture/contracts#paginate').key)));
  const copy = structuredClone(graph);
  const item = copy.entities.find(item => item.key === resolvePublicReference(copy, 'identity').key);
  item.source.file = 'src/identity.types.ts';
  assert.ok(build(copy).pages.some(page => page.route === 'docs/api/identity.types.md'));
});

test('normalizes inherited selection in the model before any renderer or guide accesses it', () => {
  const model = build();
  const child = get(model, 'widget.Part');
  const parent = get(model, 'widget');
  assert.deepEqual(child.useWhen, parent.useWhen);
  assert.strictEqual(api().getDocumentationSelection(model, 'widget.Part'), child);
});

test('local overrides keep upstream links but preserve locally authored defaults and requirements', () => {
  const model = build();
  const derived = resolvePublicReference(graph, 'Customized');
  const own = model.contracts[api().memberContractKey(derived.key, 'label')];
  assert.equal(own.inheritedFrom, api().memberContractKey(resolvePublicReference(graph, 'Original').key, 'label'));
  assert.equal(own.defaultValue, 'Local label.');
  assert.match(own.remarks[0].body, /local label context/);
});

test('preserves arbitrary Markdown fence widths and treats tag-looking code as code', () => {
  for (const fence of ['````', '~~~~']) {
    const rawDoc = ['/** Helper.', ' * @useWhen Choose it.', ' * @example Wide', ` * ${fence}ts`, ' * const text = `', ' * @docPage ../not-a-real-tag', ' * @unknown', ' * ```', ' * `;', ` * ${fence}`, ' */'].join('\n');
    const model = build(mutate('@fixture/contracts#paginate', rawDoc));
    assert.ok(!codes(model).includes('tsdoc-undefined-tag'), fence);
    assert.ok(!codes(model).includes('DOC_INVALID_PAGE'), fence);
    const examples = get(model, '@fixture/contracts#paginate').examples;
    assert.equal(examples.length, 1, fence);
    assert.match(examples[0].code, /@docPage \.\.\/not-a-real-tag\n@unknown\n```/);
  }
});

test('local callable aliases own their signature documentation instead of peer declarations', () => {
  const semantic = buildApiGraph(path.join(__dirname, 'fixtures/documentation/semantic'), config);
  const alias = resolvePublicReference(semantic, 'publicCallback');
  const model = build(semantic);
  const signature = model.contracts[alias.signatures[0].key];
  assert.match(signature.summary, /Locally documented alias/);
  assert.doesNotMatch(signature.summary, /Peer contract/);
  assert.equal(signature.source.file, alias.source.file);
});

test('documented anonymous property callbacks keep source summaries in normalized contracts', () => {
  const semantic = buildApiGraph(path.join(__dirname, 'fixtures/documentation/semantic'), config);
  const model = build(semantic);
  for (const [name, summary, line] of [
    ['tooltip', 'Tooltip for the original value.', 5],
    ['parenthesized', 'Explicit parenthesized type.', 13],
    ['union', 'Explicit union type.', 15],
    ['intersection', 'Explicit intersection type.', 17],
    ['nested', 'Explicit nearest wrapper type.', 20],
  ]) {
    const callback = semantic.entities.find(item => item.qualifiedName === `CallbackOwner.${name}`);
    assert.equal(model.contracts[callback.key].summary, summary);
    assert.deepEqual(model.contracts[callback.key].source, { file: 'src/member-docs.ts', line });
    assert.ok(!api().validateDocumentation(model).some(item => item.symbol === callback.key && item.code === 'DOC_MISSING_SUMMARY'));
  }
});

test('Omit and intersection members link their real support owner without inheriting contextual defaults', () => {
  const semantic = buildApiGraph(path.join(__dirname, 'fixtures/documentation/semantic'), config);
  const surface = resolvePublicReference(semantic, 'SurfaceProps');
  const kept = surface.variants[0].members.find(item => item.name === 'kept');
  const base = semantic.entities.find(item => item.qualifiedName === 'Base');
  assert.ok(base, 'source owner must remain available as a support entity');
  assert.equal(kept.ownerKey, base.key);
  const model = build(semantic);
  assert.equal(model.contracts[api().memberContractKey(surface.key, 'kept')].inheritedFrom, api().memberContractKey(base.key, 'kept'));
});

test('explicit docPage values reject arbitrary filename extensions', () => {
  for (const page of ['controls/example.html', 'controls/example.types', 'example.json']) {
    assert.ok(codes(build(mutate('@fixture/contracts#paginate', `/** Helper.\n * @docPage ${page}\n */`))).includes('DOC_INVALID_PAGE'));
  }
});

test('source examples are normalized once even when entity and signature share one declaration comment', () => {
  const model = build();
  const examples = model.examples.filter(example => example.source.file === 'src/pagination.ts');
  assert.equal(examples.length, 1);
  const entity = resolvePublicReference(graph, '@fixture/contracts#paginate');
  assert.strictEqual(model.contracts[entity.signatures[0].key].examples[0], get(model, '@fixture/contracts#paginate').examples[0]);
});

test('example provenance starts at the first physical code line after optional titles and blank margins', () => {
  const cases = [
    { body: ' * @example\n * ```ts\n * const value = 1;\n * ```', line: 9, code: 'const value = 1;\n' },
    { body: ' * @example Named\n * ```ts\n * const value = 1;\n * ```', line: 9, code: 'const value = 1;\n' },
    { body: ' * @example\n *\n\n * ```ts\n * const value = 1;\n * ```', line: 11, code: 'const value = 1;\n' },
    { body: ' * @example\n *\n * Named\n *\n * ```ts\n * const value = 1;\n * ```', line: 12, code: 'const value = 1;\n' },
    { body: ' * @example\n * ```ts\n *\n * const value = 1;\n * ```', line: 9, code: '\nconst value = 1;\n' },
  ];
  for (const item of cases) {
    const example = get(build(mutate('@fixture/contracts#paginate', `/** Helper.\n${item.body}\n */`)), '@fixture/contracts#paginate').examples[0];
    assert.deepEqual(example.source, { file: 'src/pagination.ts', line: item.line }, item.body);
    assert.equal(example.code, item.code, item.body);
  }
});

test('namespace module anchors use legal public access names when modules share an owner page', () => {
  const copy = structuredClone(graph);
  const template = copy.entities.find(item => item.qualifiedName === 'PagingOptions');
  for (const name of ['left', 'right']) {
    const key = `src/${name}.ts#module`;
    copy.entities.push({ ...template, key, qualifiedName: 'module', kind: 'namespace', source: { file: `src/${name}.ts`, line: 1 }, members: [] });
    copy.bindings.push({ specifier: '@fixture/contracts', exportedName: name, memberPath: [], entityKey: key });
  }
  const model = build(copy);
  assert.ok(!codes(model).includes('DOC_ANCHOR_COLLISION'));
  const page = model.pages.find(item => item.route === 'docs/api/controls/pagination.md');
  assert.equal(page.anchors['src/left.ts#module'], 'left');
  assert.equal(page.anchors['src/right.ts#module'], 'right');
});

test('retains every public alias anchor separately from the canonical entity anchor', () => {
  const model = build();
  const page = model.pages.find(item => item.route === 'docs/api/controls/pagination.md');
  for (const alias of ['paginate', 'pages', 'navigate']) {
    assert.equal(page.bindingAnchors[`@fixture/contracts#${alias}`], alias);
    assert.equal(api().resolveDocumentationReference(model, `@fixture/contracts#${alias}`).anchor, alias);
  }
});

test('model and checker agree on natural camel-case, dotted and formatted heading anchors', () => {
  const { markdownAnchors } = require('../scripts/docs/documentation.cjs');
  const model = build();
  for (const [reference, expected] of [['PagingOptions', 'pagingoptions'], ['widget.Part', 'widgetpart']]) {
    const target = api().resolveDocumentationReference(model, reference);
    assert.equal(target.anchor, expected);
    assert.ok(markdownAnchors(`## \`${reference}\`\n`).has(target.anchor));
  }
});

test('anchor collision detection uses natural punctuation removal and case folding', () => {
  for (const [names, collides] of [[['Foo.bar', 'foobar'], true], [['WIDGET.Part', 'widgetpart'], true], [['Foo-bar', 'Foo.bar'], false]]) {
    const copy = structuredClone(graph);
    const keys = ['identity', '@fixture/contracts/other#paginate'].map(reference => resolvePublicReference(copy, reference).key);
    for (const key of keys) copy.entities.find(item => item.key === key).rawDoc = '/** Anchor collision case.\n * @docPage test/anchors\n */';
    copy.bindings = copy.bindings.filter(binding => !keys.includes(binding.entityKey));
    copy.bindings.push(...names.map((name, index) => {
      const [exportedName, ...memberPath] = name.split('.');
      return { specifier: '@fixture/contracts', exportedName, memberPath, entityKey: keys[index] };
    }));
    assert.equal(codes(build(copy)).includes('DOC_ANCHOR_COLLISION'), collides, names.join(' / '));
  }
});

test('fenced code preserves starred and unstarred blank physical lines byte for byte', () => {
  const rawDoc = fs.readFileSync(path.join(fixture, 'mixed-blank-lines.jsdoc'), 'utf8');
  const expected = fs.readFileSync(path.join(fixture, 'mixed-blank-lines.expected.ts'), 'utf8');
  const model = build(mutate('@fixture/contracts#paginate', rawDoc));
  const example = get(model, '@fixture/contracts#paginate').examples[0];
  assert.equal(example.code, expected);
});

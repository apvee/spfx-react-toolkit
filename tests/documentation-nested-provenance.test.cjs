const assert = require('node:assert/strict');
const path = require('node:path');
const test = require('node:test');
const { buildApiGraph, resolvePublicReference } = require('../scripts/documentation-model.cjs');
const { buildDocumentationModel, memberContractKey, resolveDocumentationReference } = require('../scripts/documentation-contract.cjs');
const { renderApiPages } = require('../scripts/documentation-render.cjs');
const { markdownAnchors } = require('../scripts/documentation-markdown.cjs');
const root = path.join(__dirname, 'fixtures/documentation/nested-provenance');
const config = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' };
let graph, model, page;
test.before(() => {
  graph = buildApiGraph(root, config);
  assert.deepEqual(graph.diagnostics, []);
  model = buildDocumentationModel(graph, config);
  page = renderApiPages(model, config)['docs/api/index.md'];
});
const entity = name => graph.entities.find(item => item.qualifiedName === name);
const contract = (name, member) => model.contracts[memberContractKey(entity(name).key, member)];
const row = name => page.split('\n').find(line => line.startsWith('| `' + name + '`'));

test('nested configurations and method options retain their actual semantic owners and local rendered contracts', () => {
  for (const [name, members] of [
    ['Configuration.cache', ['keyFactory', 'storage']],
    ['TokenProvider.getToken.options', ['claims', 'useCachedToken']],
  ]) {
    const nested = entity(name);
    assert.deepEqual(nested.members.map(member => member.name), members);
    for (const member of nested.members) {
      assert.equal(member.ownerKey, nested.key, name + '.' + member.name);
      assert.equal(contract(name, member.name).inheritedFrom, undefined);
      assert.doesNotMatch(row(name + '.' + member.name), /Inherited from|See \[/);
    }
  }
  assert.match(row('Configuration.cache.storage'), /'session'/);
  assert.match(row('TokenProvider.getToken.options.useCachedToken'), /\| true \|/);
  assert.equal(contract('Configuration.cache', 'storage').remarks[0].body, 'Keep this nested cache setting local.');
  assert.equal(contract('Configuration.cache', 'storage').examples[0].title, 'Cache store');
  assert.match(page, /Keep this nested cache setting local\./);
  assert.match(page, /const configuration: Configuration =/);
});

test('same-line nested member comments retain their own prose without borrowing adjacent comments', () => {
  const parentMode = entity('Collision').members.find(member => member.name === 'mode');
  const nestedMode = entity('Collision.nested').members.find(member => member.name === 'mode');
  assert.equal(parentMode.source.line, nestedMode.source.line, 'exercise distinct declarations on one physical line');
  assert.ok(nestedMode.docSource, 'retain the nested same-line JSDoc source');
  assert.equal(parentMode.docSource.line, nestedMode.docSource.line, 'exercise distinct comments on one physical line');
  assert.notEqual(parentMode.rawDoc, nestedMode.rawDoc);
  assert.equal(entity('Collision').members.find(member => member.name === 'nested').rawDoc, '', 'do not borrow the adjacent parent member comment');
  assert.equal(entity('Collision.nested').members.find(member => member.name === 'plain').rawDoc, '', 'do not borrow the preceding nested member comment');
  assert.equal(contract('Collision.nested', 'mode').inheritedFrom, undefined);
  assert.equal(contract('Collision.nested', 'mode').summary, 'Nested mode. Omit for nested.');
  assert.equal(contract('Collision.nested', 'mode').defaultValue, "'nested-mode'");
  assert.equal(contract('Collision.nested', 'mode').remarks[0].body, 'Nested mode requirements.');
  assert.doesNotMatch(row('Collision.nested.mode'), /Inherited from|See \[/);
  assert.match(row('Collision.nested.mode'), /'nested-mode'/);
});

test('same-named parent members cannot consume nested defaults, remarks or examples', () => {
  const local = contract('Collision.options', 'state');
  assert.equal(local.inheritedFrom, undefined);
  assert.equal(local.defaultValue, "'nested'");
  assert.equal(local.remarks[0].body, 'Nested state requirements.');
  assert.equal(local.examples[0].title, 'Nested state');
  assert.doesNotMatch(row('Collision.options.state'), /Inherited from|See \[/);
  assert.match(page, /Nested state requirements\./);
  assert.match(page, /const collision: Collision =/);
});

test('real inherited, mapped and returned interface ownership preserves generic types and local overrides', () => {
  const base = entity('Base');
  for (const name of ['Derived', 'Mapped']) {
    const member = entity(name).members.find(member => member.name === 'value');
    assert.equal(member.ownerKey, base.key);
    assert.equal(contract(name, 'value').inheritedFrom, memberContractKey(base.key, 'value'));
    assert.match(row(name + '.value'), /Inherited from \[Base.value\]/);
  }
  assert.equal(entity('Derived').members[0].type.text, 'string | undefined');
  assert.equal(entity('Mapped').members[0].type.text, 'number | undefined');
  assert.equal(entity('Mapped').members[0].readonly, true);
  const returned = entity('getBase').signatures[0];
  assert.equal(returned.returns.members[0].type.text, 'boolean | undefined');
  assert.equal(model.contracts[memberContractKey(returned.key + '::returns', 'value')].inheritedFrom, memberContractKey(base.key, 'value'));
  assert.match(row('getBase return.value'), /Inherited from \[Base.value\]/);
  assert.equal(contract('Override', 'value').defaultValue, "'local'");
  assert.equal(contract('Override', 'value').remarks[0].body, 'Local value requirements.');
  assert.equal(contract('Override', 'value').inheritedFrom, memberContractKey(base.key, 'value'));
  for (const name of ['Configuration', 'TokenProvider', 'Collision', 'Base', 'Derived', 'Override', 'Mapped', 'getBase']) {
    assert.equal(resolvePublicReference(graph, name).key, 'src/index.ts#' + name);
    assert.ok(markdownAnchors(page).includes(resolveDocumentationReference(model, name).anchor));
  }
});

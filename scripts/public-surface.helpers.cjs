const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

/** Locate explicit package/config context without creating fixture metadata. */
function sourceContext(entryPath) {
  const repository = path.resolve(__dirname, '..');
  const config = require('./documentation.config.cjs');
  const relative = path.relative(path.join(repository, config.sourceRoot), entryPath);
  if (!relative.startsWith('..') && !path.isAbsolute(relative)) return { root: repository, config };
  for (let directory = path.dirname(path.resolve(entryPath));;) {
    if (fs.existsSync(path.join(directory, 'package.json')) && fs.existsSync(path.join(directory, 'tsconfig.json'))) {
      return { root: directory, config: { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' } };
    }
    const parent = path.dirname(directory);
    assert.notEqual(parent, directory, `Missing explicit package.json/tsconfig.json context for ${entryPath}`);
    directory = parent;
  }
}
function checkedGraph(root, config, query) {
  const graph = require('./documentation-model.cjs').buildApiGraph(root, config, query);
  assert.deepEqual(graph.diagnostics, [], 'Invalid documentation source: ' + graph.diagnostics.map(item => `${item.file}:${item.line} ${item.code}: ${item.message}`).join('\n'));
  return graph;
}
/** Project the canonical graph; default members retain physical declaration ownership. */
function projectPublicSurface(graph, bindings, root, { includeTypeMembers = false } = {}) {
  const entities = new Map(graph.entities.map(entity => [entity.key, entity]));
  function namespaceLeaf(binding) {
    if (entities.get(binding.entityKey).kind === 'namespace') return false;
    for (let index = 0; index < binding.memberPath.length; index++) {
      const parent = bindings.find(item => item.exportedName === binding.exportedName && item.memberPath.join('.') === binding.memberPath.slice(0, index).join('.'));
      if (!parent || entities.get(parent.entityKey).kind !== 'namespace') return false;
    }
    return true;
  }
  function expanded(entity, prefix, stack = new Set()) {
    if (stack.has(entity.key)) return [];
    stack = new Set([...stack, entity.key]);
    const result = [];
    for (const member of [...entity.members, ...entity.variants.flatMap(variant => variant.members)]) {
      if (member.name.startsWith('[')) continue;
      result.push({ name: prefix + '.' + member.name, jsdoc: Boolean(member.rawDoc), ownerKey: member.ownerKey });
      const owner = entities.get(member.ownerKey);
      for (const reference of member.type.references) {
        const child = entities.get(reference);
        if (member.ownerKey === entity.key && child && owner && child.source.file === owner.source.file && child.qualifiedName === owner.qualifiedName + '.' + member.name) result.push(...expanded(child, prefix + '.' + member.name, stack));
      }
    }
    return [...new Map(result.map(item => [item.name + ':' + item.ownerKey, item])).values()];
  }
  return bindings.filter(namespaceLeaf).map(binding => {
    const entity = entities.get(binding.entityKey);
    const name = [binding.exportedName, ...binding.memberPath].join('.');
    const members = includeTypeMembers && ['interface', 'type'].includes(entity.kind) ? expanded(entity, name)
      : entity.kind === 'interface' ? entity.members.filter(member => member.ownerKey === entity.key && !member.name.startsWith('['))
        .map(member => ({ name: name + '.' + member.name, jsdoc: Boolean(member.rawDoc) })) : [];
    return { name, filePath: path.resolve(root, entity.source.file), jsdoc: Boolean(entity.rawDoc), members };
  });
}
/** Batch one invocation's entry queries; rebuild on every call to observe mutations. */
function collectPublicSurfaces(entryPaths, options = {}) {
  if (!entryPaths.length) return [];
  const { root, config } = sourceContext(path.resolve(entryPaths[0]));
  const entries = entryPaths.map(entryPath => {
    const context = sourceContext(path.resolve(entryPath));
    assert.deepEqual(context, { root, config }, 'A surface query batch must use one package/config context.');
    return path.relative(root, path.resolve(entryPath)).split(path.sep).join('/');
  });
  const graph = checkedGraph(root, config, { requestedSourceEntries: entries });
  return entries.map(entry => projectPublicSurface(graph, graph.entryQueries[entry], root, options));
}
/** Collect a caller-selected source entry using the same checker as package docs. */
function collectPublicSurface(entryPath, options = {}) {
  return collectPublicSurfaces([entryPath], options)[0];
}

function assertDocumentedSurface(document, surface) {
  const mentioned = name => new RegExp('(?<![\\w.])' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?![\\w.])').test(document);
  for (const declaration of surface) {
    assert.ok(mentioned(declaration.name), `Public documentation missing ${declaration.name}`);
    assert.ok(declaration.jsdoc, `Source JSDoc missing ${declaration.name} (${declaration.filePath})`);
    for (const member of declaration.members) {
      assert.ok(mentioned(member.name), `Public documentation missing ${member.name}`);
      assert.ok(member.jsdoc, `Source JSDoc missing ${member.name} (${declaration.filePath})`);
    }
  }
}

module.exports = { collectPublicSurface, collectPublicSurfaces, assertDocumentedSurface, projectPublicSurface, sourceContext, checkedGraph };

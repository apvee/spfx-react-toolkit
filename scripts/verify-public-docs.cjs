const assert = require('node:assert/strict');
const path = require('node:path');
const fs = require('node:fs');
const config = require('./documentation.config.cjs');
const { generateArtifacts, applyArtifacts } = require('./documentation-artifacts.cjs');
const { checkedGraph, projectPublicSurface, assertDocumentedSurface } = require('./public-surface.helpers.cjs');
const { buildDocumentationModel, resolveDocumentationReference } = require('./documentation-contract.cjs');
const { projectPackageExports, findInvalidExampleImports } = require('./documentation.cjs');
const root = path.resolve(__dirname, '..');

// Composition validates the complete source contracts, examples and virtual link
// universe before the same read-only ArtifactSet comparison used by the CLI.
const artifacts = generateArtifacts(root);
const result = applyArtifacts(root, artifacts, 'check');
assert.ok(result.ok, `Stale generated documentation: ${JSON.stringify(result.diff)}. Run npm run docs:generate.`);
const graph = checkedGraph(root, config);
const model = buildDocumentationModel(graph, config);
const verified = new Set();
for (const specifier of new Set(graph.bindings.map(binding => binding.specifier))) {
  const bindings = graph.bindings.filter(binding => binding.specifier === specifier);
  for (const declaration of projectPublicSurface(graph, bindings, root, { includeTypeMembers: true })) {
    const binding = bindings.find(item => [item.exportedName, ...item.memberPath].join('.') === declaration.name);
    const identity = binding.entityKey + '#' + declaration.name;
    if (verified.has(identity)) continue;
    verified.add(identity);
    const target = resolveDocumentationReference(model, `${specifier}#${declaration.name}`);
    assert.ok(artifacts[target.page?.route], `Missing canonical reference for ${specifier}#${declaration.name}`);
    assertDocumentedSurface(artifacts[target.page.route], [declaration]);
  }
}
function markdownFiles(directory) {
  return fs.readdirSync(path.join(root, directory), { withFileTypes: true }).flatMap(entry => {
    if (entry.name === '_sources') return [];
    const route = directory + '/' + entry.name;
    return entry.isDirectory() ? markdownFiles(route) : entry.isFile() && entry.name.endsWith('.md') ? [route] : [];
  });
}
const files = ['AGENTS.md', 'README.md', config.packageRoot + '/README.md', ...markdownFiles('docs')];
const publicExports = projectPackageExports(graph, root);
assert.deepEqual(findInvalidExampleImports(root, files, publicExports), [], 'Unavailable imports in documentation examples');
for (const route of markdownFiles('docs')) {
  const text = fs.readFileSync(path.join(root, route), 'utf8');
  for (const pattern of [/\bJotai\b/i, /\batom\b/i, /\batoms\b/i]) assert.ok(!pattern.test(text), `${route} contains ${pattern}`);
}
console.log(`Public documentation verified (${Object.keys(artifacts).length} artifacts, ${publicExports.length} clean-entry exports; contracts, members, examples, imports and links).`);

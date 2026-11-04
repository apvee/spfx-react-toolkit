const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { buildApiGraph } = require('../../../scripts/docs/documentation-model.cjs');
const defaultConfig = require('../../../scripts/docs/documentation.config.cjs');
const { executableRoot, ownership } = require('../../../scripts/docs/executable-examples-render.cjs');

/** Resolve the actual owned catalogue modules' non-root toolkit imports to one
 * fixture boundary. This is browser-build input, never canonical generation input.
 * Exact aliases share the source checker's public-symbol proof; no wildcard /lib
 * mapping or separately maintained compatibility inventory is involved.
 */
function executableFixtureAliases(root, source, config = defaultConfig) {
  const graph = buildApiGraph(root, config);
  assert.deepEqual(graph.diagnostics, [], 'Fixture aliases require a valid canonical source graph.');
  const metadata = JSON.parse(fs.readFileSync(path.join(root, config.packageRoot, 'package.json'), 'utf8'));
  const clean = new Set(graph.bindings.map(binding => binding.specifier));
  const requested = new Map();
  const entities = new Map(graph.entities.map(entity => [entity.key,entity]));
  const directory = path.join(root, executableRoot);
  for (const name of fs.readdirSync(directory).filter(name => /^(?:[a-f0-9]{64}\.tsx|catalogue\.ts)$/.test(name))) {
    const text = fs.readFileSync(path.join(directory,name), 'utf8');
    assert.ok(text.startsWith(ownership + '\n'), `Fixture refuses a manually owned executable: ${name}`);
    const module = ts.createSourceFile(name, text, ts.ScriptTarget.Latest, true);
    for (const statement of module.statements) {
      if (!ts.isImportDeclaration(statement) && !ts.isExportDeclaration(statement)) continue;
      const specifier = statement.moduleSpecifier?.text;
      if (!specifier?.startsWith(metadata.name + '/')) continue;
      const named = ts.isImportDeclaration(statement) && statement.importClause?.namedBindings;
      const names = named && ts.isNamedImports(named) && !statement.importClause.name
        ? named.elements.map(element => (element.propertyName ?? element.name).text) : undefined;
      if (!requested.has(specifier)) requested.set(specifier,names ? new Set(names) : undefined);
      else if (!names) requested.set(specifier,undefined);
      else if (requested.get(specifier)) names.forEach(name => requested.get(specifier).add(name));
    }
  }
  return Object.fromEntries([...requested.keys()].sort().map(specifier => {
    assert.ok(clean.has(specifier) || Object.hasOwn(graph.compatibilityEntries,specifier), `Unapproved executable fixture import: ${specifier}`);
    const alias = '.' + specifier.slice(metadata.name.length);
    const entry = metadata.exports[alias];
    assert.ok(entry && typeof entry.types === 'string' && /^\.\/lib\/(?:[^./\\][^/\\]*\/)*[^./\\][^/\\]*\.d\.ts$/.test(entry.types), `Invalid executable fixture target: ${specifier}`);
    let stem = entry.types.slice('./lib/'.length).replace(/\.d\.ts$/, '');
    // A generic named hook imported from /hooks must not pull every unrelated
    // native SDK hook into a narrow fixture. Use the proven canonical owner only
    // when all actual imports retain their physical export names in one module.
    const names = requested.get(specifier);
    const owners = clean.has(specifier) && names ? [...names].map(name => {
      const binding = graph.bindings.find(binding => binding.specifier === specifier && binding.exportedName === name && !binding.memberPath.length);
      const entity = binding && entities.get(binding.entityKey);
      return entity?.qualifiedName === name && entity.kind !== 'namespace' ? entity.source.file : undefined;
    }) : [];
    if (owners.length && owners.every(owner => owner && owner === owners[0])) {
      stem = path.relative(config.sourceRoot,owners[0]).split(path.sep).join('/').replace(/\.tsx?$/, '');
      assert.ok(!stem.startsWith('../') && !path.isAbsolute(stem), 'Canonical fixture owner must remain inside source.');
    }
    const candidates = source ? ['.ts','.tsx'].map(extension => path.join(root,config.sourceRoot,stem + extension))
      : [path.join(root,config.packageRoot,'lib',stem + '.js')];
    const file = candidates.find(candidate => fs.existsSync(candidate));
    assert.ok(file, `Missing executable fixture implementation: ${specifier}`);
    return [specifier + '$', file];
  }));
}
module.exports = { executableFixtureAliases };

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { markdownAnchors: collectMarkdownAnchors } = require('./documentation-markdown.cjs');
const { renderPackageReadme } = require('./documentation-render.cjs');

const defaultConfig = require('./documentation.config.cjs');
const { generateArtifacts, compareArtifacts, applyArtifacts } = require('./documentation-artifacts.cjs');
const { buildDocumentationModel } = require('./documentation-contract.cjs');
const { sourceContext, checkedGraph } = require('./public-surface.helpers.cjs');

function topLevelExports(graph, bindings, root, sourceRoot, config) {
  const entities = new Map(graph.entities.map(entity => [entity.key, entity]));
  const model = buildDocumentationModel(graph, config);
  return bindings.filter(binding => !binding.memberPath.length).map(binding => {
    const entity = entities.get(binding.entityKey);
    return { name: binding.exportedName, source: path.relative(sourceRoot, path.resolve(root, entity.source.file)).split(path.sep).join('/'),
      kind: ['interface', 'type'].includes(entity.kind) ? 'Type' : entity.kind === 'function' ? 'Function' : 'Value',
      description: model.contracts[entity.key]?.summary.split('\n')[0] || '', specifier: binding.specifier };
  }).sort((a, b) => a.name.localeCompare(b.name, 'en'));
}
function collectExports(entry) {
  const { root, config } = sourceContext(path.resolve(entry));
  const route = path.relative(root, path.resolve(entry)).split(path.sep).join('/');
  const graph = checkedGraph(root, config, { requestedSourceEntries: [route] });
  return topLevelExports(graph, graph.entryQueries[route], root, path.dirname(entry), config).map(({ specifier, ...item }) => item);
}
function collectPackageExports(root, config = defaultConfig) {
  return projectPackageExports(checkedGraph(root, config), root, config);
}
function projectPackageExports(graph, root, config = defaultConfig) {
  const exports = new Map();
  for (const { specifier, ...item } of topLevelExports(graph, graph.bindings, root, path.resolve(root, config.sourceRoot), config)) {
    const previous = exports.get(item.name);
    assert.ok(!previous || previous.source === item.source, `Ambiguous public export ${item.name}`);
    if (!previous) exports.set(item.name, { ...item, imports: [] });
    exports.get(item.name).imports.push(specifier);
  }
  return [...exports.values()].sort((a, b) => a.name.localeCompare(b.name, 'en'));
}
function checkArtifacts(root, artifacts) {
  const diff = compareArtifacts(root, artifacts);
  return [...diff.missing, ...diff.changed, ...diff.orphan].sort();
}

function withoutFencedCode(source) {
  let fence;
  return source.split('\n').map(line => {
    const marker = line.match(/^\s*(`{3,}|~{3,})/);
    if (marker) {
      if (!fence) fence = marker[1];
      else if (marker[1][0] === fence[0] && marker[1].length >= fence.length) fence = undefined;
      return '';
    }
    return fence ? '' : line;
  }).join('\n');
}

function markdownAnchors(source) {
  return new Set(collectMarkdownAnchors(source));
}

function findBrokenLinks(root, files) {
  const broken = [];
  const anchorCache = new Map();
  for (const file of files) {
    const source = withoutFencedCode(fs.readFileSync(path.join(root, file), 'utf8'));
    for (const match of source.matchAll(/\]\(([^)]+)\)/g)) {
      const target = match[1].split(/\s+["']/)[0];
      if (/^(?:[a-z][a-z0-9+.-]*:|\/)/i.test(target)) continue;
      const [relative, fragment] = target.split('#');
      const resolved = path.resolve(root, path.dirname(file), decodeURIComponent(relative || path.basename(file)));
      if (!fs.existsSync(resolved)) {
        broken.push(`${file}: ${target}`);
      } else if (fragment && resolved.endsWith('.md')) {
        if (!anchorCache.has(resolved)) anchorCache.set(resolved, markdownAnchors(fs.readFileSync(resolved, 'utf8')));
        if (!anchorCache.get(resolved).has(decodeURIComponent(fragment))) broken.push(`${file}: ${target}`);
      }
    }
  }
  return broken;
}

function findInvalidExampleImports(root, files, exports) {
  const available = new Map();
  for (const item of exports) {
    for (const importPath of item.imports) {
      if (!available.has(importPath)) available.set(importPath, new Set());
      available.get(importPath).add(item.name);
    }
  }
  const packageRoot = path.join(root, 'packages/spfx-react-toolkit');
  const metadataPath = path.join(packageRoot, 'package.json');
  const metadata = fs.existsSync(metadataPath) ? JSON.parse(fs.readFileSync(metadataPath, 'utf8')) : undefined;
  const requests = [];
  const invalid = [];
  for (const file of files) {
    let fence;
    let lines = [];
    let firstLine = 0;
    for (const [index, line] of fs.readFileSync(path.join(root, file), 'utf8').split('\n').entries()) {
      const marker = line.match(/^\s*(`{3,}|~{3,})(\w*)/);
      if (!fence && marker) {
        fence = { marker: marker[1], check: /^(?:ts|tsx|typescript|js|jsx|javascript)$/.test(marker[2]) };
        firstLine = index + 2;
        lines = [];
      } else if (fence && marker && marker[1][0] === fence.marker[0] && marker[1].length >= fence.marker.length) {
        if (fence.check) {
          const source = ts.createSourceFile('example.tsx', lines.join('\n'), ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
          for (const statement of source.statements) {
            if (!ts.isImportDeclaration(statement) || !ts.isStringLiteral(statement.moduleSpecifier)) continue;
            const importPath = statement.moduleSpecifier.text;
            if (importPath !== '@apvee/spfx-react-toolkit' && !importPath.startsWith('@apvee/spfx-react-toolkit/')) continue;
            const location = `${file}:${firstLine + source.getLineAndCharacterOfPosition(statement.getStart(source)).line}`;
            const clause = statement.importClause;
            const requested = clause?.name ? ['default'] : [];
            if (clause?.namedBindings && ts.isNamedImports(clause.namedBindings)) requested.push(...clause.namedBindings.elements.map(item => (item.propertyName ?? item.name).text));
            requests.push({ importPath, requested, location });
          }
        }
        fence = undefined;
      } else if (fence) {
        lines.push(line);
      }
    }
  }
  // Resolve only metadata-declared historical imports, once per invocation,
  // through the canonical requested-source query. No query adds public imports.
  const legacyEntries = new Map();
  for (const { importPath } of requests) {
    if (available.has(importPath) || legacyEntries.has(importPath)) continue;
    const key = '.' + importPath.slice('@apvee/spfx-react-toolkit'.length);
    if (key !== './lib' && !key.startsWith('./lib/')) continue;
    const contract = metadata?.exports[key];
    const target = typeof contract === 'string' ? contract : contract?.types;
    if (typeof target !== 'string' || !/^\.\/lib\/(?:[^./\\][^/\\]*\/)*[^./\\][^/\\]*\.(?:d\.ts|js)$/.test(target)) continue;
    const stem = target.slice('./lib/'.length).replace(/\.(?:d\.ts|js)$/, '');
    const entry = ['.ts', '.tsx'].map(extension => path.join(packageRoot, 'src', stem + extension)).find(file => fs.existsSync(file));
    if (entry) legacyEntries.set(importPath, path.relative(root, entry).split(path.sep).join('/'));
  }
  if (legacyEntries.size) {
    const graph = checkedGraph(root, defaultConfig, { requestedSourceEntries: [...legacyEntries.values()] });
    for (const [importPath, entry] of legacyEntries) available.set(importPath, new Set(graph.entryQueries[entry].filter(binding => !binding.memberPath.length).map(binding => binding.exportedName)));
  }
  for (const { importPath, requested, location } of requests) {
    const names = available.get(importPath);
    if (!names) invalid.push(`${location}: unsupported public import ${importPath}`);
    else for (const name of requested) if (!names.has(name)) invalid.push(`${location}: ${importPath} does not export ${name}`);
  }
  return invalid;
}

const generatedArtifacts = generateArtifacts;

function runCli(args = process.argv.slice(2), env = process.env) {
  assert.ok(args.length === 0 || (args.length === 1 && args[0] === '--check'), 'Usage: node scripts/documentation.cjs [--check]');
  const configured = env.TOOLKIT_DOCS_MODE === undefined ? 'write' : env.TOOLKIT_DOCS_MODE;
  assert.ok(configured === 'write' || configured === 'check', `Unknown documentation mode TOOLKIT_DOCS_MODE=${configured}; expected write or check.`);
  const mode = args.includes('--check') ? 'check' : configured;
  const root = path.resolve(__dirname, '..');
  const artifacts = generateArtifacts(root);
  const result = applyArtifacts(root, artifacts, mode);
  const stale = [...result.diff.missing, ...result.diff.changed, ...result.diff.orphan];
  if (!result.ok) throw new Error(`Stale generated documentation: ${stale.join(', ')}. Run npm run docs:generate.`);
  console.log(mode === 'check' ? `Generated documentation is current (${Object.keys(artifacts).length} artifacts).` : `Updated ${stale.length} documentation artifact(s).`);
  return result;
}
if (require.main === module) {
  try { runCli(); }
  catch (error) { console.error(error.message); process.exitCode = 1; }
}
module.exports = { renderPackageReadme, collectExports, collectPackageExports, projectPackageExports, checkArtifacts, generatedArtifacts, markdownAnchors, findBrokenLinks, findInvalidExampleImports, runCli };

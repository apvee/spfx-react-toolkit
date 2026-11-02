const path = require('node:path');
const ts = require('typescript');
const { createDocumentationInputs } = require('./documentation-inputs.cjs');

/** @typedef {{packageRoot: string, sourceRoot: string, tsconfig: string, repositoryUrl?: string, repositoryRef?: string}} Config */
/** @typedef {{file: string, line: number}} DocLocation */
/** @typedef {{code: string, file: string, line: number, symbol?: string, message: string}} Diagnostic */
/** @typedef {{specifier: string, exportedName: string, entityKey: string, memberPath: string[], ownerKey?: string, rawDoc?: string, docSource?: DocLocation}} PublicBinding */
/** @typedef {{text: string, references: string[]}} TypeUse */
/** @typedef {{name: string, source: DocLocation, docSource?: DocLocation, ownerKey?: string, rawDoc: string, optional: boolean, readonly?: boolean, forbidden: boolean, declaredType?: string, type: TypeUse}} ApiMember ownerKey identifies the declaring named or anonymous type, not its lexical named ancestor. */
/** @typedef {{text: string, members: ApiMember[]}} ApiVariant */
/** @typedef {TypeUse & {members: ApiMember[], variants: ApiVariant[]}} TypeShape */
/** @typedef {{name: string, constraint?: string, default?: string}} ApiTypeParameter */
/** @typedef {{name: string, optional: boolean, rest: boolean, type: TypeUse}} ApiParameter */
/** @typedef {{key: string, source: DocLocation, docSource?: DocLocation, sourceOwned: boolean, rawDoc: string, text: string, typeParameters: ApiTypeParameter[], lexicalTypeParameters: ApiTypeParameter[], parameters: ApiParameter[], returns: TypeShape}} ApiSignature */
/** @typedef {{key: string, qualifiedName: string, kind: string, namespaceRole?: 'sourceFile'|'declaration', parameterDocumentation?: {ownerKey: string, name: string}, source: DocLocation, docSource?: DocLocation, rawDoc: string, moduleDoc: string, moduleDocSource?: DocLocation, type?: TypeUse, typeExpression?: string, typeParameters: ApiTypeParameter[], heritage: TypeUse[], signatures: ApiSignature[], members: ApiMember[], variants: ApiVariant[], references: string[], visibility: 'public'|'support'|'external'}} ApiEntity */
/** @typedef {{entities: ApiEntity[], bindings: PublicBinding[], diagnostics: Diagnostic[], entryQueries?: Record<string, PublicBinding[]>, compatibilityEntries: Record<string, string>}} ApiGraph */

const compare = (left, right) => left < right ? -1 : left > right ? 1 : 0;
const posix = value => value.split(path.sep).join('/');

/**
 * Read a package's clean export map and project its source-owned contracts with
 * the configured TypeScript checker. Peer types stay opaque: their spelling and
 * references are retained without recursively documenting the peer's API.
 * Member/parameter type uses reference entities rather than recursively copying
 * them; union variants retain effective branch members and their constraints.
 * This function does not emit files or use compiled library output.
 * @param {string} root Repository (or independent fixture) root.
 * @param {Config} config Versioned package/source/tsconfig locations, relative to root.
 * @param {{requestedSourceEntries?: string[]}} [query] Explicit in-memory source queries.
 * Query bindings are returned separately in entryQueries, keyed by the supplied
 * source path, with an empty specifier; they never advertise a package import.
 * Queries use the same guarded checker and projection, with support visibility.
 * @returns {ApiGraph} JSON-serializable semantic graph with legal public bindings.
 */
function buildApiGraph(root, config, query = {}) {
  const inputs = createDocumentationInputs(root, config);
  root = inputs.root;
  const sourceRoot = path.resolve(root, config.sourceRoot);
  const configFile = path.resolve(root, config.tsconfig);
  const metadata = JSON.parse(inputs.readFile(path.resolve(root, config.packageRoot, 'package.json')));
  const parsed = inputs.loadConfig();
  const entries = [...inputs.sourceEntries(metadata)].map(([specifier, file]) => ({ specifier, file }))
    .sort((a, b) => compare(a.specifier, b.specifier));
  const historical = [...inputs.sourceEntries(metadata, true)];
  const requested = [...new Set(query.requestedSourceEntries || [])].map(route => ({ route, file: inputs.requestedSourceEntry(route) }));
  const entryQueries = Object.fromEntries(requested.map(item => [item.route, []]));
  const options = { ...parsed.options, noEmit: true };
  const program = ts.createProgram([...new Set([...parsed.fileNames, ...entries.map(item => item.file), ...historical.map(([, file]) => file), ...requested.map(item => item.file)])], options, inputs.createCompilerHost(options));
  const checker = program.getTypeChecker();
  /** @type {Map<string, ApiEntity>} */
  const entities = new Map();
  /** @type {PublicBinding[]} */
  const bindings = [];
  const pending = new Map();
  const publicSymbols = new Set();
  const target = symbol => symbol.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(symbol) : symbol;
  function declaration(symbol) {
    const original = symbol.valueDeclaration || symbol.declarations?.[0];
    let node = original;
    // Anonymous union/intersection/conditional branches have the same lexical name as their
    // owner. Project the actual enclosing type, rather than letting the first
    // branch overwrite a named alias (or a complete parameter type).
    if (node && ts.isTypeNode(node)) {
      while (node.parent && (ts.isParenthesizedTypeNode(node.parent) || ts.isUnionTypeNode(node.parent) || ts.isIntersectionTypeNode(node.parent) || ts.isConditionalTypeNode(node.parent))) node = node.parent;
      if (node.parent && ts.isTypeAliasDeclaration(node.parent) && node.parent.type === node) return node.parent;
    }
    return node?.parent && ts.isParameter(node.parent) && node.parent.type === node ? node : original;
  }
  const isLocal = node => node && !path.relative(sourceRoot, node.getSourceFile().fileName).startsWith('..') && !path.isAbsolute(path.relative(sourceRoot, node.getSourceFile().fileName));
  function sourcePath(file) {
    const normalized = posix(file);
    const dependency = normalized.lastIndexOf('/node_modules/');
    return dependency === -1 ? posix(path.relative(root, file)) : normalized.slice(dependency + 1);
  }
  const location = node => ({ file: sourcePath(node.getSourceFile().fileName), line: node.getSourceFile().getLineAndCharacterOfPosition(node.getStart()).line + 1 });
  function docComments(node) {
    if (!node) return [];
    // Variable comments belong to the statement, namespace comments to its export.
    const owner = ts.isVariableDeclaration(node) ? node.parent.parent
      : ts.isNamespaceExport(node) ? node.parent : ts.isExportSpecifier(node) ? node.parent.parent : node;
    const local = (owner.jsDoc || node.jsDoc || []).filter(item => !item.tags?.some(tag => tag.tagName.text === 'packageDocumentation'))
      .map(item => ({ text: item.getText(), source: location(item) }));
    if (local.length || !(ts.isFunctionTypeNode(node) || ts.isConstructorTypeNode(node) || ts.isTypeLiteralNode(node) || ts.isParenthesizedTypeNode(node) || ts.isUnionTypeNode(node) || ts.isIntersectionTypeNode(node) || ts.isPropertySignature(node) || ts.isPropertyDeclaration(node))) return local;
    // TypeScript does not attach same-line type-position or member JSDoc to these nodes.
    // Preserve the nearest explicit comment on the node or its transparent
    // wrappers before falling back to the directly enclosing property contract.
    const source = node.getSourceFile();
    const transparent = current => ts.isParenthesizedTypeNode(current) || ts.isUnionTypeNode(current) || ts.isIntersectionTypeNode(current);
    let current = node;
    for (;;) {
      const attached = (current.jsDoc || []).filter(item => !item.tags?.some(tag => tag.tagName.text === 'packageDocumentation'));
      if (attached.length) return attached.map(item => ({ text: item.getText(), source: location(item) }));
      const ranges = [...(ts.getLeadingCommentRanges(source.text, current.pos) || []), ...(ts.getTrailingCommentRanges(source.text, current.pos) || [])];
      const explicit = ranges.filter(range => source.text.startsWith('/**', range.pos));
      if (explicit.length) return explicit.map(range => ({ text: source.text.slice(range.pos, range.end),
        source: { file: sourcePath(source.fileName), line: source.getLineAndCharacterOfPosition(range.pos).line + 1 } }));
      if (!current.parent || !transparent(current.parent)) break;
      current = current.parent;
    }
    // Anonymous member types have no separate named declaration. Their directly
    // enclosing property owns the contract, including through union/parenthesis
    // wrappers. Never borrow a surrounding interface/type alias description.
    const parent = current.parent;
    return parent && (ts.isPropertySignature(parent) || ts.isPropertyDeclaration(parent) || ts.isParameter(parent)) ? docComments(parent) : local;
  }
  const docs = node => docComments(node).map(item => item.text).join('\n');
  function docLocation(node) {
    const comment = docComments(node)[0];
    return comment ? { docSource: comment.source } : {};
  }
  function qualified(node, fallback) {
    const names = [];
    for (let current = node; current && !ts.isSourceFile(current); current = current.parent) {
      if (current.name) names.unshift(ts.isIdentifier(current.name) || ts.isStringLiteral(current.name) ? current.name.text : current.name.getText());
    }
    return names.join('.') || (ts.isSourceFile(node) ? 'module' : fallback);
  }
  function identity(symbol) {
    const node = declaration(symbol);
    if (!node) throw new Error(`Unresolved declaration: ${symbol.name}`);
    // Source-file modules have no declared identifier. Reserve a role marker
    // that cannot collide with a legal namespace named `module`.
    return `${location(node).file}#${ts.isSourceFile(node) ? '<module>' : qualified(node, symbol.name)}`;
  }
  function register(symbol, visibility = 'support') {
    symbol = target(symbol);
    const node = declaration(symbol);
    if (!node) return undefined;
    if (ts.isTypeAliasDeclaration(node)) symbol = checker.getSymbolAtLocation(node.name) || symbol;
    const key = identity(symbol);
    if (!entities.has(key)) {
      entities.set(key, {
        key, qualifiedName: qualified(node, symbol.name), kind: ts.isTypeNode(node) ? 'type' : symbol.flags & ts.SymbolFlags.Module ? 'namespace'
          : symbol.flags & ts.SymbolFlags.Interface ? 'interface' : symbol.flags & ts.SymbolFlags.TypeAlias ? 'type'
            : symbol.flags & ts.SymbolFlags.Function ? 'function' : 'value',
        ...(symbol.flags & ts.SymbolFlags.Module ? { namespaceRole: ts.isSourceFile(node) ? 'sourceFile' : 'declaration' } : {}),
        source: location(node), ...docLocation(node), rawDoc: docs(node), moduleDoc: '', typeParameters: [], heritage: [], signatures: [], members: [], variants: [], references: [],
        visibility: isLocal(node) ? visibility : 'external',
      });
      pending.set(key, symbol);
    }
    if (visibility === 'public' && isLocal(node)) entities.get(key).visibility = 'public';
    return key;
  }
  // The checker emits absolute import-type paths for namespace-valued returns.
  // Project those paths into the same source/dependency coordinates everywhere,
  // before signature spelling is also used as semantic identity.
  function checkerText(text) {
    return text.replace(/import\("([^"]+)"\)/g, (match, modulePath) =>
      path.isAbsolute(modulePath) ? `import(${JSON.stringify('./' + sourcePath(modulePath))})` : match);
  }
  const typeText = type => checkerText(checker.typeToString(type, undefined, ts.TypeFormatFlags.NoTruncation | ts.TypeFormatFlags.UseAliasDefinedOutsideCurrentScope));
  function typeUse(type) {
    const references = new Set();
    const visited = new Set();
    function visit(current) {
      if (visited.has(current)) return;
      visited.add(current);
      const symbol = current.aliasSymbol || current.getSymbol();
      if (symbol && !(symbol.flags & ts.SymbolFlags.TypeParameter) && declaration(symbol)) {
        const key = register(symbol);
        if (key) references.add(key);
      }
      if (current.isUnionOrIntersection()) current.types.forEach(visit);
      (current.aliasTypeArguments || []).forEach(visit);
      if (current.flags & ts.TypeFlags.Object && current.objectFlags & ts.ObjectFlags.Reference) checker.getTypeArguments(current).forEach(visit);
    }
    visit(type);
    return { text: typeText(type), references: [...references].sort(compare) };
  }
  function memberOwner(node) {
    for (let owner = node.parent; owner && !ts.isSourceFile(owner); owner = owner.parent) {
      // An anonymous object is a real declaration boundary. Its members belong
      // to the canonical nested entity, not to a surrounding named interface.
      if (ts.isTypeLiteralNode(owner)) {
        const symbol = checker.getTypeAtLocation(owner).getSymbol();
        const key = symbol && register(symbol);
        return key ? { ownerKey: key } : {};
      }
      if ((ts.isInterfaceDeclaration(owner) || ts.isTypeAliasDeclaration(owner) || ts.isClassDeclaration(owner)) && owner.name) {
        const symbol = checker.getSymbolAtLocation(owner.name);
        const key = symbol && register(symbol);
        return key ? { ownerKey: key } : {};
      }
    }
    return {};
  }
  function members(type) {
    if (!(type.flags & ts.TypeFlags.Object) && !type.isIntersection()) return [];
    return checker.getPropertiesOfType(type).filter(symbol => isLocal(declaration(symbol))).map(symbol => {
      const node = declaration(symbol);
      const declaredType = node.type?.getText();
      // TypeScript 5.3's effective mapped/synthetic property flags preserve
      // both +readonly and -readonly; source modifiers alone lose the latter.
      // This compiler adapter intentionally uses the pinned compiler's helper.
      const checkFlags = ts.getCheckFlags(symbol);
      const projected = checkFlags & (ts.CheckFlags.Mapped | ts.CheckFlags.Synthetic);
      const readonly = Boolean(checkFlags & ts.CheckFlags.Readonly) || (!projected && (
        Boolean(ts.getCombinedModifierFlags(node) & ts.ModifierFlags.Readonly) ||
        Boolean(symbol.flags & ts.SymbolFlags.GetAccessor) && !(symbol.flags & ts.SymbolFlags.SetAccessor)
      ));
      return {
        name: node.name && ts.isComputedPropertyName(node.name) ? node.name.getText() : symbol.name, source: location(node), ...docLocation(node), ...memberOwner(node), rawDoc: docs(node), optional: Boolean(symbol.flags & ts.SymbolFlags.Optional),
        readonly, forbidden: declaredType === 'never', ...(declaredType ? { declaredType } : {}),
        type: typeUse(checker.getTypeOfSymbolAtLocation(symbol, node)),
      };
    }).sort((a, b) => compare(a.name, b.name));
  }
  function shape(type) {
    return {
      ...typeUse(type), members: members(type),
      variants: type.isUnion() ? type.types.map(branch => ({ text: typeText(branch), members: members(branch) })) : [],
    };
  }
  function typeParameters(nodes) {
    return (nodes || []).map(node => ({ name: node.name.text,
      ...(node.constraint ? { constraint: node.constraint.getText() } : {}),
      ...(node.default ? { default: node.default.getText() } : {}) }));
  }
  function lexicalTypeParameters(node) {
    const scope = new Map();
    for (let owner = node.parent; owner && !ts.isSourceFile(owner); owner = owner.parent) {
      for (const parameter of typeParameters(owner.typeParameters)) {
        if (!scope.has(parameter.name)) scope.set(parameter.name, parameter);
      }
    }
    return [...scope.values()];
  }
  function signatures(type, ownerKey) {
    return checker.getSignaturesOfType(type, ts.SignatureKind.Call).map(signature => {
      const node = signature.getDeclaration();
      const text = checkerText(checker.signatureToString(signature, undefined, ts.TypeFormatFlags.NoTruncation));
      return {
        key: `${ownerKey}::${text}`, source: location(node), ...docLocation(node), sourceOwned: isLocal(node), rawDoc: docs(node), text,
        typeParameters: typeParameters(node.typeParameters), lexicalTypeParameters: lexicalTypeParameters(node),
        parameters: signature.parameters.map(parameter => {
          const parameterNode = declaration(parameter);
          return { name: parameter.name, optional: Boolean(parameter.flags & ts.SymbolFlags.Optional) || Boolean(parameterNode.questionToken || parameterNode.initializer),
            rest: Boolean(parameterNode.dotDotDotToken), type: typeUse(checker.getTypeOfSymbolAtLocation(parameter, parameterNode)) };
        }),
        returns: shape(checker.getReturnTypeOfSignature(signature)),
      };
    });
  }
  function publicBinding(symbol, specifier, exportedName, memberPath = [], ownerKey, ancestry = new Set(), destination = bindings, visibility = 'public') {
    const resolved = target(symbol);
    if (visibility === 'public') publicSymbols.add(resolved);
    const key = register(resolved, visibility);
    if (!key || ancestry.has(key)) return;
    const node = declaration(symbol);
    const rawDoc = docs(node);
    destination.push({ specifier, exportedName, entityKey: key, memberPath, ...(ownerKey ? { ownerKey } : {}),
      ...(rawDoc ? { rawDoc, docSource: docLocation(node).docSource || location(node) } : {}) });
    const next = new Set([...ancestry, key]);
    const namespace = Boolean(resolved.flags & ts.SymbolFlags.Module);
    const value = Boolean(resolved.flags & ts.SymbolFlags.Value) && !namespace;
    const children = namespace ? checker.getExportsOfModule(resolved)
      : value ? checker.getPropertiesOfType(checker.getTypeOfSymbolAtLocation(resolved, declaration(resolved))).filter(child => isLocal(declaration(child)) && !ts.isComputedPropertyName(declaration(child).name)) : [];
    for (const child of children.sort((a, b) => compare(a.name, b.name))) publicBinding(child, specifier, exportedName, [...memberPath, child.name], key, next, destination, visibility);
  }
  for (const entry of entries) {
    const source = program.getSourceFile(entry.file);
    const module = source && checker.getSymbolAtLocation(source);
    if (!module) throw new Error(`Missing TypeScript module: ${entry.file}`);
    for (const symbol of checker.getExportsOfModule(module).sort((a, b) => compare(a.name, b.name))) publicBinding(symbol, entry.specifier, symbol.name);
  }
  // Compatibility resolution is deliberately separate from advertised clean
  // bindings. Even an exact historical export fails closed when any top-level
  // symbol is private; namespace descendants qualify through the same checker.
  const compatibilityEntries = {};
  for (const [specifier, file] of historical) {
    const source = program.getSourceFile(file);
    const module = source && checker.getSymbolAtLocation(source);
    const symbols = module ? checker.getExportsOfModule(module) : [];
    if (symbols.length && symbols.every(symbol => publicSymbols.has(target(symbol)))) {
      compatibilityEntries[specifier] = posix(path.relative(root, file));
    }
  }
  for (const entry of requested) {
    const source = program.getSourceFile(entry.file);
    const module = source && checker.getSymbolAtLocation(source);
    if (!module) throw new Error(`Missing TypeScript module: ${entry.file}`);
    for (const symbol of checker.getExportsOfModule(module).sort((a, b) => compare(a.name, b.name))) {
      publicBinding(symbol, '', symbol.name, [], undefined, new Set(), entryQueries[entry.route], 'support');
    }
  }
  // New source references discovered during projection join this queue. External
  // entities are leaves; following all Fluent/React members would document peers.
  for (const [key, symbol] of pending) {
    const entity = entities.get(key);
    if (entity.visibility === 'external') continue;
    const node = declaration(symbol);
    const type = ts.isTypeNode(node) ? checker.getTypeAtLocation(node)
      : symbol.flags & (ts.SymbolFlags.Interface | ts.SymbolFlags.TypeAlias) ? checker.getDeclaredTypeOfSymbol(symbol) : checker.getTypeOfSymbolAtLocation(symbol, node);
    if (ts.isTypeNode(node) && node.parent && ts.isParameter(node.parent) && node.parent.type === node) {
      const parameter = node.parent;
      const call = parameter.parent;
      const signature = checker.getSignatureFromDeclaration(call);
      for (let owner = call; signature && owner && !ts.isSourceFile(owner); owner = owner.parent) {
        if (!owner.name) continue;
        const ownerSymbol = checker.getSymbolAtLocation(owner.name);
        const ownerKey = ownerSymbol && register(ownerSymbol);
        if (ownerKey) {
          entity.parameterDocumentation = { ownerKey: `${ownerKey}::${checkerText(checker.signatureToString(signature, undefined, ts.TypeFormatFlags.NoTruncation))}`, name: parameter.name.getText() };
          break;
        }
      }
    }
    entity.typeParameters = typeParameters(node.typeParameters);
    if (ts.isTypeAliasDeclaration(node)) entity.typeExpression = node.type.getText();
    entity.heritage = (node.heritageClauses || []).flatMap(clause => clause.types.map(base => typeUse(checker.getTypeAtLocation(base))));
    const view = shape(type);
    // Type spelling and references only; members/variants remain canonical on
    // the entity, with declared alias syntax retained in typeExpression.
    entity.type = { text: view.text, references: view.references };
    entity.members = view.members;
    entity.variants = view.variants;
    entity.signatures = signatures(type, key);
    const references = new Set([...view.references, ...entity.heritage.flatMap(base => base.references)]);
    for (const member of [...view.members, ...view.variants.flatMap(variant => variant.members)]) member.type.references.forEach(reference => references.add(reference));
    for (const signature of entity.signatures) {
      for (const parameter of signature.parameters) parameter.type.references.forEach(reference => references.add(reference));
      signature.returns.references.forEach(reference => references.add(reference));
      for (const member of signature.returns.members) member.type.references.forEach(reference => references.add(reference));
    }
    references.delete(key);
    entity.references = [...references].sort(compare);
    const moduleDocs = node.getSourceFile().statements[0]?.jsDoc || [];
    const moduleComment = moduleDocs.find(comment => comment.tags?.some(tag => ['packageDocumentation', 'module'].includes(tag.tagName.text)));
    if (moduleComment) entity.moduleDocSource = location(moduleComment);
    entity.moduleDoc = moduleDocs.filter(comment => comment.tags?.some(tag => ['packageDocumentation', 'module'].includes(tag.tagName.text))).map(comment => comment.getText()).join('\n');
  }
  const diagnostics = ts.getPreEmitDiagnostics(program).map(diagnostic => ({
    code: `TS${diagnostic.code}`,
    file: sourcePath(diagnostic.file ? diagnostic.file.fileName : configFile),
    line: diagnostic.file && diagnostic.start !== undefined ? diagnostic.file.getLineAndCharacterOfPosition(diagnostic.start).line + 1 : 1,
    message: ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n'),
  })).sort((a, b) => compare(a.file, b.file) || a.line - b.line || compare(a.code, b.code) || compare(a.message, b.message));
  return { compatibilityEntries, entities: [...entities.values()].sort((a, b) => compare(a.key, b.key)),
    bindings: bindings.sort((a, b) => compare(`${a.specifier}#${a.exportedName}.${a.memberPath.join('.')}`, `${b.specifier}#${b.exportedName}.${b.memberPath.join('.')}`)), diagnostics, ...(requested.length ? { entryQueries } : {}) };
}

/**
 * Resolve a public dotted access (or `specifier#access`) without inventing an
 * import for support types. Multiple legal bindings to the same entity agree;
 * ambiguous exports from different modules require the qualified spelling.
 * @param {ApiGraph} graph Source graph.
 * @param {string} reference Dotted public access, optionally prefixed by specifier#.
 * @returns {ApiEntity} Resolved entity; throws for missing or ambiguous references.
 */
function resolvePublicReference(graph, reference) {
  const separator = reference.indexOf('#');
  const specifier = separator === -1 ? undefined : reference.slice(0, separator);
  const access = separator === -1 ? reference : reference.slice(separator + 1);
  const matches = graph.bindings.filter(binding => (!specifier || binding.specifier === specifier) && [binding.exportedName, ...binding.memberPath].join('.') === access);
  const keys = [...new Set(matches.map(binding => binding.entityKey))];
  if (!keys.length) throw new Error(`Unknown public reference: ${reference}`);
  if (keys.length !== 1) throw new Error(`Ambiguous public reference: ${reference}; qualify it with specifier#`);
  const entity = graph.entities.find(item => item.key === keys[0]);
  if (!entity) throw new Error(`Missing entity for public reference: ${reference}`);
  return entity;
}

module.exports = { buildApiGraph, resolvePublicReference };

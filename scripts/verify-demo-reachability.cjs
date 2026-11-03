// Read-only semantic audit of the configured native sample, not its test consumers.
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { inspectNativeConfiguration } = require('./demo-reachability-native.cjs');
const normalize = file => file.split(path.sep).join('/');
const inside = (parent, file) => { const relative = path.relative(parent, file); return relative === '' || (!relative.startsWith('..') && !path.isAbsolute(relative)); };
const compare = (a, b) => a < b ? -1 : a > b ? 1 : 0;

/** Analyze without emitting, repairing inputs or creating missing output directories. */
function analyzeDemoReachability(root) {
  root = path.resolve(root);
  const app = path.join(root, 'apps/spfx-react-toolkit-test');
  const report = { roots: [], unreachableFiles: [], unusedExports: [], sourceDiagnostics: [], unresolvedDynamicReferences: [], retiredRuntimeReferences: [] };
  const relative = file => normalize(path.relative(root, file));
  const location = node => ({ file: relative(node.getSourceFile().fileName), line: node.getSourceFile().getLineAndCharacterOfPosition(node.getStart()).line + 1 });
  const problem = (node, code, message) => report.sourceDiagnostics.push({ ...location(node), code, message });
  const configFile = path.join(app, 'tsconfig.json');
  const config = ts.readConfigFile(configFile, ts.sys.readFile);
  if (config.error) throw new Error(ts.flattenDiagnosticMessageText(config.error.messageText, '\n'));
  const parsed = ts.parseJsonConfigFileContent(config.config, ts.sys, app, undefined, configFile);
  if (parsed.errors.length) throw new Error(ts.formatDiagnosticsWithColorAndContext(parsed.errors, { getCurrentDirectory: () => root, getCanonicalFileName: f => f, getNewLine: () => '\n' }));
  const options = { ...parsed.options, noEmit: true, noUnusedLocals: true, noUnusedParameters: true, allowUnreachableCode: false };
  // One Program supplies both strict diagnostics and semantic identities.
  const program = ts.createProgram(parsed.fileNames, options);
  const checker = program.getTypeChecker();
  const sourceRoot = options.rootDir || path.join(app, 'src');
  const sources = program.getSourceFiles().filter(source => inside(sourceRoot, source.fileName));
  const sourceSet = new Set(sources);
  const native = inspectNativeConfiguration({ root, app, options, program, checker, report, location });
  report.roots = native.roots.map(source => relative(source.fileName));
  const resolve = (text, source) => {
    const file = ts.resolveModuleName(text, source.fileName, options, ts.sys).resolvedModule?.resolvedFileName;
    return file ? program.getSourceFile(file) : undefined;
  };
  const exportsOf = source => source.symbol ? checker.getExportsOfModule(source.symbol) : [];
  const target = symbol => symbol && (symbol.flags & ts.SymbolFlags.Alias) ? checker.getAliasedSymbol(symbol) : symbol;
  const demands = new Map();
  const demandRoles = new Map();
  const edges = new Map();
  const evaluationEdges = new Map();
  const reachable = new Set();
  const evaluated = new Set();
  const dynamic = [];
  const reads = new Set();
  const writes = new Set();
  const readDeclarations = new Set();
  function recordRead(symbol) {
    if (!symbol) return;
    reads.add(symbol);
    for (const origin of checker.getRootSymbols(symbol)) { reads.add(origin); for (const declaration of origin.declarations || []) readDeclarations.add(declaration); }
    for (const declaration of symbol.declarations || []) readDeclarations.add(declaration);
  }
  function demand(source, name, role = 'runtime') {
    if (!source || !sourceSet.has(source)) return;
    if (!demandRoles.has(source)) demandRoles.set(source, new Map());
    const roles = demandRoles.get(source);
    if (!roles.has(name)) roles.set(name, new Set());
    if (roles.get(name).has(role)) return;
    roles.get(name).add(role);
    let names = demands.get(source);
    if (!names) demands.set(source, names = new Set());
    names.add(name);
    const symbol = exportsOf(source).find(item => item.name === name);
    if (!symbol) return;
    const declarations = symbol.declarations || [];
    for (const declaration of declarations) {
      if (ts.isExportSpecifier(declaration)) {
        const statement = declaration.parent.parent;
        if (statement.moduleSpecifier && ts.isStringLiteral(statement.moduleSpecifier)) demand(resolve(statement.moduleSpecifier.text, source), (declaration.propertyName || declaration.name).text, role);
        else {
          const local = checker.getExportSpecifierLocalTargetSymbol(declaration);
          consumeAlias(local, role);
        }
      }
    }
    // export * transit slots may share the original declaration symbol.
    for (const statement of source.statements) {
      if (!ts.isExportDeclaration(statement) || statement.exportClause || !statement.moduleSpecifier || !ts.isStringLiteral(statement.moduleSpecifier)) continue;
      const next = resolve(statement.moduleSpecifier.text, source);
      if (next && exportsOf(next).some(item => item.name === name && target(item) === target(symbol))) demand(next, name, role);
    }
    if (role === 'runtime') recordRead(target(symbol));
  }
  function consumeAlias(symbol, role = 'runtime') {
    if (!symbol) return;
    for (const declaration of symbol.declarations || []) {
      let current = declaration;
      while (current && !ts.isImportDeclaration(current) && !ts.isImportEqualsDeclaration(current) && !ts.isSourceFile(current)) current = current.parent;
      if (current && ts.isImportDeclaration(current) && ts.isStringLiteral(current.moduleSpecifier)) {
        const next = resolve(current.moduleSpecifier.text, current.getSourceFile());
        if (ts.isImportSpecifier(declaration)) demand(next, (declaration.propertyName || declaration.name).text, role);
        else if (ts.isImportClause(declaration)) demand(next, 'default', role);
      }
    }
    if (role === 'runtime') recordRead(target(symbol));
  }
  function runtimeReference(node) {
    for (let parent = node.parent; parent && !ts.isSourceFile(parent); parent = parent.parent) {
      if (ts.isExpressionWithTypeArguments(parent) && ts.isHeritageClause(parent.parent) && parent.parent.token === ts.SyntaxKind.ExtendsKeyword && ts.isClassLike(parent.parent.parent)) return true;
      if (ts.isTypeNode(parent)) return false;
    }
    return true;
  }
  function runtimeImport(node) {
    if (!node.importClause) return true;
    if (node.importClause.isTypeOnly) return false;
    const bindings = [node.importClause.name, ...(node.importClause.namedBindings ? ts.isNamespaceImport(node.importClause.namedBindings) ? [node.importClause.namedBindings.name] : node.importClause.namedBindings.elements.filter(item => !item.isTypeOnly).map(item => item.name) : [])].filter(Boolean);
    if (options.verbatimModuleSyntax || options.preserveValueImports) return bindings.length > 0;
    const symbols = new Set(bindings.map(binding => checker.getSymbolAtLocation(binding)));
    let found = false;
    function scan(current) {
      if (ts.isIdentifier(current) && !declarationName(current) && symbols.has(checker.getSymbolAtLocation(current)) && runtimeReference(current)) found = true;
      if (!found) ts.forEachChild(current, scan);
    }
    scan(node.getSourceFile());
    return found;
  }
  function runtimeExport(node) {
    if (node.isTypeOnly) return false;
    if (!node.exportClause || ts.isNamespaceExport(node.exportClause)) return true;
    return node.exportClause.elements.some(item => !item.isTypeOnly && Boolean(target(checker.getSymbolAtLocation(item.name))?.flags & ts.SymbolFlags.Value));
  }
  function dependency(source, node, text, evaluates = true) {
    if (/demoRegistry|(?:^|\/)panels\/|(?:^|\/)demos(?:\/|$)|legacyEntries/.test(text)) report.retiredRuntimeReferences.push({ ...location(node), target: text });
    const next = resolve(text, source);
    if (next && sourceSet.has(next)) { edges.get(source).add(next); if (evaluates) evaluationEdges.get(source).add(next); }
    else if ((text.startsWith('.') || text.startsWith('/')) && !next && !(text.endsWith('.module.css') && source.fileName.endsWith('.module.scss.ts') && fs.existsSync(source.fileName.slice(0, -3)))) problem(node, 'DEMO_MISSING_MODULE', `Cannot resolve ${text}.`);
    return next;
  }
  for (const source of sources) {
    edges.set(source, new Set());
    evaluationEdges.set(source, new Set());
    function scan(node) {
      if ((ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) && node.moduleSpecifier && ts.isStringLiteral(node.moduleSpecifier)) dependency(source, node.moduleSpecifier, node.moduleSpecifier.text, ts.isImportDeclaration(node) ? runtimeImport(node) : runtimeExport(node));
      if (ts.isImportTypeNode(node) && ts.isLiteralTypeNode(node.argument) && ts.isStringLiteral(node.argument.literal)) dependency(source, node.argument.literal, node.argument.literal.text, false);
      if (ts.isCallExpression(node) && (node.expression.kind === ts.SyntaxKind.ImportKeyword || (ts.isIdentifier(node.expression) && node.expression.text === 'require'))) {
        const argument = node.arguments[0];
        if (argument && ts.isStringLiteralLike(argument)) {
          const next = dependency(source, node, argument.text);
          if (next && sourceSet.has(next)) dynamic.push({ node, source: next });
        } else report.unresolvedDynamicReferences.push({ ...location(node), expression: node.getText() });
      }
      ts.forEachChild(node, scan);
    }
    scan(source);
  }
  function visit(source) {
    if (reachable.has(source)) return;
    reachable.add(source);
    for (const next of edges.get(source) || []) visit(next);
  }
  native.roots.forEach(source => { visit(source); demand(source, 'default'); });
  function evaluate(source) { if (evaluated.has(source)) return; evaluated.add(source); for (const next of evaluationEdges.get(source) || []) evaluate(next); }
  native.roots.forEach(evaluate);
  for (const source of reachable) if (!evaluated.has(source) && !source.isDeclarationFile) for (const statement of source.statements) {
    if (ts.isInterfaceDeclaration(statement) || ts.isTypeAliasDeclaration(statement) || ts.isEmptyStatement(statement) || ts.getModifiers(statement)?.some(modifier => modifier.kind === ts.SyntaxKind.DeclareKeyword)) continue;
    if (ts.isImportDeclaration(statement) && !runtimeImport(statement) || ts.isExportDeclaration(statement) && !runtimeExport(statement)) continue;
    problem(statement, 'DEMO_TYPE_ONLY_RUNTIME', 'Runtime code is retained only through an erased type dependency.');
  }
  for (const source of sources) if (!reachable.has(source) && !source.isDeclarationFile) report.unreachableFiles.push(relative(source.fileName));
  const isWrite = node => (ts.isBinaryExpression(node.parent) && node.parent.left === node && node.parent.operatorToken.kind === ts.SyntaxKind.EqualsToken);
  function declarationName(node) {
    const parent = node.parent;
    return parent.name === node && (ts.isDeclaration(parent) || ts.isPropertyAssignment(parent) || ts.isMethodDeclaration(parent)) || ts.isImportSpecifier(parent) || ts.isExportSpecifier(parent) || ts.isNamespaceImport(parent) || ts.isImportClause(parent) || ts.isBindingElement(parent) && parent.propertyName === node;
  }
  function moduleSource(expression, visited = new Set()) {
    if (!expression || visited.has(expression)) return undefined;
    visited.add(expression);
    const symbol = checker.getTypeAtLocation(expression).symbol;
    const source = symbol?.declarations?.find(ts.isSourceFile) || target(checker.getSymbolAtLocation(expression))?.declarations?.find(ts.isSourceFile);
    if (source) return source;
    if (ts.isCallExpression(expression) && ts.isIdentifier(expression.expression) && expression.expression.text === 'require' && ts.isStringLiteralLike(expression.arguments[0])) return resolve(expression.arguments[0].text, expression.getSourceFile());
    if (ts.isIdentifier(expression)) for (const declaration of checker.getSymbolAtLocation(expression)?.declarations || []) if (ts.isVariableDeclaration(declaration) && declaration.initializer) { const found = moduleSource(declaration.initializer, visited); if (found) return found; }
    return undefined;
  }
  function finiteKeys(expression, visited = new Set()) {
    if (!expression || visited.has(expression)) return undefined;
    visited.add(expression);
    if (ts.isParenthesizedExpression(expression) || ts.isAsExpression(expression) || ts.isTypeAssertionExpression(expression) || ts.isSatisfiesExpression(expression) || ts.isNonNullExpression(expression)) return finiteKeys(expression.expression, visited);
    if (ts.isIdentifier(expression)) for (const declaration of checker.getSymbolAtLocation(expression)?.declarations || []) if (ts.isVariableDeclaration(declaration) && declaration.initializer) {
      if (declaration.parent.flags & ts.NodeFlags.Const) return finiteKeys(declaration.initializer, visited);
      // Mutable variables still use their actual narrowed read-site type, but
      // an initial broad assertion cannot establish a finite namespace domain.
      let initializer = declaration.initializer;
      while (ts.isParenthesizedExpression(initializer) || ts.isSatisfiesExpression(initializer) || ts.isNonNullExpression(initializer)) initializer = initializer.expression;
      if ((ts.isAsExpression(initializer) || ts.isTypeAssertionExpression(initializer)) && !finiteKeys(initializer, visited)) return undefined;
    }
    const type = checker.getTypeAtLocation(expression);
    const types = type.isUnion() ? type.types : [type];
    return types.every(item => item.isStringLiteral() || item.isNumberLiteral()) ? types.map(item => String(item.value)) : undefined;
  }
  function memberRead(expression, names) {
    const type = checker.getTypeAtLocation(expression);
    for (const name of names) {
      recordRead(checker.getPropertyOfType(type, name));
      const source = moduleSource(expression);
      if (sourceSet.has(source)) demand(source, name);
    }
  }
  for (const source of reachable) {
    function scan(node) {
      if (ts.isIdentifier(node) && !declarationName(node)) {
        const symbol = checker.getSymbolAtLocation(node);
        if (!isWrite(node) && !(ts.isPropertyAccessExpression(node.parent) && node.parent.name === node && isWrite(node.parent))) {
          if (runtimeReference(node)) recordRead(symbol);
          if (symbol?.flags & ts.SymbolFlags.Alias) consumeAlias(symbol, runtimeReference(node) ? 'runtime' : 'type');
        }
      }
      if (ts.isQualifiedName(node)) { const owner = moduleSource(node.left); if (sourceSet.has(owner)) demand(owner, node.right.text, 'type'); }
      if (ts.isImportTypeNode(node) && ts.isLiteralTypeNode(node.argument) && ts.isStringLiteral(node.argument.literal) && node.qualifier) {
        let qualifier = node.qualifier; while (ts.isQualifiedName(qualifier)) qualifier = qualifier.left;
        demand(resolve(node.argument.literal.text, source), qualifier.text, 'type');
      }
      const jsxName = node.parent && (ts.isJsxOpeningElement(node.parent) || ts.isJsxClosingElement(node.parent) || ts.isJsxSelfClosingElement(node.parent) || ts.isJsxAttribute(node.parent) && node.parent.name === node);
      if (ts.isIdentifier(node) && !jsxName && !declarationName(node) && sourceSet.has(moduleSource(node))) {
        if (ts.isSpreadAssignment(node.parent) || ts.isSpreadElement(node.parent)) report.unresolvedDynamicReferences.push({ ...location(node), expression: node.parent.getText() });
        if (ts.isCallExpression(node.parent) && node.parent.arguments.includes(node)) {
          const signature = checker.getResolvedSignature(node.parent);
          const parameter = signature?.parameters[node.parent.arguments.indexOf(node)];
          const declaration = parameter?.valueDeclaration || parameter?.declarations?.[0];
          const type = declaration && checker.getTypeOfSymbolAtLocation(parameter, declaration);
          if (!type || type.flags & (ts.TypeFlags.Any | ts.TypeFlags.Unknown) || type.getStringIndexType()) report.unresolvedDynamicReferences.push({ ...location(node), expression: node.parent.getText() });
        }
      }
      if (ts.isPropertyAccessExpression(node)) { if (isWrite(node)) { const symbol = checker.getSymbolAtLocation(node.name); if (symbol) { writes.add(symbol); for (const origin of checker.getRootSymbols(symbol)) writes.add(origin); } } else memberRead(node.expression, [node.name.text]); }
      if (ts.isElementAccessExpression(node) && node.argumentExpression && !isWrite(node)) {
        const keys = finiteKeys(node.argumentExpression);
        if (keys) memberRead(node.expression, keys);
        else if (sourceSet.has(moduleSource(node.expression))) report.unresolvedDynamicReferences.push({ ...location(node), expression: node.getText() });
      }
      if (ts.isModuleDeclaration(node) && ts.isStringLiteral(node.name) && node.name.text.startsWith('.') && node.body && ts.isModuleBlock(node.body)) {
        const augmented = resolve(node.name.text, source);
        for (const statement of node.body.statements) if (statement.name && ts.isIdentifier(statement.name)) demand(augmented, statement.name.text);
      }
      if (ts.isObjectBindingPattern(node)) {
        const type = checker.getTypeAtLocation(node);
        for (const element of node.elements) if (!element.dotDotDotToken) recordRead(checker.getPropertyOfType(type, (element.propertyName || element.name).getText().replace(/^['"]|['"]$/g, '')));
      }
      if (ts.isVariableDeclaration(node) && ts.isObjectBindingPattern(node.name) && node.initializer) {
        for (const element of node.name.elements) if (!element.dotDotDotToken) memberRead(node.initializer, [(element.propertyName || element.name).getText().replace(/^['"]|['"]$/g, '')]);
      }
      ts.forEachChild(node, scan);
    }
    scan(source);
  }
  // A raw import returned through a finite typed load boundary is consumed by
  // actual reads on that boundary's Promise result. Type context alone is no use.
  for (const { node, source } of dynamic) {
    if (!reachable.has(node.getSourceFile())) continue;
    const direct = [...(demands.get(source) || [])];
    let cursor = node;
    while (ts.isParenthesizedExpression(cursor.parent)) cursor = cursor.parent;
    let fn = ts.isArrowFunction(cursor.parent) && cursor.parent.body === cursor ? cursor.parent
      : ts.isReturnStatement(cursor.parent) ? cursor.parent.parent.parent : undefined;
    let handled = ts.isIdentifier(node.expression) && node.expression.text === 'require' && (direct.length > 0 || ts.isExpressionStatement(cursor.parent));
    if (fn && (ts.isArrowFunction(fn) || ts.isFunctionExpression(fn) || ts.isMethodDeclaration(fn))) {
      const contextual = checker.getContextualType(fn);
      for (const signature of contextual?.getCallSignatures() || []) {
        const result = checker.getPromisedTypeOfPromise(checker.getReturnTypeOfSignature(signature));
        if (!result || result.flags & (ts.TypeFlags.Any | ts.TypeFlags.Unknown) || result.getStringIndexType()) continue;
        for (const property of result.getProperties()) if (reads.has(property) || property.declarations?.some(declaration => readDeclarations.has(declaration))) { demand(source, property.name); handled = true; }
      }
    }
    // Inferred import().then callbacks retain real module symbols; a finite
    // direct read resolves through aliases, brackets and destructuring above.
    if (ts.isPropertyAccessExpression(cursor.parent) && cursor.parent.expression === cursor && cursor.parent.name.text === 'then') handled = direct.length > 0 || (demands.get(source)?.size || 0) > 0;
    if (!handled) report.unresolvedDynamicReferences.push({ ...location(node), expression: node.getText() });
  }
  for (const resource of native.resources) if (resource.symbol && !reads.has(resource.symbol) && !resource.symbol.declarations?.some(declaration => readDeclarations.has(declaration))) report.sourceDiagnostics.push({ file: resource.file, line: resource.line, code: 'DEMO_UNUSED_RESOURCE_MEMBER', message: `Native dictionary member ${resource.name} has no application read.` });
  const artifactMembers = new Set();
  require('./demo-reachability-metadata.cjs').proveMetadataReplicas({ root, reachable, resolve, checker, demand: (source, name) => demand(source, name, 'artifact'), problem, artifactMembers });
  require('./demo-reachability-members.cjs').inspectMembers({ checker, reachable, reads, writes, readDeclarations, artifactMembers, problem });
  for (const diagnostic of ts.getPreEmitDiagnostics(program)) {
    if (diagnostic.file && !sourceSet.has(diagnostic.file)) continue;
    report.sourceDiagnostics.push({ file: diagnostic.file ? relative(diagnostic.file.fileName) : relative(configFile), line: diagnostic.file && diagnostic.start !== undefined ? diagnostic.file.getLineAndCharacterOfPosition(diagnostic.start).line + 1 : 1, code: `TS${diagnostic.code}`, message: ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n') });
  }
  for (const source of reachable) for (const symbol of exportsOf(source)) {
    const roles = demandRoles.get(source)?.get(symbol.name);
    if (roles?.size === 1 && roles.has('type') && target(symbol)?.flags & ts.SymbolFlags.Value) problem(symbol.declarations?.[0] || source, 'DEMO_TYPE_ONLY_EXPORT', `Export ${symbol.name} has only an erased type consumer, with no runtime body demand.`);
    if (!demands.get(source)?.has(symbol.name)) {
      const declaration = symbol.declarations?.[0] || source;
      report.unusedExports.push({ file: relative(source.fileName), symbol: symbol.name, line: location(declaration).line });
    }
  }
  for (const key of Object.keys(report)) {
    report[key] = [...new Map(report[key].map(item => [JSON.stringify(item), item])).values()];
    report[key].sort((a, b) => typeof a === 'string' ? compare(a, b) : compare(a.file, b.file) || a.line - b.line || compare(JSON.stringify(a), JSON.stringify(b)));
  }
  return report;
}
function runCli(args = []) {
  let root = path.resolve(__dirname, '..');
  try {
    if (args.length) {
      if (args.length !== 2 || args[0] !== '--root' || !args[1]) throw new Error('Usage: verify-demo-reachability.cjs [--root <repository>]');
      root = path.resolve(args[1]);
    }
    const report = analyzeDemoReachability(root);
    let failures = 0;
    for (const [category, items] of Object.entries(report)) {
      if (category === 'roots') continue;
      for (const item of items) {
        failures += 1;
        console.error(typeof item === 'string' ? `${item}:1 ${category}` : `${item.file}:${item.line} ${item.code || category} ${item.message || item.symbol || item.expression || item.target}`);
      }
    }
    console.log(`Demo reachability: ${report.roots.length} native roots, ${failures} problems.`);
    if (failures) process.exitCode = 1;
  } catch (error) { console.error(error.message); process.exitCode = 1; }
}
if (require.main === module) runCli(process.argv.slice(2));
module.exports = { analyzeDemoReachability, runCli };

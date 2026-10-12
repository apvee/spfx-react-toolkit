const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { collectPublicSurface } = require('./public-surface.helpers.cjs');
const slash = value => value.split(path.sep).join('/');
function sourceFiles(directory) {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap(item => item.isDirectory()
    ? sourceFiles(path.join(directory, item.name)) : /\.tsx?$/.test(item.name) && !item.name.endsWith('.d.ts') ? [path.join(directory, item.name)] : []).sort();
}
function auditImportGraph({ sourceRoot, entrypoints }) {
  sourceRoot = path.resolve(sourceRoot);
  const files = sourceFiles(sourceRoot);
  const relative = file => slash(path.relative(sourceRoot, file));
  const options = { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ESNext, moduleResolution: ts.ModuleResolutionKind.NodeJs, jsx: ts.JsxEmit.React, skipLibCheck: true };
  const program = ts.createProgram(files, options);
  const checker = program.getTypeChecker();
  const emittedSources = new Map();
  program.emit(undefined, (name, text, bom, error, sources) => { if (name.endsWith('.js')) for (const source of sources ?? []) emittedSources.set(source.fileName, text); });
  const edges = [], modules = [], issues = [], publicExports = [];
  function resolve(from, specifier) {
    const result = ts.resolveModuleName(specifier, from, options, ts.sys).resolvedModule;
    return result && (specifier.startsWith('.') ? relative(result.resolvedFileName) : specifier);
  }
  function runtimeExportNames(file, stack = []) {
    if (stack.includes(file)) return [];
    const emittedText = emittedSources.get(file);
    if (emittedText === undefined) return [];
    const emitted = ts.createSourceFile(file + '.js', emittedText, ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
    const names = new Set();
    const importedNamespaces = new Map();
    const bindingNames = node => node.importClause ? [node.importClause.name,
      ...(node.importClause.namedBindings && ts.isNamespaceImport(node.importClause.namedBindings)
        ? [node.importClause.namedBindings.name] : node.importClause.namedBindings?.elements?.map(element => element.name) ?? [])].filter(Boolean) : [];
    const emittedBindings = new Set(emitted.statements.filter(ts.isImportDeclaration).flatMap(node => bindingNames(node).map(name => name.text)));
    for (const node of program.getSourceFile(file).statements.filter(ts.isImportDeclaration)) {
      for (const name of bindingNames(node)) {
        if (!emittedBindings.has(name.text)) continue;
        const symbol = checker.getSymbolAtLocation(name);
        const canonical = symbol && (symbol.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(symbol) : symbol);
        const declaration = canonical?.valueDeclaration ?? canonical?.declarations?.[0];
        if (declaration && ts.isSourceFile(declaration)) importedNamespaces.set(name.text,
          runtimeExportNames(declaration.fileName, [...stack, file]));
      }
    }
    for (const node of emitted.statements) {
      if (ts.isExportDeclaration(node)) {
        const target = node.moduleSpecifier && resolve(file, node.moduleSpecifier.text);
        const leaves = target && node.moduleSpecifier.text.startsWith('.') ? runtimeExportNames(path.join(sourceRoot, target), [...stack, file]) : [];
        if (!node.exportClause) for (const name of leaves) { if (name !== 'default') names.add(name); }
        else if (ts.isNamespaceExport(node.exportClause)) for (const name of leaves) names.add(node.exportClause.name.text + '.' + name);
        else for (const element of node.exportClause.elements) {
          const localName = (element.propertyName ?? element.name).text;
          const localNamespace = !node.moduleSpecifier && importedNamespaces.get(localName);
          if (localNamespace) {
            for (const name of localNamespace) names.add(element.name.text + '.' + name);
            continue;
          }
          const nested = leaves.filter(name => name.startsWith(localName + '.'));
          if (nested.length) for (const name of nested) names.add(element.name.text + name.slice(localName.length));
          else names.add(element.name.text);
        }
      } else if (node.modifiers?.some(modifier => modifier.kind === ts.SyntaxKind.ExportKeyword)) {
        if (ts.isVariableStatement(node)) for (const declaration of node.declarationList.declarations) names.add(declaration.name.getText(emitted));
        else if (node.name) names.add(node.name.text);
      } else if (ts.isExportAssignment(node)) names.add('default');
    }
    return [...names];
  }
  for (const file of files) {
    const source = program.getSourceFile(file);
    // TypeScript's emitted ESNext imports establish whether an unmarked type
    // import/re-export is actually erased. Static syntax alone cannot do this.
    const emitted = ts.createSourceFile(file + '.js', emittedSources.get(file) ?? '', ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
    const emittedSpecifiers = new Set(emitted.statements.filter(n => (ts.isImportDeclaration(n) || ts.isExportDeclaration(n)) && n.moduleSpecifier).map(n => n.moduleSpecifier.text));
    const ownEdges = [];
    for (const node of source.statements) {
      if (!(ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) || !node.moduleSpecifier) continue;
      const specifier = node.moduleSpecifier.text;
      const target = resolve(file, specifier);
      const markedType = node.isTypeOnly || node.importClause?.isTypeOnly || node.exportClause?.elements?.every(e => e.isTypeOnly) || node.importClause?.namedBindings?.elements?.every(e => e.isTypeOnly);
      const sourceBindings = node.importClause ? [node.importClause.name, ...(node.importClause.namedBindings && ts.isNamespaceImport(node.importClause.namedBindings) ? [node.importClause.namedBindings.name] : node.importClause.namedBindings?.elements?.map(e => e.name) ?? [])].filter(Boolean).map(name => name.text) : [];
      const emittedBindings = emitted.statements.filter(n => ts.isImportDeclaration(n) && n.moduleSpecifier.text === specifier).flatMap(n => [n.importClause?.name?.text, ...(n.importClause?.namedBindings && ts.isNamespaceImport(n.importClause.namedBindings) ? [n.importClause.namedBindings.name.text] : n.importClause?.namedBindings?.elements?.map(e => e.name.text) ?? [])]);
      const erasedBindings = sourceBindings.length && sourceBindings.every(name => !emittedBindings.includes(name));
      const erasedReexports = ts.isExportDeclaration(node) && node.exportClause && ts.isNamedExports(node.exportClause) && node.exportClause.elements.every(element => {
        const symbol = checker.getSymbolAtLocation(element.name);
        const canonical = symbol && (symbol.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(symbol) : symbol);
        return canonical && !(canonical.flags & ts.SymbolFlags.Value);
      });
      const kind = markedType || erasedBindings || erasedReexports || !emittedSpecifiers.has(specifier) ? 'type-only'
        : ts.isExportDeclaration(node) ? 'reexport' : !node.importClause ? 'bare-import' : 'import';
      const edge = { from: relative(file), specifier, resolvedTarget: target ?? null, kind, declarationKind: ts.isImportDeclaration(node) ? 'import' : 'reexport' };
      edges.push(edge); ownEdges.push(edge);
      if (!target) issues.push({ ...edge, kind: 'unresolved-module' });
      if (kind === 'import' && specifier.startsWith('.') && target && /(^|\/)index\.tsx?$/.test(target)) issues.push({ ...edge, kind: 'internal-barrel-import' });
    }
    const initializations = [];
    for (const node of source.statements) {
      if (ts.isVariableStatement(node)) for (const declaration of node.declarationList.declarations) {
        if (declaration.initializer) initializations.push({ kind: ts.SyntaxKind[declaration.initializer.kind], name: declaration.name.getText(source), expression: declaration.initializer.getText(source) });
      }
      if (ts.isIfStatement(node) || ts.isTryStatement(node) || ts.isForStatement(node) || ts.isForOfStatement(node) || ts.isWhileStatement(node) || ts.isExpressionStatement(node) || ts.isClassDeclaration(node) && node.members.some(m => m.modifiers?.some(x => x.kind === ts.SyntaxKind.StaticKeyword) && m.initializer)) initializations.push({ kind: ts.SyntaxKind[node.kind], expression: node.getText(source) });
    }
    const registrations = ownEdges.filter(e => e.kind === 'bare-import' && !e.specifier.startsWith('.'));
    const allocations = initializations.filter(i => /CallExpression|NewExpression|ObjectLiteralExpression|ArrayLiteralExpression|ExpressionStatement|IfStatement|TryStatement|ForStatement|ForOfStatement|WhileStatement/.test(i.kind));
    const effects = registrations.length ? { classification: 'external-registration', rationale: 'External bare imports require manual ownership review; module evaluation may register external state.' }
      : allocations.length ? { classification: 'review-required', rationale: 'Top-level calls/allocations require manual review; static analysis does not certify purity.' }
      : { classification: 'no-observed-external-effect', rationale: 'No top-level calls/allocations or external bare imports observed; this is not a purity certificate.' };
    modules.push({ path: relative(file), externalImports: ownEdges.filter(e => !e.specifier.startsWith('.')), topLevelInitializations: initializations, effects, directConsumers: [] });
    if (registrations.length) issues.push({ kind: 'external-registration', from: relative(file), specifiers: registrations.map(e => e.specifier) });
    if (allocations.length) issues.push({ kind: 'top-level-allocation', from: relative(file), names: allocations.map(i => i.name ?? i.kind) });
  }
  for (const module of modules) module.directConsumers = [...new Set(edges.filter(e => e.resolvedTarget === module.path).map(e => e.from))].sort();
  const runtimeEdges = edges.filter(e => e.kind !== 'type-only');
  function dependencies(start) {
    const found = new Set(), visited = new Set();
    function visit(current) {
      if (visited.has(current)) return;
      visited.add(current);
      for (const edge of runtimeEdges.filter(e => e.from === current && e.resolvedTarget)) {
        if (!edge.specifier.startsWith('.')) found.add(edge.specifier);
        else visit(edge.resolvedTarget);
      }
    }
    visit(start); return [...found].sort();
  }
  const cycleKeys = new Set();
  function cycleVisit(current, stack) {
    if (stack.includes(current)) {
      const cycle = [...stack.slice(stack.indexOf(current)), current];
      const key = [...new Set(cycle)].sort().join('|');
      if (!cycleKeys.has(key)) { cycleKeys.add(key); issues.push({ kind: 'runtime-cycle', modules: cycle }); }
      return;
    }
    for (const edge of runtimeEdges.filter(e => e.from === current && e.specifier.startsWith('.') && e.resolvedTarget)) cycleVisit(edge.resolvedTarget, [...stack, current]);
  }
  // Resolve aliases through TypeScript, retaining namespace qualification. The
  // existing surface collector is reused as a cross-check where its syntax
  // supports the entry (it does not handle exported imported bindings).
  for (const entrypoint of entrypoints) {
    const file = path.resolve(sourceRoot, entrypoint);
    const source = program.getSourceFile(file);
    if (!source) { issues.push({ kind: 'missing-entrypoint', entrypoint }); continue; }
    const entrySymbol = checker.getSymbolAtLocation(source);
    const runtimeNames = new Set(runtimeExportNames(file));
    const seen = new Set();
    function visit(symbol, prefix, forceType = false) {
      if (seen.has(symbol)) return;
      seen.add(symbol);
      for (const exported of checker.getExportsOfModule(symbol)) {
        const declarations = exported.declarations ?? [];
        const explicitType = forceType || declarations.some(d => ts.isExportSpecifier(d) && (d.isTypeOnly || d.parent.parent.isTypeOnly));
        const canonical = exported.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(exported) : exported;
        const declaration = canonical.valueDeclaration ?? canonical.declarations?.[0];
        if (!declaration) { issues.push({ kind: 'unresolved-export', entrypoint, name: prefix + exported.name }); continue; }
        if (ts.isSourceFile(declaration)) visit(canonical, prefix + exported.name + '.', explicitType);
        else {
          const canonicalSource = relative(declaration.getSourceFile().fileName);
          const kind = !explicitType && runtimeNames.has(prefix + exported.name) && canonical.flags & ts.SymbolFlags.Value ? 'runtime' : 'type';
          publicExports.push({ entrypoint, name: prefix + exported.name, canonicalSource, kind, dependencies: kind === 'runtime' ? dependencies(canonicalSource) : [] });
        }
      }
      seen.delete(symbol);
    }
    if (entrySymbol) visit(entrySymbol, '');
    try {
      const collected = collectPublicSurface(file);
      for (const declaration of collected) if (!publicExports.some(e => e.entrypoint === entrypoint && e.name === declaration.name)) issues.push({ kind: 'surface-discrepancy', entrypoint, name: declaration.name });
    } catch (error) { issues.push({ kind: 'surface-collector-fallback', entrypoint, reason: error.message }); }
  }
  // Traverse each graph once, avoiding exponential re-walks of diamond imports.
  const finished = new Set(), active = [];
  function visitCycle(current) {
    if (active.includes(current)) { cycleVisit(current, active); return; }
    if (finished.has(current)) return;
    active.push(current);
    for (const edge of runtimeEdges.filter(e => e.from === current && e.specifier.startsWith('.') && e.resolvedTarget)) visitCycle(edge.resolvedTarget);
    active.pop(); finished.add(current);
  }
  for (const module of modules) visitCycle(module.path);
  return { schemaVersion: 1, publicExports, modules, edges, issues };
}
// Enforced source boundaries are separate from the descriptive audit. Effects
// metadata owns bare external registrations; local allocations are not thereby
// claimed pure, and runtime renderers/registries must remain feature-local.
function hasImmediateRendererCall(expression, allowLocalStyles = false) {
  const source = ts.createSourceFile('initializer.ts', `const value = (${expression});`, ts.ScriptTarget.Latest, true);
  const renderer = /^(createDOMRenderer|createRenderer|makeStyles|insertCSSRules|registerRenderer|createRegistry|registerFeature)$/;
  function unwrap(node) { return ts.isParenthesizedExpression(node) ? unwrap(node.expression) : node; }
  function callables(node) {
    node=unwrap(node);
    if(ts.isArrowFunction(node) || ts.isFunctionExpression(node)) return [node];
    if(ts.isConditionalExpression(node)) return [...callables(node.whenTrue), ...callables(node.whenFalse)];
    if(!ts.isCallExpression(node)) return [];
    return callables(node.expression).flatMap(factory => {
      const returned=ts.isBlock(factory.body)
        ? factory.body.statements.find(statement=>ts.isReturnStatement(statement))?.expression
        : factory.body;
      return returned ? callables(returned) : [];
    });
  }
  function visit(node) {
    // Function bodies only run when directly invoked here. A returned factory
    // remains deferred even when its outer wrapper is immediately invoked.
    if (ts.isArrowFunction(node) || ts.isFunctionExpression(node) || ts.isFunctionDeclaration(node)) return false;
    if (ts.isMethodDeclaration(node) || ts.isGetAccessorDeclaration(node) || ts.isSetAccessorDeclaration(node)) {
      // Creating an object evaluates computed keys, but not method/accessor
      // bodies or their parameter defaults.
      return ts.isComputedPropertyName(node.name) && visit(node.name.expression);
    }
    if (ts.isCallExpression(node)) {
      const callee = unwrap(node.expression);
      const invoked=callables(callee);
      const name = ts.isIdentifier(callee) ? callee.text : ts.isPropertyAccessExpression(callee) ? callee.name.text : '';
      if (renderer.test(name) && !(allowLocalStyles && name === 'makeStyles')) return true;
      for (const factory of invoked) {
        const immediateDefault = factory.parameters.some((parameter, index) => {
          const argument = node.arguments[index] && unwrap(node.arguments[index]);
          const absent = !argument || ts.isIdentifier(argument) && argument.text === 'undefined';
          return absent && parameter.initializer && visit(parameter.initializer);
        });
        if (immediateDefault || visit(factory.body)) return true;
      }
    }
    return ts.forEachChild(node, visit) === true;
  }
  return visit(source);
}
function checkSourceRules(report, sideEffects) {
  const issues = [];
  const classified = Array.isArray(sideEffects) ? new Set(sideEffects) : new Set();
  for (const edge of report.edges) {
    if (edge.declarationKind === 'reexport' || !['import', 'bare-import', 'type-only'].includes(edge.kind)) continue;
    if (edge.specifier.startsWith('.') && /(^|\/)index\.tsx?$/.test(edge.resolvedTarget || '')) issues.push({ ...edge, kind: 'internal-barrel-import' });
    if (/^@apvee\/spfx-react-toolkit(?:\/|$)/.test(edge.specifier)) issues.push({ ...edge, kind: 'internal-package-import' });
  }
  for (const module of report.modules) {
    const registrations = module.externalImports.filter(edge => edge.kind === 'bare-import');
    const emittedPath = './lib/' + module.path.replace(/\.tsx?$/, '.js');
    if (registrations.length && !classified.has(emittedPath)) issues.push({ kind: 'unclassified-external-registration', from: module.path, specifiers: registrations.map(edge => edge.specifier) });
    if (module.topLevelInitializations.some(item => hasImmediateRendererCall(item.expression, ['controls/formatted-text/formatted-text.styles.internal.ts', 'controls/auto-skeleton/auto-skeleton.styles.internal.ts', 'controls/state-message/state-message.styles.internal.ts', 'controls/group-header.internal.tsx', 'controls/settings-group/settings-group.styles.internal.ts', 'controls/accordion-group/accordion-group.styles.internal.ts', 'controls/form/controls/form-controls.styles.internal.ts', 'controls/form/composition/form-sections.styles.internal.ts', 'controls/data-grid/data-grid.styles.internal.ts', 'controls/web-part-shell/web-part-shell.styles.internal.ts', 'controls/web-part-shell/title/web-part-title-styles.internal.ts', 'controls/web-part-shell/title/web-part-title-editor.internal.tsx', 'controls/web-part-shell/title/web-part-title-toolbar.internal.tsx', 'controls/stage/stage.styles.internal.ts', 'controls/pagination/pagination.styles.internal.ts', 'controls/description-list/description-list.styles.internal.ts'].includes(module.path)))) issues.push({ kind: 'package-global-renderer', from: module.path });
  }
  return issues;
}
function inspectBuild(sourceRoot, libRoot) {
  const expected = new Set(sourceFiles(sourceRoot).flatMap(file => {
    const stem = slash(path.relative(sourceRoot, file)).replace(/\.tsx?$/, '');
    return [stem + '.js', stem + '.d.ts'];
  }));
  const actual = [];
  function walk(dir) { for (const item of fs.readdirSync(dir, { withFileTypes: true })) item.isDirectory() ? walk(path.join(dir, item.name)) : actual.push(slash(path.relative(libRoot, path.join(dir, item.name)))); }
  walk(libRoot);
  const stale = actual.filter(file => /\.(js|d\.ts)$/.test(file) && !expected.has(file));
  const missing = [...expected].filter(file => !actual.includes(file));
  const commonJS = actual.filter(file => file.endsWith('.js') && /\b(?:module\.exports|exports\.)/.test(fs.readFileSync(path.join(libRoot, file), 'utf8')));
  return { stale, missing, commonJS, expectedFiles: expected.size, actualFiles: actual.length };
}
if (require.main === module) {
  const root = path.resolve(__dirname, '..');
  const sourceRoot = path.join(root, 'packages/spfx-react-toolkit/src');
  const output = path.resolve(process.argv[2] || '.docs/maintenance/evidence/tree-shaking/initial/import-graph.json');
  const report = auditImportGraph({ sourceRoot, entrypoints: ['index.ts', 'core/index.ts', 'hooks/index.ts', 'services/index.ts', 'helpers/index.ts', 'helpers/styles/index.ts', 'controls/index.ts', 'controls/formatted-text/index.ts', 'controls/relative-time/index.ts', 'controls/auto-skeleton/index.ts', 'controls/state-message/index.ts', 'controls/settings-group/index.ts', 'controls/accordion-group/index.ts', 'controls/form/index.ts', 'controls/data-grid/index.ts', 'controls/web-part-shell/index.ts', 'controls/stage/index.ts', 'controls/pagination/index.ts', 'controls/description-list/index.ts'] });
  const manifest = JSON.parse(fs.readFileSync(path.join(root, 'packages/spfx-react-toolkit/package.json')));
  report.peerDependencyContracts = manifest.peerDependencies;
  report.runtimeDependencyContracts = manifest.dependencies;
  report.sourceRuleIssues = checkSourceRules(report, manifest.sideEffects);
  for (const module of report.modules) {
    for (const edge of module.externalImports) {
      const packageName = edge.specifier.startsWith('@') ? edge.specifier.split('/').slice(0, 2).join('/') : edge.specifier.split('/')[0];
      edge.peerPackage = packageName;
      edge.peerRange = manifest.peerDependencies[packageName] ?? null;
      edge.runtimeDependencyRange = manifest.dependencies?.[packageName];
    }
    // These classifications record the manual source review for this baseline.
    // Unknown/new effectful modules retain review-required rather than inheriting
    // a broad declaration that all toolkit code is pure.
    const observed = module.topLevelInitializations;
    const hasRegistration = module.externalImports.some(e => e.kind === 'bare-import');
    let manual = { classification: 'definitions-only', rationale: 'Manual review: declarations and imported bindings; no observed top-level allocation/registration. This does not certify transitive dependency purity.' };
    if (hasRegistration) manual = { classification: 'feature-registration', rationale: 'Manual review: PnP feature modules augment prototypes; retain evaluation only for consumers of the owning feature.' };
    else if (observed.length && /^helpers\/styles\//.test(module.path)) manual = { classification: 'local-allocation', rationale: 'Manual review: descriptors/recipes copy and freeze local data; brand Symbol, WeakMap cache and priority tables allocate local state. No renderer/CSS/DOM work runs at module evaluation; shared brand/cache identity must survive consumer composition.' };
    else if (observed.length && ['core/context.internal.tsx', 'core/state.internal.tsx'].includes(module.path)) manual = { classification: 'local-context-allocation', rationale: 'Manual review: React.createContext allocates local context objects and development displayName assignments mutate those objects only; provider runtime stores remain instance-owned.' };
    else if (observed.length && module.path === 'core/teams-runtime.internal.ts') manual = { classification: 'local-cache-allocation', rationale: 'Manual source review: WeakMap keys SDK identity and holds per-store subscriptions. SDK handler registration and theme requests are deferred to acquisition; cleanup removes store references. No window registry or module-evaluation SDK call.' };
    else if (observed.length && module.path === 'helpers/spfx-api-permission-precheck.helpers.ts') manual = { classification: 'local-allocation', rationale: 'Manual review: Sets of expected audiences/statuses are local lookup data; requests and diagnostics run only in exported functions.' };
    else if (observed.length && module.path === 'controls/formatted-text/formatted-text-content.internal.tsx') manual = { classification: 'local-parser-allocation', rationale: 'Manual review: unified processor configures its own parser/plugins only; no DOM, renderer, external registration or text cache.' };
    else if (observed.length && module.path === 'controls/auto-skeleton/auto-skeleton.styles.internal.ts') manual = { classification: 'local-style-allocation', rationale: 'makeStyles defines local rules; renderer and insertion happen only through mounted hooks, no package renderer.' };
    else if (observed.length && module.path === 'controls/formatted-text/formatted-text.styles.internal.ts') manual = { classification: 'local-style-allocation', rationale: 'Manual review: makeStyles defines component-owned rules; renderer/context/CSS work happens through the hook, no package-wide renderer or registry.' };
    else if (observed.length && module.path === 'controls/description-list/description-list.tsx') manual = { classification: 'local-component-allocation', rationale: 'Manual review: createContext, forwardRef and Object.assign define local components/context only. Rendering uses host styles; no subscriptions, external registrations or package renderer.' };
    else if (observed.length && module.path === 'controls/formatted-text/formatted-text.tsx') manual = { classification: 'local-component-allocation', rationale: 'Manual review: forwardRef creates a local component; descriptors/styles and parsing run when it renders.' };
    else if (observed.length && module.path === 'helpers/spfx-time-zone.internal.ts') manual = { classification: 'local-data-allocation', rationale: 'Audited readonly IANA lookup is a local object literal, no DOM/renderer/requests/registration.' };
    else if (observed.length && module.path === 'controls/auto-skeleton/auto-skeleton.tsx') manual = { classification: 'local-component-allocation', rationale: 'forwardRef defines the component; DOM measurement, subscriptions and geometry classes run only when mounted with per-instance cleanup.' };
    else if (observed.length && module.path === 'controls/relative-time/relative-time.tsx') manual = { classification: 'local-component-allocation', rationale: 'forwardRef defines a component; formatter/timers/visibility hooks run only when mounted, with per-instance cleanup.' };
    else if (observed.length && module.path === 'controls/state-message/state-message.styles.internal.ts') manual = { classification: 'local-style-allocation', rationale: 'Component-owned makeStyles definitions; insertion only through mounted hooks with host renderer.' };
    else if (observed.length && module.path === 'controls/state-message/state-message.tsx') manual = { classification: 'local-component-allocation', rationale: 'forwardRef and immutable descriptors allocate local definitions; theme reading and artwork rendering only run when mounted.' };
    else if (observed.length && ['controls/group-header.internal.tsx', 'controls/settings-group/settings-group.styles.internal.ts', 'controls/accordion-group/accordion-group.styles.internal.ts', 'controls/form/controls/form-controls.styles.internal.ts', 'controls/form/composition/form-sections.styles.internal.ts', 'controls/data-grid/data-grid.styles.internal.ts', 'controls/web-part-shell/web-part-shell.styles.internal.ts', 'controls/web-part-shell/title/web-part-title-styles.internal.ts', 'controls/web-part-shell/title/web-part-title-editor.internal.tsx', 'controls/web-part-shell/title/web-part-title-toolbar.internal.tsx', 'controls/stage/stage.styles.internal.ts', 'controls/pagination/pagination.styles.internal.ts', 'controls/description-list/description-list.styles.internal.ts'].includes(module.path)) manual = { classification: 'local-style-allocation', rationale: 'Manual review: makeStyles defines component-local token/selector rules; useSx reads the host renderer only during mounted rendering, with no package renderer or external registration.' };
    else if (observed.length && ['controls/settings-group/settings-group.tsx', 'controls/accordion-group/accordion-group.tsx'].includes(module.path)) manual = { classification: 'local-component-allocation', rationale: 'Manual review: forwardRef and Object.assign allocate local compound component definitions; immutable responsive descriptors hold data, while parsing/context/motion run only on render.' };
    else if (observed.length && ['controls/form/form-context.internal.tsx', 'controls/form/field-path.internal.tsx', 'controls/settings-field-context.internal.tsx'].includes(module.path)) manual = { classification: 'local-context-allocation', rationale: 'Manual review: React contexts allocate local identity; methods, prefixes, disabled scopes and settings label relationships are supplied by mounted providers. No external registration or renderer creation.' };
    else if (observed.length && module.path === 'controls/form/form-methods.internal.ts') manual = { classification: 'local-cache-allocation', rationale: 'Manual review: WeakMap stores metadata keyed by individual RHF methods; reset wrappers and validators are installed per form instance. The legacy null file literal denotes cleared native selection. No package registry, renderer or external registration.' };
    else if (observed.length && ['controls/form/controls/form-controls.styles.internal.ts', 'controls/form/composition/form-sections.styles.internal.ts', 'controls/data-grid/data-grid.styles.internal.ts'].includes(module.path)) manual = { classification: 'local-style-allocation', rationale: 'Manual review: makeStyles defines adapter-owned rules. CSS insertion runs only through mounted hooks using the shared host renderer; there is no package-wide renderer.' };
    else if (observed.length && ['controls/form/form.tsx', 'controls/form/form-error.tsx', 'controls/form/form-reset.tsx', 'controls/form/form-submit.tsx',
      'controls/form/controls/form-input.tsx', 'controls/form/controls/form-textarea.tsx', 'controls/form/controls/form-search-box.tsx',
      'controls/form/controls/form-checkbox.tsx', 'controls/form/controls/form-switch.tsx', 'controls/form/controls/form-spin-button.tsx',
      'controls/form/controls/form-slider.tsx', 'controls/form/controls/form-select.tsx', 'controls/form/controls/form-radio-group.tsx',
      'controls/form/controls/form-dropdown.tsx', 'controls/form/controls/form-dropdown-multi.tsx', 'controls/form/controls/form-checkbox-group.tsx',
      'controls/form/controls/form-combobox.tsx', 'controls/form/controls/form-combobox-multi.tsx', 'controls/form/controls/form-tag-picker.tsx',
      'controls/form/controls/form-rating.tsx', 'controls/form/controls/form-color-picker.tsx', 'controls/form/controls/form-swatch-picker.tsx',
      'controls/form/controls/form-date-picker.tsx', 'controls/form/controls/form-time-picker.tsx', 'controls/form/controls/form-date-time.tsx',
      'controls/form/controls/form-file-input.tsx'].includes(module.path)) manual = { classification: 'local-component-allocation', rationale: 'Manual review: forwardRef and the complete compound suite allocate local component definitions. Field registration, parsing, validation, subscriptions and browser work run only for mounted forms. No external registration, package renderer or PURE masking; retaining the complete family is measured separately.' };
    else if (observed.length && module.path === 'controls/form/composition/form-array-section.tsx') manual = { classification: 'local-data-allocation', rationale: 'Manual review: default labels and announcement formatters are local data/functions. Add/remove/reorder, DOM focus and announcements run only for mounted array sections; no renderer or external registration.' };
    else if (observed.length && module.path === 'controls/data-grid/data-grid.tsx') manual = { classification: 'local-component-allocation', rationale: 'Manual source review: forwardRef/Object.assign construct the dynamic compound suite and its private per-family descriptor metadata; frozen loading rows and neutral sort data are local. Rendering, parsing, DOM events and host renderer use occur only when mounted. No external registry, renderer creation or PURE masking.' };
    else if (observed.length && module.path === 'controls/data-grid/state-validation.internal.ts') manual = { classification: 'local-data-allocation', rationale: 'Manual source review: Object.freeze stores default strings and local formatter functions. Validation and per-grid resolution occur on invocation; no external registrations or renderer creation.' };
    else if (observed.length && module.path === 'controls/web-part-shell/web-part-shell.tsx') manual = { classification: 'local-component-allocation', rationale: 'Manual source review: forwardRef defines Shell. Its existing provider, DOM work and title loader execute only for mounted instances; no package renderer or registry.' };
    else if (observed.length && module.path === 'controls/web-part-shell/title/web-part-title.internal.tsx') manual = { classification: 'local-lazy-component-allocation', rationale: 'Manual source review: React.lazy creates stable component identity without invoking its import loader. Editor chunk loading occurs only for visible Edit titles; no engine creation at evaluation.' };
    else if (observed.length && module.path === 'controls/web-part-shell/title/web-part-title-codec.internal.ts') manual = { classification: 'local-data-allocation', rationale: 'Manual source review: two Sets contain only parser tag names. HTML parsing runs on invocation using the supplied document; no DOM, renderer or external registration at evaluation.' };
    else if (observed.length && module.path === 'controls/web-part-shell/title/web-part-title-extensions.internal.ts') manual = { classification: 'local-extension-allocation', rationale: 'Manual source review: Tiptap Mark/Extension/HardBreak factories define private schema objects in the async module. Plugins and engine instances are created only by mounted editor calls; no global feature registry or renderer.' };
    else if (observed.length && module.path === 'controls/pagination/pagination.internal.ts') manual = { classification: 'local-allocation', rationale: 'Manual source review: fixed English defaults allocate local label data; validation and label resolution run only during rendering. No external registration, renderer or CSS insertion.' };
    else if (observed.length && module.path === 'controls/pagination/pagination.tsx') manual = { classification: 'local-component-allocation', rationale: 'Manual source review: forwardRef and displayName allocate a local component; navigation, host styling and event requests run only when mounted. No package renderer or external registration.' };
    else if (observed.length && module.path === 'controls/stage/stage-context.internal.tsx') manual = { classification: 'local-context-allocation', rationale: 'Manual source review: React.createContext defines local identity; IDs, description registration and controller state belong to mounted Stage instances. No renderer or external registration.' };
    else if (observed.length && ['controls/stage/stage.tsx', 'controls/stage/stage-header.tsx', 'controls/stage/stage-body.tsx', 'controls/stage/stage-footer.tsx'].includes(module.path)) manual = { classification: 'local-component-allocation', rationale: 'Manual source review: forwardRef/Object.assign define local compound components. Context, portals, DOM inert synchronization, focus, styles and requests execute only when mounted. No global renderer or feature registry.' };
    else if (module.effects.classification === 'review-required') manual = { classification: 'review-required', rationale: 'New/unclassified top-level effects require an explicit review checkpoint.' };
    module.effects.manual = manual;
  }
  report.build = inspectBuild(sourceRoot, path.join(root, 'packages/spfx-react-toolkit/lib'));
  fs.mkdirSync(path.dirname(output), { recursive: true }); fs.writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  console.log(`Audit: ${report.modules.length} modules; ${report.publicExports.length} exports; ${report.issues.length} observations; build ${JSON.stringify(report.build)}`);
  if (report.sourceRuleIssues.length || report.build.stale.length || report.build.missing.length || report.build.commonJS.length || report.issues.some(i => ['missing-entrypoint', 'unresolved-module', 'unresolved-export', 'surface-discrepancy'].includes(i.kind))) process.exitCode = 1;
}
module.exports = { auditImportGraph, inspectBuild, checkSourceRules };

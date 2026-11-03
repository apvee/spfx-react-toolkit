const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');

/** Checks only authored sample source; toolkit compound controls/native Fluent slots are allowed. */
function assertDemoControls(directory) {
  const violations = [];
  const files = [];
  function collect(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const file = path.join(dir, entry.name);
      if (entry.isDirectory()) collect(file);
      else if (/\.tsx?$/.test(file)) files.push(file);
    }
  }
  collect(directory);
  const program = ts.createProgram(files, { noResolve: true, noLib: true, jsx: ts.JsxEmit.React });
  const checker = program.getTypeChecker();
  for (const file of files) {
    const source = program.getSourceFile(file);
    const providerBindings = new Set();
    const providerNamespaces = new Set();
    for (const statement of source.statements) {
      if (!ts.isImportDeclaration(statement) || !ts.isStringLiteral(statement.moduleSpecifier)
        || !['@fluentui/react-components', '@fluentui/react-provider'].includes(statement.moduleSpecifier.text)) continue;
      const bindings = statement.importClause?.namedBindings;
      if (bindings && ts.isNamedImports(bindings)) for (const binding of bindings.elements) {
        if ((binding.propertyName || binding.name).text === 'FluentProvider') providerBindings.add(checker.getSymbolAtLocation(binding.name));
      }
      if (bindings && ts.isNamespaceImport(bindings)) providerNamespaces.add(checker.getSymbolAtLocation(bindings.name));
    }
    const isProvider = tag => ts.isIdentifier(tag)
      ? tag.text === 'FluentProvider' || providerBindings.has(checker.getSymbolAtLocation(tag))
      : ts.isPropertyAccessExpression(tag) && tag.name.text === 'FluentProvider' && providerNamespaces.has(checker.getSymbolAtLocation(tag.expression));
    // This native host has no outer Shell theme owner. Recognize only its one
    // provider-owned projection, with binding provenance rather than a filename exemption.
    const nativeTopFile = path.relative(directory, file).split(path.sep).join('/') === 'components/shared/ApplicationCustomizerExampleHost.tsx';
    const hostHooks = new Set(), hostProviders = new Set();
    for (const statement of source.statements) {
      if (!ts.isImportDeclaration(statement) || statement.moduleSpecifier.text !== '@apvee/spfx-react-toolkit') continue;
      const bindings = statement.importClause?.namedBindings;
      if (!bindings || !ts.isNamedImports(bindings)) continue;
      for (const binding of bindings.elements) {
        const name = (binding.propertyName || binding.name).text;
        if (name === 'useSPFxFluent9ThemeInfo') hostHooks.add(checker.getSymbolAtLocation(binding.name));
        if (name === 'SPFxApplicationCustomizerProvider') hostProviders.add(checker.getSymbolAtLocation(binding.name));
      }
    }
    const jsxNodes = [];
    const gather = node => { if (ts.isJsxOpeningElement(node) || ts.isJsxSelfClosingElement(node)) jsxNodes.push(node); ts.forEachChild(node, gather); };
    gather(source);
    const enclosingFunction = node => ts.findAncestor(node, ts.isFunctionDeclaration);
    const hostThemeProjection = (node, attribute) => {
      if (!nativeTopFile || !attribute.initializer || !ts.isJsxExpression(attribute.initializer) || !ts.isIdentifier(attribute.initializer.expression)) return false;
      const themed = jsxNodes.filter(item => isProvider(item.tagName) && item.attributes.properties.some(prop => ts.isJsxAttribute(prop) && prop.name.text === 'theme'));
      if (themed.length !== 1 || themed[0] !== node || node.attributes.properties.some(ts.isJsxSpreadAttribute)) return false;
      const component = enclosingFunction(node);
      if (!component?.body || node.parent !== component.body.statements.find(ts.isReturnStatement)?.expression) return false;
      const themeSymbol = checker.getSymbolAtLocation(attribute.initializer.expression);
      const binding = themeSymbol?.valueDeclaration;
      if (!binding || !ts.isBindingElement(binding) || (binding.propertyName || binding.name).getText(source) !== 'theme') return false;
      const declaration = binding.parent.parent;
      if (!ts.isVariableDeclaration(declaration) || !(declaration.parent.flags & ts.NodeFlags.Const) || enclosingFunction(declaration) !== component ||
        !declaration.initializer || !ts.isCallExpression(declaration.initializer) || declaration.initializer.arguments.length ||
        !hostHooks.has(checker.getSymbolAtLocation(declaration.initializer.expression))) return false;
      let written = false;
      const writes = node => {
        const target = ts.isBinaryExpression(node) && node.operatorToken.kind >= ts.SyntaxKind.FirstAssignment && node.operatorToken.kind <= ts.SyntaxKind.LastAssignment ? node.left :
          (ts.isPrefixUnaryExpression(node) || ts.isPostfixUnaryExpression(node)) && [ts.SyntaxKind.PlusPlusToken, ts.SyntaxKind.MinusMinusToken].includes(node.operator) ? node.operand : undefined;
        let base = target;
        while (base && (ts.isPropertyAccessExpression(base) || ts.isElementAccessExpression(base))) base = base.expression;
        if (base && checker.getSymbolAtLocation(base) === themeSymbol) written = true;
        ts.forEachChild(node, writes);
      };
      writes(component);
      if (written) return false;
      const nativeProviders = jsxNodes.filter(item => ts.isIdentifier(item.tagName) && hostProviders.has(checker.getSymbolAtLocation(item.tagName)));
      if (nativeProviders.length !== 1) return false;
      const nativeProvider = nativeProviders[0], wrapper = enclosingFunction(nativeProvider);
      if (!wrapper?.body || wrapper.name?.text !== 'ApplicationCustomizerExampleHost' || !ts.getModifiers(wrapper)?.some(modifier => modifier.kind === ts.SyntaxKind.ExportKeyword)) return false;
      const instanceBinding = wrapper.parameters[0]?.name;
      if (!instanceBinding || !ts.isObjectBindingPattern(instanceBinding)) return false;
      const instance = instanceBinding.elements.find(item => (item.propertyName || item.name).getText(source) === 'instance');
      const input = nativeProvider.attributes.properties.find(prop => ts.isJsxAttribute(prop) && prop.name.text === 'instance');
      if (!instance || !input?.initializer || !ts.isJsxExpression(input.initializer) || !input.initializer.expression ||
        checker.getSymbolAtLocation(input.initializer.expression) !== checker.getSymbolAtLocation(instance.name) || nativeProvider.attributes.properties.some(ts.isJsxSpreadAttribute)) return false;
      const calls = jsxNodes.filter(item => ts.isIdentifier(item.tagName) && checker.getSymbolAtLocation(item.tagName) === checker.getSymbolAtLocation(component.name));
      return calls.length === 1 && Boolean(ts.findAncestor(calls[0], ancestor => ancestor === nativeProvider.parent));
    };
    const browserDialogs = new Set(['confirm', 'prompt', 'alert']);
    const isUnbound = identifier => ts.isIdentifier(identifier) && !checker.getSymbolAtLocation(identifier)?.declarations?.length;
    const browserDialogName = target => {
      if (ts.isIdentifier(target)) return isUnbound(target) ? target.text : undefined;
      if (!ts.isPropertyAccessExpression(target) && !ts.isElementAccessExpression(target)) return undefined;
      if (!isUnbound(target.expression) || !['window', 'globalThis'].includes(target.expression.text)) return undefined;
      if (ts.isPropertyAccessExpression(target)) return target.name.text;
      return target.argumentExpression && ts.isStringLiteral(target.argumentExpression) ? target.argumentExpression.text : undefined;
    };
    const record = (node, text) => {
      const location = source.getLineAndCharacterOfPosition(node.getStart(source));
      violations.push(`${path.relative(directory, file)}:${location.line + 1}: ${text}`);
    };
    const iconRegions = new Set(['icon', 'expandIcon', 'collapseIcon', 'actionIcon']);
    const hasPlaceholderIcon = (node, seen = new Set()) => {
      if ((ts.isJsxOpeningElement(node) || ts.isJsxSelfClosingElement(node)) && ['svg', 'i'].includes(node.tagName.getText(source))) return true;
      if ((ts.isStringLiteral(node) || ts.isJsxText(node)) && /[\u2190-\u21ff\u25a0-\u27bf]|\p{Extended_Pictographic}/u.test(node.text)) return true;
      if (ts.isIdentifier(node)) {
        const symbol = checker.getSymbolAtLocation(node);
        if (symbol && !seen.has(symbol)) {
          seen.add(symbol);
          for (const declaration of symbol.declarations || []) {
            if (ts.isVariableDeclaration(declaration) && declaration.initializer && hasPlaceholderIcon(declaration.initializer, seen)) return true;
          }
        }
      }
      return !!ts.forEachChild(node, child => hasPlaceholderIcon(child, seen));
    };
    function visit(node) {
      if ((ts.isJsxAttribute(node) || ts.isPropertyAssignment(node)) && iconRegions.has(node.name.getText(source)) && node.initializer && hasPlaceholderIcon(node.initializer)) record(node, 'placeholder icon: use a Fluent SVG icon');
      if (ts.isImportDeclaration(node) && ts.isStringLiteral(node.moduleSpecifier)) {
        const name = node.moduleSpecifier.text;
        if (name === '@fluentui/react' || name.startsWith('@fluentui/react/')) record(node, 'Fluent UI 8 demo import');
        const bindings = node.importClause?.namedBindings;
        if (bindings && ts.isNamedImports(bindings)) for (const binding of bindings.elements) {
          if (/^(webLightTheme|webDarkTheme|teamsLightTheme|teamsDarkTheme|teamsHighContrastTheme|getTeamsFluentTheme|createTheme)$/.test((binding.propertyName || binding.name).text)) record(binding, 'local demo theme');
        }
      }
      if (ts.isJsxOpeningElement(node) || ts.isJsxSelfClosingElement(node)) {
        const tag = node.tagName.getText(source);
        if (['input', 'select', 'textarea', 'button', 'details', 'summary'].includes(tag)) record(node, `authored native interactive <${tag}>`);
        if (nativeTopFile && isProvider(node.tagName) && node.attributes.properties.some(ts.isJsxSpreadAttribute)) record(node, 'host theme projection cannot use spread props');
        for (const attribute of node.attributes.properties) if (ts.isJsxAttribute(attribute)) {
          if (attribute.name.text === 'style') record(attribute, 'authored inline style');
          if (isProvider(node.tagName) && attribute.name.text === 'theme' && !hostThemeProjection(node, attribute)) record(attribute, 'panel-owned theme');
        }
      }
      if (ts.isCallExpression(node)) {
        const nativeName = browserDialogName(node.expression);
        if (browserDialogs.has(nativeName)) record(node, `native browser dialog ${nativeName}`);
      }
      ts.forEachChild(node, visit);
    }
    visit(source);
  }
  if (violations.length) throw new Error(`Demo controls contract violations:\n${violations.join('\n')}`);
  return files.length;
}
module.exports = { assertDemoControls };

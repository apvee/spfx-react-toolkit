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
        for (const attribute of node.attributes.properties) if (ts.isJsxAttribute(attribute)) {
          if (attribute.name.text === 'style') record(attribute, 'authored inline style');
          if (isProvider(node.tagName) && attribute.name.text === 'theme') record(attribute, 'panel-owned theme');
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

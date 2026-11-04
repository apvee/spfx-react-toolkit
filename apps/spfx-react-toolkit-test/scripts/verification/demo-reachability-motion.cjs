// Pinned Fluent native-root motion handoff, including React 17's inert delta.
// A generic cloneElement<P> annotation alone grants no member demand.
const ts = require('typescript');
function inspectNativeMotion({ checker, reachable, markConcrete }) {
  const target = symbol => symbol?.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(symbol) : symbol;
  const symbol = node => node && target(checker.getSymbolAtLocation(ts.isPropertyAccessExpression(node) ? node.name : node));
  const reactCall = (node, name) => ts.isCallExpression(node) && symbol(node.expression)?.name === name && symbol(node.expression).declarations?.some(declaration => declaration.getSourceFile().fileName.includes('/node_modules/@types/react/'));
  function unwrap(node, visited = new Set()) {
    if (!node || visited.has(node)) return node;
    visited.add(node);
    if (ts.isParenthesizedExpression(node) || ts.isAsExpression(node)) return unwrap(node.expression, visited);
    if (ts.isIdentifier(node)) {
      const declaration = checker.getSymbolAtLocation(node)?.valueDeclaration;
      if (declaration && ts.isVariableDeclaration(declaration) && declaration.initializer) return unwrap(declaration.initializer, visited);
    }
    return node;
  }
  function genuinePanel(tag) {
    const component = symbol(tag);
    if (component?.name !== 'AccordionPanel') return false;
    const owner = component.declarations?.find(declaration => declaration.getSourceFile().fileName.includes('/node_modules/@fluentui/react-accordion/'))?.getSourceFile();
    const slots = owner?.symbol && checker.getExportsOfModule(owner.symbol).find(item => item.name === 'AccordionPanelSlots');
    const declaration = slots?.declarations?.find(ts.isTypeAliasDeclaration);
    if (!declaration || !ts.isTypeLiteralNode(declaration.type)) return false;
    const root = declaration.type.members.find(member => member.name?.getText() === 'root');
    const motion = declaration.type.members.find(member => member.name?.getText() === 'collapseMotion');
    const slot = root?.type && ts.isTypeReferenceNode(root.type) && root.type.typeArguments?.[0];
    return Boolean(motion && slot && ts.isTypeReferenceNode(slot) && slot.typeArguments?.length === 1 && ts.isLiteralTypeNode(slot.typeArguments[0]) && ts.isStringLiteral(slot.typeArguments[0].literal) && slot.typeArguments[0].literal.text === 'div');
  }
  for (const source of reachable) {
    function scan(node) {
      if (ts.isJsxOpeningElement(node) || ts.isJsxSelfClosingElement(node)) {
        if (genuinePanel(node.tagName)) {
          const attribute = node.attributes.properties.find(item => ts.isJsxAttribute(item) && item.name.text === 'collapseMotion');
          const value = attribute?.initializer && ts.isJsxExpression(attribute.initializer) && unwrap(attribute.initializer.expression);
          const children = value && ts.isObjectLiteralExpression(value) && value.properties.find(item => ts.isPropertyAssignment(item) && item.name.getText() === 'children');
          const callback = children && unwrap(children.initializer);
          if (callback && ts.isFunctionLike(callback) && callback.parameters.length === 2 && callback.parameters.every(parameter => ts.isIdentifier(parameter.name))) inspectCallback(callback);
        }
      }
      ts.forEachChild(node, scan);
    }
    scan(source);
  }
  function inspectCallback(callback) {
    const component = checker.getSymbolAtLocation(callback.parameters[0].name);
    const props = checker.getSymbolAtLocation(callback.parameters[1].name);
    const sourceMember = (node, name) => node && ts.isPropertyAccessExpression(node) && node.name.text === name && checker.getSymbolAtLocation(node.expression) === props;
    function inspectReturn(expression) {
      expression = unwrap(expression);
      if (!expression || !ts.isJsxElement(expression) || checker.getSymbolAtLocation(expression.openingElement.tagName) !== component) return;
      const retained = expression.openingElement.attributes.properties.some(attribute => ts.isJsxAttribute(attribute) && attribute.name.text === 'unmountOnExit' && attribute.initializer && ts.isJsxExpression(attribute.initializer) && attribute.initializer.expression?.kind === ts.SyntaxKind.FalseKeyword);
      if (!retained) return;
      for (const child of expression.children) {
        const conditional = ts.isJsxExpression(child) && unwrap(child.expression);
        if (!conditional || !ts.isConditionalExpression(conditional) || !reactCall(conditional.condition, 'isValidElement') || !sourceMember(conditional.condition.arguments[0], 'children') || !sourceMember(conditional.whenFalse, 'children')) continue;
        const clone = unwrap(conditional.whenTrue);
        if (!reactCall(clone, 'cloneElement') || !sourceMember(clone.arguments[0], 'children')) continue;
        const delta = unwrap(clone.arguments[1]);
        if (!delta || !ts.isObjectLiteralExpression(delta)) continue;
        for (const member of delta.properties) {
          if (!ts.isPropertyAssignment(member) || member.name.getText() !== 'inert') continue;
          const policy = unwrap(member.initializer);
          if (policy && ts.isConditionalExpression(policy) && sourceMember(policy.condition, 'visible') && ts.isIdentifier(policy.whenTrue) && policy.whenTrue.text === 'undefined' && ts.isStringLiteral(policy.whenFalse) && policy.whenFalse.text === '') markConcrete(member);
        }
      }
    }
    if (!ts.isBlock(callback.body)) inspectReturn(callback.body);
    else { function returns(node) { if (ts.isFunctionLike(node)) return; if (ts.isReturnStatement(node) && node.expression) inspectReturn(node.expression); ts.forEachChild(node, returns); } returns(callback.body); }
  }
}
module.exports = { inspectNativeMotion };

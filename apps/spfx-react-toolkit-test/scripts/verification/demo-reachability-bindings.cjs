// The toolkit's real field/family boundaries carry finite serialized data keys.
// These relations are scoped to resolved imported factories and their own data,
// never to a property spelling or an unrelated contextual Record.
const ts = require('typescript');
function inspectBindings({ checker, reachable, selectData, projectMapping, problem }) {
  const official = new Map();
  const target = symbol => symbol?.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(symbol) : symbol;
  for (const source of reachable) for (const statement of source.statements) {
    if (!ts.isImportDeclaration(statement) || !ts.isStringLiteral(statement.moduleSpecifier) || !/^@apvee\/spfx-react-toolkit(?:\/|$)/.test(statement.moduleSpecifier.text)) continue;
    const module = checker.getSymbolAtLocation(statement.moduleSpecifier);
    if (!module) continue;
    for (const symbol of checker.getExportsOfModule(module)) if (['createForm', 'createDataGrid', 'DataGrid', 'Form'].includes(symbol.name)) {
      const actual = target(symbol);
      if (actual?.declarations?.every(declaration => !reachable.has(declaration.getSourceFile()))) official.set(actual, symbol.name);
    }
  }
  const familyBySymbol = new Map(), methods = new Map(), builders = new Map();
  const families = [];
  const allNodes = [];
  for (const source of reachable) {
    function collect(node) { allNodes.push(node); ts.forEachChild(node, collect); }
    collect(source);
  }
  function familyOf(expression) {
    if (!expression) return undefined;
    if (ts.isPropertyAccessExpression(expression)) return familyOf(expression.expression);
    const symbol = checker.getSymbolAtLocation(expression);
    if (familyBySymbol.has(symbol)) return familyBySymbol.get(symbol);
    const imported = official.get(target(symbol));
    if (imported === 'DataGrid' || imported === 'Form') {
      let family = families.find(item => item.source === expression.getSourceFile() && item.direct && item.kind === (imported === 'Form' ? 'form' : 'grid'));
      if (!family) { family = { kind: imported === 'Form' ? 'form' : 'grid', source: expression.getSourceFile(), direct: true, fields: new Set(), data: [] }; families.push(family); }
      return family;
    }
    return undefined;
  }
  const property = (object, name) => object && ts.isObjectLiteralExpression(object) ? object.properties.find(item => item.name?.getText().replace(/^["']|["']$/g, '') === name)?.initializer : undefined;
  for (const node of allNodes) if (ts.isVariableDeclaration(node) && ts.isIdentifier(node.name) && node.initializer && ts.isCallExpression(node.initializer)) {
    const name = official.get(target(checker.getSymbolAtLocation(node.initializer.expression)));
    if (name === 'createForm' || name === 'createDataGrid') {
      const family = { kind: name === 'createForm' ? 'form' : 'grid', source: node.getSourceFile(), fields: new Set(), data: [] };
      families.push(family); familyBySymbol.set(checker.getSymbolAtLocation(node.name), family);
    }
  }
  for (const node of allNodes) if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression)) {
    const family = familyOf(node.expression.expression);
    if (!family) continue;
    if (family.kind === 'form' && node.expression.name.text === 'useForm' && ts.isVariableDeclaration(node.parent)) {
      const defaults = property(node.arguments[0], 'defaultValues');
      if (defaults) methods.set(checker.getSymbolAtLocation(node.parent.name), { family, defaults });
    }
    if (family.kind === 'grid' && node.expression.name.text === 'defineColumns') {
      const callback = node.arguments[0];
      if (callback && (ts.isArrowFunction(callback) || ts.isFunctionExpression(callback)) && callback.parameters[0] && ts.isIdentifier(callback.parameters[0].name)) builders.set(checker.getSymbolAtLocation(callback.parameters[0].name), family);
    }
  }
  function keys(expression) {
    if (!expression) return undefined;
    if (ts.isStringLiteralLike(expression)) return [expression.text];
    if (ts.isJsxExpression(expression)) return keys(expression.expression);
    if (ts.isTemplateExpression(expression)) {
      let text = expression.head.text;
      for (const span of expression.templateSpans) {
        const type = checker.getTypeAtLocation(span.expression);
        if (!(type.flags & ts.TypeFlags.NumberLike)) return undefined;
        text += '*' + span.literal.text;
      }
      return [text];
    }
    const type = checker.getTypeAtLocation(expression), arms = type.isUnion() ? type.types : [type];
    return arms.every(arm => arm.isStringLiteral()) ? arms.map(arm => arm.value) : undefined;
  }
  function addKeys(family, expression) {
    const values = keys(expression);
    if (!values) { problem(expression, 'DEMO_UNRESOLVED_BINDING', 'A genuine toolkit field binding needs a literal or finite typed data path.'); return; }
    for (const value of values) family.fields.add(value);
  }
  function props(opening) {
    const signature = checker.getResolvedSignature(opening), parameter = signature?.parameters[0];
    return parameter && checker.getTypeOfSymbolAtLocation(parameter, opening);
  }
  for (const node of allNodes) {
    if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression)) {
      const family = builders.get(checker.getSymbolAtLocation(node.expression.expression));
      if (family) {
        const field = property(node.arguments[0], 'field');
        if (field) addKeys(family, field);
      }
    }
    if (!ts.isJsxOpeningElement(node) && !ts.isJsxSelfClosingElement(node)) continue;
    const family = familyOf(node.tagName);
    if (!family) continue;
    const attributes = node.attributes.properties.filter(ts.isJsxAttribute);
    const attribute = name => attributes.find(item => item.name.text === name);
    const value = name => { const initializer = attribute(name)?.initializer; return initializer && ts.isJsxExpression(initializer) ? initializer.expression : initializer; };
    if (family.kind === 'form') {
      const name = value('name'); if (name) addKeys(family, name);
      if (ts.isPropertyAccessExpression(node.tagName) && node.tagName.name.text === 'Root') {
        const source = value('methods'), bound = source && methods.get(checker.getSymbolAtLocation(source));
        if (bound?.family === family) family.data.push(bound.defaults);
      }
    } else {
      const field = value('field'); if (field) addKeys(family, field);
      const items = value('items'); if (items) family.data.push(items);
      const rowId = value('getRowId');
      if (rowId) {
        let declaration = checker.getSymbolAtLocation(rowId)?.valueDeclaration;
        if (declaration && ts.isVariableDeclaration(declaration)) declaration = declaration.initializer;
        if (declaration && ts.isFunctionLike(declaration) && declaration.parameters[0] && declaration.body) {
          const parameter = checker.getSymbolAtLocation(declaration.parameters[0].name);
          function callbackKeys(child) {
            if (ts.isPropertyAccessExpression(child) && checker.getSymbolAtLocation(child.expression) === parameter) family.fields.add(child.name.text);
            if (ts.isElementAccessExpression(child) && checker.getSymbolAtLocation(child.expression) === parameter) addKeys(family, child.argumentExpression);
            ts.forEachChild(child, callbackKeys);
          }
          callbackKeys(declaration.body);
        }
      }
      const mapping = value('values'), contract = props(node)?.getProperty('values');
      if (mapping && contract?.declarations?.every(declaration => !reachable.has(declaration.getSourceFile()))) {
        const expected = checker.getTypeOfSymbolAtLocation(contract, node);
        projectMapping(mapping, expected);
      }
    }
  }
  for (const node of allNodes) if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression) && node.expression.name.text === 'reset') {
    const bound = methods.get(checker.getSymbolAtLocation(node.expression.expression));
    if (bound && bound.family.data.length && node.arguments[0]) bound.family.data.push(node.arguments[0]);
  }
  for (const family of families) {
    // A detached useForm or a field-shaped local lookalike has no bound data.
    if (family.kind === 'form' && !family.data.length) continue;
    for (const data of family.data) selectData(data, [...family.fields].map(field => field.split('.')));
  }
}
module.exports = { inspectBindings };

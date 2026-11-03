// Concrete member demand. Interface declarations are relations, not executable
// members. Reads retain their narrowed declaration identities (including unions).
const ts = require('typescript');
function inspectMembers({ checker, reachable, reads, writes, readDeclarations, artifactMembers, problem }) {
  const outside = declaration => !reachable.has(declaration.getSourceFile());
  const used = symbol => symbol && (reads.has(symbol) || symbol.declarations?.some(declaration => readDeclarations.has(declaration)));
  function variants(type, object) {
    if (!type?.isUnion()) return type ? [type] : [];
    const tags = object.properties.filter(ts.isPropertyAssignment).filter(property => ts.isStringLiteral(property.initializer) || property.initializer.kind === ts.SyntaxKind.TrueKeyword || property.initializer.kind === ts.SyntaxKind.FalseKeyword);
    return type.types.filter(arm => tags.every(property => {
      const member = checker.getPropertyOfType(arm, property.name.getText().replace(/^['"]|['"]$/g, ''));
      if (!member) return true;
      const memberType = checker.getTypeOfSymbolAtLocation(member, property);
      if (memberType.isStringLiteral()) return memberType.value === property.initializer.text;
      return true;
    }));
  }
  function contextualMember(object, name) {
    const contextual = checker.getContextualType(object);
    return variants(contextual, object).map(type => checker.getPropertyOfType(type, name)).filter(Boolean);
  }
  function finiteType(type, deep = false, visited = new Set()) {
    if (!type || visited.has(type) || type.symbol?.declarations?.some(ts.isSourceFile)) return;
    visited.add(type);
    const arms = type.isUnion() ? type.types : [type];
    for (const arm of arms) {
      for (const property of arm.getProperties()) {
        if (!property.declarations?.some(declaration => reachable.has(declaration.getSourceFile()))) continue;
        // Object enumeration and JSON visit own enumerable data, not the
        // non-enumerable methods/accessors installed on a class prototype.
        if (property.declarations?.some(declaration => ts.isClassLike(declaration.parent) && (ts.isMethodDeclaration(declaration) || ts.isAccessor(declaration)))) continue;
        const at = property.valueDeclaration || property.declarations?.[0];
        if (deep && at && checker.getTypeOfSymbolAtLocation(property, at).getCallSignatures().length) continue;
        reads.add(property);
        for (const declaration of property.declarations || []) readDeclarations.add(declaration);
        if (deep && at) finiteType(checker.getTypeOfSymbolAtLocation(property, at), true, visited);
      }
      if (deep) finiteType(checker.getIndexTypeOfType(arm, ts.IndexKind.Number), true, visited);
    }
  }
  const serializedParameters = new Set();
  function finiteRead(expression, deep = false) {
    finiteType(checker.getTypeAtLocation(expression), deep);
    if (deep && ts.isIdentifier(expression)) {
      const symbol = checker.getSymbolAtLocation(expression);
      if (symbol?.declarations?.some(ts.isParameter)) serializedParameters.add(symbol);
    }
  }
  const relations = new Map();
  const origins = new Map();
  const stateSetters = new Map();
  const expressionSymbol = expression => ts.isShorthandPropertyAssignment(expression.parent) ? checker.getShorthandAssignmentValueSymbol(expression.parent) : checker.getSymbolAtLocation(ts.isPropertyAccessExpression(expression) ? expression.name : expression);
  const origin = (symbol, expression) => { if (symbol && expression) { if (!origins.has(symbol)) origins.set(symbol, new Set()); origins.get(symbol).add(expression); } };
  const seen = new Set();
  const add = (from, to) => {
    if (!from || !to || from === to) return;
    if (!relations.has(from)) relations.set(from, new Set());
    relations.get(from).add(to);
  };
  function link(source, target, depth = 0, externalBoundary = false) {
    if (!source || !target || depth > 8 || source === target) return;
    const key = `${source.id}:${target.id}:${externalBoundary}`;
    if (seen.has(key)) return;
    seen.add(key);
    if (source.flags & (ts.TypeFlags.Any | ts.TypeFlags.Unknown) || target.flags & (ts.TypeFlags.Any | ts.TypeFlags.Unknown)) return;
    const sourcePromise = checker.getPromisedTypeOfPromise(source), targetPromise = checker.getPromisedTypeOfPromise(target);
    if (sourcePromise && targetPromise) link(sourcePromise, targetPromise, depth + 1, externalBoundary);
    const sources = source.isUnion() ? source.types : [source];
    const targets = target.isUnion() ? target.types : [target];
    for (const sourceArm of sources) for (const targetArm of targets) {
      // A literal discriminant excludes a sibling branch. This preserves the
      // request field difference between error and pending/success Outcomes.
      const concreteProperties = sourceArm.getProperties().filter(property => property.declarations?.some(declaration => reachable.has(declaration.getSourceFile())));
      const incompatible = concreteProperties.some(property => {
        const other = targetArm.getProperty(property.name);
        if (!other || !property.declarations?.length || !other.declarations?.length) return false;
        const a = checker.getTypeOfSymbolAtLocation(property, property.declarations[0]);
        const b = checker.getTypeOfSymbolAtLocation(other, other.declarations[0]);
        return a.isStringLiteral() && b.isStringLiteral() && a.value !== b.value;
      });
      if (incompatible) continue;
      for (const property of concreteProperties) {
        const other = targetArm.getProperty(property.name);
        if (!other) continue;
        add(other, property);
        const at = property.valueDeclaration || property.declarations?.[0];
        const bt = other.valueDeclaration || other.declarations?.[0];
        if (at) link(checker.getTypeOfSymbolAtLocation(property, at), checker.getTypeOfSymbolAtLocation(other, bt || at), depth + 1, externalBoundary);
      }
      const sourceCalls = sourceArm.getCallSignatures();
      const targetCalls = targetArm.getCallSignatures();
      for (const a of sourceCalls) for (const b of targetCalls) if (a.declaration && reachable.has(a.declaration.getSourceFile()) || b.declaration && reachable.has(b.declaration.getSourceFile())) {
        link(checker.getReturnTypeOfSignature(a), checker.getReturnTypeOfSignature(b), depth + 1, externalBoundary);
        a.parameters.forEach((parameter, index) => { const other = b.parameters[index]; add(parameter, other); if (other && parameter.valueDeclaration && other.valueDeclaration) link(checker.getTypeOfSymbolAtLocation(other, other.valueDeclaration), checker.getTypeOfSymbolAtLocation(parameter, parameter.valueDeclaration), depth + 1); });
      }
      const a = checker.getIndexTypeOfType(sourceArm, ts.IndexKind.Number);
      const b = checker.getIndexTypeOfType(targetArm, ts.IndexKind.Number);
      if (a && b) link(a, b, depth + 1, externalBoundary);
    }
  }
  for (const source of reachable) {
    function assignments(node) {
      if (ts.isObjectLiteralExpression(node) || ts.isArrowFunction(node) || ts.isFunctionExpression(node) || node.parent && ts.isJsxExpression(node.parent) && node.parent.expression === node) {
        const contextual = checker.getContextualType(node);
        let external = false;
        if (node.parent && ts.isJsxExpression(node.parent) && ts.isJsxAttribute(node.parent.parent)) {
          const opening = node.parent.parent.parent.parent;
          const declaration = checker.getResolvedSignature(opening)?.declaration;
          external = Boolean(declaration && outside(declaration));
        }
        if (contextual) link(checker.getTypeAtLocation(node), contextual, 0, external);
      }
      if (ts.isObjectLiteralExpression(node) && ts.isArrayLiteralExpression(node.parent)) link(checker.getTypeAtLocation(node), checker.getIndexTypeOfType(checker.getTypeAtLocation(node.parent), ts.IndexKind.Number));
      if (ts.isVariableDeclaration(node) && ts.isArrayBindingPattern(node.name) && node.initializer && ts.isCallExpression(node.initializer)) {
        const declaration = checker.getResolvedSignature(node.initializer)?.declaration;
        const first = node.name.elements[0];
        if (declaration?.name?.getText() === 'useState' && declaration.getSourceFile().fileName.includes('/node_modules/@types/react/') && first && ts.isBindingElement(first) && ts.isIdentifier(first.name)) {
          const state = checker.getSymbolAtLocation(first.name); origin(state, node.initializer.arguments[0]);
          const setter = node.name.elements[1]; if (setter && ts.isBindingElement(setter) && ts.isIdentifier(setter.name)) stateSetters.set(checker.getSymbolAtLocation(setter.name), state);
        }
      }
      if (ts.isVariableDeclaration(node) && node.initializer) { link(checker.getTypeAtLocation(node.initializer), checker.getTypeAtLocation(node.name)); if (ts.isIdentifier(node.name)) origin(checker.getSymbolAtLocation(node.name), node.initializer); }
      if (ts.isVariableDeclaration(node) && ts.isObjectBindingPattern(node.name) && node.initializer) for (const element of node.name.elements) if (!element.dotDotDotToken && ts.isIdentifier(element.name)) origin(checker.getSymbolAtLocation(element.name), element);
      if (ts.isPropertyAssignment(node)) origin(checker.getSymbolAtLocation(node.name), node.initializer);
      if (ts.isPropertyDeclaration(node) && node.initializer) origin(checker.getSymbolAtLocation(node.name), node.initializer);
      if (ts.isShorthandPropertyAssignment(node)) origin(checker.getSymbolAtLocation(node.name), node.name);
      if (ts.isMethodDeclaration(node) && node.name?.getText() === 'getDerivedStateFromError' && ts.isClassDeclaration(node.parent) && ts.getModifiers(node)?.some(modifier => modifier.kind === ts.SyntaxKind.StaticKeyword)) {
        const instance = checker.getTypeAtLocation(node.parent), state = instance.getProperty('state');
        const genuine = (checker.getBaseTypes(instance) || []).some(base => base.symbol?.name === 'Component' && base.symbol.declarations?.some(declaration => declaration.getSourceFile().fileName.includes('/node_modules/@types/react/')));
        const signature = checker.getSignatureFromDeclaration(node);
        if (genuine && state && signature) link(checker.getReturnTypeOfSignature(signature), checker.getTypeOfSymbolAtLocation(state, node));
      }
      if (ts.isReturnStatement(node) && node.expression) {
        let owner = node.parent;
        while (owner && !ts.isFunctionLike(owner)) owner = owner.parent;
        const signature = owner && checker.getSignatureFromDeclaration(owner);
        let external = false;
        if (owner && ts.isMethodDeclaration(owner) && ts.isClassDeclaration(owner.parent)) {
          const instance = checker.getTypeAtLocation(owner.parent);
          external = (checker.getBaseTypes(instance) || []).some(base => base.getProperty(owner.name.getText())?.declarations?.some(outside));
        }
        if (signature) link(checker.getTypeAtLocation(node.expression), checker.getReturnTypeOfSignature(signature), 0, external);
      }
      if (ts.isBinaryExpression(node) && node.operatorToken.kind === ts.SyntaxKind.EqualsToken) {
        link(checker.getTypeAtLocation(node.right), checker.getTypeAtLocation(node.left));
        if (ts.isIdentifier(node.left) || ts.isPropertyAccessExpression(node.left) && checker.getSymbolAtLocation(node.left.name)?.declarations?.some(declaration => reachable.has(declaration.getSourceFile()))) origin(checker.getSymbolAtLocation(ts.isPropertyAccessExpression(node.left) ? node.left.name : node.left), node.right);
      }
      if (ts.isJsxSpreadAttribute(node)) {
        const opening = node.parent.parent;
        const signature = checker.getResolvedSignature(opening);
        const parameter = signature?.parameters[0];
        const declaration = parameter?.valueDeclaration || parameter?.declarations?.[0];
        if (declaration) link(checker.getTypeAtLocation(node.expression), checker.getTypeOfSymbolAtLocation(parameter, declaration), 0, outside(declaration));
        if (opening.tagName && ts.isIdentifier(opening.tagName) && /^[a-z]/.test(opening.tagName.text)) for (const property of checker.getTypeAtLocation(node.expression).getProperties()) if (/^(data|aria)-/.test(property.name)) reads.add(property);
      }
      if (ts.isJsxOpeningElement(node) || ts.isJsxSelfClosingElement(node)) {
        const signature = checker.getResolvedSignature(node), parameter = signature?.parameters[0];
        const declaration = signature?.declaration;
        if (declaration && reachable.has(declaration.getSourceFile())) origin(parameter, node.attributes);
      }
      if (ts.isCallExpression(node)) {
        const signature = checker.getResolvedSignature(node);
        node.arguments.forEach((argument, index) => {
          const parameter = signature?.parameters[index];
          const declaration = parameter?.valueDeclaration || parameter?.declarations?.[0];
          if (declaration) { link(checker.getTypeAtLocation(argument), checker.getTypeOfSymbolAtLocation(parameter, declaration), 0, outside(declaration)); if (!outside(declaration)) origin(parameter, argument); }
        });
        // Griffel's mapped slot keys have no declaration symbols. The genuine
        // installed makeStyles signature maps each input slot to the matching
        // class key, while each descriptor is a GriffelStyle external contract.
        const declaration = signature?.declaration;
        if (declaration?.name?.getText() === 'makeStyles' && declaration.getSourceFile().fileName.includes('/node_modules/@griffel/')) {
          const argument = node.arguments[0];
          const output = checker.getTypeAtLocation(node).getCallSignatures()[0];
          const classes = output && checker.getReturnTypeOfSignature(output);
          if (argument && classes) for (const property of checker.getTypeAtLocation(argument).getProperties()) add(classes.getProperty(property.name), property);
        }
      }
      ts.forEachChild(node, assignments);
    }
    assignments(source);
  }
  function propagate() {
    const queue = [...reads];
    const declared = new Map();
    for (const symbol of relations.keys()) for (const declaration of symbol.declarations || []) { if (!declared.has(declaration)) declared.set(declaration, new Set()); declared.get(declaration).add(symbol); }
    for (const [declaration, symbols] of declared) if (readDeclarations.has(declaration)) queue.push(...symbols);
    const done = new Set();
    while (queue.length) {
      const symbol = queue.pop();
      if (done.has(symbol)) continue;
      done.add(symbol);
      reads.add(symbol);
      for (const declaration of symbol.declarations || []) { readDeclarations.add(declaration); for (const related of declared.get(declaration) || []) if (!done.has(related)) queue.push(related); }
      for (const root of checker.getRootSymbols(symbol)) if (!done.has(root)) queue.push(root);
      for (const next of relations.get(symbol) || []) if (!done.has(next)) queue.push(next);
    }
  }
  for (const source of reachable) {
    function updates(node) {
      if (ts.isCallExpression(node)) {
        const state = stateSetters.get(checker.getSymbolAtLocation(node.expression));
        const value = node.arguments[0];
        if (state && value) {
          if (ts.isFunctionLike(value) && value.body) {
            if (!ts.isBlock(value.body)) origin(state, value.body);
            else { function returns(child) { if (ts.isFunctionLike(child)) return; if (ts.isReturnStatement(child) && child.expression) origin(state, child.expression); ts.forEachChild(child, returns); } returns(value.body); }
          } else origin(state, value);
        }
      }
      ts.forEachChild(node, updates);
    }
    updates(source);
  }
  const identities = new Map();
  for (const symbol of new Set([...relations.keys(), ...origins.keys()])) for (const declaration of symbol.declarations || []) { if (!identities.has(declaration)) identities.set(declaration, new Set()); identities.get(declaration).add(symbol); }
  function consumeObject(expression, deep = false, visited = new Set()) {
    if (!expression || visited.has(expression)) return;
    visited.add(expression);
    const type = checker.getTypeAtLocation(expression);
    if (type.symbol?.declarations?.some(ts.isSourceFile)) return;
    if (ts.isObjectLiteralExpression(expression)) {
      for (const property of expression.properties) {
        if (ts.isSpreadAssignment(property)) { consumeObject(property.expression, deep, visited); continue; }
        if (!property.name) continue;
        if (deep && (ts.isMethodDeclaration(property) || ts.isPropertyAssignment(property) && checker.getTypeAtLocation(property.initializer).getCallSignatures().length)) continue;
        const symbol = checker.getSymbolAtLocation(property.name);
        if (!symbol) continue;
        reads.add(symbol);
        for (const declaration of symbol?.declarations || []) readDeclarations.add(declaration);
        if (deep && ts.isPropertyAssignment(property)) consumeObject(property.initializer, true, visited);
        if (deep && ts.isShorthandPropertyAssignment(property)) consumeObject(property.name, true, visited);
      }
      return;
    }
    if (ts.isArrayLiteralExpression(expression)) {
      if (deep)
        for (const element of expression.elements) consumeObject(element, true, visited);
      return;
    }
    if (ts.isAsExpression(expression) || ts.isTypeAssertionExpression(expression) || ts.isParenthesizedExpression(expression) || ts.isNonNullExpression(expression)) {
      consumeObject(expression.expression, deep, visited);
      return;
    }
    if (ts.isConditionalExpression(expression)) {
      consumeObject(expression.whenTrue, deep, visited);
      consumeObject(expression.whenFalse, deep, visited);
      return;
    }
    if (ts.isCallExpression(expression)) {
      const declaration = checker.getResolvedSignature(expression)?.declaration;
      if (declaration && reachable.has(declaration.getSourceFile()) && declaration.body) {
        if (!ts.isBlock(declaration.body)) consumeObject(declaration.body, deep, visited);
        else {
          function returns(node) {
            if (ts.isFunctionLike(node)) return;
            if (ts.isReturnStatement(node) && node.expression) consumeObject(node.expression, deep, visited);
            ts.forEachChild(node, returns);
          }
          returns(declaration.body);
        }
      }
    }
    const symbol = ts.isShorthandPropertyAssignment(expression.parent) ? checker.getShorthandAssignmentValueSymbol(expression.parent) : checker.getSymbolAtLocation(ts.isPropertyAccessExpression(expression) ? expression.name : expression);
    const queue = symbol ? [symbol] : [], done = new Set();
    while (queue.length) {
      const current = queue.pop();
      if (!current || done.has(current)) continue;
      done.add(current);
      for (const value of origins.get(current) || []) consumeObject(value, deep, visited);
      for (const next of relations.get(current) || []) queue.push(next);
      for (const declaration of current.declarations || [])
        for (const equivalent of identities.get(declaration) || [])
          queue.push(equivalent);
      for (const root of checker.getRootSymbols(current))
        if (root !== current) queue.push(root);
    }
    finiteRead(expression, deep);
  }
  function objectOrigins(expression, visited = new Set()) {
    if (!expression || visited.has(expression)) return [];
    visited.add(expression);
    if (ts.isBindingElement(expression)) return selectedBindingValues(expression).flatMap(value => objectOrigins(value, visited));
    if (ts.isObjectLiteralExpression(expression) || ts.isNewExpression(expression) || ts.isJsxAttributes(expression)) return [expression];
    if (ts.isArrayLiteralExpression(expression)) return expression.elements.flatMap(element => objectOrigins(ts.isSpreadElement(element) ? element.expression : element, visited));
    if (ts.isAsExpression(expression) || ts.isTypeAssertionExpression(expression) || ts.isParenthesizedExpression(expression) || ts.isNonNullExpression(expression)) return objectOrigins(expression.expression, visited);
    if (ts.isConditionalExpression(expression)) return [...objectOrigins(expression.whenTrue, visited), ...objectOrigins(expression.whenFalse, visited)];
    if (ts.isBinaryExpression(expression) && [ts.SyntaxKind.QuestionQuestionToken, ts.SyntaxKind.BarBarToken, ts.SyntaxKind.AmpersandAmpersandToken].includes(expression.operatorToken.kind)) return [...objectOrigins(expression.left, visited), ...objectOrigins(expression.right, visited)];
    if (ts.isArrowFunction(expression) || ts.isFunctionExpression(expression)) return functionResults(expression, visited);
    if (ts.isCallExpression(expression)) {
      const declaration = checker.getResolvedSignature(expression)?.declaration;
      if (declaration?.getSourceFile().fileName.includes('/typescript/lib/') && ts.isPropertyAccessExpression(expression.expression) && expression.expression.expression.getText() === 'Object' && ['freeze', 'seal'].includes(expression.expression.name.text)) return objectOrigins(expression.arguments[0], visited);
      if (declaration && reachable.has(declaration.getSourceFile()) && declaration.body) return functionResults(declaration, visited);
      const implementations = callableOrigins(expression.expression);
      if (implementations.length) return implementations.flatMap(implementation => functionResults(implementation, visited));
      if (declaration?.getSourceFile().fileName.includes('/typescript/lib/') && ts.isPropertyAccessExpression(expression.expression) && ['filter', 'slice', 'sort', 'reverse'].includes(expression.expression.name.text)) return objectOrigins(expression.expression.expression, visited);
    }
    const symbol = expressionSymbol(expression);
    const queue = symbol ? [symbol] : [], done = new Set(), found = [];
    while (queue.length) {
      const current = queue.pop();
      if (!current || done.has(current)) continue;
      done.add(current);
      for (const value of origins.get(current) || []) found.push(...objectOrigins(value, visited));
      for (const next of relations.get(current) || []) queue.push(next);
      for (const declaration of current.declarations || [])
        for (const equivalent of identities.get(declaration) || [])
          queue.push(equivalent);
      for (const root of checker.getRootSymbols(current))
        if (root !== current) queue.push(root);
    }
    return found;
  }
  function selectedBindingValues(element) {
    const declaration = element.parent.parent;
    if (!ts.isVariableDeclaration(declaration) || !declaration.initializer) return [];
    const name = (element.propertyName || element.name).getText().replace(/^["']|["']$/g, '');
    return objectOrigins(declaration.initializer).flatMap(object => ts.isObjectLiteralExpression(object) ? object.properties.filter(property => property.name?.getText().replace(/^["']|["']$/g, '') === name).map(property => ts.isPropertyAssignment(property) ? property.initializer : ts.isShorthandPropertyAssignment(property) ? property.name : undefined).filter(Boolean) : []);
  }
  function callableOrigins(expression, visited = new Set()) {
    if (!expression || visited.has(expression)) return [];
    visited.add(expression);
    if (ts.isFunctionLike(expression) && expression.body) return [expression];
    if (ts.isCallExpression(expression)) {
      const declaration = checker.getResolvedSignature(expression)?.declaration;
      if (declaration?.name?.getText() === 'useCallback' && declaration.getSourceFile().fileName.includes('/node_modules/@types/react/')) return callableOrigins(expression.arguments[0], visited);
    }
    const symbol = expressionSymbol(expression);
    const queue = symbol ? [symbol] : [], done = new Set(), found = [];
    while (queue.length) {
      const current = queue.pop();
      if (!current || done.has(current)) continue;
      done.add(current);
      for (const value of origins.get(current) || []) found.push(...callableOrigins(value, visited));
      for (const declaration of current.declarations || []) {
        if (ts.isFunctionLike(declaration) && declaration.body) found.push(declaration);
        for (const equivalent of identities.get(declaration) || []) queue.push(equivalent);
      }
      for (const next of relations.get(current) || []) queue.push(next);
      for (const root of checker.getRootSymbols(current))
        if (root !== current) queue.push(root);
    }
    return found;
  }
  function functionResults(declaration, visited) {
    if (!ts.isBlock(declaration.body)) return objectOrigins(declaration.body, visited);
    const found = [];
    function scan(node) { if (ts.isFunctionLike(node)) return; if (ts.isReturnStatement(node) && node.expression) found.push(...objectOrigins(node.expression, visited)); ts.forEachChild(node, scan); }
    scan(declaration.body); return found;
  }
  const markConcrete = member => { const symbol = checker.getSymbolAtLocation(member.name); if (symbol) { reads.add(symbol); for (const declaration of symbol.declarations || []) readDeclarations.add(declaration); } };
  require('./demo-reachability-motion.cjs').inspectNativeMotion({ checker, reachable, markConcrete });
  function selectData(expression, paths) {
    for (const object of objectOrigins(expression)) {
      if (!ts.isObjectLiteralExpression(object)) continue;
      for (const member of object.properties) {
        if (ts.isSpreadAssignment(member)) { selectData(member.expression, paths); continue; }
        if (!member.name || !ts.isPropertyAssignment(member) && !ts.isShorthandPropertyAssignment(member)) continue;
        const name = member.name.getText().replace(/^["']|["']$/g, '');
        const selected = paths.map(parts => parts[0] === '*' ? parts.slice(1) : parts).filter(parts => parts[0] === name);
        if (!selected.length) continue;
        markConcrete(member);
        const value = ts.isPropertyAssignment(member) ? member.initializer : member.name;
        const rest = selected.map(parts => parts.slice(1));
        if (rest.some(parts => !parts.length)) {
          // Selected field values have an actual SDK transport/classification
          // role. Preserve their plain payload; callable business is still read-audited.
          for (const payload of objectOrigins(value)) if (ts.isObjectLiteralExpression(payload)) for (const property of payload.properties) if (ts.isPropertyAssignment(property) && !checker.getTypeAtLocation(property.initializer).getCallSignatures().length) markConcrete(property);
        } else selectData(value, rest);
      }
    }
  }
  function projectMapping(expression, expected) {
    const arms = expected.isUnion() ? expected.types : [expected];
    const records = arms.filter(type => type.objectFlags & ts.ObjectFlags.Mapped && !type.getStringIndexType());
    for (const object of objectOrigins(expression)) {
      if (ts.isObjectLiteralExpression(object)) for (const member of object.properties) {
        if (!member.name || !ts.isPropertyAssignment(member)) continue;
        const name = member.name.getText().replace(/^["']|["']$/g, '');
        const contract = records.map(type => type.getProperty(name)).find(Boolean);
        if (!contract) continue;
        markConcrete(member);
        consumeExternal(member.initializer, checker.getTypeOfSymbolAtLocation(contract, member));
      }
      if (ts.isNewExpression(object) && object.expression.getText() === 'Map') {
        const map = arms.find(type => type.symbol?.name === 'ReadonlyMap');
        const args = map && checker.getTypeArguments(map);
        const keyType = args?.[0], presentation = args?.[1];
        const domain = keyType && (keyType.isUnion() ? keyType.types : [keyType]);
        const allowed = domain?.every(type => type.isStringLiteral() || type.isNumberLiteral()) && new Set(domain.map(type => String(type.value)));
        const entries = object.arguments?.[0];
        if (allowed && presentation && entries && ts.isArrayLiteralExpression(entries)) for (const pair of entries.elements) if (ts.isArrayLiteralExpression(pair) && (ts.isStringLiteral(pair.elements[0]) || ts.isNumericLiteral(pair.elements[0]))) {
          if (allowed.has(pair.elements[0].text)) consumeExternal(pair.elements[1], presentation);
          else problem(pair.elements[0], 'DEMO_UNUSED_MAPPING_KEY', 'Mapping key is outside the resolved finite data domain.');
        }
      }
    }
  }
  function consumeExternal(expression, expected, visited = new Set(), omitted = new Set()) {
    if (!expression || !expected || expected.flags & (ts.TypeFlags.Any | ts.TypeFlags.Unknown | ts.TypeFlags.Never)) return;
    const key = `${expression.getSourceFile().fileName}:${expression.pos}:${expression.end}:${expression.kind}:${expected.id}:${[...omitted].join(',')}`;
    if (visited.has(key)) return;
    visited.add(key);
    if (ts.isBindingElement(expression)) {
      for (const value of selectedBindingValues(expression)) consumeExternal(value, expected, visited, omitted);
      return;
    }
    // Optional SDK fields are unions with undefined. Querying the union itself
    // loses array indices and callable signatures; project its concrete arms.
    if (expected.isUnion()) {
      for (const arm of expected.types) consumeExternal(expression, arm, visited, omitted);
      return;
    }
    if (ts.isParenthesizedExpression(expression) || ts.isAsExpression(expression) || ts.isTypeAssertionExpression(expression) || ts.isNonNullExpression(expression)) {
      consumeExternal(expression.expression, expected, visited, omitted);
      return;
    }
    if (ts.isConditionalExpression(expression)) {
      consumeExternal(expression.whenTrue, expected, visited, omitted);
      consumeExternal(expression.whenFalse, expected, visited, omitted);
      return;
    }
    if (ts.isBinaryExpression(expression) && [ts.SyntaxKind.QuestionQuestionToken, ts.SyntaxKind.BarBarToken, ts.SyntaxKind.AmpersandAmpersandToken].includes(expression.operatorToken.kind)) {
      consumeExternal(expression.left, expected, visited, omitted);
      consumeExternal(expression.right, expected, visited, omitted);
      return;
    }
    const promised = checker.getPromisedTypeOfPromise(expected);
    if (ts.isCallExpression(expression) && promised) {
      const declaration = checker.getResolvedSignature(expression)?.declaration;
      if (declaration?.getSourceFile().fileName.includes('/typescript/lib/') && ts.isPropertyAccessExpression(expression.expression)) {
        if (expression.expression.name.text === 'then') {
          const callback = expression.arguments[0];
          if (callback && ts.isFunctionLike(callback)) externalReturns(callback, promised, visited);
        }
        if (expression.expression.name.text === 'resolve') consumeExternal(expression.arguments[0], promised, visited);
      }
    }
    if (ts.isFunctionLike(expression)) {
      for (const signature of expected.getCallSignatures()) externalReturns(expression, checker.getReturnTypeOfSignature(signature), visited);
      return;
    }
    for (const signature of expected.getCallSignatures())
      for (const implementation of callableOrigins(expression))
        externalReturns(implementation, checker.getReturnTypeOfSignature(signature), visited);
    if (ts.isArrayLiteralExpression(expression)) {
      const element = checker.getIndexTypeOfType(expected, ts.IndexKind.Number);
      if (element) for (const value of expression.elements) ts.isSpreadElement(value) ? consumeExternal(value.expression, expected, visited) : consumeExternal(value, element, visited);
      return;
    }
    if (ts.isJsxAttributes(expression)) {
      for (const attribute of expression.properties) {
        if (ts.isJsxSpreadAttribute(attribute)) {
          consumeExternal(attribute.expression, expected, visited, omitted);
          continue;
        }
        if (!attribute.initializer || !ts.isJsxExpression(attribute.initializer) || !attribute.initializer.expression || omitted.has(attribute.name.text)) continue;
        const contract = expected.getProperty(attribute.name.text);
        if (contract) consumeExternal(attribute.initializer.expression, checker.getTypeOfSymbolAtLocation(contract, attribute), visited);
      }
      return;
    }
    if (ts.isCallExpression(expression) && ts.isPropertyAccessExpression(expression.expression)) {
      const declaration = checker.getResolvedSignature(expression)?.declaration;
      if (declaration?.getSourceFile().fileName.includes('/typescript/lib/')) {
        if (expression.expression.expression.getText() === 'Object' && ['freeze', 'seal'].includes(expression.expression.name.text)) consumeExternal(expression.arguments[0], expected, visited, omitted);
        if (['filter', 'slice', 'sort', 'reverse'].includes(expression.expression.name.text)) consumeExternal(expression.expression.expression, expected, visited, omitted);
        if (expression.expression.name.text === 'map') {
          const element = checker.getIndexTypeOfType(expected, ts.IndexKind.Number), callback = expression.arguments[0];
          if (element && callback && ts.isFunctionLike(callback)) externalReturns(callback, element, visited);
        }
      }
      if (declaration?.name?.getText() === 'useMemo' && declaration.getSourceFile().fileName.includes('/node_modules/@types/react/')) {
        const callback = expression.arguments[0];
        if (callback && ts.isFunctionLike(callback)) externalReturns(callback, expected, visited);
      }
    }
    if (ts.isObjectLiteralExpression(expression)) {
      const arms = variants(expected, expression);

      for (const member of expression.properties) {
        if (ts.isSpreadAssignment(member)) {
          consumeExternal(member.expression, expected, visited, omitted);
          continue;
        }
        if (!member.name) continue;
        const name = member.name.getText().replace(/^["']|["']$/g, '');
        if (omitted.has(name)) continue;
        for (const arm of arms) {
          const contract = arm.getProperty(name);
          const native = (/^(data|aria)-/.test(name) || name === 'inert') && arm.getProperty('className')?.declarations?.some(declaration => declaration.getSourceFile().fileName.includes('/node_modules/@types/react/'));
          const index = checker.getIndexTypeOfType(arm, ts.IndexKind.String);
          const indexed = index && !(index.flags & (ts.TypeFlags.Any | ts.TypeFlags.Unknown | ts.TypeFlags.TypeParameter)) && (checker.getIndexInfosOfType(arm).some(info => info.declaration && outside(info.declaration)) || arm.aliasSymbol?.declarations?.some(outside));
          const mapped = contract && !contract.declarations?.length && arm.objectFlags & ts.ObjectFlags.Mapped;
          if (native || indexed || contract && (contract.declarations?.some(outside) || mapped)) markConcrete(member);
          const supplied = ts.isPropertyAssignment(member) ? member.initializer : ts.isShorthandPropertyAssignment(member) ? member.name : undefined;
          if (indexed && !contract && supplied) consumeExternal(supplied, index, visited);
          if (contract) {
            const value = ts.isPropertyAssignment(member) ? member.initializer : ts.isShorthandPropertyAssignment(member) ? member.name : ts.isMethodDeclaration(member) ? member : undefined;
            if (value) consumeExternal(value, checker.getTypeOfSymbolAtLocation(contract, member), visited);
          }
        }
      }
      return;
    }
    if (ts.isCallExpression(expression)) {
      const declaration = checker.getResolvedSignature(expression)?.declaration;
      if (declaration && reachable.has(declaration.getSourceFile()) && declaration.body) externalReturns(declaration, expected, visited);
      for (const implementation of callableOrigins(expression.expression)) externalReturns(implementation, expected, visited);
    }
    if (ts.isPropertyAccessExpression(expression)) for (const object of objectOrigins(expression.expression)) {
      if (ts.isObjectLiteralExpression(object)) {
        const member = object.properties.find(property => property.name?.getText().replace(/^["']|["']$/g, '') === expression.name.text);
        if (member && ts.isPropertyAssignment(member)) consumeExternal(member.initializer, expected, visited, omitted);
      }
      if (ts.isJsxAttributes(object)) for (const attribute of object.properties) {
        if (ts.isJsxAttribute(attribute) && attribute.name.text === expression.name.text && attribute.initializer && ts.isJsxExpression(attribute.initializer)) consumeExternal(attribute.initializer.expression, expected, visited, omitted);
      }
    }
    const symbol = expressionSymbol(expression);
    const queue = symbol ? [symbol] : [], done = new Set();
    while (queue.length) {
      const current = queue.pop();
      if (!current || done.has(current)) continue;
      done.add(current);
      for (const value of origins.get(current) || []) consumeExternal(value, expected, visited, omitted);
      for (const next of relations.get(current) || []) queue.push(next);
      for (const declaration of current.declarations || [])
        for (const equivalent of identities.get(declaration) || [])
          queue.push(equivalent);
      for (const root of checker.getRootSymbols(current))
        if (root !== current) queue.push(root);
    }
  }
  function externalReturns(declaration, expected, visited) {
    if (!declaration.body) return;
    if (!ts.isBlock(declaration.body)) { consumeExternal(declaration.body, expected, visited); return; }
    function scan(node) { if (ts.isFunctionLike(node)) return; if (ts.isReturnStatement(node) && node.expression) consumeExternal(node.expression, expected, visited); ts.forEachChild(node, scan); }
    scan(declaration.body);
  }
  for (const source of reachable) {
    function boundaries(node) {
      if (ts.isCallExpression(node) || ts.isNewExpression(node)) {
        const signature = checker.getResolvedSignature(node), declaration = signature?.declaration;
        if (ts.isCallExpression(node) && declaration?.name?.getText() === 'createElement' && declaration.getSourceFile().fileName.includes('/node_modules/@types/react/') && node.arguments[0] && node.arguments[1]) {
          for (const component of checker.getTypeAtLocation(node.arguments[0]).getCallSignatures()) {
            const props = component.parameters[0]; if (props) consumeExternal(node.arguments[1], checker.getTypeOfSymbolAtLocation(props, node.arguments[0]));
          }
        }
        if (declaration && outside(declaration) && !(declaration.name?.getText() === 'makeStyles' && declaration.getSourceFile().fileName.includes('/node_modules/@griffel/'))) (node.arguments || []).forEach((argument, index) => {
          const parameter = signature.parameters[Math.min(index, signature.parameters.length - 1)], at = parameter?.valueDeclaration || parameter?.declarations?.[0];
          // Use the declared boundary, not a generic substitution inferred
          // from this very argument. Otherwise useMemo<T>(() => object) or an
          // identity<T>(object) would turn an annotation into a fake SDK read.
          const declared = at?.name && checker.getSymbolAtLocation(at.name);
          if (at) {
            const contract = checker.getTypeOfSymbolAtLocation(declared || parameter, at);
            consumeExternal(argument, at.dotDotDotToken && !ts.isSpreadElement(argument) ? checker.getIndexTypeOfType(contract, ts.IndexKind.Number) : contract);
          }
        });
      }
      if (ts.isJsxOpeningElement(node) || ts.isJsxSelfClosingElement(node)) {
        const signature = checker.getResolvedSignature(node), declaration = signature?.declaration, parameter = signature?.parameters[0];
        if (declaration && outside(declaration) && parameter) {
          const expected = checker.getTypeOfSymbolAtLocation(parameter, node);
          const attributes = node.attributes.properties;
          attributes.forEach((attribute, index) => {
            if (ts.isJsxSpreadAttribute(attribute)) {
              const omitted = new Set(attributes.slice(index + 1).filter(ts.isJsxAttribute).map(item => item.name.text));
              if (ts.isJsxOpeningElement(node) && node.parent.children.some(child => !ts.isJsxText(child) || child.text.trim())) omitted.add('children');
              consumeExternal(attribute.expression, expected, new Set(), omitted);
            } else if (attribute.initializer && ts.isJsxExpression(attribute.initializer) && attribute.initializer.expression) {
              const member = expected.getProperty(attribute.name.text);
              if (member) consumeExternal(attribute.initializer.expression, checker.getTypeOfSymbolAtLocation(member, node));
            }
          });
        }
      }
      if (ts.isMethodDeclaration(node) && ts.isClassDeclaration(node.parent) && node.body) {
        const type = checker.getTypeAtLocation(node.parent), inherited = (checker.getBaseTypes(type) || []).map(base => base.getProperty(node.name.getText())).find(symbol => symbol?.declarations?.some(outside));
        if (inherited) { const signature = checker.getSignatureFromDeclaration(node); if (signature) externalReturns(node, checker.getReturnTypeOfSignature(signature), new Set()); }
      }
      ts.forEachChild(node, boundaries);
    }
    boundaries(source);
  }
  require('./demo-reachability-bindings.cjs').inspectBindings({ checker, reachable, selectData, projectMapping, problem });
  // Mapping projection adds real external presentation relations. Include their
  // fresh symbols in the declaration identity index before final propagation.
  for (const symbol of new Set([...relations.keys(), ...origins.keys()])) for (const declaration of symbol.declarations || []) { if (!identities.has(declaration)) identities.set(declaration, new Set()); identities.get(declaration).add(symbol); }
  for (const source of reachable) {
    function scans(node) {
      if (ts.isSpreadAssignment(node) || ts.isSpreadElement(node)) finiteRead(node.expression);
      if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression)) {
        const owner = checker.getSymbolAtLocation(node.expression.expression);
        if (owner?.declarations?.every(outside) && (['Object', 'JSON'].includes(node.expression.expression.getText()) && ['keys', 'entries', 'values', 'stringify', 'assign'].includes(node.expression.name.text) || node.expression.expression.getText() === 'console' && node.expression.name.text === 'table')) for (const argument of node.arguments) consumeObject(argument, ['stringify', 'table'].includes(node.expression.name.text));
      }
      ts.forEachChild(node, scans);
    }
    scans(source);
  }
  // A local unknown-typed formatter can serialize an actual finite argument.
  // Follow only its resolved positional parameter, never a spelling/global name.
  for (const source of reachable) {
    function serializedCalls(node) {
      if (ts.isCallExpression(node)) {
        const signature = checker.getResolvedSignature(node);
        node.arguments.forEach((argument, index) => { if (serializedParameters.has(signature?.parameters[index])) finiteRead(argument, true); });
      }
      ts.forEachChild(node, serializedCalls);
    }
    serializedCalls(source);
  }
  propagate();
  for (const source of reachable) {
    function styles(node) {
      if (ts.isCallExpression(node)) {
        const signature = checker.getResolvedSignature(node), declaration = signature?.declaration;
        if (declaration?.name?.getText() === 'makeStyles' && declaration.getSourceFile().fileName.includes('/node_modules/@griffel/')) {
          const object = node.arguments[0], parameter = signature.parameters[0];
          if (object && ts.isObjectLiteralExpression(object) && parameter) {
            const contract = checker.getTypeOfSymbolAtLocation(parameter, node);
            for (const member of object.properties) if (member.name && ts.isPropertyAssignment(member) && used(checker.getSymbolAtLocation(member.name))) {
              const expected = contract.getProperty(member.name.getText());
              if (expected) consumeExternal(member.initializer, checker.getTypeOfSymbolAtLocation(expected, node));
            }
          }
        }
      }
      ts.forEachChild(node, styles);
    }
    styles(source);
  }
  propagate();
  for (const source of reachable) {
    function inspect(node) {
      const concrete = ts.isPropertyAssignment(node) || ts.isShorthandPropertyAssignment(node) || ts.isMethodDeclaration(node) || ts.isPropertyDeclaration(node) || ts.isGetAccessorDeclaration(node) || ts.isSetAccessorDeclaration(node) || ts.isParameter(node) && ts.isConstructorDeclaration(node.parent) && ts.isParameterPropertyDeclaration(node, node.parent);
      if (concrete && node.name && !artifactMembers.has(node)) {
        const name = ts.isIdentifier(node.name) || ts.isStringLiteral(node.name) || ts.isNumericLiteral(node.name) ? node.name.text : undefined;
        const symbol = checker.getSymbolAtLocation(node.name);
        let live = ts.isSetAccessorDeclaration(node) ? [...writes].some(write => write === symbol || write.declarations?.includes(node)) : used(symbol);
        if (!live && ts.isObjectLiteralExpression(node.parent) && name) {
          const contextual = checker.getContextualType(node.parent);
          live = contextualMember(node.parent, name).some(property => used(property));
          if (!live && contextual?.aliasSymbol?.name === 'GriffelStyle' && contextual.aliasSymbol.declarations?.some(declaration => declaration.getSourceFile().fileName.includes('/node_modules/@griffel/'))) live = true;
        }
        if (!live && (ts.isClassDeclaration(node.parent) || ts.isClassExpression(node.parent)) && name) {
          const type = checker.getTypeAtLocation(node.parent);
          for (const base of checker.getBaseTypes(type) || []) {
            const inherited = checker.getPropertyOfType(base, name);
            if (inherited?.declarations?.some(outside)) live = true;
            if (!live && base.symbol?.name === 'Component' && base.symbol.declarations?.some(declaration => declaration.getSourceFile().fileName.includes('/node_modules/@types/react/'))) {
              const module = base.symbol.parent;
              const componentClass = module && checker.getExportsOfModule(module).find(symbol => symbol.name === 'ComponentClass');
              const contract = componentClass && checker.getDeclaredTypeOfSymbol(componentClass).getProperty(name);
              if (contract?.declarations?.some(outside)) live = true;
            }
          }
        }
        if (!live && name) problem(node.name, 'DEMO_UNUSED_MEMBER', `Concrete ${ts.isSetAccessorDeclaration(node) ? 'setter' : 'member'} ${name} has no ${ts.isSetAccessorDeclaration(node) ? 'write' : 'read'} or resolved external contract.`);
      }
      ts.forEachChild(node, inspect);
    }
    inspect(source);
  }
}
module.exports = { inspectMembers };

// Physical source proof for the mandatory exact-copy metadata export slot.
// This does not build a canonical model or infer public API routes from outputs.
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { createHash } = require('node:crypto');
const { parseContract } = require('../../../../scripts/docs/documentation-contract.cjs');
const { readMetadata } = require('../../../../scripts/docs/executable-examples-model.cjs');
const { ownership } = require('../../../../scripts/docs/executable-examples-render.cjs');
function proveMetadataReplicas({ root, reachable, resolve, checker, demand, problem, artifactMembers }) {
  function literal(node) {
    if (!node) return undefined;
    if (ts.isStringLiteral(node) || ts.isNumericLiteral(node)) return ts.isNumericLiteral(node) ? Number(node.text) : node.text;
    if (ts.isArrayLiteralExpression(node)) return node.elements.map(literal);
    if (ts.isObjectLiteralExpression(node)) {
      const object = {};
      for (const property of node.properties) if (ts.isPropertyAssignment(property) && (ts.isIdentifier(property.name) || ts.isStringLiteral(property.name))) object[property.name.text] = literal(property.initializer);
      return object;
    }
    return undefined;
  }
  const proven = new Map();
  const topEntries = new Map();
  const acceptedCatalogues = new Map();
  for (const source of reachable) {
    function scan(node) {
      if (ts.isObjectLiteralExpression(node)) {
        const data = literal(node);
        if (data.id && data.ownerKey && data.source && data.metadata && data.title) {
          const loader = node.properties.find(item => item.name?.getText() === 'load');
          const imports = [];
          function findImports(child) { if (ts.isCallExpression(child) && child.expression.kind === ts.SyntaxKind.ImportKeyword && ts.isStringLiteral(child.arguments[0])) imports.push(child.arguments[0]); ts.forEachChild(child, findImports); }
          if (loader) findImports(loader);
          const output = imports.length === 1 ? resolve(imports[0].text, source) : undefined;
          try {
            if (!output || !reachable.has(output)) throw new Error('Replica is not reachable through one literal catalogue loader.');
            const separator = data.ownerKey.lastIndexOf('#');
            const file = data.ownerKey.slice(0, separator);
            const name = data.ownerKey.slice(separator + 1);
            const absolute = path.resolve(root, file);
            if (separator < 1 || path.relative(root, absolute).startsWith('..') || data.source.file !== file) throw new Error('Invalid physical source owner.');
            const physical = ts.createSourceFile(absolute, fs.readFileSync(absolute, 'utf8'), ts.ScriptTarget.Latest, true, absolute.endsWith('.tsx') ? ts.ScriptKind.TSX : ts.ScriptKind.TS);
            const comments = [];
            if (name === '<module>') {
              for (const doc of physical.statements[0]?.jsDoc || []) if (doc.tags?.some(tag => ['packageDocumentation', 'module'].includes(tag.tagName.text))) comments.push(doc);
            } else {
              function owners(child, prefix = '') {
                let qualified = prefix;
                if (child.name && (ts.isIdentifier(child.name) || ts.isStringLiteral(child.name))) qualified = prefix ? prefix + '.' + child.name.text : child.name.text;
                if (qualified === name) for (const doc of (ts.isVariableDeclaration(child) ? child.parent.parent.jsDoc : child.jsDoc) || []) comments.push(doc);
                ts.forEachChild(child, nested => owners(nested, qualified));
              }
              owners(physical);
            }
            const diagnostics = [];
            const examples = comments.flatMap(comment => parseContract(comment.getText(), { file, line: physical.getLineAndCharacterOfPosition(comment.getStart()).line + 1 }, data.ownerKey, diagnostics).examples).filter(example => example.title === data.title);
            if (diagnostics.length || examples.length !== 1 || examples[0].language !== 'tsx') throw new Error('Missing or ambiguous physical TSX owner/title contract.');
            const example = examples[0];
            if (example.source.line !== data.source.line || output.text !== ownership + '\n' + example.code) throw new Error('Exact copied code/source position disagree with physical owner.');
            const id = createHash('sha256').update(JSON.stringify([data.ownerKey, data.title])).digest('hex');
            if (data.id !== id || path.basename(output.fileName) !== id + '.tsx') throw new Error('Owner/title identity and lazy target disagree.');
            const sourceLoader = node.properties.find(item => item.name?.getText().replace(/^["']|["']$/g, '') === 'loadSource');
            const sourceFactory = sourceLoader && ts.isPropertyAssignment(sourceLoader) ? sourceLoader.initializer : undefined;
            const sourceImport = sourceFactory && ts.isArrowFunction(sourceFactory) && sourceFactory.parameters.length === 0 ? sourceFactory.body : undefined;
            if (!sourceImport || !ts.isCallExpression(sourceImport) || sourceImport.expression.kind !== ts.SyntaxKind.ImportKeyword || sourceImport.arguments.length !== 1 || !ts.isStringLiteral(sourceImport.arguments[0]) || sourceImport.arguments[0].text !== './' + id + '.source') throw new Error('Source replica requires its own direct literal loader and chunk identity.');
            const sourceScanner = ts.createScanner(ts.ScriptTarget.Latest, false, ts.LanguageVariant.Standard, sourceImport.getText());
            const sourceComments = [];
            for (let token = sourceScanner.scan(); token !== ts.SyntaxKind.EndOfFileToken; token = sourceScanner.scan()) {
              if (token === ts.SyntaxKind.SingleLineCommentTrivia || token === ts.SyntaxKind.MultiLineCommentTrivia) sourceComments.push({ kind: token, text: sourceScanner.getTokenText() });
            }
            if (sourceComments.length !== 1 || sourceComments[0].kind !== ts.SyntaxKind.MultiLineCommentTrivia || sourceComments[0].text !== `/* webpackChunkName: 'spfx-jsdoc-source-${id}' */`) throw new Error('Source replica requires only its exact webpack chunk block comment.');
            const sourceOutput = resolve(sourceImport.arguments[0].text, source);
            if (!sourceOutput || !reachable.has(sourceOutput) || path.normalize(sourceOutput.fileName) !== path.normalize(path.join(path.dirname(output.fileName), id + '.source.ts'))) throw new Error('Source replica must be the reachable same-ID sibling of its Demo.');
            if (sourceOutput.text !== ownership + '\n' + `export const code = ${JSON.stringify(example.code)};\n`) throw new Error('Source replica must export only the exact canonical JSON string.');
            // Integrity reads source bytes without demanding its code export or
            // loader member. Runtime consumers must keep those slots live.
            const symbol = checker.getExportsOfModule(output.symbol).find(item => item.name === 'demo');
            const actual = symbol?.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(symbol) : symbol;
            const metadata = readMetadata(actual?.valueDeclaration);
            if (!metadata || JSON.stringify(metadata) !== JSON.stringify(data.metadata)) throw new Error('Exact metadata contract disagrees with catalogue.');
            if (!proven.has(source)) proven.set(source, []);
            proven.get(source).push({ node, data });
            if (metadata.hostAdapter === 'applicationCustomizerTop') { if (!topEntries.has(source)) topEntries.set(source, []); topEntries.get(source).push(output); }
            demand(output, 'demo');
            for (const member of actual.valueDeclaration.initializer.properties) artifactMembers.add(member);
            // These exact catalogue provenance fields are read by this physical
            // integrity contract, separately from application presentation.
            for (const member of node.properties) if (member.name && ['id', 'ownerKey', 'title', 'source', 'metadata'].includes(member.name.getText().replace(/^["']|["']$/g, ''))) {
              artifactMembers.add(member);
              if (member.initializer && ts.isObjectLiteralExpression(member.initializer)) for (const child of member.initializer.properties) artifactMembers.add(child);
            }
          } catch (error) { problem(node, 'DEMO_METADATA_PROVENANCE', error.message); }
        }
      }
      ts.forEachChild(node, scan);
    }
    scan(source);
  }
  for (const [source, entries] of proven) {
    const arrays = source.statements.filter(ts.isVariableStatement).flatMap(statement => statement.declarationList.declarations);
    const declaration = arrays.find(item => item.name.getText() === 'executableExamples');
    if (!declaration) continue;
    try {
      const fields = ['id', 'ownerKey', 'title', 'groupRoute', 'groupKey', 'featureKey', 'featureTitle', 'publicNames', 'source', 'metadata', 'load', 'loadSource'];
      if (!declaration.initializer || !ts.isArrayLiteralExpression(declaration.initializer) || declaration.initializer.elements.length !== entries.length || !declaration.initializer.elements.every(node => entries.some(entry => entry.node === node))) throw new Error('Catalogue array must contain exactly the physically proven entries.');
      for (const entry of entries) {
        const keys = entry.node.properties.map(property => property.name?.getText().replace(/^["']|["']$/g, '')).sort();
        if (JSON.stringify(keys) !== JSON.stringify([...fields].sort())) throw new Error('Catalogue entry differs from the exact finite provenance schema.');
        if (!Array.isArray(entry.data.publicNames) || !entry.data.publicNames.every(name => typeof name === 'string') || !['groupRoute', 'groupKey', 'featureKey', 'featureTitle'].every(name => typeof entry.data[name] === 'string' && entry.data[name])) throw new Error('Catalogue classification fields must be finite source-projection literals.');
      }
      const hierarchy = arrays.find(item => item.name.getText() === 'executableGroups')?.initializer;
      if (!hierarchy || !ts.isArrayLiteralExpression(hierarchy)) throw new Error('Missing literal catalogue hierarchy.');
      const referenced = new Set();
      const property = (object, name) => object.properties.find(item => item.name?.getText() === name)?.initializer;
      for (const group of hierarchy.elements) {
        const key = literal(property(group, 'key'));
        if (typeof literal(property(group, 'title')) !== 'string' || group.properties.map(item => item.name?.getText()).sort().join(',') !== 'features,key,title') throw new Error('Invalid finite group schema.');
        for (const member of group.properties) artifactMembers.add(member);
        const features = property(group, 'features');
        if (!ts.isArrayLiteralExpression(features)) throw new Error('Invalid group features.');
        for (const feature of features.elements) {
          const featureKey = literal(property(feature, 'key')), title = literal(property(feature, 'title'));
          if (typeof title !== 'string' || feature.properties.map(item => item.name?.getText()).sort().join(',') !== 'key,scenarios,title') throw new Error('Invalid finite feature schema.');
          for (const member of feature.properties) artifactMembers.add(member);
          const scenarios = property(feature, 'scenarios');
          if (!ts.isArrayLiteralExpression(scenarios)) throw new Error('Invalid feature scenarios.');
          for (const reference of scenarios.elements) {
            if (!ts.isElementAccessExpression(reference) || reference.expression.getText() !== 'executableExamples' || !ts.isNumericLiteral(reference.argumentExpression)) throw new Error('Hierarchy needs exact finite entry references.');
            const index = Number(reference.argumentExpression.text), entry = entries[index];
            if (!entry || referenced.has(index) || entry.data.groupKey !== key || entry.data.featureKey !== featureKey || entry.data.featureTitle !== title) throw new Error('Catalogue hierarchy and entry classification disagree.');
            referenced.add(index);
          }
        }
      }
      if (referenced.size !== entries.length) throw new Error('Catalogue hierarchy omits an accepted entry.');
      demand(source, 'executableExamples');
      acceptedCatalogues.set(source, { declaration });
      for (const { node } of entries) for (const member of node.properties) if (['groupRoute', 'groupKey', 'featureKey', 'featureTitle', 'publicNames'].includes(member.name?.getText().replace(/^["']|["']$/g, ''))) artifactMembers.add(member);
    } catch (error) { problem(declaration, 'DEMO_CATALOGUE_CONTRACT', error.message); }
  }
  const sameDeclaration = (a, b) => a && b && a.declarations?.some(declaration => b.declarations?.includes(declaration));
  function signalOf(props, at, outerExact) {
    if (!props || props.flags & (ts.TypeFlags.Any | ts.TypeFlags.Unknown | ts.TypeFlags.Never) || checker.getIndexInfosOfType(props).length) return undefined;
    if (outerExact && props.getProperties().some(property => !['render', 'resetRevision', 'applicationCustomizerTop'].includes(property.name) || !(property.flags & ts.SymbolFlags.Optional))) return undefined;
    const signal = props.getProperty('applicationCustomizerTop');
    if (!signal || !(signal.flags & ts.SymbolFlags.Optional)) return undefined;
    const type = checker.getNonNullableType(checker.getTypeOfSymbolAtLocation(signal, at));
    if (type.flags & (ts.TypeFlags.Any | ts.TypeFlags.Unknown | ts.TypeFlags.Never) || checker.getIndexInfosOfType(type).length || type.getCallSignatures().length || type.getConstructSignatures().length) return undefined;
    const fields = type.getProperties();
    if (fields.length !== 1 || fields[0].name !== 'placeholder' || fields[0].flags & ts.SymbolFlags.Optional || !fields[0].declarations?.some(node => ts.getCombinedModifierFlags(node) & ts.ModifierFlags.Readonly)) return undefined;
    const literal = checker.getTypeOfSymbolAtLocation(fields[0], at);
    return literal.isStringLiteral() && literal.value === 'Top' ? fields[0] : undefined;
  }
  for (const [catalogue, outputs] of topEntries) {
    const accepted = acceptedCatalogues.get(catalogue);
    if (!accepted) continue;
    let valid = true;
    for (const output of outputs) {
      const exported = checker.getExportsOfModule(output.symbol).find(symbol => symbol.name === 'Demo');
      const actual = exported?.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(exported) : exported;
      const signatures = actual && checker.getTypeOfSymbolAtLocation(actual, output).getCallSignatures();
      if (!signatures?.length || signatures.some(signature => !signature.parameters[0] || !signalOf(checker.getNonNullableType(checker.getTypeOfSymbolAtLocation(signature.parameters[0], output)), output, true))) { problem(output, 'DEMO_TOP_CONTRACT', 'Accepted Top source requires the exact optional readonly prerequisite and bounded outer props.'); valid = false; }
    }
    if (!valid) continue;
    const exportedProps = checker.getExportsOfModule(catalogue.symbol).find(symbol => symbol.name === 'ExecutableExampleProps');
    const common = exportedProps && checker.getDeclaredTypeOfSymbol(exportedProps);
    const commonLeaf = signalOf(common, catalogue, false);
    const entries = checker.getTypeAtLocation(accepted.declaration.name);
    const entry = checker.getIndexTypeOfType(entries, ts.IndexKind.Number);
    const load = entry?.getProperty('load');
    if (!commonLeaf || !load) { problem(catalogue, 'DEMO_TOP_CONTRACT', 'Accepted catalogue is missing its exact common prerequisite/loader channel.'); continue; }
    for (const source of reachable) {
      function supply(node) {
        if (ts.isJsxAttribute(node) && node.name.text === 'applicationCustomizerTop' && node.initializer && ts.isJsxExpression(node.initializer) && node.initializer.expression && ts.isObjectLiteralExpression(node.initializer.expression)) {
          const object = node.initializer.expression;
          const contextual = checker.getNonNullableType(checker.getContextualType(object) || checker.getTypeAtLocation(object));
          if (!sameDeclaration(contextual.getProperty('placeholder'), commonLeaf)) return;
          const opening = node.parent.parent;
          const component = checker.getSymbolAtLocation(opening.tagName)?.valueDeclaration;
          const factory = component && ts.isVariableDeclaration(component) && component.initializer && ts.isCallExpression(component.initializer) ? checker.getResolvedSignature(component.initializer)?.declaration : undefined;
          let actualLoader = false;
          if (factory?.body && reachable.has(factory.getSourceFile())) {
            function consumer(child) { if (ts.isCallExpression(child) && ts.isPropertyAccessExpression(child.expression) && sameDeclaration(checker.getSymbolAtLocation(child.expression.name), load)) actualLoader = true; ts.forEachChild(child, consumer); }
            consumer(factory.body);
          }
          if (!actualLoader) return;
          const member = object.properties[0];
          if (object.properties.length !== 1 || !ts.isPropertyAssignment(member) || member.name.getText() !== 'placeholder' || !ts.isStringLiteral(member.initializer) || member.initializer.text !== 'Top') problem(object, 'DEMO_TOP_CONTRACT', 'Actual Top loader channel must supply exactly placeholder: Top.');
          else artifactMembers.add(member);
        }
        ts.forEachChild(node, supply);
      }
      supply(source);
    }
  }
}
module.exports = { proveMetadataReplicas };

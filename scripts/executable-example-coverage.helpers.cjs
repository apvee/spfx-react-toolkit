/** Read generated literal metadata only. Loaders are never evaluated/imported. */
function collectExecutableCoverage(text) {
  const ts = require('typescript');
  const source = ts.createSourceFile('catalogue.ts', text, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
  let entries;
  for (const statement of source.statements) {
    if (!ts.isVariableStatement(statement)) continue;
    for (const declaration of statement.declarationList.declarations) {
      if (ts.isIdentifier(declaration.name) && declaration.name.text === 'executableExamples') entries = declaration.initializer;
    }
  }
  if (!entries || !ts.isArrayLiteralExpression(entries)) throw new Error('Generated catalogue must declare executableExamples as a literal array.');
  const nativeOwners = new Set();
  const ids = new Set(), identities = new Set(), features = new Set(), symbols = new Set(), associations = new Set();
  for (const entry of entries.elements) {
    if (!ts.isObjectLiteralExpression(entry)) throw new Error('Generated scenario must be a literal object.');
    const fields = new Map(entry.properties.filter(ts.isPropertyAssignment).map(property => [property.name.text, property.initializer]));
    function string(name) {
      const value = fields.get(name);
      if (!value || !ts.isStringLiteral(value) || !value.text) throw new Error(`Generated scenario requires literal ${name}.`);
      return value.text;
    }
    const id = string('id'), identity = JSON.stringify([string('ownerKey'), string('title')]);
    if (identities.has(identity)) throw new Error(`Duplicate scenario identity: ${identity}`);
    if (ids.has(id)) throw new Error(`Duplicate scenario id: ${id}`);
    identities.add(identity); ids.add(id); features.add(string('featureKey'));
    const metadata = fields.get('metadata');
    if (metadata && ts.isObjectLiteralExpression(metadata) && metadata.properties.some(property =>
      ts.isPropertyAssignment(property) && property.name.text === 'hostAdapter' &&
      ts.isStringLiteral(property.initializer) && property.initializer.text === 'propertyPane')) nativeOwners.add(string('ownerKey'));
    const names = fields.get('publicNames');
    if (!names || !ts.isArrayLiteralExpression(names)) throw new Error('Generated scenario requires literal publicNames.');
    for (const name of names.elements) {
      if (!ts.isStringLiteral(name) || !name.text) throw new Error('Generated public name must be a nonempty literal.');
      symbols.add(name.text); associations.add(JSON.stringify([id, name.text]));
    }
  }
  return { featureCount: features.size, scenarioCount: ids.size, symbols, associationCount: associations.size, nativeOwners };
}


module.exports = { collectExecutableCoverage };

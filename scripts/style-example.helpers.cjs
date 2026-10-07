const assert = require('node:assert/strict');
const ts = require('typescript');

function parse(source) { return ts.createSourceFile('example.tsx', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX); }
function walk(node, callback) { callback(node); ts.forEachChild(node, child => walk(child, callback)); }
function property(object, name) {
  return object.properties.find(item => ts.isPropertyAssignment(item) && item.name.getText().replace(/['"]/g, '') === name)?.initializer;
}

/** Static linkage complements behavioral demo tests; it cannot establish browser outcomes. */
function assertStyleExampleCoverage(surface, registrySource, panelSource, catalogSource) {
  let entry;
  walk(parse(registrySource), node => {
    if (ts.isObjectLiteralExpression(node) && property(node, 'key')?.text === 'styles') entry = node;
  });
  assert.ok(entry, 'Styles lazy demo entry missing');
  assert.match(property(entry, 'component')?.getText() ?? '', /lazyPanel[\s\S]*import\([\s\S]*StylesPanel/, 'Styles panel must be lazy');
  const registrySymbols = new Set();
  walk(property(entry, 'coverage'), node => {
    if (ts.isPropertyAssignment(node) && node.name.getText() === 'symbol') registrySymbols.add(node.initializer.text);
  });
  const styleSymbols = new Set(surface.filter(item => !item.name.startsWith('Sx')).map(item => item.name.split('.')[0]));
  for (const symbol of styleSymbols) assert.ok(registrySymbols.has(symbol), `Styles registry missing ${symbol}`);

  const families = new Map();
  walk(parse(catalogSource), node => {
    if (ts.isCallExpression(node) && node.expression.getText() === 'family' && ts.isStringLiteral(node.arguments[0])) {
      families.set(node.arguments[0].text, node.arguments[1]);
    }
  });
  for (const item of surface) {
    if (item.name.startsWith('Sx') || /^(hover|active|focusVisible|responsive\.|viewport\.)/.test(item.name)) continue;
    const split = item.name.lastIndexOf('.');
    const namespace = item.name.slice(0, split);
    const member = item.name.slice(split + 1);
    const family = families.get(namespace);
    assert.ok(family, `Interactive style catalog missing ${namespace}`);
    if (ts.isObjectLiteralExpression(family)) {
      assert.equal(property(family, member)?.getText(), item.name, `Interactive style catalog missing ${item.name}`);
    } else {
      assert.equal(family.getText(), namespace, `Interactive style catalog does not expose ${item.name}`);
    }
  }
  for (const symbol of ['hover', 'active', 'focusVisible', 'responsive', 'viewport']) {
    if (!styleSymbols.has(symbol)) continue;
    assert.match(panelSource, new RegExp('\\b' + symbol + '\\s*(?:\\[|\\.|\\()'), `Interactive panel missing ${symbol} scope use`);
  }
  const thresholdChoices = new Map();
  walk(parse(panelSource), node => {
    if (!ts.isJsxSelfClosingElement(node) && !ts.isJsxOpeningElement(node)) return;
    const id = node.attributes.properties.find(attribute => ts.isJsxAttribute(attribute) && attribute.name.text === 'id')?.initializer;
    const choices = node.attributes.properties.find(attribute => ts.isJsxAttribute(attribute) && attribute.name.text === 'choices')?.initializer;
    if (id && ts.isStringLiteral(id) && choices && ts.isJsxExpression(choices) && choices.expression && ts.isArrayLiteralExpression(choices.expression)) {
      thresholdChoices.set(id.text, new Set(choices.expression.elements.filter(ts.isStringLiteral).map(element => element.text)));
    }
  });
  for (const item of surface.filter(declaration => /^(responsive|viewport)\./.test(declaration.name))) {
    const [namespace, threshold] = item.name.split('.');
    assert.ok(thresholdChoices.get(namespace + '-threshold')?.has(threshold), `Interactive panel missing ${item.name} threshold choice`);
  }
  assert.match(panelSource, /\buseSx\s*\(/, 'Interactive panel must use the hook');
  assert.match(panelSource, /stylesCatalog\s*\[/, 'Interactive panel must consume the catalog');
  assert.match(panelSource, /onChange\s*=/, 'Interactive panel must expose selection actions');
  assert.match(panelSource, /className\s*=\s*\{sx\([^}]*\bdescriptor\b/, 'Interactive preview must apply the selected descriptor through sx');
}
module.exports = { assertStyleExampleCoverage };

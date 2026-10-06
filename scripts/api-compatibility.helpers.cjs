const assert = require('node:assert/strict');
const ts = require('typescript');

const printer = ts.createPrinter({ removeComments: true });
const approvedAdditions = new Map([
  ['hooks/index.d.ts', './useSPFxSiteKeyValueStore'],
  ['services/index.d.ts', './spfx-site-key-value-store.service'],
]);

function parseDeclaration(text) {
  return ts.createSourceFile('module.d.ts', text, ts.ScriptTarget.Latest, true);
}

function assertCompatibleDeclaration(moduleName, currentText, baselineText, options = {}) {
  const current = parseDeclaration(currentText);
  const approvedPath = approvedAdditions.get(moduleName);
  let additionCount = 0;
  const historicalStatements = current.statements.filter(statement => {
    if (approvedPath && ts.isExportDeclaration(statement) &&
        !statement.exportClause && !statement.isTypeOnly &&
        !statement.assertClause && !statement.attributes &&
        !statement.modifiers?.length && statement.moduleSpecifier &&
        ts.isStringLiteral(statement.moduleSpecifier) &&
        statement.moduleSpecifier.text === approvedPath) {
      additionCount += 1;
      return false;
    }
    return true;
  });
  assert.ok(additionCount <= 1, `Duplicate approved export: ${moduleName}`);
  if (approvedPath && options.requireApprovedAdditions) {
    assert.equal(additionCount, 1, `Missing approved export: ${moduleName}`);
  }
  const historicalCurrent = ts.factory.updateSourceFile(current, historicalStatements);
  assert.equal(
    printer.printFile(historicalCurrent),
    printer.printFile(parseDeclaration(baselineText)),
    `Declaration contract changed: ${moduleName}`
  );
}

module.exports = { assertCompatibleDeclaration };

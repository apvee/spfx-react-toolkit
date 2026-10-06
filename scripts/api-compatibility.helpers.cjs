const assert = require('node:assert/strict');
const ts = require('typescript');

const printer = ts.createPrinter({ removeComments: true });
const approvedAdditions = new Map([
  ['hooks/index.d.ts', './useSPFxSiteKeyValueStore'],
  ['services/index.d.ts', './spfx-site-key-value-store.service'],
]);
const pnpHookPaths = ['./useSPFxPnPListById', './useSPFxPnPListByUrl', './useSPFxPnPListByPath'];
const pnpServiceModule = 'services/spfx-pnp-list.service.d.ts';

function parseDeclaration(text) {
  return ts.createSourceFile('module.d.ts', text, ts.ScriptTarget.Latest, true);
}

// Keep these approvals explicit: only this readonly discriminated union and
// the factory's second parameter may differ from the historical service.
const selectorSource = parseDeclaration(`export type SPFxPnPListSelector =
  | { readonly kind: 'title'; readonly title: string }
  | { readonly kind: 'id'; readonly id: string }
  | { readonly kind: 'url'; readonly serverRelativeUrl: string }
  | { readonly kind: 'path'; readonly webRelativePath: string };`);
const parameterSource = parseDeclaration('declare function approved(listTarget: string | SPFxPnPListSelector): void;');
const printNode = (node, source) => printer.printNode(ts.EmitHint.Unspecified, node, source);
const approvedSelector = printNode(selectorSource.statements[0], selectorSource);
const approvedParameter = printNode(parameterSource.statements[0].parameters[0], parameterSource);

function isPlainExportAll(statement, modulePath) {
  return ts.isExportDeclaration(statement) &&
    !statement.exportClause && !statement.isTypeOnly &&
    !statement.assertClause && !statement.attributes &&
    !statement.modifiers?.length && statement.moduleSpecifier &&
    ts.isStringLiteral(statement.moduleSpecifier) &&
    statement.moduleSpecifier.text === modulePath;
}

function assertCompatibleDeclaration(moduleName, currentText, baselineText, options = {}) {
  const current = parseDeclaration(currentText);
  const baseline = parseDeclaration(baselineText);
  const approvedPath = approvedAdditions.get(moduleName);
  const pnpPaths = moduleName === 'hooks/index.d.ts' ? pnpHookPaths : [];
  const exportCounts = new Map([...(approvedPath ? [approvedPath] : []), ...pnpPaths].map(modulePath => [modulePath, 0]));
  let selectorCount = 0;
  let factoryCount = 0;
  const historicalStatements = [];
  for (const statement of current.statements) {
    const exportPath = [...exportCounts.keys()].find(modulePath => isPlainExportAll(statement, modulePath));
    if (exportPath) {
      exportCounts.set(exportPath, exportCounts.get(exportPath) + 1);
      continue;
    }
    if (moduleName === pnpServiceModule) {
      if (ts.isTypeAliasDeclaration(statement) && printNode(statement, current) === approvedSelector) {
        selectorCount += 1;
        continue;
      }
      if (ts.isFunctionDeclaration(statement) && statement.name?.text === 'createSPFxPnPListService' &&
          statement.parameters[1] && printNode(statement.parameters[1], current) === approvedParameter) {
        factoryCount += 1;
        const historicalFactory = baseline.statements.find(node => ts.isFunctionDeclaration(node) && node.name?.text === 'createSPFxPnPListService');
        assert.ok(historicalFactory?.parameters[1], 'Missing historical PnP list factory baseline');
        historicalStatements.push(ts.factory.updateFunctionDeclaration(
          statement, statement.modifiers, statement.asteriskToken, statement.name, statement.typeParameters,
          [statement.parameters[0], historicalFactory.parameters[1], ...statement.parameters.slice(2)], statement.type, statement.body
        ));
        continue;
      }
    }
    historicalStatements.push(statement);
  }
  for (const [modulePath, count] of exportCounts) {
    assert.ok(count <= 1, `Duplicate approved export: ${moduleName}: ${modulePath}`);
  }
  if (approvedPath && options.requireApprovedAdditions) {
    assert.equal(exportCounts.get(approvedPath), 1, `Missing approved export: ${moduleName}`);
  }
  if (options.requirePnPListAdditions) {
    for (const modulePath of pnpPaths) {
      assert.equal(exportCounts.get(modulePath), 1, `Missing approved PnP list export: ${moduleName}: ${modulePath}`);
    }
    if (moduleName === pnpServiceModule) {
      assert.equal(selectorCount, 1, 'Missing approved PnP list selector');
      assert.equal(factoryCount, 1, 'Missing approved PnP list factory widening');
    }
  }
  assert.ok(selectorCount <= 1, 'Duplicate approved PnP list selector');
  assert.ok(factoryCount <= 1, 'Duplicate approved PnP list factory');
  const historicalCurrent = ts.factory.updateSourceFile(current, historicalStatements);
  assert.equal(
    printer.printFile(historicalCurrent),
    printer.printFile(baseline),
    `Declaration contract changed: ${moduleName}`
  );
}

module.exports = { assertCompatibleDeclaration };

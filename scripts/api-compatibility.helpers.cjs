const assert = require('node:assert/strict');
const ts = require('typescript');
const { assertEntrypointContract } = require('./package-entrypoints.cjs');

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
  const stablePaths = moduleName === 'hooks/index.d.ts' ? ['./useStableCallback'] : [];
  const sxPaths = moduleName === 'hooks/index.d.ts' ? ['./useSx'] : moduleName === 'helpers/index.d.ts' ? ['./styles'] : [];
  const exportCounts = new Map([...(approvedPath ? [approvedPath] : []), ...pnpPaths, ...stablePaths, ...sxPaths].map(modulePath => [modulePath, 0]));
  let selectorCount = 0;
  let factoryCount = 0;
  let websRegistrationCount = 0;
  const historicalStatements = [];
  for (const statement of current.statements) {
    const exportPath = [...exportCounts.keys()].find(modulePath => isPlainExportAll(statement, modulePath));
    if (exportPath) {
      exportCounts.set(exportPath, exportCounts.get(exportPath) + 1);
      continue;
    }
    if (moduleName === pnpServiceModule) {
      // The standalone list service needs this single additive registration.
      // Named/type imports, attributes and every other module remain historical.
      if (ts.isImportDeclaration(statement) && !statement.importClause &&
          !statement.assertClause && !statement.attributes && !statement.modifiers?.length &&
          ts.isStringLiteral(statement.moduleSpecifier) && statement.moduleSpecifier.text === '@pnp/sp/webs') {
        websRegistrationCount += 1;
        continue;
      }
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
  if (options.requireSxAdditions) {
    for (const modulePath of sxPaths) {
      assert.equal(exportCounts.get(modulePath), 1, `Missing approved style export: ${moduleName}: ${modulePath}`);
    }
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
  assert.ok(websRegistrationCount <= 1, 'Duplicate approved PnP list webs registration');
  if (moduleName === pnpServiceModule && options.requirePnPListWebsRegistration) {
    assert.equal(websRegistrationCount, 1, 'Missing approved PnP list webs registration');
  }
  const historicalCurrent = ts.factory.updateSourceFile(current, historicalStatements);
  assert.equal(
    printer.printFile(historicalCurrent),
    printer.printFile(baseline),
    `Declaration contract changed: ${moduleName}`
  );
}

// Explicit additive peer contract: shared instances are supplied by the host.
// Historical fixtures are immutable, including the original Fluent dependencies.
function assertCompatiblePackageContract(manifest, originalPackage) {
  const sharedFluentPeers = {
    '@fluentui/react-migration-v8-v9': originalPackage.dependencies['@fluentui/react-migration-v8-v9'],
    '@fluentui/react-theme': originalPackage.dependencies['@fluentui/react-theme'],
  };
  assert.deepEqual(sharedFluentPeers, {
    '@fluentui/react-migration-v8-v9': '^9.9.12', '@fluentui/react-theme': '^9.2.0',
  });
  assert.deepEqual(manifest.peerDependencies, {
    ...originalPackage.peerDependencies, ...sharedFluentPeers,
    '@fluentui/react-utilities': '^9.25.1',
    '@griffel/core': '^1.19.2', '@griffel/react': '^1.5.30',
    '@fluentui/react-shared-contexts': '^9.25.2',
  }, 'Runtime peer contracts changed outside the approved shared Fluent/Griffel peers');
  assert.deepEqual(manifest.peerDependenciesMeta ?? {}, originalPackage.peerDependenciesMeta ?? {},
    'Mandatory peers must not become optional');
  const expectedDependencies = { ...originalPackage.dependencies };
  assert.equal(expectedDependencies.tslib, '2.3.1', 'Unexpected historical tslib contract');
  delete expectedDependencies.tslib;
  for (const name of Object.keys(sharedFluentPeers)) delete expectedDependencies[name];
  assert.deepEqual(manifest.dependencies ?? {}, expectedDependencies,
    'Runtime dependencies changed outside the approved tslib cleanup and shared Fluent peers');
  assert.equal(manifest.private, false);
  assert.deepEqual(manifest.files, [...originalPackage.files, 'lib/styles/**/*'], 'Published path allowlist changed outside the styles facade');
  assert.equal(manifest.name, '@apvee/spfx-react-toolkit');
  assert.equal(manifest.version, '2.1.0');
  assert.equal(manifest.main, 'lib/index.js');
  assert.equal(manifest.types, 'lib/index.d.ts');
  assertEntrypointContract(manifest);
}

module.exports = { assertCompatibleDeclaration, assertCompatiblePackageContract };

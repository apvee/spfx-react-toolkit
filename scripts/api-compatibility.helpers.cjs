const assert = require('node:assert/strict');
const ts = require('typescript');
const { assertEntrypointContract } = require('./package-entrypoints.cjs');

const printer = ts.createPrinter({ removeComments: true });
const approvedAdditions = new Map([
  ['index.d.ts', './controls'],
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
const ianaSource = parseDeclaration('interface Approved { readonly ianaTimeZone?: string; }');
const approvedIanaProperty = printNode(ianaSource.statements[0].members[0], ianaSource);
const searchOperatorSource = parseDeclaration("interface Approved { refinerOperator?: 'and' | 'or'; }");
const approvedSearchOperator = printNode(searchOperatorSource.statements[0].members[0], searchOperatorSource);
const searchOperatorInterfaces = new Map([
  ['hooks/useSPFxPnPSearch.d.ts', ['UseSPFxPnPSearchOptions']],
  ['services/spfx-pnp-search.service.d.ts', ['SPFxPnPSearchOptions', 'SPFxPnPSearchRequestOptions']],
]);

// Exact additive scoped permission declarations; historical signatures and types
// are compared unchanged after removing only these approved AST statements.
const scopedPermissionModule = 'hooks/useSPFxCrossSitePermissions.d.ts';
const scopedPermissionSource = parseDeclaration(`
import type { SPFxScopedPermissionsOptions } from '../services/spfx-permission-target.types';
export type { SPFxScopedPermissionsOptions } from '../services/spfx-permission-target.types';
export interface SPFxScopedPermissionsInfo {
  readonly permissions: SPPermission | undefined;
  readonly hasPermission: (permission: SPPermission) => boolean;
  readonly isLoading: boolean;
  readonly error: Error | undefined;
}
export declare function useSPFxCrossSitePermissions(options: SPFxScopedPermissionsOptions): SPFxScopedPermissionsInfo;
`);
const approvedScopedStatements = scopedPermissionSource.statements.map(statement => printNode(statement, scopedPermissionSource));

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
  const configurationPaths = moduleName === 'hooks/index.d.ts' ? ['./useSPFxConfigurationStore']
    : moduleName === 'services/index.d.ts' ? ['./spfx-configuration-store.service'] : [];
  const exportCounts = new Map([...(approvedPath ? [approvedPath] : []), ...pnpPaths, ...stablePaths, ...sxPaths, ...configurationPaths].map(modulePath => [modulePath, 0]));
  const scopedCounts = approvedScopedStatements.map(() => 0);
  let selectorCount = 0;
  let factoryCount = 0;
  let websRegistrationCount = 0;
  const historicalStatements = [];
  let ianaCount = 0;
  const isLocaleHook = moduleName === 'hooks/useSPFxLocaleInfo.d.ts';
  const isLocaleHelper = moduleName === 'helpers/spfx-page-context.helpers.d.ts';
  const searchOperatorCounts = new Map((searchOperatorInterfaces.get(moduleName) ?? []).map(name => [name, 0]));
  function historicalMembers(members) {
    return members.filter(member => {
      if (printNode(member, current) !== approvedIanaProperty) return true;
      ianaCount++; return false;
    });
  }
  for (const statement of current.statements) {
    if (moduleName === scopedPermissionModule) {
      const approvedIndex = approvedScopedStatements.indexOf(printNode(statement, current));
      if (approvedIndex >= 0) { scopedCounts[approvedIndex]++; continue; }
    }
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
    if (ts.isInterfaceDeclaration(statement) && searchOperatorCounts.has(statement.name.text)) {
      const members = statement.members.filter(member => {
        if (printNode(member, current) !== approvedSearchOperator) return true;
        searchOperatorCounts.set(statement.name.text, searchOperatorCounts.get(statement.name.text) + 1);
        return false;
      });
      historicalStatements.push(ts.factory.updateInterfaceDeclaration(statement, statement.modifiers, statement.name,
        statement.typeParameters, statement.heritageClauses, members));
    } else if (isLocaleHook && ts.isInterfaceDeclaration(statement) && statement.name.text === 'SPFxLocaleInfo') {
      historicalStatements.push(ts.factory.updateInterfaceDeclaration(statement, statement.modifiers, statement.name,
        statement.typeParameters, statement.heritageClauses, historicalMembers(statement.members)));
    } else if (isLocaleHelper && ts.isFunctionDeclaration(statement) && statement.name?.text === 'getSPFxLocaleInfo' && ts.isTypeLiteralNode(statement.type)) {
      historicalStatements.push(ts.factory.updateFunctionDeclaration(statement, statement.modifiers, statement.asteriskToken,
        statement.name, statement.typeParameters, statement.parameters, ts.factory.createTypeLiteralNode(historicalMembers(statement.type.members)), statement.body));
    } else historicalStatements.push(statement);
  }
  for (const [modulePath, count] of exportCounts) {
    assert.ok(count <= 1, `Duplicate approved export: ${moduleName}: ${modulePath}`);
  }
  if (approvedPath && ((moduleName === 'index.d.ts' && options.requireControlsAdditions) || (moduleName !== 'index.d.ts' && options.requireApprovedAdditions))) {
    assert.equal(exportCounts.get(approvedPath), 1, `Missing approved export: ${moduleName}`);
  }
  if (options.requireSxAdditions) {
    for (const modulePath of sxPaths) {
      assert.equal(exportCounts.get(modulePath), 1, `Missing approved style export: ${moduleName}: ${modulePath}`);
    }
  }
  if (options.requireConfigurationStoreAdditions) {
    for (const modulePath of configurationPaths) {
      assert.equal(exportCounts.get(modulePath), 1, `Missing approved configuration export: ${moduleName}: ${modulePath}`);
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
  assert.ok(ianaCount <= 1, 'Duplicate approved IANA property');
  if (options.requireIanaAdditions && (isLocaleHook || isLocaleHelper)) assert.equal(ianaCount, 1, 'Missing approved IANA property');
  for (const [name, count] of searchOperatorCounts) {
    assert.ok(count <= 1, `Duplicate approved search refiner operator: ${name}`);
    if (options.requireSearchRefinerOperator) assert.equal(count, 1, `Missing approved search refiner operator: ${name}`);
  }
  if (moduleName === scopedPermissionModule) {
    for (const [index, count] of scopedCounts.entries()) {
      assert.ok(count <= 1, `Duplicate approved scoped permission declaration: ${index}`);
      if (options.requireScopedPermissionsAdditions) assert.equal(count, 1, `Missing approved scoped permission declaration: ${index}`);
    }
  }
  const historicalCurrent = ts.factory.updateSourceFile(current, historicalStatements);
  assert.equal(
    printer.printFile(historicalCurrent),
    printer.printFile(baseline),
    `Declaration contract changed: ${moduleName}`
  );
}

// Explicit additive peer contract: shared Fluent/Griffel/Form instances are supplied by the host.
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
    '@microsoft/sp-property-pane': '>=1.18.0 <2.0.0',
    '@fluentui/react-utilities': '^9.26.7',
    '@griffel/core': '^1.21.4', '@griffel/react': '^1.7.8',
    '@fluentui/react-shared-contexts': '^9.26.4',
    '@fluentui/react-components': '^9.74.9',
    'react-hook-form': '>=7.87.0 <8.0.0',
    '@fluentui/react-datepicker-compat': '>=0.6.16 <1.0.0',
    '@fluentui/react-timepicker-compat': '>=0.4.21 <1.0.0',
    '@fluentui/react-icons': '>=2.0.312 <3.0.0',
    '@fluentui/react-charts': '^9.3.27',
    '@fluentui/react-theme': '^9.2.2',
    '@fluentui/react-migration-v8-v9': '^9.10.20',
  }, 'Runtime peer contracts changed outside the approved shared Fluent/Griffel/Form peers');
  assert.deepEqual(manifest.peerDependenciesMeta ?? {}, originalPackage.peerDependenciesMeta ?? {},
    'Mandatory peers must not become optional');
  const expectedDependencies = { ...originalPackage.dependencies };
  assert.equal(expectedDependencies.tslib, '2.3.1', 'Unexpected historical tslib contract');
  delete expectedDependencies.tslib;
  for (const name of Object.keys(sharedFluentPeers)) delete expectedDependencies[name];
  Object.assign(expectedDependencies, {
    unified: '11.0.5', 'remark-parse': '11.0.0', 'remark-gfm': '4.0.1',
    '@hookform/resolvers': '5.9.1', '@standard-schema/spec': '1.1.0',
    // Exact private engine versions approved by the Shell dependency checkpoint.
    '@tiptap/core': '3.30.3',
    '@tiptap/pm': '3.30.3',
    '@tiptap/react': '3.30.3',
    '@tiptap/extensions': '3.30.3',
    '@tiptap/extension-document': '3.30.3',
    '@tiptap/extension-text': '3.30.3',
    '@tiptap/extension-paragraph': '3.30.3',
    '@tiptap/extension-heading': '3.30.3',
    '@tiptap/extension-hard-break': '3.30.3',
    '@tiptap/extension-bold': '3.30.3',
    '@tiptap/extension-italic': '3.30.3',
    '@tiptap/extension-underline': '3.30.3',
    '@tiptap/extension-bubble-menu': '3.30.3',
    '@tiptap/extension-floating-menu': '3.30.3',
  });
  assert.deepEqual(manifest.dependencies ?? {}, expectedDependencies,
    'Runtime dependencies changed outside the approved tslib cleanup, shared Fluent peers, Form resolver/schema dependencies and exact private Shell engine dependencies');
  assert.equal(manifest.private, false);
  assert.deepEqual(manifest.files, [...originalPackage.files, 'lib/styles/**/*', 'lib/controls/**/*', 'THIRD-PARTY-NOTICES.md', 'lib/property-pane/**/*', 'cli/spfx-react-toolkit.cjs', 'localization/resources.json', 'lib/localization/**/*'], 'Published path allowlist changed outside approved public entries and native localization metadata');
  assert.equal(manifest.name, '@apvee/spfx-react-toolkit');
  assert.equal(manifest.version, '2.1.0');
  assert.equal(manifest.main, 'lib/index.js');
  assert.equal(manifest.types, 'lib/index.d.ts');
  assertEntrypointContract(manifest);
}

module.exports = { assertCompatibleDeclaration, assertCompatiblePackageContract };

const assert = require('node:assert/strict');
const test = require('node:test');
const baseline = require('./fixtures/api-baseline.json');
const { assertCompatibleDeclaration } = require('../scripts/api-compatibility.helpers.cjs');

test('preserves every historical declaration, ignoring comments and whitespace', () => {
  for (const [name, snapshot] of Object.entries(baseline)) {
    assert.doesNotThrow(() => assertCompatibleDeclaration(name, snapshot.declaration, snapshot.declaration));
    assert.doesNotThrow(() => assertCompatibleDeclaration(name, `\n/* updated documentation */\n${snapshot.declaration}\n`, snapshot.declaration));
  }
});

const tenantModules = [
  'hooks/useSPFxTenantKeyValueStore.d.ts',
  'services/spfx-tenant-key-value-store.service.d.ts',
];
for (const name of tenantModules) {
  const original = baseline[name].declaration;
  for (const [label, before, after] of [
    ['generic default', '<T = unknown>', '<T = string>'],
    ['readonly property', 'readonly key: string;', 'key: string;'],
    ['parameter type', 'key: string', 'key: number'],
    ['optional parameter', 'description?: string', 'description: string'],
  ]) {
    test(`rejects changed tenant ${label} in ${name}`, () => {
      assert.ok(original.includes(before));
      assert.throws(() => assertCompatibleDeclaration(name, original.replace(before, after), original), /Declaration contract changed/);
    });
  }
}

const allowedBarrels = [
  ['hooks/index.d.ts', './useSPFxSiteKeyValueStore', './useSPFxTenantKeyValueStore'],
  ['services/index.d.ts', './spfx-site-key-value-store.service', './spfx-tenant-key-value-store.service'],
];
for (const [name, addition, historical] of allowedBarrels) {
  const original = baseline[name].declaration;
  const exportAll = `export * from '${addition}';`;
  test(`permits one approved export-all in ${name}`, () => {
    assert.doesNotThrow(() => assertCompatibleDeclaration(name, `${original}\n${exportAll}`, original));
    assert.doesNotThrow(() => assertCompatibleDeclaration(name, `export /* comment */ * from "${addition}";\n${original}`, original));
  });
  for (const [label, current] of [
    ['duplicate approved exports', `${original}\n${exportAll}\n${exportAll}`],
    ['removed historical export', original.replace(`export * from '${historical}';`, '')],
    ['replaced historical export', original.replace(`export * from '${historical}';`, exportAll)],
    ['unrelated export', `${original}\nexport * from './unexpected';`],
    ['wrong path', `${original}\nexport * from '${addition}.js';`],
    ['named export', `${original}\nexport { Example } from '${addition}';`],
    ['namespace export', `${original}\nexport * as Example from '${addition}';`],
    ['type-only export', `${original}\nexport type * from '${addition}';`],
    ['export with attributes', `${original}\nexport * from '${addition}' with { type: 'json' };`],
    ['unrelated declaration alongside approved export', `${original}\n${exportAll}\nexport declare const unexpected: string;`],
  ]) {
    test(`rejects ${label} in ${name}`, () => {
      assert.throws(() => assertCompatibleDeclaration(name, current, original));
    });
  }
  test(`rejects the approved ${name} path in other historical modules`, () => {
    for (const otherName of ['index.d.ts', 'helpers/index.d.ts', ...tenantModules, ...allowedBarrels.map(([barrel]) => barrel).filter(barrel => barrel !== name)]) {
      const otherOriginal = baseline[otherName].declaration;
      assert.throws(() => assertCompatibleDeclaration(otherName, `${otherOriginal}\n${exportAll}`, otherOriginal), /Declaration contract changed/);
    }
  });
}

for (const [name, addition] of allowedBarrels) {
  test(`finished feature requires its approved export in ${name}`, () => {
    const original = baseline[name].declaration;
    assert.throws(() => assertCompatibleDeclaration(name, original, original, { requireApprovedAdditions: true }), /Missing approved export/);
    assert.doesNotThrow(() => assertCompatibleDeclaration(name, `${original}\nexport * from '${addition}';`, original, { requireApprovedAdditions: true }));
  });
}

for (const [missingName] of allowedBarrels) {
  test(`integrated verify-api rejects a finished feature missing ${missingName} export`, () => {
    const fs = require('node:fs');
    const path = require('node:path');
    const scriptPath = path.resolve(__dirname, '../scripts/verify-api.cjs');
    const localRequire = require('node:module').createRequire(scriptPath);
    const declarationsRoot = path.resolve(__dirname, '../packages/spfx-react-toolkit/lib') + path.sep;
    const fileBoundary = {
      ...fs,
      existsSync(file) {
        if (file.endsWith('/src/hooks/useSPFxPnPListById.ts')) return false;
        return file.endsWith('/src/hooks/useSPFxSiteKeyValueStore.ts') || fs.existsSync(file);
      },
      readFileSync(file, options) {
        if (file.startsWith(declarationsRoot)) {
          const name = file.slice(declarationsRoot.length);
          if (baseline[name]) {
            const addition = allowedBarrels.find(([barrel]) => barrel === name)?.[1];
            return baseline[name].declaration + (addition && name !== missingName ? `\nexport * from '${addition}';` : '');
          }
        }
        return fs.readFileSync(file, options);
      },
    };
    const requireBoundary = key => key === 'node:fs' ? fileBoundary : localRequire(key);
    const verify = new Function('require', '__dirname', fs.readFileSync(scriptPath, 'utf8'));
    assert.throws(() => verify(requireBoundary, path.dirname(scriptPath)), /Missing approved export/);
  });
}

const pnpHookPaths = ['./useSPFxPnPListById', './useSPFxPnPListByUrl', './useSPFxPnPListByPath'];
const pnpExports = pnpHookPaths.map(module => `export * from '${module}';`).join('\n');
const pnpServiceName = 'services/spfx-pnp-list.service.d.ts';
const selectorDeclaration = `export type SPFxPnPListSelector =
  | { readonly kind: 'title'; readonly title: string }
  | { readonly kind: 'id'; readonly id: string }
  | { readonly kind: 'url'; readonly serverRelativeUrl: string }
  | { readonly kind: 'path'; readonly webRelativePath: string };`;
const widenedService = baseline[pnpServiceName].declaration.replace('listTitle: string', 'listTarget: string | SPFxPnPListSelector');
const approvedService = `${selectorDeclaration}\n${widenedService}`;
const pnpRequired = { requirePnPListAdditions: true };

test('permits exactly the PnP selector and factory widening with or without completion requirements', () => {
  for (const options of [{}, pnpRequired]) {
    assert.doesNotThrow(() => assertCompatibleDeclaration(pnpServiceName, approvedService, baseline[pnpServiceName].declaration, options));
  }
});

test('PnP completion requirements reject missing selector and missing factory widening separately', () => {
  for (const current of [baseline[pnpServiceName].declaration, widenedService, `${selectorDeclaration}\n${baseline[pnpServiceName].declaration}`]) {
    assert.throws(() => assertCompatibleDeclaration(pnpServiceName, current, baseline[pnpServiceName].declaration, pnpRequired), /Missing approved PnP list/);
  }
});

for (const [label, before, after] of [
  ['mutable selector', 'readonly title: string', 'title: string'],
  ['selector discriminant', "kind: 'id'", "kind: 'guid'"],
  ['optional selector value', 'readonly id: string', 'readonly id?: string'],
  ['selector value type', 'readonly webRelativePath: string', 'readonly webRelativePath: number'],
  ['selector extra field', 'readonly title: string', 'readonly title: string; readonly extra: boolean'],
  ['selector union order', "kind: 'title'; readonly title: string", "kind: 'id'; readonly id: string"],
  ['factory parameter name', 'listTarget:', 'listTitle:'],
  ['factory target optionality', 'listTarget:', 'listTarget?:'],
  ['factory target narrowing', 'string | SPFxPnPListSelector', 'SPFxPnPListSelector'],
  ['factory extra union', 'string | SPFxPnPListSelector', 'string | SPFxPnPListSelector | number'],
  ['factory generic', 'createSPFxPnPListService<T = unknown>', 'createSPFxPnPListService<T = string>'],
  ['factory client parameter', 'sp: SPFI', 'sp: unknown'],
  ['factory third parameter', 'defaultPageSize?: number', 'defaultPageSize: number'],
  ['factory return', '): SPFxPnPListService<T>', '): SPFxPnPListService<unknown>'],
  ['service method return', 'getById: (id: number) => Promise<T>', 'getById: (id: number) => Promise<T | undefined>'],
  ['query generic result', 'readonly items: T[]', 'readonly items: unknown[]'],
  ['batch result readonly', 'readonly value: TValue', 'value: TValue'],
  ['historical import', "import type { IItems } from '@pnp/sp/items';", ''],
]) {
  test(`rejects unapproved PnP ${label}`, () => {
    assert.ok(approvedService.includes(before));
    assert.throws(() => assertCompatibleDeclaration(pnpServiceName, approvedService.replace(before, after), baseline[pnpServiceName].declaration));
  });
}
for (const [label, current] of [
  ['duplicate selector', `${approvedService}\n${selectorDeclaration}`],
  ['duplicate factory', `${approvedService}\nexport declare function createSPFxPnPListService<T = unknown>(sp: SPFI, listTarget: string | SPFxPnPListSelector, defaultPageSize?: number): SPFxPnPListService<T>;`],
  ['unrelated declaration', `${approvedService}\nexport type Unrelated = string;`],
]) {
  test(`rejects PnP ${label}`, () => {
    assert.throws(() => assertCompatibleDeclaration(pnpServiceName, current, baseline[pnpServiceName].declaration));
  });
}

test('permits all PnP export-alls independently of site-store completion requirements', () => {
  const name = 'hooks/index.d.ts';
  const original = baseline[name].declaration;
  assert.doesNotThrow(() => assertCompatibleDeclaration(name, `${original}\n${pnpExports}`, original, pnpRequired));
  assert.throws(() => assertCompatibleDeclaration(name, `${original}\n${pnpExports}`, original, { ...pnpRequired, requireApprovedAdditions: true }), /Missing approved export/);
  assert.doesNotThrow(() => assertCompatibleDeclaration(name, `${original}\nexport * from './useSPFxSiteKeyValueStore';\n${pnpExports}`, original, { ...pnpRequired, requireApprovedAdditions: true }));
});

for (const module of pnpHookPaths) {
  test(`requires the finished PnP export ${module}`, () => {
    const original = baseline['hooks/index.d.ts'].declaration;
    assert.throws(() => assertCompatibleDeclaration('hooks/index.d.ts', `${original}\n${pnpExports.replace(`export * from '${module}';`, '')}`, original, pnpRequired), /Missing approved PnP list/);
    assert.doesNotThrow(() => assertCompatibleDeclaration('hooks/index.d.ts', `${original}\nexport * from '${module}';`, original));
  });
  for (const [label, exportText] of [
    ['duplicate', `export * from '${module}';`],
    ['wrong path', `export * from '${module}.js';`],
    ['named', `export { Example } from '${module}';`],
    ['namespace', `export * as Example from '${module}';`],
    ['type-only', `export type * from '${module}';`],
    ['attributes', `export * from '${module}' with { type: 'json' };`],
  ]) {
    test(`rejects PnP ${label} export ${module}`, () => {
      const original = baseline['hooks/index.d.ts'].declaration;
      assert.throws(() => assertCompatibleDeclaration('hooks/index.d.ts', `${original}\n${pnpExports}\n${exportText}`, original));
    });
  }
  test(`rejects PnP export ${module} outside the hook barrel`, () => {
    for (const name of ['index.d.ts', 'services/index.d.ts', pnpServiceName, 'hooks/useSPFxPnPList.d.ts']) {
      assert.throws(() => assertCompatibleDeclaration(name, `${baseline[name].declaration}\nexport * from '${module}';`, baseline[name].declaration));
    }
  });
}

test('historical title hook cannot widen even when PnP additions are required', () => {
  const name = 'hooks/useSPFxPnPList.d.ts';
  const original = baseline[name].declaration;
  assert.throws(() => assertCompatibleDeclaration(name, original.replace('listTitle: string', 'listTitle: string | SPFxPnPListSelector'), original, pnpRequired));
});

for (const missing of [...pnpHookPaths, 'selector', 'factory']) {
  test(`integrated verify-api requires finished PnP ${missing} independently of site store`, () => {
    const fs = require('node:fs');
    const path = require('node:path');
    const scriptPath = path.resolve(__dirname, '../scripts/verify-api.cjs');
    const localRequire = require('node:module').createRequire(scriptPath);
    const declarationsRoot = path.resolve(__dirname, '../packages/spfx-react-toolkit/lib') + path.sep;
    const fileBoundary = {
      ...fs,
      existsSync(file) {
        if (file.endsWith('/src/hooks/useSPFxSiteKeyValueStore.ts')) return false;
        if (file.endsWith('/src/hooks/useSPFxPnPListById.ts')) return true;
        return fs.existsSync(file);
      },
      readFileSync(file, options) {
        if (file.startsWith(declarationsRoot)) {
          const name = file.slice(declarationsRoot.length);
          if (name === 'hooks/index.d.ts') return `${baseline[name].declaration}\n${pnpExports.replace(`export * from '${missing}';`, '')}`;
          if (name === pnpServiceName) {
            if (missing === 'selector') return widenedService;
            if (missing === 'factory') return `${selectorDeclaration}\n${baseline[name].declaration}`;
            return approvedService;
          }
          if (baseline[name]) return baseline[name].declaration;
        }
        return fs.readFileSync(file, options);
      },
    };
    const verify = new Function('require', '__dirname', fs.readFileSync(scriptPath, 'utf8'));
    assert.throws(() => verify(key => key === 'node:fs' ? fileBoundary : localRequire(key), path.dirname(scriptPath)), /Missing approved PnP list/);
  });
}

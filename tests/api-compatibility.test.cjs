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

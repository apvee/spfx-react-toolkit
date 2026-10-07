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
        if (file.endsWith('/src/hooks/useSx.ts')) return false;
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
test('permits only the approved stable callback export in the hooks barrel', () => {
  const name = 'hooks/index.d.ts';
  const original = baseline[name].declaration;
  const addition = "export * from './useStableCallback';";
  assert.doesNotThrow(() => assertCompatibleDeclaration(name, `${original}\n${addition}`, original));
  assert.throws(() => assertCompatibleDeclaration(name, `${original}\n${addition}\n${addition}`, original), /Duplicate approved export/);
  assert.throws(() => assertCompatibleDeclaration(name, `${original}\nexport * from './useStableCallbackOther';`, original), /Declaration contract changed/);
  const root = baseline['index.d.ts'].declaration;
  assert.throws(() => assertCompatibleDeclaration('index.d.ts', `${root}\n${addition}`, root), /Declaration contract changed/);
});
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
        if (file.endsWith('/src/hooks/useSx.ts')) return false;
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

const styleBarrels = [
  ['hooks/index.d.ts', './useSx', './useSPFxContext'],
  ['helpers/index.d.ts', './styles', './spfx-storage.helpers'],
];
for (const [name, addition, historical] of styleBarrels) {
  const original = baseline[name].declaration;
  const exportAll = `export * from '${addition}';`;
  test(`permits exactly the approved style addition ${name}: ${addition}`, () => {
    assert.doesNotThrow(() => assertCompatibleDeclaration(name, `${original}\n${exportAll}`, original));
    assert.doesNotThrow(() => assertCompatibleDeclaration(name, `${original}\n${exportAll}`, original, { requireSxAdditions: true }));
  });
  test(`requires the finished style addition ${name}: ${addition}`, () => {
    assert.throws(() => assertCompatibleDeclaration(name, original, original, { requireSxAdditions: true }), /Missing approved style export/);
  });
  for (const [label, current] of [
    ['duplicate', `${original}\n${exportAll}\n${exportAll}`],
    ['wrong path', `${original}\nexport * from '${addition}.js';`],
    ['named', `${original}\nexport { Example } from '${addition}';`],
    ['namespace', `${original}\nexport * as Example from '${addition}';`],
    ['type-only', `${original}\nexport type * from '${addition}';`],
    ['attributes', `${original}\nexport * from '${addition}' with { type: 'json' };`],
    ['historical removal', `${original.replace(`export * from '${historical}';`, '')}\n${exportAll}`],
    ['unrelated addition', `${original}\n${exportAll}\nexport type Extra = number;`],
  ]) {
    test(`rejects style ${label} in ${name}`, () => {
      if (label === 'historical removal') assert.ok(original.includes(`export * from '${historical}';`));
      assert.throws(() => assertCompatibleDeclaration(name, current, original));
    });
  }
  test(`rejects the style path ${addition} in every other historical module`, () => {
    for (const [otherName, snapshot] of Object.entries(baseline)) {
      if (otherName === name) continue;
      assert.throws(() => assertCompatibleDeclaration(otherName, `${snapshot.declaration}\n${exportAll}`, snapshot.declaration));
    }
  });
}

test('style completion checks preserve previous callback, site-store and PnP approvals', () => {
  const name = 'hooks/index.d.ts';
  const original = baseline[name].declaration;
  const all = `${original}\n${pnpExports}\nexport * from './useSPFxSiteKeyValueStore';\nexport * from './useStableCallback';\nexport * from './useSx';`;
  assert.doesNotThrow(() => assertCompatibleDeclaration(name, all, original, {
    requireApprovedAdditions: true, requirePnPListAdditions: true, requireSxAdditions: true,
  }));
});

const originalPackage = require('./fixtures/package-baseline.json');
const currentPackage = require('../packages/spfx-react-toolkit/package.json');
const { assertCompatiblePackageContract } = require('../scripts/api-compatibility.helpers.cjs');
test('permits only the approved package peer additions and previous dependency cleanup', () => {
  assert.doesNotThrow(() => assertCompatiblePackageContract(currentPackage, originalPackage));
});
for (const [label, mutate] of [
  ['historical peer range', manifest => { manifest.peerDependencies.react = '^18.0.0'; }],
  ['new Griffel core peer range', manifest => { manifest.peerDependencies['@griffel/core'] = '^2.0.0'; }],
  ['new Griffel React peer removal', manifest => { delete manifest.peerDependencies['@griffel/react']; }],
  ['new shared-context peer range', manifest => { manifest.peerDependencies['@fluentui/react-shared-contexts'] = '*'; }],
  ['unexpected peer', manifest => { manifest.peerDependencies['@fluentui/react-provider'] = '^9.22.8'; }],
  ['private Griffel runtime dependency', manifest => { manifest.dependencies = { '@griffel/core': '1.19.2' }; }],
  ['optional mandatory peer', manifest => { manifest.peerDependenciesMeta = { '@griffel/core': { optional: true } }; }],
  ['package allowlist', manifest => { manifest.files.push('src/**/*'); }],
  ['deep-import restriction', manifest => { manifest.exports = { '.': './lib/index.js' }; }],
  ['entry point', manifest => { manifest.main = 'src/index.ts'; }],
]) {
  test(`rejects unapproved package ${label}`, () => {
    const candidate = JSON.parse(JSON.stringify(currentPackage));
    mutate(candidate);
    assert.throws(() => assertCompatiblePackageContract(candidate, originalPackage));
  });
}

for (const [missingName] of styleBarrels) {
  test(`integrated verify-api requires finished style ${missingName}`, () => {
    const fs = require('node:fs');
    const path = require('node:path');
    const scriptPath = path.resolve(__dirname, '../scripts/verify-api.cjs');
    const localRequire = require('node:module').createRequire(scriptPath);
    const declarationsRoot = path.resolve(__dirname, '../packages/spfx-react-toolkit/lib') + path.sep;
    const fileBoundary = {
      ...fs,
      existsSync(file) {
        if (file.endsWith('/src/hooks/useSPFxSiteKeyValueStore.ts') || file.endsWith('/src/hooks/useSPFxPnPListById.ts')) return false;
        if (file.endsWith('/src/hooks/useSx.ts')) return true;
        return fs.existsSync(file);
      },
      readFileSync(file, options) {
        if (file.startsWith(declarationsRoot)) {
          const name = file.slice(declarationsRoot.length);
          if (baseline[name]) {
            const addition = styleBarrels.find(([barrel]) => barrel === name)?.[1];
            return baseline[name].declaration + (addition && name !== missingName ? `\nexport * from '${addition}';` : '');
          }
        }
        return fs.readFileSync(file, options);
      },
    };
    const verify = new Function('require', '__dirname', fs.readFileSync(scriptPath, 'utf8'));
    assert.throws(() => verify(key => key === 'node:fs' ? fileBoundary : localRequire(key), path.dirname(scriptPath)), /Missing approved style export/);
  });
}

for (const marker of [
  'The build failed because a task wrote output to stderr.',
  'Exiting with exit code: 1',
  '\u001b[31mError - [webpack] compilation failed\u001b[0m',
]) {
  test(`package gate rejects explicit subprocess failure with wrapper status zero: ${marker}`, () => {
    const fs = require('node:fs');
    const path = require('node:path');
    const scriptPath = path.resolve(__dirname, '../scripts/verify-package.cjs');
    const localRequire = require('node:module').createRequire(scriptPath);
    const childBoundary = {
      spawnSync: () => ({ status: 0, signal: null, stdout: marker, stderr: '' }),
      execFileSync: () => { throw new Error('Unexpected pack after failed build'); },
    };
    const verify = new Function('require', '__dirname', 'process', fs.readFileSync(scriptPath, 'utf8'));
    const processBoundary = { env: {}, stdout: { write() {} }, stderr: { write() {} } };
    assert.throws(() => verify(key => key === 'node:child_process' ? childBoundary : localRequire(key), path.dirname(scriptPath), processBoundary), /reported subprocess failure despite wrapper status 0/);
  });
}

test('package gate keeps visible data-age warnings distinct from actual subprocess failure', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const scriptPath = path.resolve(__dirname, '../scripts/verify-package.cjs');
  const localRequire = require('node:module').createRequire(scriptPath);
  const childBoundary = {
    spawnSync: () => ({ status: 0, signal: null, stdout: '', stderr: '[baseline-browser-mapping] The data in this module is over two months old.' }),
    execFileSync: () => { throw new Error('Build accepted; pack reached'); },
  };
  let visibleWarning = '';
  const verify = new Function('require', '__dirname', 'process', fs.readFileSync(scriptPath, 'utf8'));
  const processBoundary = { env: {}, stdout: { write() {} }, stderr: { write(value) { visibleWarning += value; } } };
  assert.throws(() => verify(key => key === 'node:child_process' ? childBoundary : localRequire(key), path.dirname(scriptPath), processBoundary), /Build accepted; pack reached/);
  assert.match(visibleWarning, /baseline-browser-mapping/);
});

test('metadata preparation leaves malformed configuration and evaluation errors fatal', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const script = fs.readFileSync(path.resolve(__dirname, '../scripts/build-metadata-preload.cjs'), 'utf8');
  function invoke(resolveError, evaluationError) {
    const fakeRequire = () => { if (evaluationError) throw evaluationError; return () => []; };
    fakeRequire.resolve = name => { if (resolveError) throw resolveError; return name; };
    const requireBoundary = key => key === 'node:module' ? { createRequire: () => fakeRequire } : require(key);
    const mockModule = { exports: {} };
    new Function('require', 'module', script)(requireBoundary, mockModule);
    return mockModule.exports.prepareMetadata();
  }
  assert.throws(() => invoke(Object.assign(new Error('invalid package metadata'), { code: 'ERR_INVALID_PACKAGE_CONFIG' })), /invalid package metadata/);
  assert.throws(() => invoke(undefined, Object.assign(new Error('metadata evaluation failed'), { code: 'MODULE_NOT_FOUND' })), /metadata evaluation failed/);
});

test('Gulp metadata guard prepares only the parent Gulp entry point', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const script = fs.readFileSync(path.resolve(__dirname, '../scripts/gulp-build-metadata-preload.cjs'), 'utf8');
  for (const [entry, expected] of [
    ['/consumer/node_modules/.bin/gulp', 1],
    ['/consumer/node_modules/gulp/bin/gulp.js', 1],
    ['/runtime/npm/bin/npm-cli.js', 0],
    ['/consumer/node_modules/typescript/bin/tsc', 0],
    ['/consumer/node_modules/eslint/bin/eslint.js', 0],
    ['/consumer/task-runner.js', 0],
    [undefined, 0],
  ]) {
    let prepared = 0;
    let installed = 0;
    const requireBoundary = key => {
      if (key === './build-metadata-preload.cjs') return { installMetadataAdvisoryRouting: () => { installed += 1; }, prepareMetadata: () => { prepared += 1; } };
      return require(key);
    };
    new Function('require', 'process', script)(requireBoundary, { argv: ['node', entry] });
    assert.equal(prepared, expected, `Unexpected metadata preparation for ${entry}`);
    assert.equal(installed, 1, `Workers must install exact advisory routing for ${entry}`);
  }
});

test('exact locked metadata advisories stay visible while every other diagnostic retains its channel', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const script = fs.readFileSync(path.resolve(__dirname, '../scripts/build-metadata-preload.cjs'), 'utf8');
  const baselineWarning = '[baseline-browser-mapping] The data in this module is over two months old.  To ensure accurate Baseline data, please update: `npm i baseline-browser-mapping@latest -D`';
  const browserWarning = months => `Browserslist: browsers data (caniuse-lite) is ${months} ${months === 1 ? 'month' : 'months'} old. Please run:\n  npx update-browserslist-db@latest\n  Why you should do it regularly: https://github.com/browserslist/update-db#readme`;
  const visible = [];
  const stderr = [];
  const originalError = (...args) => stderr.push(args);
  const consoleBoundary = { warn: (...args) => stderr.push(args), error: originalError };
  const stdoutBoundary = { write: text => visible.push(text) };
  const stderrBoundary = { write: text => stderr.push([text]) };
  const originalStderrWrite = stderrBoundary.write;
  const moduleBoundary = { exports: {} };
  let metadataResolutions = 0;
  const requireBoundary = key => key === 'node:module'
    ? { createRequire: () => { metadataResolutions += 1; throw new Error('Unexpected eager metadata import'); } } : require(key);
  new Function('require', 'module', 'console', 'process', script)(requireBoundary, moduleBoundary,
    consoleBoundary, { stdout: stdoutBoundary, stderr: stderrBoundary });
  moduleBoundary.exports.installMetadataAdvisoryRouting();
  moduleBoundary.exports.installMetadataAdvisoryRouting();
  assert.equal(metadataResolutions, 0, 'Workers must not import/query metadata eagerly');
  for (const message of [baselineWarning, browserWarning(1), browserWarning(12), browserWarning(121)]) consoleBoundary.warn(message);
  assert.equal(visible.length, 4);
  assert.equal(stderr.length, 0);
  assert.equal(visible[0], `[verification metadata advisory] ${baselineWarning}\n`);
  assert.equal(visible[2], `[verification metadata advisory] ${browserWarning(12)}\n`);
  const unknown = [
    ['Unknown compiler warning'],
    [baselineWarning + ' Additional failure'],
    ['prefix ' + baselineWarning],
    [browserWarning(12) + '\nAdditional error'],
    [browserWarning(12) + '\n'],
    [browserWarning(12).replace('update-browserslist-db@latest', 'other-command')],
    [baselineWarning, 'extra argument'],
    [{ message: baselineWarning }],
  ];
  for (const args of unknown) consoleBoundary.warn(...args);
  assert.deepEqual(stderr, unknown);
  assert.equal(visible.length, 4, 'Unknown diagnostics must not move to stdout');
  assert.equal(consoleBoundary.error, originalError, 'console.error must remain untouched');
  assert.equal(stderrBoundary.write, originalStderrWrite, 'Raw stderr writer must remain untouched');
  consoleBoundary.error('Compiler error');
  stderrBoundary.write('Raw subprocess stderr');
  assert.deepEqual(stderr.slice(-2), [['Compiler error'], ['Raw subprocess stderr']]);
});

test('advisory routing preserves fatal behavior of an unknown warning monitor', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const script = fs.readFileSync(path.resolve(__dirname, '../scripts/build-metadata-preload.cjs'), 'utf8');
  const moduleBoundary = { exports: {} };
  const consoleBoundary = { warn() { throw new Error('Strict warning monitor rejected the diagnostic'); } };
  new Function('require', 'module', 'console', 'process', script)(require, moduleBoundary, consoleBoundary,
    { stdout: { write() { throw new Error('Unknown warning was routed'); } } });
  moduleBoundary.exports.installMetadataAdvisoryRouting();
  assert.throws(() => consoleBoundary.warn('Unexpected compiler warning'), /Strict warning monitor rejected/);
});

function withMetadataPackageFixture(packageContents, moduleSource, check) {
  const fs = require('node:fs');
  const path = require('node:path');
  const os = require('node:os');
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'sx-metadata-package-'));
  const packageFile = path.join(directory, 'node_modules/baseline-browser-mapping/package.json');
  try {
    fs.writeFileSync(path.join(directory, 'package.json'), JSON.stringify({ name: 'metadata-test-host', private: true }));
    if (packageContents !== undefined) {
      fs.mkdirSync(path.dirname(packageFile), { recursive: true });
      fs.writeFileSync(packageFile, packageContents);
      if (moduleSource !== undefined) fs.writeFileSync(path.join(path.dirname(packageFile), 'index.js'), moduleSource);
    }
    const source = fs.readFileSync(path.resolve(__dirname, '../scripts/build-metadata-preload.cjs'), 'utf8');
    const moduleBoundary = { exports: {} };
    new Function('require', 'module', 'process', source)(require, moduleBoundary, { cwd: () => directory });
    check(moduleBoundary.exports.prepareMetadata, { directory, packageFile });
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
}

test('real metadata package lookup tolerates genuinely absent requested packages', () => {
  withMetadataPackageFixture(undefined, undefined, (prepare, { directory }) => {
    const path = require('node:path');
    const fromHost = require('node:module').createRequire(path.join(directory, 'package.json'));
    for (const name of ['baseline-browser-mapping', 'browserslist']) {
      assert.throws(() => fromHost.resolve(name), error => error.code === 'MODULE_NOT_FOUND' && error.message.startsWith(`Cannot find module '${name}'\n`));
    }
    assert.doesNotThrow(prepare);
  });
});

test('real installed metadata package with a missing main remains fatal', () => {
  withMetadataPackageFixture(JSON.stringify({ name: 'baseline-browser-mapping', main: 'missing.js' }), undefined,
    (prepare, { packageFile }) => {
      assert.throws(prepare, error => error.code === 'MODULE_NOT_FOUND' && error.path === packageFile && error.message.includes('missing.js'));
    });
});

test('real installed metadata package with a missing export remains fatal', () => {
  withMetadataPackageFixture(JSON.stringify({ name: 'baseline-browser-mapping', exports: './missing.js' }), undefined,
    prepare => assert.throws(prepare, error => error.code === 'MODULE_NOT_FOUND' && error.message.includes('missing.js')));
});

test('real installed metadata package with malformed configuration remains fatal', () => {
  withMetadataPackageFixture('{ malformed JSON', undefined,
    prepare => assert.throws(prepare, error => error.code === 'ERR_INVALID_PACKAGE_CONFIG'));
});

test('real metadata module evaluation errors remain fatal', () => {
  withMetadataPackageFixture(JSON.stringify({ name: 'baseline-browser-mapping', main: 'index.js' }),
    "require('missing-metadata-evaluation-dependency');",
    prepare => assert.throws(prepare, error => error.code === 'MODULE_NOT_FOUND' && error.message.includes('missing-metadata-evaluation-dependency')));
});


test('real installed metadata package without a default entry remains fatal', () => {
  withMetadataPackageFixture(JSON.stringify({ name: 'baseline-browser-mapping' }), undefined,
    prepare => assert.throws(prepare, error => error.code === 'MODULE_NOT_FOUND' && error.message.startsWith("Cannot find module 'baseline-browser-mapping'\n")));
});

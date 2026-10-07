// A real, external SPFx consumer: no workspace symlinks, aliases or shared node_modules.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {execFileSync, spawnSync} = require('node:child_process');
const ts = require('typescript');
const { collectPublicSurface } = require('./public-surface.helpers.cjs');
const {createRequire} = require('node:module');
const root = path.resolve(__dirname, '..');
const library = path.join(root, 'packages/spfx-react-toolkit');
const app = path.join(root, 'apps/spfx-react-toolkit-test');
const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'spfx-tarball-consumer-'));
const evidence = process.env.SPFX_PACKAGE_EVIDENCE
  ? path.resolve(process.env.SPFX_PACKAGE_EVIDENCE) : path.join(temp, 'evidence');
fs.mkdirSync(evidence, { recursive: true });
const commandResults = [];
const stripAnsi = text => text.replace(/\x1b\[[0-?]*[ -/]*[@-~]/g, '');
function run(args, cwd, { prepareMetadata = false } = {}) {
  let command = 'npm';
  const env = { ...process.env };
  let commandArgs = args;
  if (prepareMetadata) {
    // The guarded preload prepares metadata only in Gulp and installs exact,
    // visible advisory routing in inherited compiler/minifier workers.
    const script = JSON.parse(fs.readFileSync(path.join(cwd, 'package.json'), 'utf8')).scripts[args[1]];
    const gulpTasks = { 'bundle:ship': ['bundle', '--ship'], 'package:solution': ['package-solution', '--ship'] };
    assert.ok(args[0] === 'run' && gulpTasks[args[1]], 'Metadata preparation requires a known strict Gulp task');
    assert.equal(script, 'gulp ' + gulpTasks[args[1]].join(' '), 'Consumer Gulp script changed; inspect before preparing');
    command = process.execPath;
    const preload = path.join(root, 'scripts/gulp-build-metadata-preload.cjs');
    env.NODE_OPTIONS = [env.NODE_OPTIONS, '--require', JSON.stringify(preload)].filter(Boolean).join(' ');
    commandArgs = ['--require', preload,
      createRequire(path.join(cwd, 'package.json')).resolve('gulp/bin/gulp.js'), ...gulpTasks[args[1]]];
  }
  const result = spawnSync(command, commandArgs, { cwd, env, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
  const stdout = result.stdout ?? '';
  const stderr = result.stderr ?? '';
  process.stdout.write(stdout);
  process.stderr.write(stderr);
  const output = stripAnsi(stdout + '\n' + stderr);
  const failureMarkers = output.match(/The build failed[^\n]*|Exiting with exit code:\s*[1-9]\d*[^\n]*|Error - [^\n]*/g) ?? [];
  const warnings = output.split('\n').filter(line => /\bwarn(?:ing)?\b|Browserslist:|\[baseline-browser-mapping\]/i.test(line));
  const label = `${String(commandResults.length + 1).padStart(2, '0')}-${args.join('-').replace(/[^a-zA-Z0-9-]/g, '')}`;
  fs.writeFileSync(path.join(evidence, label + '.stdout.log'), stdout);
  fs.writeFileSync(path.join(evidence, label + '.stderr.log'), stderr);
  commandResults.push({ args, command, commandArgs, cwd, prepareMetadata, metadataAdvisoryRouting: prepareMetadata, status: result.status, signal: result.signal,
    error: result.error?.message, failureMarkers, warnings });
  fs.writeFileSync(path.join(evidence, 'commands.json'), JSON.stringify(commandResults, null, 2));
  assert.ifError(result.error);
  assert.equal(result.status, 0, `npm ${args.join(' ')} failed; evidence: ${evidence}`);
  assert.deepEqual(failureMarkers, [], `npm ${args.join(' ')} reported subprocess failure despite wrapper status ${result.status}; evidence: ${evidence}`);
  return stdout;
}

function assertPackedImports(packedPaths) {
  for (const file of packedPaths) {
    const extension = file.endsWith('.d.ts') ? '.d.ts' : file.endsWith('.js') ? '.js' : undefined;
    if (!extension) continue;
    const source = ts.createSourceFile(file, fs.readFileSync(path.join(library, file), 'utf8'), ts.ScriptTarget.Latest, true);
    function check(specifier) {
      if (!specifier.startsWith('.')) return;
      const base = path.posix.normalize(path.posix.join(path.posix.dirname(file), specifier.replace(/\.js$/, '')));
      const candidates = [base + extension, base + '/index' + extension];
      assert.ok(candidates.some(candidate => packedPaths.has(candidate)),
        `Missing packed ${extension} dependency for ${file}: ${specifier}`);
    }
    function visit(node) {
      if ((ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) && node.moduleSpecifier && ts.isStringLiteral(node.moduleSpecifier)) check(node.moduleSpecifier.text);
      if (ts.isImportTypeNode(node) && ts.isLiteralTypeNode(node.argument) && ts.isStringLiteral(node.argument.literal)) check(node.argument.literal.text);
      if (ts.isCallExpression(node) && node.arguments.length && ts.isStringLiteral(node.arguments[0]) &&
          (node.expression.kind === ts.SyntaxKind.ImportKeyword || (ts.isIdentifier(node.expression) && node.expression.text === 'require'))) check(node.arguments[0].text);
      ts.forEachChild(node, visit);
    }
    visit(source);
  }
}

function assertEmittedStyleJSDoc(packedPaths) {
  const styleEntry = path.join(library, 'src/helpers/styles/index.ts');
  const surface = [...collectPublicSurface(styleEntry), ...collectPublicSurface(path.join(library, 'src/hooks/useSx.ts'))];
  function findDeclaration(source, name) {
    return source.statements.find(node => node.name?.text === name ||
      (ts.isVariableStatement(node) && node.declarationList.declarations.some(declaration => declaration.name.text === name)));
  }
  const normalizeDocs = (node, source) => (node.jsDoc ?? []).map(doc => doc.getText(source).replace(/\s+/g, ' ').trim()).join(' ');
  for (const declaration of surface) {
    const relative = path.relative(path.join(library, 'src'), declaration.filePath).replace(/\.ts$/, '.d.ts');
    const packedFile = 'lib/' + relative.split(path.sep).join('/');
    assert.ok(packedPaths.has(packedFile), `Missing documented style declaration: ${packedFile}`);
    const emittedSource = ts.createSourceFile(packedFile, fs.readFileSync(path.join(library, packedFile), 'utf8'), ts.ScriptTarget.Latest, true);
    const originalSource = ts.createSourceFile(declaration.filePath, fs.readFileSync(declaration.filePath, 'utf8'), ts.ScriptTarget.Latest, true);
    const name = declaration.name.split('.').pop();
    const emitted = findDeclaration(emittedSource, name);
    const original = findDeclaration(originalSource, name);
    assert.ok(emitted && original, `Missing emitted declaration: ${declaration.name}`);
    assert.ok(normalizeDocs(emitted, emittedSource), `Missing emitted JSDoc: ${declaration.name}`);
    assert.equal(normalizeDocs(emitted, emittedSource), normalizeDocs(original, originalSource), `JSDoc changed during emission: ${declaration.name}`);
    for (const member of declaration.members) {
      const memberName = member.name.split('.').pop();
      const originalMember = original.members.find(node => node.name?.getText(originalSource) === memberName);
      const emittedMember = emitted.members.find(node => node.name?.getText(emittedSource) === memberName);
      assert.ok(emittedMember && normalizeDocs(emittedMember, emittedSource), `Missing emitted member JSDoc: ${member.name}`);
      assert.equal(normalizeDocs(emittedMember, emittedSource), normalizeDocs(originalMember, originalSource), `Member JSDoc changed during emission: ${member.name}`);
    }
  }
  fs.writeFileSync(path.join(evidence, 'style-declarations.json'), JSON.stringify({ declarations: surface.map(item => item.name), count: surface.length }, null, 2));
  return surface.length;
}
try {
  run(['run','build:library'],root);
  const packed = JSON.parse(execFileSync('npm',['pack','--json','--pack-destination',temp],{cwd:library,encoding:'utf8'}))[0];
  assert.equal(packed.name,'@apvee/spfx-react-toolkit');
  for (const file of packed.files) {
    assert.ok(/^(package\.json|README\.md|LICENSE|lib\/(index\.[^/]+|(core|hooks|services|helpers|utils)\/.+))$/.test(file.path), `Unexpected package file: ${file.path}`);
  }
  const packedPaths = new Set(packed.files.map(file => file.path));
  fs.writeFileSync(path.join(evidence, 'pack.json'), JSON.stringify(packed, null, 2));
  assertPackedImports(packedPaths);
  const documentedDeclarations = assertEmittedStyleJSDoc(packedPaths);
  // Both the public entry points and their persistence/client dependencies must ship.
  for (const module of [
    'hooks/useSx',
    'helpers/styles/index',
    'helpers/styles/types',
    'hooks/useStableCallback',
    'hooks/useSPFxSiteKeyValueStore',
    'services/spfx-site-key-value-store.service',
    'services/spfx-list-key-value-store.internal',
    'services/spfx-site-key-value-store.internal',
    'helpers/spfx-tenant-value.helpers',
    'hooks/useSPFxSPHttpClient',
    'hooks/useSPFxPageContext',
    'hooks/useSPFxServiceScope',
    'hooks/useAsyncInvoke.internal'
  ]) {
    for (const extension of ['js', 'd.ts']) {
      const file = `lib/${module}.${extension}`;
      assert.ok(packedPaths.has(file), `Missing site store package dependency: ${file}`);
    }
  }
  // The selector resolver and shared hook runtime are required at runtime.
  for (const module of [
    'services/spfx-pnp-list.service',
    'services/spfx-pnp-list-target.internal',
    'hooks/useSPFxPnPList',
    'hooks/useSPFxPnPList.internal',
    'hooks/useSPFxPnPListById',
    'hooks/useSPFxPnPListByUrl',
    'hooks/useSPFxPnPListByPath',
    'hooks/useSPFxPnPContext'
  ]) {
    for (const extension of ['js', 'd.ts']) {
      const file = `lib/${module}.${extension}`;
      assert.ok(packedPaths.has(file), `Missing PnP list package dependency: ${file}`);
    }
  }
  const consumer = path.join(temp,'consumer'); fs.mkdirSync(consumer);
  for (const name of ['src','config','sharepoint','teams','gulpfile.js','tsconfig.json','.eslintrc.js']) fs.cpSync(path.join(app,name),path.join(consumer,name),{recursive:true});
  const manifest = JSON.parse(fs.readFileSync(path.join(app,'package.json'),'utf8'));
  manifest.name='isolated-spfx-toolkit-consumer';
  assert.equal(manifest.workspaces, undefined, 'External consumer cannot contain workspaces');
  manifest.dependencies['@apvee/spfx-react-toolkit']='file:'+path.join(temp,packed.filename);
  fs.writeFileSync(path.join(consumer,'package.json'),JSON.stringify(manifest,null,2));
  // Seed locked transitive versions while npm replaces workspace identities with the tarball.
  const lock=JSON.parse(fs.readFileSync(path.join(root,'package-lock.json'),'utf8'));
  lock.name=manifest.name;lock.packages['']={name:manifest.name,version:manifest.version,dependencies:manifest.dependencies,devDependencies:manifest.devDependencies,engines:manifest.engines};
  delete lock.packages['apps/spfx-react-toolkit-test'];delete lock.packages['packages/spfx-react-toolkit'];
  for (const [key,value] of Object.entries(lock.packages)) if(value.link) delete lock.packages[key];
  fs.writeFileSync(path.join(consumer,'package-lock.json'),JSON.stringify(lock,null,2));
  const tsconfig=JSON.parse(fs.readFileSync(path.join(consumer,'tsconfig.json'),'utf8'));
  assert.equal(tsconfig.compilerOptions.paths, undefined, 'External consumer cannot use workspace path aliases');
  tsconfig.compilerOptions.typeRoots=['./node_modules/@types','./node_modules/@microsoft'];
  fs.writeFileSync(path.join(consumer,'tsconfig.json'),JSON.stringify(tsconfig,null,2));
  fs.writeFileSync(path.join(consumer,'src/consumer-compatibility.ts'),`import { SPFxWebPartProvider, useSPFxProperties, createSPFxPnPListService, createScopedSPFxStorageKey } from '@apvee/spfx-react-toolkit';\nimport { useSPFxContext } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxContext';\nimport type { SPFxProviderProps } from '@apvee/spfx-react-toolkit/lib/core/types';\nexport const contracts = { SPFxWebPartProvider, useSPFxProperties, createSPFxPnPListService, createScopedSPFxStorageKey, useSPFxContext };\nexport type ConsumerProviderProps = SPFxProviderProps;\nimport { useSPFxEnvironmentInfo, useSPFxLocaleInfo, useSPFxLocalStorage, useSPFxPerformance, SPFxStorageHook, SPFxPerfResult } from '@apvee/spfx-react-toolkit';\nexport function useDocumentedAPIContracts(): { isTeams: boolean; locale: string; uiLocale: string; stored: boolean; save: SPFxStorageHook<{ enabled: boolean }>['setValue']; remove: () => void; timed: () => Promise<SPFxPerfResult<boolean>> } {\n  const environment = useSPFxEnvironmentInfo();\n  const locale = useSPFxLocaleInfo();\n  const storage = useSPFxLocalStorage('prefs', { enabled: true });\n  const performance = useSPFxPerformance();\n  return { isTeams: environment.isTeams, locale: locale.locale, uiLocale: locale.uiLocale, stored: storage.value.enabled, save: storage.setValue, remove: storage.remove, timed: () => performance.time('probe', () => storage.value.enabled) };\n}\n`);
  // Compile-only contracts: these exported functions are never invoked by this gate.
  fs.writeFileSync(path.join(consumer,'src/site-store-compatibility.ts'),`import type { SPHttpClient } from '@microsoft/sp-http';
import {
  useSPFxSiteKeyValueStore, createSPFxSiteKeyValueStoreService,
  SPFxSiteKeyValueStoreItem, SPFxSiteKeyValueStoreResult,
  SPFxSiteKeyValueStoreServiceItem, SPFxSiteKeyValueStoreService,
  useSPFxTenantKeyValueStore, createSPFxTenantKeyValueStoreService,
  SPFxTenantKeyValueStoreItem, SPFxTenantKeyValueStoreResult,
  SPFxTenantKeyValueStoreServiceItem, SPFxTenantKeyValueStoreService
} from '@apvee/spfx-react-toolkit';
import {
  useSPFxSiteKeyValueStore as useDeepSiteStore,
  SPFxSiteKeyValueStoreItem as DeepSiteItem,
  SPFxSiteKeyValueStoreResult as DeepSiteResult
} from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxSiteKeyValueStore';
import {
  createSPFxSiteKeyValueStoreService as createDeepSiteService,
  SPFxSiteKeyValueStoreServiceItem as DeepSiteServiceItem,
  SPFxSiteKeyValueStoreService as DeepSiteService
} from '@apvee/spfx-react-toolkit/lib/services/spfx-site-key-value-store.service';
import {
  useSPFxTenantKeyValueStore as useDeepTenantStore,
  SPFxTenantKeyValueStoreItem as DeepTenantItem,
  SPFxTenantKeyValueStoreResult as DeepTenantResult
} from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxTenantKeyValueStore';
import {
  createSPFxTenantKeyValueStoreService as createDeepTenantService,
  SPFxTenantKeyValueStoreServiceItem as DeepTenantServiceItem,
  SPFxTenantKeyValueStoreService as DeepTenantService
} from '@apvee/spfx-react-toolkit/lib/services/spfx-tenant-key-value-store.service';
import { serializeValue, deserializeValue, escapeODataValue } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxTenantKeyValueStore.serialization.internal';
import { getListApiUrl, TENANT_KEY_VALUE_STORE_LIST_TITLE } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxTenantKeyValueStore.sharepoint.internal';

type Preference = { enabled: boolean };
export const storeContracts = {
  useSPFxSiteKeyValueStore, createSPFxSiteKeyValueStoreService, useDeepSiteStore, createDeepSiteService,
  useSPFxTenantKeyValueStore, createSPFxTenantKeyValueStoreService, useDeepTenantStore, createDeepTenantService,
  serializeValue, deserializeValue, escapeODataValue, getListApiUrl, TENANT_KEY_VALUE_STORE_LIST_TITLE
};

export function useStoreContracts(): { site: DeepSiteResult; tenant: DeepTenantResult; flags: boolean[]; errors: (Error | undefined)[] } {
  const site: SPFxSiteKeyValueStoreResult = useSPFxSiteKeyValueStore();
  const tenant: SPFxTenantKeyValueStoreResult = useSPFxTenantKeyValueStore();
  return { site, tenant, flags: [site.isLoading, site.isWriting, site.canWrite, site.isReady], errors: [site.error, site.writeError] };
}

export async function siteServiceContracts(client: SPHttpClient, url: string): Promise<DeepSiteServiceItem<Preference> | undefined> {
  const service: SPFxSiteKeyValueStoreService = createSPFxSiteKeyValueStoreService(client);
  const deepService: DeepSiteService = createDeepSiteService(client);
  const allowed: boolean = await deepService.canCurrentUserWrite(url);
  await service.ensureListReady(url);
  const item: SPFxSiteKeyValueStoreServiceItem<Preference> | undefined = await service.get<Preference>('prefs', url);
  const unknownItem: SPFxSiteKeyValueStoreServiceItem | undefined = await service.get('other', url);
  const items: SPFxSiteKeyValueStoreServiceItem<unknown>[] = await service.list(url);
  await service.save<Preference>('prefs', { enabled: allowed }, url, 'Description');
  await service.save('other', unknownItem ? unknownItem.value : items.length, url);
  await service.remove('other', url);
  return item;
}

export async function siteHookContracts(store: SPFxSiteKeyValueStoreResult): Promise<DeepSiteItem<Preference> | undefined> {
  const item: SPFxSiteKeyValueStoreItem<Preference> | undefined = await store.get<Preference>('prefs');
  const unknownItem: SPFxSiteKeyValueStoreItem | undefined = await store.get('other');
  const items: SPFxSiteKeyValueStoreItem<unknown>[] = await store.list();
  await store.save<Preference>('prefs', { enabled: true }, 'Description');
  await store.save('other', unknownItem ? unknownItem.value : items.length);
  await store.remove('other');
  return item;
}

export function readonlySiteItems(item: SPFxSiteKeyValueStoreItem<Preference>, serviceItem: SPFxSiteKeyValueStoreServiceItem<Preference>): readonly [string, Preference, string | undefined, number] {
  // @ts-expect-error Public hook item keys remain readonly.
  item.key = 'changed';
  // @ts-expect-error Public hook item values remain readonly.
  item.value = { enabled: false };
  // @ts-expect-error Public hook item descriptions remain readonly.
  item.description = undefined;
  // @ts-expect-error Public hook item IDs remain readonly.
  item.id = 0;
  // @ts-expect-error Public service item keys remain readonly.
  serviceItem.key = 'changed';
  // @ts-expect-error Public service item values remain readonly.
  serviceItem.value = { enabled: false };
  // @ts-expect-error Public service item descriptions remain readonly.
  serviceItem.description = undefined;
  // @ts-expect-error Public service item IDs remain readonly.
  serviceItem.id = 0;
  return [item.key, serviceItem.value, item.description, serviceItem.id];
}

export async function tenantServiceContracts(client: SPHttpClient, catalogUrl: string): Promise<DeepTenantServiceItem<Preference> | undefined> {
  const service: SPFxTenantKeyValueStoreService = createSPFxTenantKeyValueStoreService(client);
  const deepService: DeepTenantService = createDeepTenantService(client);
  await deepService.ensureListReady(catalogUrl);
  const item: SPFxTenantKeyValueStoreServiceItem<Preference> | undefined = await service.get<Preference>('prefs', catalogUrl);
  const items: SPFxTenantKeyValueStoreServiceItem<unknown>[] = await service.list(catalogUrl);
  await service.save<Preference>('prefs', { enabled: items.length > 0 }, catalogUrl, 'Description');
  await service.remove('prefs', catalogUrl);
  return item;
}

export async function tenantHookContracts(store: SPFxTenantKeyValueStoreResult): Promise<DeepTenantItem<Preference> | undefined> {
  const item: SPFxTenantKeyValueStoreItem<Preference> | undefined = await store.get<Preference>('prefs');
  const items: SPFxTenantKeyValueStoreItem<unknown>[] = await store.list();
  await store.save<Preference>('prefs', { enabled: items.length > 0 }, 'Description');
  await store.remove('prefs');
  return item;
}
`);
  fs.writeFileSync(path.join(consumer, 'src/pnp-list-compatibility.ts'), `import type { SPFI } from '@pnp/sp';
import {
  createSPFxPnPListService, useSPFxPnPList, useSPFxPnPListById,
  useSPFxPnPListByUrl, useSPFxPnPListByPath,
  SPFxPnPListSelector, SPFxPnPListService, SPFxPnPListQueryResult,
  SPFxPnPListBatchResult, SPFxPnPListInfo, PnPContextInfo
} from '@apvee/spfx-react-toolkit';
import {
  createSPFxPnPListService as createDeepService,
  SPFxPnPListSelector as DeepSelector,
  SPFxPnPListService as DeepService,
  SPFxPnPListQueryResult as DeepQueryResult,
  SPFxPnPListBatchResult as DeepBatchResult
} from '@apvee/spfx-react-toolkit/lib/services/spfx-pnp-list.service';
import { useSPFxPnPList as useDeepTitle, SPFxPnPListInfo as DeepInfo } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxPnPList';
import { useSPFxPnPListById as useDeepId } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxPnPListById';
import { useSPFxPnPListByUrl as useDeepUrl } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxPnPListByUrl';
import { useSPFxPnPListByPath as useDeepPath } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxPnPListByPath';

type Item = { Id: number; Title: string };
function consumeProbeValues(...values: unknown[]): number { return values.length; }
type LegacyFactory = <T = unknown>(sp: SPFI, listTitle: string, defaultPageSize?: number) => SPFxPnPListService<T>;
export const legacyFactories: LegacyFactory[] = [createSPFxPnPListService, createDeepService];

export async function serviceContracts(sp: SPFI): Promise<DeepQueryResult<Item>> {
  const targets: (string | SPFxPnPListSelector)[] = [
    'Tasks', { kind: 'title', title: 'Tasks' },
    { kind: 'id', id: '11111111-2222-3333-4444-555555555555' },
    { kind: 'url', serverRelativeUrl: '/sites/team/Lists/Tasks' },
    { kind: 'path', webRelativePath: 'Lists/Tasks' }
  ];
  const deepTargets: (string | DeepSelector)[] = targets;
  const services: DeepService<Item>[] = targets.map(target => createSPFxPnPListService<Item>(sp, target, 50));
  services.push(...deepTargets.map(target => createDeepService<Item>(sp, target, 50)));
  const unknownService: SPFxPnPListService<unknown> = createSPFxPnPListService(sp, 'Tasks');
  const unknownDeep: DeepService<unknown> = createDeepService(sp, { kind: 'path', webRelativePath: 'Tasks' });
  const unknownItem: unknown = await unknownService.getById(1);
  await unknownDeep.query();
  for (const service of services) {
    const queried: SPFxPnPListQueryResult<Item> = await service.query(items => items.select('Id', 'Title'), { pageSize: 50 });
    const more: DeepQueryResult<Item> = await service.loadMore(items => items.top(50), 50, queried.nextSkip);
    const item: Item = await service.getById(1);
    const id: number = await service.create({ Title: item.Title });
    const updated: void = await service.update(id, { Title: 'Changed' });
    const removed: void = await service.remove(id);
    const created: SPFxPnPListBatchResult<number[]> = await service.createBatch([{ Title: 'Batch' }]);
    const changed: DeepBatchResult<void> = await service.updateBatch([{ id, item: { Title: 'Batch changed' } }]);
    const deleted: SPFxPnPListBatchResult<void> = await service.removeBatch(created.value);
    const errors: unknown[] = deleted.errors;
    const error: Error | undefined = changed.summaryError;
    consumeProbeValues(unknownItem, more, updated, removed, errors, error);
    // @ts-expect-error Item IDs stay numeric.
    await service.getById('1');
    // @ts-expect-error Generic item fields retain their declared types.
    await service.create({ Title: 1 });
  }
  return services[0].query();
}

export function useListContracts(context: PnPContextInfo): DeepInfo<Item>[] {
  const results: SPFxPnPListInfo<Item>[] = [
    useSPFxPnPList<Item>('Tasks'), useSPFxPnPList<Item>('Tasks', { pageSize: 50 }, context),
    useDeepTitle<Item>('Tasks'), useDeepTitle<Item>('Tasks', { pageSize: 50 }, context),
    useSPFxPnPListById<Item>('guid'), useSPFxPnPListById<Item>('guid', { pageSize: 50 }, context),
    useDeepId<Item>('guid'), useDeepId<Item>('guid', { pageSize: 50 }, context),
    useSPFxPnPListByUrl<Item>('/Lists/Tasks'), useSPFxPnPListByUrl<Item>('/Lists/Tasks', { pageSize: 50 }, context),
    useDeepUrl<Item>('/Lists/Tasks'), useDeepUrl<Item>('/Lists/Tasks', { pageSize: 50 }, context),
    useSPFxPnPListByPath<Item>('Lists/Tasks'), useSPFxPnPListByPath<Item>('Lists/Tasks', { pageSize: 50 }, context),
    useDeepPath<Item>('Lists/Tasks'), useDeepPath<Item>('Lists/Tasks', { pageSize: 50 }, context)
  ];
  const unknownResults: SPFxPnPListInfo<unknown>[] = [
    useSPFxPnPList('Tasks'), useDeepTitle('Tasks'),
    useSPFxPnPListById('guid'), useDeepId('guid'),
    useSPFxPnPListByUrl('/Lists/Tasks'), useDeepUrl('/Lists/Tasks'),
    useSPFxPnPListByPath('Lists/Tasks'), useDeepPath('Lists/Tasks')
  ];
  consumeProbeValues(unknownResults);
  return results;
}

export async function hookResults(hook: SPFxPnPListInfo<Item>): Promise<Item[]> {
  const items: Item[] = hook.items;
  const queried: Item[] = await hook.query(query => query.top(50), { pageSize: 50 });
  const more: Item[] = await hook.loadMore();
  const item: Item | undefined = await hook.getById(1);
  const id: number = await hook.create({ Title: 'New' });
  const updated: void = await hook.update(id, { Title: 'Changed' });
  const removed: void = await hook.remove(id);
  const ids: number[] = await hook.createBatch([{ Title: 'Batch' }]);
  const batchUpdated: void = await hook.updateBatch([{ id, item: { Title: 'Changed' } }]);
  const batchRemoved: void = await hook.removeBatch(ids);
  const refreshed: void = await hook.refetch();
  hook.clearError();
  const flags: boolean[] = [hook.loading, hook.loadingMore, hook.hasMore, hook.isEmpty];
  const error: Error | undefined = hook.error;
  consumeProbeValues(items, more, item, updated, removed, batchUpdated, batchRemoved, refreshed, flags, error);
  // @ts-expect-error Hook generic results cannot lose their item fields.
  const strings: string[] = queried;
  // @ts-expect-error Hook item fields retain their declared types.
  await hook.update(id, { Title: false });
  consumeProbeValues(strings);
  return queried;
}

export function invalidTargets(sp: SPFI): void {
  for (const factory of [createSPFxPnPListService, createDeepService]) {
    // @ts-expect-error Unsupported selector discriminant.
    factory(sp, { kind: 'guid', id: 'guid' });
    // @ts-expect-error Title value must be a string.
    factory(sp, { kind: 'title', title: 1 });
    // @ts-expect-error GUID value must be a string.
    factory(sp, { kind: 'id', id: 1 });
    // @ts-expect-error URL value must be a string.
    factory(sp, { kind: 'url', serverRelativeUrl: false });
    // @ts-expect-error Path value must be a string.
    factory(sp, { kind: 'path', webRelativePath: null });
    // @ts-expect-error Selector requires its matching value.
    factory(sp, { kind: 'id' });
    // @ts-expect-error A title selector cannot carry GUID fields.
    factory(sp, { kind: 'title', title: 'Tasks', id: 'guid' });
    // @ts-expect-error A GUID selector cannot carry URL fields.
    factory(sp, { kind: 'id', id: 'guid', serverRelativeUrl: '/Lists/Tasks' });
    // @ts-expect-error A URL selector cannot carry path fields.
    factory(sp, { kind: 'url', serverRelativeUrl: '/Lists/Tasks', webRelativePath: 'Lists/Tasks' });
    // @ts-expect-error A path selector cannot carry title fields.
    factory(sp, { kind: 'path', webRelativePath: 'Lists/Tasks', title: 'Tasks' });
    // @ts-expect-error Targets cannot be numeric.
    factory(sp, 1);
  }
  for (const titleHook of [useSPFxPnPList, useDeepTitle]) {
    // @ts-expect-error Historical title hooks retain their string argument.
    titleHook({ kind: 'title', title: 'Tasks' });
  }
  for (const hook of [useSPFxPnPListById, useDeepId, useSPFxPnPListByUrl, useDeepUrl, useSPFxPnPListByPath, useDeepPath]) {
    // @ts-expect-error New hook inputs are strings.
    hook(1);
    // @ts-expect-error New hook inputs are values rather than selector objects.
    hook({ kind: 'id', id: 'guid' });
    // @ts-expect-error Hook page sizes remain numeric.
    hook('value', { pageSize: '50' });
    // @ts-expect-error Hook context must be a PnP context.
    hook('value', undefined, {});
  }
}

export function readonlySelectors(selector: Extract<SPFxPnPListSelector, { kind: 'title' }>, deep: Extract<DeepSelector, { kind: 'path' }>): void {
  // @ts-expect-error Selector discriminants remain readonly.
  selector.kind = 'title';
  // @ts-expect-error Selector values remain readonly.
  selector.title = 'Changed';
  // @ts-expect-error Deep selector values remain readonly.
  deep.webRelativePath = 'Changed';
}
`);
  fs.writeFileSync(path.join(consumer, 'src/stable-callback-compatibility.ts'), `import { useStableCallback } from '@apvee/spfx-react-toolkit';
import { useStableCallback as useDeepStableCallback } from '@apvee/spfx-react-toolkit/lib/hooks/useStableCallback';
import { useEventCallback } from '@fluentui/react-utilities';

export const aliasContract: typeof useEventCallback = useStableCallback;
export const deepAliasContract: typeof useEventCallback = useDeepStableCallback;

export function useCallbackContracts(): { text: string; pending: Promise<{ id: number }>; invalid: number } {
  const sync = useStableCallback((id: number, label: string) => label + id);
  const asyncCallback = useDeepStableCallback(async (id: number) => ({ id }));
  const text: string = sync(1, 'Item');
  const pending: Promise<{ id: number }> = asyncCallback(1);
  // @ts-expect-error Callback arguments retain their types.
  sync('1', 'Item');
  // @ts-expect-error The callback does not accept a dependency array.
  useStableCallback(() => 1, []);
  // @ts-expect-error Sync return types remain inferred.
  const invalid: number = sync(1, 'Item');
  return { text, pending, invalid };
}
`);
  fs.mkdirSync(path.join(consumer, 'compatibility'));
  for (const mode of ['public', 'deep']) {
    fs.copyFileSync(path.join(root, `tests/fixtures/sx-package-${mode}.ts`), path.join(consumer, `compatibility/sx-package-${mode}.ts`));
  }
  run(['install','--no-audit','--no-fund'],consumer);
  assert.equal(fs.lstatSync(path.join(consumer,'node_modules/@apvee/spfx-react-toolkit')).isSymbolicLink(),false);
  const hostRequire = createRequire(path.join(consumer, 'package.json'));
  const libraryRequire = createRequire(path.join(consumer, 'node_modules/@apvee/spfx-react-toolkit/package.json'));
  const sharedPackages = ['@griffel/core', '@griffel/react', '@fluentui/react-shared-contexts',
    '@fluentui/react-migration-v8-v9', '@fluentui/react-theme', '@fluentui/react-utilities'];
  const sharedPaths = {};
  for (const name of sharedPackages) {
    const hostPath = fs.realpathSync(hostRequire.resolve(name + '/package.json'));
    assert.equal(hostPath, fs.realpathSync(libraryRequire.resolve(name + '/package.json')), `${name} must be shared with the consumer`);
    assert.equal(fs.realpathSync(hostRequire.resolve(name)), fs.realpathSync(libraryRequire.resolve(name)), `${name} must share its runtime entry`);
    assert.ok(hostPath.startsWith(fs.realpathSync(path.join(consumer, 'node_modules')) + path.sep), `${name} must reside inside the isolated consumer`);
    const installedManifest = JSON.parse(fs.readFileSync(hostPath, 'utf8'));
    assert.equal(installedManifest.version, lock.packages['node_modules/' + name].version, `${name} changed from the locked toolchain`);
    sharedPaths[name] = hostPath;
  }
  for (const name of ['typescript', 'react', 'react-dom', '@microsoft/sp-build-web', '@microsoft/sp-core-library']) {
    const installed = hostRequire(name + '/package.json');
    assert.equal(installed.version, lock.packages['node_modules/' + name].version, `${name} changed from the locked consumer toolchain`);
  }
  // The actual renderer registry must also be the same object when resolved
  // through the toolkit, Griffel React and the host, not merely matching versions.
  const griffelReactRequire = createRequire(sharedPaths['@griffel/react']);
  const core = hostRequire('@griffel/core');
  assert.ok(core.DEFINITION_LOOKUP_TABLE, 'Missing Griffel definition registry');
  for (const from of [libraryRequire, griffelReactRequire]) {
    assert.equal(from('@griffel/core').DEFINITION_LOOKUP_TABLE, core.DEFINITION_LOOKUP_TABLE,
      'Duplicate Griffel definition registry');
  }
  const providerRequire = createRequire(hostRequire.resolve('@fluentui/react-provider/package.json'));
  assert.equal(fs.realpathSync(providerRequire.resolve('@fluentui/react-shared-contexts')),
    fs.realpathSync(hostRequire.resolve('@fluentui/react-shared-contexts')), 'FluentProvider must use the shared contexts');
  const packageTree = run(['ls', '--parseable', '--all', ...sharedPackages], consumer);
  const installedPaths = new Map(sharedPackages.map(name => [name, new Set()]));
  for (const installedPath of packageTree.trim().split('\n')) {
    const packageFile = path.join(installedPath, 'package.json');
    if (!fs.existsSync(packageFile)) continue;
    const installed = JSON.parse(fs.readFileSync(packageFile, 'utf8'));
    if (installedPaths.has(installed.name)) installedPaths.get(installed.name).add(fs.realpathSync(packageFile));
  }
  for (const [name, paths] of installedPaths) {
    assert.deepEqual([...paths], [sharedPaths[name]], `Duplicated installed shared package: ${name}`);
  }
  fs.writeFileSync(path.join(evidence, 'shared-packages.json'), JSON.stringify(sharedPaths, null, 2));
  run(['ls','react','react-dom','@microsoft/sp-core-library','@pnp/sp', ...sharedPackages,'tslib'],consumer);
  // SPFx generates Sass declarations during the bundle. A clean consumer has
  // no .scss.ts files yet, so generate them before the standalone typecheck.
  const consumerFailures = [];
  // Run independent checks even when a completed ship build reports the known
  // metadata warning failure. Every command retains its strict result, and the
  // aggregate gate still fails; emitted artifacts alone are never a ship pass.
  const consumerSteps = [
    { args: ['run', 'bundle:ship'], prepareMetadata: true },
    { args: ['run', 'typecheck'] },
    { args: ['exec', '--', 'tsc', '--noEmit', '--strict', '--skipLibCheck', '--target', 'es2020', '--module', 'esnext', '--moduleResolution', 'node',
      'compatibility/sx-package-public.ts', 'compatibility/sx-package-deep.ts'] },
    { args: ['run', 'package:solution'], prepareMetadata: true },
  ];
  for (const step of consumerSteps) {
    try { run(step.args, consumer, step); }
    catch (error) { consumerFailures.push({ args: step.args, message: error.message }); }
  }
  fs.writeFileSync(path.join(evidence, 'consumer-result.json'), JSON.stringify({
    consumer, packedFiles: packed.files.length, packedBytes: packed.size,
    documentedDeclarations, failures: consumerFailures,
  }, null, 2));
  assert.deepEqual(consumerFailures, [], `Isolated consumer verification failed; inspect every command in ${evidence}`);
  console.log(`Tarball consumer passed: ${packed.files.length} files; ${packed.size} bytes; ${documentedDeclarations} preserved JSDoc declarations; TypeScript, public/deep imports, SPFx ship bundle and solution. Consumer: ${consumer}`);
} finally {
  if(process.env.SPFX_KEEP_CONSUMER !== '1') fs.rmSync(temp,{recursive:true,force:true});
}

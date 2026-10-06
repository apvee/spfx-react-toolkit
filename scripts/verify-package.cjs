// A real, external SPFx consumer: no workspace symlinks, aliases or shared node_modules.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {execFileSync} = require('node:child_process');
const {createRequire} = require('node:module');
const root = path.resolve(__dirname, '..');
const library = path.join(root, 'packages/spfx-react-toolkit');
const app = path.join(root, 'apps/spfx-react-toolkit-test');
const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'spfx-tarball-consumer-'));
const run = (args,cwd) => execFileSync('npm',args,{cwd,stdio:'inherit'});
try {
  run(['run','build:library'],root);
  const packed = JSON.parse(execFileSync('npm',['pack','--json','--pack-destination',temp],{cwd:library,encoding:'utf8'}))[0];
  assert.equal(packed.name,'@apvee/spfx-react-toolkit');
  for (const file of packed.files) {
    assert.ok(/^(package\.json|README\.md|LICENSE|lib\/(index\.[^/]+|(core|hooks|services|helpers|utils)\/.+))$/.test(file.path), `Unexpected package file: ${file.path}`);
  }
  const packedPaths = new Set(packed.files.map(file => file.path));
  // Both the public entry points and their persistence/client dependencies must ship.
  for (const module of [
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
  const consumer = path.join(temp,'consumer'); fs.mkdirSync(consumer);
  for (const name of ['src','config','sharepoint','teams','gulpfile.js','tsconfig.json','.eslintrc.js']) fs.cpSync(path.join(app,name),path.join(consumer,name),{recursive:true});
  const manifest = JSON.parse(fs.readFileSync(path.join(app,'package.json'),'utf8'));
  manifest.name='isolated-spfx-toolkit-consumer';
  manifest.dependencies['@apvee/spfx-react-toolkit']='file:'+path.join(temp,packed.filename);
  fs.writeFileSync(path.join(consumer,'package.json'),JSON.stringify(manifest,null,2));
  // Seed locked transitive versions while npm replaces workspace identities with the tarball.
  const lock=JSON.parse(fs.readFileSync(path.join(root,'package-lock.json'),'utf8'));
  lock.name=manifest.name;lock.packages['']={name:manifest.name,version:manifest.version,dependencies:manifest.dependencies,devDependencies:manifest.devDependencies,engines:manifest.engines};
  delete lock.packages['apps/spfx-react-toolkit-test'];delete lock.packages['packages/spfx-react-toolkit'];
  for (const [key,value] of Object.entries(lock.packages)) if(value.link) delete lock.packages[key];
  fs.writeFileSync(path.join(consumer,'package-lock.json'),JSON.stringify(lock,null,2));
  const tsconfig=JSON.parse(fs.readFileSync(path.join(consumer,'tsconfig.json'),'utf8'));
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
  run(['install','--no-audit','--no-fund'],consumer);
  assert.equal(fs.lstatSync(path.join(consumer,'node_modules/@apvee/spfx-react-toolkit')).isSymbolicLink(),false);
  const hostRequire = createRequire(path.join(consumer, 'package.json'));
  const libraryRequire = createRequire(path.join(consumer, 'node_modules/@apvee/spfx-react-toolkit/package.json'));
  for (const name of ['@fluentui/react-migration-v8-v9', '@fluentui/react-theme']) {
    assert.equal(fs.realpathSync(hostRequire.resolve(name + '/package.json')),
      fs.realpathSync(libraryRequire.resolve(name + '/package.json')), `${name} must be shared with the consumer`);
  }
  run(['ls','react','react-dom','@microsoft/sp-core-library','@pnp/sp','@fluentui/react-migration-v8-v9','@fluentui/react-theme','tslib'],consumer);
  // SPFx generates Sass declarations during the bundle. A clean consumer has
  // no .scss.ts files yet, so generate them before the standalone typecheck.
  run(['run','bundle:ship'],consumer);
  run(['run','typecheck'],consumer);
  run(['run','package:solution'],consumer);
  console.log(`Tarball consumer passed: ${packed.files.length} files; ${packed.size} bytes; TypeScript, public/deep imports, SPFx ship bundle and solution. Consumer: ${consumer}`);
} finally {
  if(process.env.SPFX_KEEP_CONSUMER !== '1') fs.rmSync(temp,{recursive:true,force:true});
}

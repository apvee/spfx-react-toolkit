import { SPHttpClient } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';

import { createSPFxListKeyValueStoreService } from './spfx-list-key-value-store.internal';
import { createSiteStoreSetup, normalizeSiteStoreUrl, siteStoreListUrl } from './spfx-site-key-value-store.internal';

/** A stored site collection key and its deserialized value. */
export interface SPFxSiteKeyValueStoreServiceItem<T = unknown> {
  readonly key: string;
  readonly value: T;
  readonly description: string | undefined;
  readonly id: number;
}

/** Non-React CRUD and permission checks for a collection's root-web store. */
export interface SPFxSiteKeyValueStoreService {
  ensureListReady: (siteCollectionUrl: string) => Promise<void>;
  get: <T = unknown>(key: string, siteCollectionUrl: string) => Promise<SPFxSiteKeyValueStoreServiceItem<T> | undefined>;
  list: (siteCollectionUrl: string) => Promise<SPFxSiteKeyValueStoreServiceItem<unknown>[]>;
  save: <T = unknown>(key: string, value: T, siteCollectionUrl: string, description?: string) => Promise<void>;
  remove: (key: string, siteCollectionUrl: string) => Promise<void>;
  canCurrentUserWrite: (siteCollectionUrl: string) => Promise<boolean>;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function readPermissionMask(data: unknown): Record<string, unknown> | undefined {
  if (!isRecord(data)) return undefined;
  if ('d' in data) {
    return isRecord(data.d) && isRecord(data.d.EffectiveBasePermissions) ? data.d.EffectiveBasePermissions : undefined;
  }
  if ('EffectiveBasePermissions' in data) {
    return isRecord(data.EffectiveBasePermissions) ? data.EffectiveBasePermissions : undefined;
  }
  return data;
}

function permissionWord(value: unknown): number | undefined {
  if (typeof value !== 'number' && (typeof value !== 'string' || !/^\d+$/.test(value))) return undefined;
  const number = Number(value);
  return Number.isInteger(number) && number >= 0 && number <= 0xffffffff ? number : undefined;
}

/** Creates an isolated service; pass pageContext.site.absoluteUrl as the collection URL. */
export function createSPFxSiteKeyValueStoreService(spHttpClient: SPHttpClient): SPFxSiteKeyValueStoreService {
  const store = createSPFxListKeyValueStoreService(spHttpClient, 'site');
  const setup = createSiteStoreSetup(spHttpClient);
  const canCurrentUserWrite = async (siteCollectionUrl: string): Promise<boolean> => {
    try {
      const url = normalizeSiteStoreUrl(siteCollectionUrl);
      const exists = await setup.exists(url);
      const response = await spHttpClient.get(
        `${exists ? siteStoreListUrl(url) : `${url}/_api/web`}/effectiveBasePermissions`, SPHttpClient.configurations.v1
      );
      if (!response.ok) return false;
      const data: unknown = await response.json();
      const record = readPermissionMask(data);
      if (!record) return false;
      const High = permissionWord(record.High);
      const Low = permissionWord(record.Low);
      if (High === undefined || Low === undefined) return false;
      const permissions = new SPPermission({ High, Low });
      return permissions.hasAllPermissions(SPPermission.addListItems, SPPermission.editListItems, SPPermission.deleteListItems) &&
        (exists || permissions.hasPermission(SPPermission.manageLists));
    } catch { return false; }
  };
  return { ...store, canCurrentUserWrite };
}

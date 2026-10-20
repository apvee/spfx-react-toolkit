import { SPHttpClient } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';

import { createSPFxListKeyValueStoreService } from './spfx-list-key-value-store.internal';
import { createSiteStoreSetup, normalizeSiteStoreUrl, siteStoreListUrl } from './spfx-site-key-value-store.internal';

/**
 * Stored key and deserialized value. T is an expected shape, not runtime validation; legacy tenant serialization can coerce JSON-like strings and does not preserve Date/bigint types.
 *
 * @example
 * ```ts
 * const item: SPFxSiteKeyValueStoreServiceItem<boolean> | undefined = await store.get('featureEnabled', siteCollectionUrl);
 * console.log(item?.value);
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxSiteKeyValueStoreServiceItem<T = unknown> {
  /**
   * Stored Title/key, preserving its stored casing.
   */
  readonly key: string;
  /**
   * Deserialized value; caller must validate untrusted content.
   */
  readonly value: T;
  /**
   * Optional metadata; empty descriptions are returned as undefined.
   */
  readonly description: string | undefined;
  /**
   * SharePoint list item ID.
   */
  readonly id: number;
}

/**
 * Root-web SiteKeyValueStore operations shared by all subsites. Reads verify schema without provisioning, and only confirmed list absence is treated as empty. Requests reject on failed schema/item/page operations.
 *
 * @example
 * ```ts
 * const store: SPFxSiteKeyValueStoreService = createSPFxSiteKeyValueStoreService(client);
 * await store.save('featureEnabled', true, siteCollectionUrl);
 * ```
 */
export interface SPFxSiteKeyValueStoreService {
  /**
   * Provisions or repairs required list fields; concurrent setup is shared per URL within this instance. Failed setup can retry.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @returns Resolves when setup succeeds.
   * @throws Provisioning or incompatible-schema failures.
   */
  ensureListReady: (siteCollectionUrl: string) => Promise<void>;
  /**
   * Reads a key without provisioning.
   * @param key - Stored key to match.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @returns Deserialized item, or undefined for an absent key or confirmed absent list.
   * @throws Failed read requests or incompatible schema.
   */
  get: <T = unknown>(key: string, siteCollectionUrl: string) => Promise<SPFxSiteKeyValueStoreServiceItem<T> | undefined>;
  /**
   * Reads every page in Title order.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @returns Items, or an empty array for a confirmed absent list.
   * @throws Failed reads, schema or continuation validation failures.
   */
  list: (siteCollectionUrl: string) => Promise<SPFxSiteKeyValueStoreServiceItem<unknown>[]>;
  /**
   * Creates or updates a key after ensuring setup. Updates use wildcard ETags and last-writer-wins semantics.
   * @param key - Stored key.
   * @param value - Value serialized by legacy tenant serialization.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @param description - Omitted preserves an existing description; empty string clears it.
   * @returns Resolves after the write succeeds.
   * @throws Setup, serialization or request failures.
   */
  save: <T = unknown>(key: string, value: T, siteCollectionUrl: string, description?: string) => Promise<void>;
  /**
   * Deletes without provisioning; confirmed absent list/key is a no-op.
   * @param key - Key to remove.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @returns Resolves after deletion or confirmed absence.
   * @throws Request or schema failures.
   */
  remove: (key: string, siteCollectionUrl: string) => Promise<void>;
  /**
   * Checks effective Add/Edit/Delete Items on the list; an absent list also needs Manage Lists on the root web. This indicator does not authorize writes or schema changes.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @returns False for empty URLs or failed lookups; otherwise the effective-permission indicator.
   */
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

/**
 * Creates an isolated, non-React site store. Construction performs no request; SharePoint enforces actual grants.
 * @param spHttpClient - Caller-owned authenticated SharePoint client.
 * @returns Explicit store operations.
 *
 * @example
 * ```ts
 * const store = createSPFxSiteKeyValueStoreService(context.spHttpClient);
 * await store.save('featureEnabled', true, context.pageContext.site.absoluteUrl);
 * ```
 */
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

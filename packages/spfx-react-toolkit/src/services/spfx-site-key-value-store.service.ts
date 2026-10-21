import { SPHttpClient } from '@microsoft/sp-http';

import { createSPFxListKeyValueStoreService } from './spfx-list-key-value-store.internal';
import { canWriteKeyValueStore } from './spfx-site-key-value-store.internal';

/**
 * Stored key and deserialized value. T is an expected shape, not runtime validation; historical data keeps its JSON-first interpretation. New text preserves its value; bigint reads as exact decimal text and Date as ISO text. Legacy empty REST Note fields still read as empty strings.
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
   * Creates or updates a key after ensuring setup. Uses the observed item ETag; conflicts (HTTP 412) and missing ETags reject without a wildcard overwrite or blind retry. Creation collisions allow one validated conditional update and expose its actual error. This protects the service snapshot, not a consumer's earlier get().
   * @param key - Stored key.
   * @param value - Value to encode: empty/JSON-like text is JSON-quoted, bigint becomes quoted decimal text and Date remains ISO text. Existing data is not rewritten. Circular values, nested bigint and invalid Date throw before provisioning; T does not validate runtime content.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @param description - Omitted preserves an existing description; empty string clears it.
   * @returns Resolves after the write succeeds.
   * @throws Setup, serialization, missing ETag, request or conflict failures.
   */
  save: <T = unknown>(key: string, value: T, siteCollectionUrl: string, description?: string) => Promise<void>;
  /**
   * Deletes using an observed ETag without provisioning; confirmed absent list/key is a no-op. Snapshot/deletion 404 is accepted only after rechecking key absence; concurrent changes reject.
   * @param key - Key to remove.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @returns Resolves after deletion or confirmed absence.
   * @throws Request/schema failures, missing ETag or concurrent changes.
   */
  remove: (key: string, siteCollectionUrl: string) => Promise<void>;
  /**
   * Checks effective Add/Edit/Delete Items on the list; an absent list also needs Manage Lists on the root web. This indicator does not authorize writes or schema changes.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @returns False for empty URLs or failed lookups; otherwise the effective-permission indicator.
   */
  canCurrentUserWrite: (siteCollectionUrl: string) => Promise<boolean>;
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
  const canCurrentUserWrite = (siteCollectionUrl: string): Promise<boolean> =>
    canWriteKeyValueStore(spHttpClient, siteCollectionUrl, 'site');
  return { ...store, canCurrentUserWrite };
}

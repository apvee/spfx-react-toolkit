/**
 * Share tenant settings through a caller-selected app-catalog list with observed-ETag mutations.
 *
 * Keys are forwarded unchanged; SharePoint field/unique-column comparison
 * determines equality, while returned Title casing is preserved. URLs are
 * trimmed and stripped of trailing slashes, not validated as collection/catalog
 * addresses. Empty URLs reject before requests. Reads validate schema without
 * enforcing or repairing uniqueness: get rejects duplicates, list preserves
 * historical duplicate rows. Only the list existence probe's 404 is absence;
 * field/item/page 404 and access failures reject. Services have no value cache,
 * React provider dependency or disposal method.
 *
 * @packageDocumentation
 * @docPage services/tenant-key-value-store
 */
import { SPHttpClient } from '@microsoft/sp-http';

import { createSPFxListKeyValueStoreService } from './spfx-list-key-value-store.internal';

/**
 * Stored key and deserialized value. T is an expected shape, not runtime validation; existing raw values keep legacy JSON-first interpretation. New strings preserve their text; Date reads as ISO text and bigint as exact decimal text. A legacy REST null Note reads as null here, whereas the site store reads it as empty text.
 *
 * @example
 * ```ts
 * import { SPFxTenantKeyValueStoreService, SPFxTenantKeyValueStoreServiceItem } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(catalogUrl: string, store: SPFxTenantKeyValueStoreService) {
 *   const item: SPFxTenantKeyValueStoreServiceItem<boolean> | undefined = await store.get('featureEnabled', catalogUrl);
 *   console.log(item?.value);
 * }
 * ```
 * @template T - Expected application data shape; defaults to unknown and performs no runtime validation.
 */
export interface SPFxTenantKeyValueStoreServiceItem<T = unknown> {
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
 * TenantKeyValueStore operations in an explicitly supplied app catalog. Reads and remove never provision or repair schema; only save/ensureListReady can do so. Requests reject on failed schema/read/write operations. A confirmed missing list is empty.
 *
 * @example
 * ```ts
 * import type { SPHttpClient } from '@microsoft/sp-http';
 * import { SPFxTenantKeyValueStoreService, createSPFxTenantKeyValueStoreService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(client: SPHttpClient, catalogUrl: string) {
 *   const store: SPFxTenantKeyValueStoreService = createSPFxTenantKeyValueStoreService(client);
 *   await store.save('featureEnabled', true, catalogUrl);
 * }
 * ```
 */
export interface SPFxTenantKeyValueStoreService {
  /**
   * Provisions the hidden generic TenantKeyValueStore (BaseTemplate 100, Hidden/NoCrawl true), or repairs Title (Text), Value and Description (Note), requiring indexed unique Title. Normalized catalog URLs share setup within this instance. Failed setup can retry; existing duplicates or incompatible fields reject without deleting data. Separate instances recover list/field creation races only through bounded validated rereads; 401/403 never recover.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Resolves when setup succeeds.
   * @throws Provisioning or incompatible-schema failures.
   */
  ensureListReady: (catalogUrl: string) => Promise<void>;
  /**
   * Reads a unique key after read-only schema validation; duplicate matches reject.
   * @typeParam T - Expected deserialized value; defaults to unknown and performs no runtime validation.
   * @param key - Stored key to match.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Deserialized item, or undefined for a confirmed absent list/key.
   * @throws Failed probes/reads, malformed data or incompatible schema; 401/403/5xx are never absence.
   */
  get: <T = unknown>(key: string, catalogUrl: string) => Promise<SPFxTenantKeyValueStoreServiceItem<T> | undefined>;
  /**
   * Reads every server page in Title order. Continuation links must stay on the same origin/list items path and cannot repeat; partial results are not returned after a failure.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Complete items, including historical duplicates, or an empty array for a confirmed absent list.
   * @throws Failed reads, malformed rows or unsafe/repeated continuation links, without partial results.
   */
  list: (catalogUrl: string) => Promise<SPFxTenantKeyValueStoreServiceItem<unknown>[]>;
  /**
   * Creates or updates a key after ensuring setup. Updates require the ETag of the service's item snapshot; concurrent changes reject (HTTP 412), without a blind retry. Creation collisions allow one validated conditional update. This is not compare-and-swap against a previous get() call.
   * @typeParam T - Input value shape; defaults to unknown and does not guarantee serializability.
   * @param key - Stored key.
   * @param value - Value to serialize. New JSON-like strings are JSON-quoted, bigint becomes quoted decimal text, Date stays ISO text; other inputs retain the legacy format. Circular objects, nested bigint and invalid Date reject before provisioning. T does not validate stored data.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @param description - Omitted preserves an existing description; empty string clears it.
   * @returns Resolves after the write succeeds.
   * @throws Setup, serialization or request failures.
   */
  save: <T = unknown>(key: string, value: T, catalogUrl: string, description?: string) => Promise<void>;
  /**
   * Deletes a key without provisioning, using an observed item ETag. A confirmed absent list/key is a no-op; concurrent changes reject. A vanished snapshot/deletion is accepted only after rechecking key absence, never by deleting a recreated key.
   * @param key - Key to remove.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Resolves after deletion or confirmed absence.
   * @throws Request/schema failures, missing ETag or concurrent changes.
   */
  remove: (key: string, catalogUrl: string) => Promise<void>;
}

/**
 * Creates an isolated, non-React tenant store. Construction performs no request; SharePoint enforces actual grants.
 * @param spHttpClient - Caller-owned authenticated SharePoint client.
 * @returns Explicit store operations.
 *
 * @example Basic usage
 * ```ts
 * import type { WebPartContext } from '@microsoft/sp-webpart-base';
 * import { createSPFxTenantKeyValueStoreService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(context: WebPartContext, catalogUrl: string) {
 *   const store = createSPFxTenantKeyValueStoreService(context.spHttpClient);
 *   await store.save('featureEnabled', true, catalogUrl);
 * }
 * ```
 * @useWhen You share application settings through a TenantKeyValueStore list at an explicit app-catalog URL.
 *
 *
 * @example Preserve numeric-looking text in a tenant list
 * ```ts
 * import type { SPHttpClient } from '@microsoft/sp-http';
 *
 * import { createSPFxTenantKeyValueStoreService } from '@apvee/spfx-react-toolkit';
 * async function demonstrate(context: { spHttpClient: SPHttpClient }, catalogUrl: string) {
 *   const store = createSPFxTenantKeyValueStoreService(context.spHttpClient);
 *   await store.save('my-app:version', '123', catalogUrl);
 *   const version = await store.get<unknown>('my-app:version', catalogUrl);
 *   if (typeof version?.value === 'string') console.log(version.value);
 * }
 * ```
 */
export function createSPFxTenantKeyValueStoreService(
  spHttpClient: SPHttpClient
): SPFxTenantKeyValueStoreService {
  return createSPFxListKeyValueStoreService(spHttpClient, 'tenant');
}

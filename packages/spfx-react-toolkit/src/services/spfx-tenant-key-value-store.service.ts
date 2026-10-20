import { SPHttpClient } from '@microsoft/sp-http';

import { createSPFxListKeyValueStoreService } from './spfx-list-key-value-store.internal';

/**
 * Stored key and deserialized value. T is an expected shape, not runtime validation; legacy tenant serialization can coerce JSON-like strings and does not preserve Date/bigint types.
 *
 * @example
 * ```ts
 * const item: SPFxTenantKeyValueStoreServiceItem<boolean> | undefined = await store.get('featureEnabled', catalogUrl);
 * console.log(item?.value);
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
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
 * TenantKeyValueStore operations in an explicitly supplied app catalog. Reads do not provision; save and remove ensure the list exists. Requests reject on failure.
 *
 * @example
 * ```ts
 * const store: SPFxTenantKeyValueStoreService = createSPFxTenantKeyValueStoreService(client);
 * await store.save('featureEnabled', true, catalogUrl);
 * ```
 */
export interface SPFxTenantKeyValueStoreService {
  /**
   * Provisions or repairs required list fields; concurrent setup is shared per URL within this instance. Failed setup can retry.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Resolves when setup succeeds.
   * @throws Provisioning or incompatible-schema failures.
   */
  ensureListReady: (catalogUrl: string) => Promise<void>;
  /**
   * Reads a key without provisioning.
   * @param key - Stored key to match.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Deserialized item, or undefined for an absent key.
   * @throws Failed read requests, including an absent list.
   */
  get: <T = unknown>(key: string, catalogUrl: string) => Promise<SPFxTenantKeyValueStoreServiceItem<T> | undefined>;
  /**
   * Reads the first server page in Title order; continuation links are not followed.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Items.
   * @throws Failed reads.
   */
  list: (catalogUrl: string) => Promise<SPFxTenantKeyValueStoreServiceItem<unknown>[]>;
  /**
   * Creates or updates a key after ensuring setup. Updates use wildcard ETags and last-writer-wins semantics.
   * @param key - Stored key.
   * @param value - Value serialized by legacy tenant serialization.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @param description - Omitted preserves an existing description; empty string clears it.
   * @returns Resolves after the write succeeds.
   * @throws Setup, serialization or request failures.
   */
  save: <T = unknown>(key: string, value: T, catalogUrl: string, description?: string) => Promise<void>;
  /**
   * Ensures the list exists, then deletes a key; an absent key is a no-op.
   * @param key - Key to remove.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Resolves after deletion or confirmed absence.
   * @throws Request or provisioning failures.
   */
  remove: (key: string, catalogUrl: string) => Promise<void>;
}

/**
 * Creates an isolated, non-React tenant store. Construction performs no request; SharePoint enforces actual grants.
 * @param spHttpClient - Caller-owned authenticated SharePoint client.
 * @returns Explicit store operations.
 *
 * @example
 * ```ts
 * const store = createSPFxTenantKeyValueStoreService(context.spHttpClient);
 * await store.save('featureEnabled', true, catalogUrl);
 * ```
 */
export function createSPFxTenantKeyValueStoreService(
  spHttpClient: SPHttpClient
): SPFxTenantKeyValueStoreService {
  return createSPFxListKeyValueStoreService(spHttpClient, 'tenant');
}

// useSPFxTenantKeyValueStore.ts
// Hook to manage tenant-wide key-value pairs using a hidden SharePoint list in the tenant app catalog

import { useState, useCallback, useEffect, useRef } from 'react';
import { useAppCatalogUrl } from './useAppCatalogUrl.internal';
import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';

// ═══════════════════════════════════════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════════════════════════════════════

const LIST_TITLE = 'TenantKeyValueStore';

// ═══════════════════════════════════════════════════════════════════════════
// PURE FUNCTIONS (extracted for stability and testability)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Escape a string value for use in OData filter expressions.
 * Single quotes must be doubled to prevent injection.
 *
 * @param value - Raw string value
 * @returns Escaped string safe for OData $filter
 */
function escapeODataValue(value: string): string {
  return value.replace(/'/g, "''");
}

/**
 * Serialize a value for storage.
 * - Primitives (string, number, boolean, null, bigint) → String(value)
 * - Date → ISO 8601 string
 * - Objects/arrays → JSON.stringify()
 *
 * @param value - Value to serialize
 * @returns Serialized string representation
 */
function serializeValue(value: unknown): string {
  if (value === null) return String(value);
  if (value instanceof Date) return value.toISOString();
  const type = typeof value;
  if (type === 'string' || type === 'number' || type === 'boolean' || type === 'bigint') {
    return String(value);
  }
  return JSON.stringify(value);
}

/**
 * Deserialize a stored string value back to a typed value.
 * Attempts JSON.parse first; falls back to raw string.
 *
 * @param rawValue - Stored string value
 * @returns Parsed value
 */
function deserializeValue<T>(rawValue: string): T {
  try {
    return JSON.parse(rawValue) as T;
  } catch {
    return rawValue as unknown as T;
  }
}

/**
 * Build the list REST API base URL.
 *
 * @param catalogUrl - Tenant app catalog absolute URL
 * @returns REST API URL for the TenantKeyValueStore list
 */
function getListApiUrl(catalogUrl: string): string {
  return `${catalogUrl}/_api/web/lists/getByTitle('${LIST_TITLE}')`;
}

/**
 * Create a multiline text (Note) field on the list.
 *
 * @param client - SPHttpClient instance
 * @param listApiUrl - REST API URL for the target list
 * @param fieldTitle - Internal name / title for the new field
 */
async function createField(
  client: SPHttpClient,
  listApiUrl: string,
  fieldTitle: string
): Promise<void> {
  const response: SPHttpClientResponse = await client.post(
    `${listApiUrl}/fields`,
    SPHttpClient.configurations.v1,
    {
      body: JSON.stringify({
        FieldTypeKind: 3, // Note (multiline text)
        Title: fieldTitle
      })
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to create ${fieldTitle} field: ${response.statusText}. ${errorText}`);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// TYPES
// ═══════════════════════════════════════════════════════════════════════════

/**
 * A single item from the tenant key-value store
 */
export interface SPFxTenantKeyValueStoreItem<T = unknown> {
  /** Property key (stored in Title column) */
  readonly key: string;
  /** Deserialized property value */
  readonly value: T;
  /** Optional description metadata */
  readonly description: string | undefined;
  /** SharePoint list item ID (internal use) */
  readonly id: number;
}

/**
 * Return type for useSPFxTenantKeyValueStore hook
 */
export interface SPFxTenantKeyValueStoreResult {
  /**
   * Loading state for read operations (`get`, `list`).
   * True during any active read call.
   */
  readonly isLoading: boolean;

  /**
   * Last error from read operations.
   * Cleared on next successful read.
   */
  readonly error: Error | undefined;

  /**
   * Loading state for write operations (`save`, `remove`).
   * True during any active write call.
   */
  readonly isWriting: boolean;

  /**
   * Last error from write operations.
   * Cleared on next successful write.
   */
  readonly writeError: Error | undefined;

  /**
   * Whether the current user has permission to write to the store.
   * Requires Site Collection Administrator role on the tenant app catalog.
   * Computed asynchronously on first operation.
   */
  readonly canWrite: boolean;

  /**
   * Computed state: true when the SPHttpClient is available and
   * the hook is ready to perform operations.
   */
  readonly isReady: boolean;

  /**
   * Get a single item by key.
   * Returns undefined if the key does not exist or the list has not been provisioned.
   *
   * @param key - Property key to look up
   * @returns The item with deserialized value, or undefined
   *
   * @example
   * ```tsx
   * const store = useSPFxTenantKeyValueStore();
   *
   * const handleLoad = async () => {
   *   const item = await store.get<string>('apiEndpoint');
   *   if (item) {
   *     console.log(item.value); // "https://api.example.com"
   *   }
   * };
   * ```
   */
  readonly get: <T = unknown>(key: string) => Promise<SPFxTenantKeyValueStoreItem<T> | undefined>;

  /**
   * List all items in the store.
   * Returns an empty array if the list has not been provisioned.
   * Items are sorted by key (Title) ascending.
   *
   * @returns Array of all items with deserialized values
   *
   * @example
   * ```tsx
   * const store = useSPFxTenantKeyValueStore();
   *
   * const items = await store.list();
   * items.forEach(item => {
   *   console.log(`${item.key}: ${typeof item.value === 'object' ? JSON.stringify(item.value) : item.value}`);
   * });
   * ```
   */
  readonly list: () => Promise<SPFxTenantKeyValueStoreItem<unknown>[]>;

  /**
   * Save (create or update) a key-value pair.
   * Auto-provisions the hidden list on first write if it doesn't exist.
   *
   * Smart serialization:
   * - Primitives (string, number, boolean, null, bigint) → String(value)
   * - Date objects → ISO 8601 string
   * - Objects/arrays → JSON.stringify(value)
   *
   * @param key - Property key
   * @param value - Value to store
   * @param description - Optional description metadata
   *
   * @example
   * ```tsx
   * const store = useSPFxTenantKeyValueStore();
   *
   * // String value
   * await store.save<string>('apiEndpoint', 'https://api.example.com', 'Production API');
   *
   * // Complex object
   * await store.save<FeatureFlags>('featureFlags', { enableChat: true, maxUsers: 1000 });
   * ```
   */
  readonly save: <T = unknown>(key: string, value: T, description?: string) => Promise<void>;

  /**
   * Remove an item by key.
   * No-op if the key does not exist or the list has not been provisioned.
   *
   * @param key - Property key to remove
   *
   * @example
   * ```tsx
   * const store = useSPFxTenantKeyValueStore();
   *
   * await store.remove('deprecatedSetting');
   * ```
   */
  readonly remove: (key: string) => Promise<void>;
}

/**
 * SharePoint list item response shape
 */
interface IListItemResponse {
  Id: number;
  Title: string;
  Value: string;
  Description?: string;
}

/**
 * SharePoint list items collection response (odata=nometadata via SPHttpClient default)
 */
interface IListItemsResponse {
  value: IListItemResponse[];
}

/**
 * SharePoint fields check response (odata=nometadata via SPHttpClient default)
 */
interface IFieldsCheckResponse {
  value: Array<{ InternalName: string }>;
}

// ═══════════════════════════════════════════════════════════════════════════
// HOOK
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Hook to manage tenant-wide key-value pairs backed by a hidden SharePoint list.
 *
 * The store uses a hidden list named `TenantKeyValueStore` in the tenant app catalog.
 * The list is auto-provisioned on the first operation (get, list, save, or remove), with:
 * - Title column: key (indexed, unique)
 * - Value column: multiline text (Note)
 * - Description column: multiline text (Note)
 *
 * This hook provides an alternative to tenant properties (StorageEntity) for scenarios
 * requiring read/write access via REST, since Microsoft has blocked the
 * SetStorageEntity and RemoveStorageEntity REST API endpoints.
 *
 * Features:
 * - CRUD operations (get, list, save, remove) on tenant-scoped key-value data
 * - Auto-provisioning of the hidden list on first operation (any CRUD call)
 * - Smart provisioning: ref-based fast path (zero API calls after first check),
 *   field introspection only when list already exists, Title uniqueness only on creation
 * - Smart serialization for primitives, Date, and complex objects
 * - Permission checking (canWrite flag via IsSiteAdmin)
 * - Unique key enforcement (Indexed + EnforceUniqueValues on Title)
 * - Concurrent provisioning protection (mutex ref)
 * - Memory leak safe with mounted state tracking
 * - Idempotent remove (no-op if key or list doesn't exist)
 *
 * Requirements:
 * - Tenant app catalog must be provisioned
 * - Read: Any authenticated user
 * - Write/Remove: Site Collection Administrator role on tenant app catalog site
 *
 * @returns Object with loading states, error states, permission flag, and CRUD functions
 *
 * @example Basic usage
 * ```tsx
 * function TenantConfigPanel() {
 *   const store = useSPFxTenantKeyValueStore();
 *
 *   const [endpoint, setEndpoint] = React.useState<string>('');
 *
 *   React.useEffect(() => {
 *     store.get<string>('apiEndpoint').then(item => {
 *       if (item) setEndpoint(item.value);
 *     });
 *   }, []);
 *
 *   const handleSave = async () => {
 *     await store.save<string>('apiEndpoint', endpoint, 'Production API endpoint');
 *   };
 *
 *   if (!store.isReady) return <Spinner />;
 *
 *   return (
 *     <Stack tokens={{ childrenGap: 10 }}>
 *       <TextField value={endpoint} onChange={(_, v) => setEndpoint(v ?? '')} />
 *       <PrimaryButton onClick={handleSave} disabled={store.isWriting}>
 *         {store.isWriting ? 'Saving...' : 'Save'}
 *       </PrimaryButton>
 *     </Stack>
 *   );
 * }
 * ```
 *
 * @example Heterogeneous value types
 * ```tsx
 * const store = useSPFxTenantKeyValueStore();
 *
 * // String
 * await store.save<string>('appVersion', '2.1.0');
 * const ver = await store.get<string>('appVersion');
 *
 * // Number
 * await store.save<number>('maxUploadSize', 10485760);
 * const size = await store.get<number>('maxUploadSize');
 *
 * // Complex object
 * interface FeatureFlags { enableChat: boolean; maxUsers: number; }
 * await store.save<FeatureFlags>('featureFlags', { enableChat: true, maxUsers: 500 });
 * const flags = await store.get<FeatureFlags>('featureFlags');
 * ```
 *
 * @example Dashboard listing all properties
 * ```tsx
 * function AllPropertiesView() {
 *   const store = useSPFxTenantKeyValueStore();
 *   const [items, setItems] = React.useState<SPFxTenantKeyValueStoreItem<unknown>[]>([]);
 *
 *   React.useEffect(() => {
 *     store.list().then(setItems);
 *   }, []);
 *
 *   return (
 *     <DetailsList
 *       items={items.map(i => ({
 *         key: i.key,
 *         value: typeof i.value === 'object' ? JSON.stringify(i.value) : String(i.value),
 *         description: i.description ?? '',
 *       }))}
 *       columns={[
 *         { key: 'key', name: 'Key', fieldName: 'key', minWidth: 150 },
 *         { key: 'value', name: 'Value', fieldName: 'value', minWidth: 200 },
 *         { key: 'desc', name: 'Description', fieldName: 'description', minWidth: 200 },
 *       ]}
 *     />
 *   );
 * }
 * ```
 *
 * @example Permission-aware UI
 * ```tsx
 * function TenantStoreManager() {
 *   const store = useSPFxTenantKeyValueStore();
 *
 *   if (!store.canWrite) {
 *     return (
 *       <MessageBar messageBarType={MessageBarType.info}>
 *         Read-only access. Contact your SharePoint administrator for write permissions.
 *       </MessageBar>
 *     );
 *   }
 *
 *   return <EditablePropertyPanel store={store} />;
 * }
 * ```
 */
export function useSPFxTenantKeyValueStore(): SPFxTenantKeyValueStoreResult {
  const { spHttpClient, discoverAppCatalogUrl, checkWritePermission, isMountedRef } = useAppCatalogUrl();

  // State management
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  const [isWriting, setIsWriting] = useState<boolean>(false);
  const [writeError, setWriteError] = useState<Error | undefined>(undefined);
  const [canWrite, setCanWrite] = useState<boolean>(false);

  // Provisioning state
  const listProvisionedRef = useRef<boolean>(false);
  const provisioningPromiseRef = useRef<Promise<void> | undefined>(undefined);
  const permissionCheckedRef = useRef<boolean>(false);

  // ─────────────────────────────────────────────────────────────────────────
  // Internal helpers
  // ─────────────────────────────────────────────────────────────────────────



  /**
   * Check and update write permission (executed once)
   */
  const ensurePermissionChecked = useCallback((catalogUrl: string): void => {
    if (permissionCheckedRef.current) return;
    permissionCheckedRef.current = true;
    checkWritePermission(catalogUrl)
      .then(hasPermission => {
        if (isMountedRef.current) {
          setCanWrite(hasPermission);
        }
      })
      .catch(() => {
        if (isMountedRef.current) {
          setCanWrite(false);
        }
      });
  }, [checkWritePermission, isMountedRef]);



  /**
   * Ensure the hidden list exists and has all required fields.
   * Uses a ref-based fast path (zero API calls after first successful check)
   * and a mutex to prevent concurrent provisioning.
   *
   * Flow:
   * 1. Fast path: `listProvisionedRef` is true → return immediately
   * 2. Mutex: reuse in-flight promise if already provisioning
   * 3. Lightweight existence check (1 API call): GET list?$select=Id
   *    - 200 (exists) → field introspection, create missing fields only
   *    - 404 (missing) → full provision: create list + fields + Title uniqueness
   * 4. Mark ref true on success, cleanup mutex on success or failure
   */
  const ensureListReady = useCallback((
    client: SPHttpClient,
    catalogUrl: string
  ): Promise<void> => {
    // Fast path: already verified — zero API calls
    if (listProvisionedRef.current) return Promise.resolve();

    // Mutex: reuse in-flight provisioning promise
    if (provisioningPromiseRef.current) return provisioningPromiseRef.current;

    const doProvision = async (): Promise<void> => {
      const listApiUrl = getListApiUrl(catalogUrl);

      // ── Step 1: Lightweight existence check (single API call) ──────
      const listResponse: SPHttpClientResponse = await client.get(
        `${listApiUrl}?$select=Id`,
        SPHttpClient.configurations.v1
      );

      const listExists = listResponse.status !== 404;

      if (listExists && !listResponse.ok) {
        throw new Error(`Failed to check list existence: ${listResponse.statusText}`);
      }

      if (listExists) {
        // ── Step 2: List exists → field introspection ──────────────
        const fieldsResponse: SPHttpClientResponse = await client.get(
          `${listApiUrl}/fields?$filter=InternalName eq 'Value' or InternalName eq 'Description'&$select=InternalName`,
          SPHttpClient.configurations.v1
        );

        if (!fieldsResponse.ok) {
          throw new Error(`Failed to check list fields: ${fieldsResponse.statusText}`);
        }

        const fields: IFieldsCheckResponse = await fieldsResponse.json();
        const existingFields = fields.value.map(f => f.InternalName);
        const hasValue = existingFields.includes('Value');
        const hasDescription = existingFields.includes('Description');

        // All fields present → list is fully ready
        if (hasValue && hasDescription) {
          return;
        }

        // Create only the missing fields
        if (!hasValue) {
          await createField(client, listApiUrl, 'Value');
        }
        if (!hasDescription) {
          await createField(client, listApiUrl, 'Description');
        }

        // Skip Title uniqueness — list already existed, constraint is either
        // already set or intentionally not re-applied on field repair
        return;
      }

      // ── Step 3: List does not exist → full provision ─────────────
      // 3a. Create list
      const createListResponse: SPHttpClientResponse = await client.post(
        `${catalogUrl}/_api/web/lists`,
        SPHttpClient.configurations.v1,
        {
          body: JSON.stringify({
            BaseTemplate: 100,
            Title: LIST_TITLE,
            Hidden: true,
            NoCrawl: true
          })
        }
      );

      if (!createListResponse.ok) {
        const errorText = await createListResponse.text();
        throw new Error(`Failed to create list: ${createListResponse.statusText}. ${errorText}`);
      }

      // 3b. Create Value + Description fields (list is fresh — no need to check)
      await createField(client, getListApiUrl(catalogUrl), 'Value');
      await createField(client, getListApiUrl(catalogUrl), 'Description');

      // 3c. Set Title field as Indexed + EnforceUniqueValues (only on creation)
      const titleResp: SPHttpClientResponse = await client.post(
        `${getListApiUrl(catalogUrl)}/fields/getByInternalNameOrTitle('Title')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'X-HTTP-Method': 'MERGE',
            'If-Match': '*'
          },
          body: JSON.stringify({
            Indexed: true,
            EnforceUniqueValues: true
          })
        }
      );

      if (!titleResp.ok) {
        // Non-fatal: unique constraint may already exist
        console.warn('Failed to set Title uniqueness constraint. It may already be configured.');
      }
    };

    // All ref mutations are synchronous — no await between read and write.
    // doProvision() never touches refs; .then() callbacks are separate scopes.
    const cleanupMutex = (): void => { provisioningPromiseRef.current = undefined; };
    const promise = doProvision()
      .then(() => { listProvisionedRef.current = true; cleanupMutex(); })
      .catch((err: Error) => { cleanupMutex(); throw err; });
    provisioningPromiseRef.current = promise;
    return promise;
  }, []);

  // ─────────────────────────────────────────────────────────────────────────
  // Eager initialization: verify/provision list as soon as SPHttpClient is
  // available, overlapping with component render. By the time the user
  // triggers an operation, listProvisionedRef is already true → zero latency.
  // The mutex in ensureListReady handles the race if an operation fires
  // while this init is still in-flight.
  // ─────────────────────────────────────────────────────────────────────────
  useEffect(() => {
    if (!spHttpClient) return;

    discoverAppCatalogUrl()
      .then(catalogUrl => {
        ensurePermissionChecked(catalogUrl);
        return ensureListReady(spHttpClient, catalogUrl);
      })
      .catch(() => {
        // Silently fail — operations will retry on demand via their own ensureListReady call
      });
  }, [spHttpClient, discoverAppCatalogUrl, ensurePermissionChecked, ensureListReady]);

  /**
   * Find an item by key, returning the raw list item or undefined.
   */
  const findItemByKey = useCallback(async (
    client: SPHttpClient,
    catalogUrl: string,
    key: string
  ): Promise<IListItemResponse | undefined> => {
    const safeKey = escapeODataValue(key);
    const response: SPHttpClientResponse = await client.get(
      `${getListApiUrl(catalogUrl)}/items?$filter=Title eq '${safeKey}'&$select=Id,Title,Value,Description&$top=1`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to find item: ${response.statusText}`);
    }

    const data: IListItemsResponse = await response.json();
    return data.value.length > 0 ? data.value[0] : undefined;
  }, []);

  /**
   * Map a raw SharePoint list item to an SPFxTenantKeyValueStoreItem.
   */
  function mapItem<T>(raw: IListItemResponse): SPFxTenantKeyValueStoreItem<T> {
    return {
      key: raw.Title,
      value: deserializeValue<T>(raw.Value),
      description: raw.Description || undefined,
      id: raw.Id,
    };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // Public operations
  // ─────────────────────────────────────────────────────────────────────────

  const get = useCallback(async <T = unknown>(key: string): Promise<SPFxTenantKeyValueStoreItem<T> | undefined> => {
    if (!spHttpClient) {
      throw new Error('SPHttpClient not available. Cannot get property.');
    }

    setIsLoading(true);
    setError(undefined);

    try {
      const catalogUrl = await discoverAppCatalogUrl();

      // Ensure list exists (auto-provision if needed; fallback-safe for read-only users)
      try {
        await ensureListReady(spHttpClient, catalogUrl);
      } catch {
        // Provisioning may fail for read-only users — treat as "list not yet created"
        return undefined;
      }

      const raw = await findItemByKey(spHttpClient, catalogUrl, key);

      return raw ? mapItem<T>(raw) : undefined;
    } catch (err) {
      if (isMountedRef.current) {
        const capturedError = err instanceof Error ? err : new Error(String(err));
        setError(capturedError);
        console.error('Failed to get tenant key-value:', capturedError);
      }
      return undefined;
    } finally {
      if (isMountedRef.current) {
        setIsLoading(false);
      }
    }
  }, [spHttpClient, discoverAppCatalogUrl, ensureListReady, findItemByKey, isMountedRef]);

  const list = useCallback(async (): Promise<SPFxTenantKeyValueStoreItem<unknown>[]> => {
    if (!spHttpClient) {
      throw new Error('SPHttpClient not available. Cannot list properties.');
    }

    setIsLoading(true);
    setError(undefined);

    try {
      const catalogUrl = await discoverAppCatalogUrl();

      // Ensure list exists (auto-provision if needed; fallback-safe for read-only users)
      try {
        await ensureListReady(spHttpClient, catalogUrl);
      } catch {
        // Provisioning may fail for read-only users — treat as "list not yet created"
        return [];
      }

      const response: SPHttpClientResponse = await spHttpClient.get(
        `${getListApiUrl(catalogUrl)}/items?$select=Id,Title,Value,Description&$orderby=Title`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to list items: ${response.statusText}`);
      }

      const data: IListItemsResponse = await response.json();
      return data.value.map(raw => mapItem<unknown>(raw));
    } catch (err) {
      if (isMountedRef.current) {
        const capturedError = err instanceof Error ? err : new Error(String(err));
        setError(capturedError);
        console.error('Failed to list tenant key-values:', capturedError);
      }
      return [];
    } finally {
      if (isMountedRef.current) {
        setIsLoading(false);
      }
    }
  }, [spHttpClient, discoverAppCatalogUrl, ensureListReady, isMountedRef]);

  const save = useCallback(async <T = unknown>(
    key: string,
    value: T,
    description?: string
  ): Promise<void> => {
    if (!spHttpClient) {
      throw new Error('SPHttpClient not available. Cannot save property.');
    }

    setIsWriting(true);
    setWriteError(undefined);

    try {
      const catalogUrl = await discoverAppCatalogUrl();

      // Ensure list is ready (auto-provision on first operation)
      await ensureListReady(spHttpClient, catalogUrl);

      const serializedValue = serializeValue(value);

      // Check if key already exists
      const existing = await findItemByKey(spHttpClient, catalogUrl, key);

      if (existing) {
        // Update existing item
        const updateResponse: SPHttpClientResponse = await spHttpClient.post(
          `${getListApiUrl(catalogUrl)}/items(${existing.Id})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'X-HTTP-Method': 'MERGE',
              'If-Match': '*'
            },
            body: JSON.stringify({
              Value: serializedValue,
              Description: description ?? existing.Description ?? ''
            })
          }
        );

        if (!updateResponse.ok) {
          const errorText = await updateResponse.text();
          throw new Error(`Failed to update item: ${updateResponse.statusText}. ${errorText}`);
        }
      } else {
        // Create new item
        const createResponse: SPHttpClientResponse = await spHttpClient.post(
          `${getListApiUrl(catalogUrl)}/items`,
          SPHttpClient.configurations.v1,
          {
            body: JSON.stringify({
              Title: key,
              Value: serializedValue,
              Description: description ?? ''
            })
          }
        );

        if (!createResponse.ok) {
          const errorText = await createResponse.text();
          throw new Error(`Failed to create item: ${createResponse.statusText}. ${errorText}`);
        }
      }
    } catch (err) {
      if (isMountedRef.current) {
        const capturedError = err instanceof Error ? err : new Error(String(err));
        setWriteError(capturedError);
        console.error('Failed to save tenant key-value:', capturedError);
      }
      throw err;
    } finally {
      if (isMountedRef.current) {
        setIsWriting(false);
      }
    }
  }, [spHttpClient, discoverAppCatalogUrl, ensureListReady, findItemByKey, isMountedRef]);

  const remove = useCallback(async (key: string): Promise<void> => {
    if (!spHttpClient) {
      throw new Error('SPHttpClient not available. Cannot remove property.');
    }

    setIsWriting(true);
    setWriteError(undefined);

    try {
      const catalogUrl = await discoverAppCatalogUrl();

      // Ensure list is ready (auto-provision on first operation)
      await ensureListReady(spHttpClient, catalogUrl);

      // Find the item
      const existing = await findItemByKey(spHttpClient, catalogUrl, key);

      if (!existing) {
        return; // Item doesn't exist — idempotent no-op
      }

      // Delete the item
      const deleteResponse: SPHttpClientResponse = await spHttpClient.post(
        `${getListApiUrl(catalogUrl)}/items(${existing.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'X-HTTP-Method': 'DELETE',
            'If-Match': '*'
          }
        }
      );

      if (!deleteResponse.ok) {
        const errorText = await deleteResponse.text();
        throw new Error(`Failed to remove item: ${deleteResponse.statusText}. ${errorText}`);
      }
    } catch (err) {
      if (isMountedRef.current) {
        const capturedError = err instanceof Error ? err : new Error(String(err));
        setWriteError(capturedError);
        console.error('Failed to remove tenant key-value:', capturedError);
      }
      throw err;
    } finally {
      if (isMountedRef.current) {
        setIsWriting(false);
      }
    }
  }, [spHttpClient, discoverAppCatalogUrl, ensureListReady, findItemByKey, isMountedRef]);

  // Computed: ready when client is available
  const isReady = spHttpClient !== undefined;

  return {
    isLoading,
    error,
    isWriting,
    writeError,
    canWrite,
    isReady,
    get,
    list,
    save,
    remove,
  };
}

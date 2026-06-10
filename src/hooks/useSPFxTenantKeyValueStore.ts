// useSPFxTenantKeyValueStore.ts
// Hook to manage tenant-wide key-value pairs using a hidden SharePoint list in the tenant app catalog

import { useState, useCallback, useEffect, useMemo, useRef } from 'react';
import { useAppCatalogUrl } from './useAppCatalogUrl.internal';
import { createSPFxTenantKeyValueStoreService } from '../services/spfx-tenant-key-value-store.service';

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
 * - Concurrent provisioning protection (service mutex)
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
    const tenantKeyValueStoreService = useMemo(
        () => spHttpClient ? createSPFxTenantKeyValueStoreService(spHttpClient) : undefined,
        [spHttpClient]
    );

    // State management
    const [isLoading, setIsLoading] = useState<boolean>(false);
    const [error, setError] = useState<Error | undefined>(undefined);
    const [isWriting, setIsWriting] = useState<boolean>(false);
    const [writeError, setWriteError] = useState<Error | undefined>(undefined);
    const [canWrite, setCanWrite] = useState<boolean>(false);

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

    // ─────────────────────────────────────────────────────────────────────────
    // Eager initialization: verify/provision list as soon as SPHttpClient is
    // available, overlapping with component render. By the time the user
    // triggers an operation, the service cache is already warm → zero latency.
    // The service mutex handles the race if an operation fires while this init
    // is still in-flight.
    // ─────────────────────────────────────────────────────────────────────────
    useEffect(() => {
        if (!tenantKeyValueStoreService) return;

        discoverAppCatalogUrl()
            .then(catalogUrl => {
                ensurePermissionChecked(catalogUrl);
                return tenantKeyValueStoreService.ensureListReady(catalogUrl);
            })
            .catch(() => {
                // Silently fail — operations will retry on demand via their own ensureListReady call
            });
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, ensurePermissionChecked]);

    // ─────────────────────────────────────────────────────────────────────────
    // Public operations
    // ─────────────────────────────────────────────────────────────────────────

    const get = useCallback(async <T = unknown>(key: string): Promise<SPFxTenantKeyValueStoreItem<T> | undefined> => {
        if (!tenantKeyValueStoreService) {
            throw new Error('SPHttpClient not available. Cannot get property.');
        }

        setIsLoading(true);
        setError(undefined);

        try {
            const catalogUrl = await discoverAppCatalogUrl();

            // Ensure list exists (auto-provision if needed; fallback-safe for read-only users)
            try {
                await tenantKeyValueStoreService.ensureListReady(catalogUrl);
            } catch {
                // Provisioning may fail for read-only users — treat as "list not yet created"
                return undefined;
            }

            return tenantKeyValueStoreService.get<T>(key, catalogUrl);
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
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef]);

    const list = useCallback(async (): Promise<SPFxTenantKeyValueStoreItem<unknown>[]> => {
        if (!tenantKeyValueStoreService) {
            throw new Error('SPHttpClient not available. Cannot list properties.');
        }

        setIsLoading(true);
        setError(undefined);

        try {
            const catalogUrl = await discoverAppCatalogUrl();

            // Ensure list exists (auto-provision if needed; fallback-safe for read-only users)
            try {
                await tenantKeyValueStoreService.ensureListReady(catalogUrl);
            } catch {
                // Provisioning may fail for read-only users — treat as "list not yet created"
                return [];
            }

            return tenantKeyValueStoreService.list(catalogUrl);
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
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef]);

    const save = useCallback(async <T = unknown>(
        key: string,
        value: T,
        description?: string
    ): Promise<void> => {
        if (!tenantKeyValueStoreService) {
            throw new Error('SPHttpClient not available. Cannot save property.');
        }

        setIsWriting(true);
        setWriteError(undefined);

        try {
            const catalogUrl = await discoverAppCatalogUrl();
            await tenantKeyValueStoreService.save<T>(key, value, catalogUrl, description);
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
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef]);

    const remove = useCallback(async (key: string): Promise<void> => {
        if (!tenantKeyValueStoreService) {
            throw new Error('SPHttpClient not available. Cannot remove property.');
        }

        setIsWriting(true);
        setWriteError(undefined);

        try {
            const catalogUrl = await discoverAppCatalogUrl();
            await tenantKeyValueStoreService.remove(key, catalogUrl);
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
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef]);

    // Computed: ready when client is available
    const isReady = spHttpClient !== undefined;

    return useMemo(() => ({
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
    }), [isLoading, error, isWriting, writeError, canWrite, isReady, get, list, save, remove]);
}

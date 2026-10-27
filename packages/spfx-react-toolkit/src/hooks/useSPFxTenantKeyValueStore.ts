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
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const store = useSPFxTenantKeyValueStore();
 * const onLoad = async () => {
 *   const item: SPFxTenantKeyValueStoreItem<boolean> | undefined = await store.get('featureEnabled');
 *   console.log(item?.value);
 * };
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxTenantKeyValueStoreItem<T = unknown> {
    /** Property key (stored in Title column) */
    readonly key: string;
    /** Deserialized property value; the generic does not validate runtime data or preserve all legacy serialization types */
    readonly value: T;
    /** Optional description metadata */
    readonly description: string | undefined;
    /** SharePoint list item ID (internal use) */
    readonly id: number;
}

/**
 * Return type for useSPFxTenantKeyValueStore hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const store: SPFxTenantKeyValueStoreResult = useSPFxTenantKeyValueStore();
 * const onLoad = async () => console.log(await store.get<boolean>('featureEnabled'));
 * ```
 */
export interface SPFxTenantKeyValueStoreResult {
    /**
     * Loading state for read operations (`get`, `list`).
     * True during any active read call.
     */
    readonly isLoading: boolean;

    /**
     * Read error from the latest-started read; cleared when another read starts. Discovery, access, schema, item and paging failures publish this error and return the read fallback.
     */
    readonly error: Error | undefined;

    /**
     * Loading state for write operations (`save`, `remove`).
     * True during any active write call.
     */
    readonly isWriting: boolean;

    /**
     * Write error from the latest-started write; cleared when another write starts.
     */
    readonly writeError: Error | undefined;

    /**
     * Effective Add/Edit/Delete Items indicator for the catalog list. When absent, catalog web Manage Lists is also required. Initially false; checked on discovery and refreshed after operations. Failed checks return false; latest check wins. SharePoint enforces writes and schema repair can require additional permissions.
     */
    readonly canWrite: boolean;

    /**
     * Computed state: true when the SPHttpClient is available and
     * operations can be requested. Does not imply completed discovery, an existing list or write authorization.
     */
    readonly isReady: boolean;

    /**
     * Reads without provisioning or repairing the list. Only a confirmed absent list/key is empty; discovery/access/schema/read failures publish error and return undefined.
     * @param key - Stored property key.
     * @returns Item or undefined for an absent list/key or failed read; inspect error to distinguish failures.
     * @throws When the SharePoint client is unavailable.
     */
    readonly get: <T = unknown>(key: string) => Promise<SPFxTenantKeyValueStoreItem<T> | undefined>;

    /**
     * Reads every server page in Title order without provisioning. Discovery/access/schema/item/paging failures publish error and return []; no partial results are returned.
     * @returns All items, an empty array for a confirmed absent list, or an empty fallback with error.
     * @throws When the SharePoint client is unavailable.
     */
    readonly list: () => Promise<SPFxTenantKeyValueStoreItem<unknown>[]>;

    /**
     * Save (create or update) a key-value pair.
     * Auto-provisions the hidden list on first write if it doesn't exist.
     *
     * Smart serialization:
     * - Empty strings and strings that parse as JSON are JSON-quoted to preserve their text
     * - bigint becomes quoted decimal text and reads as string without precision loss
     * - Numbers/booleans and legacy null input retain their string representation
     * - Date objects → ISO 8601 string
     * - Objects/arrays → JSON.stringify(value)
     *
     * @param key - Property key
     * @param value - Value to store
     * @param description - Optional description metadata
     *
     * @example
     * ```tsx
     * // Inside a React component beneath a matching host-specific SPFx provider:
     * const store = useSPFxTenantKeyValueStore();
     *
     * // String value
     * // Call this handler from an event or an effect; do not request during render.
     * const run = async () => {
     *   await store.save<string>('apiEndpoint', 'https://api.example.com', 'Production API');
     *
     *   // Complex object
     *   await store.save('featureFlags', { enableChat: true, maxUsers: 1000 });
     * };
     * ```
     * @returns Resolves after a successful write; updates preserve an omitted description and empty string clears it.
     * @throws Client, catalog discovery, schema, serialization, missing ETag or write/conflict failures; operation failures also publish writeError. No wildcard overwrite is performed.
     */
    readonly save: <T = unknown>(key: string, value: T, description?: string) => Promise<void>;

    /**
     * Deletes without provisioning or schema repair. A confirmed absent list/key is a no-op; changes since the service snapshot reject using ETag preconditions.
     * @param key - Key to remove.
     * @returns Resolves after successful deletion or absent key.
     * @throws Client, discovery, schema, missing ETag or deletion/conflict failures, also captured in writeError.
     */
    readonly remove: (key: string) => Promise<void>;
}

// ═══════════════════════════════════════════════════════════════════════════
// HOOK
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Manages the hidden TenantKeyValueStore list in the tenant app catalog beneath an SPFx provider.
 * Mount discovers the catalog and checks effective permissions without provisioning. Reads/removes
 * never create or repair schema; save ensures validated fields and unique indexed keys. isReady
 * means client availability only. canWrite is an asynchronous effective-permission indicator.
 * New string writes preserve text and bigint reads as exact decimal text; Date remains ISO text.
 * Existing raw data retains its legacy JSON-first interpretation, and T performs no validation.
 * Read failures publish error and return fallbacks; write failures publish writeError and reject.
 * ETag preconditions detect changes during mutations, not edits since a caller's earlier get().
 * Context changes reset state; obsolete/unmounted completions cannot publish state.
 * @returns Client readiness, read/write state, permission indicator and CRUD methods.
 * @throws When called outside an SPFx provider.
 *
 * @example
 * ```tsx
 * function TenantSetting() {
 *   const store = useSPFxTenantKeyValueStore();
 *   const load = async () => console.log(await store.get<boolean>('featureEnabled'));
 *   return <button onClick={load} disabled={!store.isReady}>Load setting</button>;
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

    const operationRef = useRef({ service: tenantKeyValueStoreService, discover: discoverAppCatalogUrl,
        reads: 0, writes: 0, readId: 0, writeId: 0, permissionId: 0 });
    if (operationRef.current.service !== tenantKeyValueStoreService ||
        operationRef.current.discover !== discoverAppCatalogUrl) {
        operationRef.current = { service: tenantKeyValueStoreService, discover: discoverAppCatalogUrl,
            reads: 0, writes: 0, readId: 0, writeId: 0, permissionId: 0 };
    }
    const currentOperation = operationRef.current;
    useEffect(() => {
        setIsLoading(false);
        setIsWriting(false);
        setError(undefined);
        setWriteError(undefined);
        setCanWrite(false);
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl]);

    // ─────────────────────────────────────────────────────────────────────────
    // Internal helpers
    // ─────────────────────────────────────────────────────────────────────────



    /**
     * Refresh the effective-permission indicator; only the latest check may publish.
     */
    const ensurePermissionChecked = useCallback((catalogUrl: string): void => {
        const identity = currentOperation;
        const request = ++identity.permissionId;
        checkWritePermission(catalogUrl)
            .then(hasPermission => {
                if (isMountedRef.current && operationRef.current === identity && request === identity.permissionId) {
                    setCanWrite(hasPermission);
                }
            })
            .catch(() => {
                if (isMountedRef.current && operationRef.current === identity && request === identity.permissionId) {
                    setCanWrite(false);
                }
            });
    }, [checkWritePermission, isMountedRef, currentOperation]);

    // Discovery and permission lookup are read-only; provisioning belongs to save.
    useEffect(() => {
        if (!tenantKeyValueStoreService) return;

        discoverAppCatalogUrl()
            .then(catalogUrl => {
                if (!isMountedRef.current || operationRef.current !== currentOperation) return;
                ensurePermissionChecked(catalogUrl);
            })
            .catch(() => {
                // Explicit operations retry discovery and publish failures in their own error channel.
            });
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, ensurePermissionChecked, isMountedRef, currentOperation]);

    // ─────────────────────────────────────────────────────────────────────────
    // Public operations
    // ─────────────────────────────────────────────────────────────────────────

    const get = useCallback(async <T = unknown>(key: string): Promise<SPFxTenantKeyValueStoreItem<T> | undefined> => {
        if (!tenantKeyValueStoreService) {
            throw new Error('SPHttpClient not available. Cannot get property.');
        }

        const identity = currentOperation;
        const request = ++identity.readId;
        const isCurrent = (): boolean => isMountedRef.current && operationRef.current === identity &&
            identity.service === tenantKeyValueStoreService;
        identity.reads++;
        if (isCurrent()) { setIsLoading(true); setError(undefined); }

        let catalogUrl: string | undefined;
        try {
            catalogUrl = await discoverAppCatalogUrl();

            return await tenantKeyValueStoreService.get<T>(key, catalogUrl);
        } catch (err) {
            if (isCurrent() && request === identity.readId) {
                const capturedError = err instanceof Error ? err : new Error(String(err));
                setError(capturedError);
                console.error('Failed to get tenant key-value:', capturedError);
            }
            return undefined;
        } finally {
            identity.reads--;
            if (isCurrent()) {
                setIsLoading(identity.reads > 0);
                if (catalogUrl) ensurePermissionChecked(catalogUrl);
            }
        }
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef, currentOperation, ensurePermissionChecked]);

    const list = useCallback(async (): Promise<SPFxTenantKeyValueStoreItem<unknown>[]> => {
        if (!tenantKeyValueStoreService) {
            throw new Error('SPHttpClient not available. Cannot list properties.');
        }

        const identity = currentOperation;
        const request = ++identity.readId;
        const isCurrent = (): boolean => isMountedRef.current && operationRef.current === identity &&
            identity.service === tenantKeyValueStoreService;
        identity.reads++;
        if (isCurrent()) { setIsLoading(true); setError(undefined); }

        let catalogUrl: string | undefined;
        try {
            catalogUrl = await discoverAppCatalogUrl();

            return await tenantKeyValueStoreService.list(catalogUrl);
        } catch (err) {
            if (isCurrent() && request === identity.readId) {
                const capturedError = err instanceof Error ? err : new Error(String(err));
                setError(capturedError);
                console.error('Failed to list tenant key-values:', capturedError);
            }
            return [];
        } finally {
            identity.reads--;
            if (isCurrent()) {
                setIsLoading(identity.reads > 0);
                if (catalogUrl) ensurePermissionChecked(catalogUrl);
            }
        }
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef, currentOperation, ensurePermissionChecked]);

    const save = useCallback(async <T = unknown>(
        key: string,
        value: T,
        description?: string
    ): Promise<void> => {
        if (!tenantKeyValueStoreService) {
            throw new Error('SPHttpClient not available. Cannot save property.');
        }

        const identity = currentOperation;
        const request = ++identity.writeId;
        const isCurrent = (): boolean => isMountedRef.current && operationRef.current === identity &&
            identity.service === tenantKeyValueStoreService;
        identity.writes++;
        if (isCurrent()) { setIsWriting(true); setWriteError(undefined); }

        let catalogUrl: string | undefined;
        try {
            catalogUrl = await discoverAppCatalogUrl();
            await tenantKeyValueStoreService.save<T>(key, value, catalogUrl, description);
        } catch (err) {
            if (isCurrent() && request === identity.writeId) {
                const capturedError = err instanceof Error ? err : new Error(String(err));
                setWriteError(capturedError);
                console.error('Failed to save tenant key-value:', capturedError);
            }
            throw err;
        } finally {
            identity.writes--;
            if (isCurrent()) {
                setIsWriting(identity.writes > 0);
                if (catalogUrl) ensurePermissionChecked(catalogUrl);
            }
        }
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef, currentOperation, ensurePermissionChecked]);

    const remove = useCallback(async (key: string): Promise<void> => {
        if (!tenantKeyValueStoreService) {
            throw new Error('SPHttpClient not available. Cannot remove property.');
        }

        const identity = currentOperation;
        const request = ++identity.writeId;
        const isCurrent = (): boolean => isMountedRef.current && operationRef.current === identity &&
            identity.service === tenantKeyValueStoreService;
        identity.writes++;
        if (isCurrent()) { setIsWriting(true); setWriteError(undefined); }

        let catalogUrl: string | undefined;
        try {
            catalogUrl = await discoverAppCatalogUrl();
            await tenantKeyValueStoreService.remove(key, catalogUrl);
        } catch (err) {
            if (isCurrent() && request === identity.writeId) {
                const capturedError = err instanceof Error ? err : new Error(String(err));
                setWriteError(capturedError);
                console.error('Failed to remove tenant key-value:', capturedError);
            }
            throw err;
        } finally {
            identity.writes--;
            if (isCurrent()) {
                setIsWriting(identity.writes > 0);
                if (catalogUrl) ensurePermissionChecked(catalogUrl);
            }
        }
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef, currentOperation, ensurePermissionChecked]);

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

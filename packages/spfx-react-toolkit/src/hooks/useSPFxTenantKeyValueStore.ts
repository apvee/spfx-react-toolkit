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
     * Read error from the latest-started read; cleared when another read starts. Setup failures inside get/list return fallbacks without setting this error.
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
     * App catalog IsSiteAdmin indicator, checked asynchronously during mount/discovery or operations. SharePoint still enforces writes.
     */
    readonly canWrite: boolean;

    /**
     * Computed state: true when the SPHttpClient is available and
     * the hook is ready to perform operations.
     */
    readonly isReady: boolean;

    /**
     * Ensures setup, then reads a key. Setup failure returns undefined without publishing error; discovery/read failures publish error and return undefined.
     * @param key - Stored property key.
     * @returns Item or undefined for absent keys or failed reads/setup.
     * @throws When the SharePoint client is unavailable.
     */
    readonly get: <T = unknown>(key: string) => Promise<SPFxTenantKeyValueStoreItem<T> | undefined>;

    /**
     * Ensures setup, then reads the first server page in Title order; does not follow continuation links. Setup failure returns [] silently; discovery/read failures publish error and return [].
     * @returns First page of items, or an empty fallback.
     * @throws When the SharePoint client is unavailable.
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
     * // Inside a React component beneath a matching host-specific SPFx provider:
     * const store = useSPFxTenantKeyValueStore();
     *
     * // String value
     * // Call this handler from an event or an effect; do not request during render.
     * const run = async () => {
     *   await store.save<string>('apiEndpoint', 'https://api.example.com', 'Production API');
     *
     *   // Complex object
     *   await store.save<FeatureFlags>('featureFlags', { enableChat: true, maxUsers: 1000 });
     * };
     * ```
     * @returns Resolves after a successful write; updates preserve an omitted description and empty string clears it.
     * @throws Client, catalog discovery, provisioning, serialization or write failures; operation failures also publish writeError.
     */
    readonly save: <T = unknown>(key: string, value: T, description?: string) => Promise<void>;

    /**
     * Ensures list setup before deleting a key; an absent key is a no-op.
     * @param key - Key to remove.
     * @returns Resolves after successful deletion or absent key.
     * @throws Client, discovery, provisioning or deletion failures, also captured in writeError.
     */
    readonly remove: (key: string) => Promise<void>;
}

// ═══════════════════════════════════════════════════════════════════════════
// HOOK
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Manages a hidden TenantKeyValueStore list in the tenant app catalog. Discovery, permission lookup and setup run on mount; each operation can also ensure setup. Service readiness is cached per catalog URL. isReady indicates client availability, not completed discovery or write authorization. SharePoint enforces actual grants; canWrite uses the catalog IsSiteAdmin flag. Values use legacy tenant serialization: JSON-like strings may be coerced and Date/bigint types do not round-trip. Read setup failures return fallbacks silently; other read failures publish error. Write failures reject. Context changes reset state and obsolete/unmounted completions cannot publish it.
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

    const permissionCheckedRef = useRef<boolean>(false);
    const operationRef = useRef({ service: tenantKeyValueStoreService, discover: discoverAppCatalogUrl,
        reads: 0, writes: 0, readId: 0, writeId: 0 });
    if (operationRef.current.service !== tenantKeyValueStoreService ||
        operationRef.current.discover !== discoverAppCatalogUrl) {
        operationRef.current = { service: tenantKeyValueStoreService, discover: discoverAppCatalogUrl,
            reads: 0, writes: 0, readId: 0, writeId: 0 };
        permissionCheckedRef.current = false;
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
     * Check and update write permission (executed once)
     */
    const ensurePermissionChecked = useCallback((catalogUrl: string): void => {
        if (permissionCheckedRef.current) return;
        permissionCheckedRef.current = true;
        const identity = currentOperation;
        checkWritePermission(catalogUrl)
            .then(hasPermission => {
                if (isMountedRef.current && operationRef.current === identity) {
                    setCanWrite(hasPermission);
                }
            })
            .catch(() => {
                if (isMountedRef.current && operationRef.current === identity) {
                    setCanWrite(false);
                }
            });
    }, [checkWritePermission, isMountedRef, currentOperation]);

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
                if (!isMountedRef.current || operationRef.current !== currentOperation) return;
                ensurePermissionChecked(catalogUrl);
                return tenantKeyValueStoreService.ensureListReady(catalogUrl);
            })
            .catch(() => {
                // Silently fail — operations will retry on demand via their own ensureListReady call
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

        try {
            const catalogUrl = await discoverAppCatalogUrl();

            // Ensure list exists (auto-provision if needed; fallback-safe for read-only users)
            try {
                await tenantKeyValueStoreService.ensureListReady(catalogUrl);
            } catch {
                // Provisioning may fail for read-only users — treat as "list not yet created"
                return undefined;
            }

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
            if (isCurrent()) setIsLoading(identity.reads > 0);
        }
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef, currentOperation]);

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

        try {
            const catalogUrl = await discoverAppCatalogUrl();

            // Ensure list exists (auto-provision if needed; fallback-safe for read-only users)
            try {
                await tenantKeyValueStoreService.ensureListReady(catalogUrl);
            } catch {
                // Provisioning may fail for read-only users — treat as "list not yet created"
                return [];
            }

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
            if (isCurrent()) setIsLoading(identity.reads > 0);
        }
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef, currentOperation]);

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

        try {
            const catalogUrl = await discoverAppCatalogUrl();
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
            if (isCurrent()) setIsWriting(identity.writes > 0);
        }
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef, currentOperation]);

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

        try {
            const catalogUrl = await discoverAppCatalogUrl();
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
            if (isCurrent()) setIsWriting(identity.writes > 0);
        }
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef, currentOperation]);

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

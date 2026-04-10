// useSPFxTenantProperty.ts
// Hook to read tenant-wide properties using SharePoint StorageEntity API (read-only)

import { useState, useCallback, useEffect } from 'react';
import { useAppCatalogUrl } from './useAppCatalogUrl.internal';
import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * Return type for useSPFxTenantProperty hook
 */
export interface SPFxTenantPropertyResult<T> {
  /** 
   * The loaded property value from tenant app catalog.
   * Undefined if not loaded yet or on error.
   */
  readonly data: T | undefined;
  
  /** 
   * Property description metadata (optional).
   * SharePoint StorageEntity only supports description, not comment.
   */
  readonly description: string | undefined;
  
  /** 
   * Loading state for read operations.
   * True during initial load or manual load() calls.
   */
  readonly isLoading: boolean;
  
  /** 
   * Last error from read operations.
   * Cleared on successful load.
   */
  readonly error: Error | undefined;
  
  /** 
   * Manually load/reload the property from tenant app catalog.
   * Updates data, description, isLoading, and error states.
   * 
   * @returns Promise that resolves when load completes
   * 
   * @example
   * ```tsx
   * const { data, load } = useSPFxTenantProperty<string>('appVersion', false);
   * 
   * // Load on button click
   * <button onClick={load}>Refresh</button>
   * ```
   */
  readonly load: () => Promise<void>;
  
  /** 
   * Computed state: true if data is loaded successfully.
   * Equivalent to: !isLoading && !error && data !== undefined
   * 
   * Useful for conditional rendering:
   * ```tsx
   * if (!isReady) return <Spinner />;
   * return <div>{data}</div>;
   * ```
   */
  readonly isReady: boolean;
}

/**
 * SharePoint StorageEntity response interface
 */
interface IStorageEntity {
  Value: string;
  Description?: string;
}

/**
 * Hook to read tenant-wide properties using SharePoint StorageEntity API (read-only)
 * 
 * Provides read access to tenant-scoped properties stored in the SharePoint tenant
 * app catalog. Properties are accessible across all sites in the tenant.
 * 
 * **Note:** Write and remove operations have been removed because Microsoft has
 * blocked the SetStorageEntity and RemoveStorageEntity REST API endpoints.
 * Tenant properties can only be managed via PowerShell (Set-PnPStorageEntity,
 * Remove-PnPStorageEntity) or the SharePoint Management Shell.
 * For a read/write key-value store at tenant level, use `useSPFxTenantKeyValueStore`.
 * 
 * Features:
 * - Tenant-wide centralized storage (not site-specific)
 * - Smart deserialization for primitives and complex objects
 * - Optional metadata (description only - SharePoint limitation)
 * - Type-safe with TypeScript generics
 * - Memory leak safe with mounted state tracking
 * - Automatic app catalog URL discovery
 * 
 * Requirements:
 * - Tenant app catalog must be provisioned
 * - Any authenticated user can read tenant properties
 * 
 * @param key - Unique property key (e.g., 'appVersion', 'apiEndpoint', 'featureFlags')
 * @param autoFetch - Whether to automatically load property on mount. Default: true
 * 
 * @returns Object with data, metadata, loading state, error state, and load function
 * 
 * @example Basic usage - string property
 * ```tsx
 * function VersionDisplay() {
 *   const { data, isLoading, error } = useSPFxTenantProperty<string>('appVersion');
 *   
 *   if (isLoading) return <Spinner label="Loading version..." />;
 *   if (error) return <MessageBar messageBarType={MessageBarType.error}>
 *     Failed to load: {error.message}
 *   </MessageBar>;
 *   
 *   return <Text>Current Version: {data ?? 'Not Set'}</Text>;
 * }
 * ```
 * 
 * @example Complex object with JSON
 * ```tsx
 * interface FeatureFlags {
 *   enableChat: boolean;
 *   enableAnalytics: boolean;
 *   maxUsers: number;
 * }
 * 
 * const { data, isLoading } = useSPFxTenantProperty<FeatureFlags>('featureFlags');
 * 
 * // Returns parsed object (stored as JSON string via PowerShell)
 * if (data?.enableChat) {
 *   return <ChatPanel />;
 * }
 * ```
 * 
 * @example With metadata viewing
 * ```tsx
 * function PropertyViewer() {
 *   const { data, description, isLoading } = 
 *     useSPFxTenantProperty<string>('appConfig');
 *   
 *   if (isLoading) return <Spinner />;
 *   
 *   return (
 *     <Stack tokens={{ childrenGap: 5 }}>
 *       <Text variant="large">Value: {data}</Text>
 *       {description && <Text variant="small">Description: {description}</Text>}
 *     </Stack>
 *   );
 * }
 * ```
 * 
 * @example Lazy loading with manual trigger
 * ```tsx
 * const { data, load, isLoading } = useSPFxTenantProperty<Config>(
 *   'appConfig',
 *   false  // Don't auto-fetch
 * );
 * 
 * return (
 *   <div>
 *     <button onClick={load} disabled={isLoading}>
 *       {isLoading ? 'Loading...' : 'Load Config'}
 *     </button>
 *     {data && <ConfigDisplay config={data} />}
 *   </div>
 * );
 * ```
 * 
 * @example Multi-property dashboard
 * ```tsx
 * function TenantDashboard() {
 *   const version = useSPFxTenantProperty<string>('appVersion');
 *   const maintenance = useSPFxTenantProperty<boolean>('maintenanceMode');
 *   const config = useSPFxTenantProperty<AppConfig>('appConfig');
 *   
 *   const isLoading = version.isLoading || maintenance.isLoading || config.isLoading;
 *   
 *   if (isLoading) return <Spinner label="Loading dashboard..." />;
 *   
 *   return (
 *     <Stack tokens={{ childrenGap: 20 }}>
 *       <Text variant="xxLarge">Tenant Configuration</Text>
 *       <Label>App Version: {version.data ?? 'Not Set'}</Label>
 *       <Label>Maintenance Mode: {maintenance.data ? 'ON' : 'OFF'}</Label>
 *       {config.data && <pre>{JSON.stringify(config.data, null, 2)}</pre>}
 *     </Stack>
 *   );
 * }
 * ```
 */
export function useSPFxTenantProperty<T = unknown>(
  key: string,
  autoFetch: boolean = true
): SPFxTenantPropertyResult<T> {
  const { spHttpClient, discoverAppCatalogUrl, isMountedRef } = useAppCatalogUrl();
  
  // State management
  const [data, setData] = useState<T | undefined>(undefined);
  const [description, setDescription] = useState<string | undefined>(undefined);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  
  /**
   * Deserialize value from storage
   * - Try JSON.parse first
   * - If fails, return raw string (will be cast to T by TypeScript)
   */
  const deserializeValue = useCallback((rawValue: string): T => {
    try {
      return JSON.parse(rawValue) as T;
    } catch {
      // Not valid JSON, assume it's a primitive value
      return rawValue as T;
    }
  }, []);
  
  /**
   * Load property from tenant app catalog
   */
  const load = useCallback(async (): Promise<void> => {
    if (!spHttpClient) {
      console.warn('SPHttpClient not available yet. Skipping load.');
      return;
    }
    
    if (!key) {
      console.warn('key is required. Skipping load.');
      return;
    }
    
    setIsLoading(true);
    setError(undefined);
    
    try {
      // Discover app catalog URL
      const catalogUrl = await discoverAppCatalogUrl();
      
      // Read property
      const response: SPHttpClientResponse = await spHttpClient.get(
        `${catalogUrl}/_api/web/GetStorageEntity('${encodeURIComponent(key)}')`,
        SPHttpClient.configurations.v1
      );
      
      if (!response.ok) {
        throw new Error(`Failed to read property: ${response.statusText}`);
      }
      
      const entity: IStorageEntity = await response.json();
      
      if (isMountedRef.current) {
        if (entity.Value) {
          setData(deserializeValue(entity.Value));
          setDescription(entity.Description);
        } else {
          // Property doesn't exist or has no value
          setData(undefined);
          setDescription(undefined);
        }
      }
    } catch (err) {
      if (isMountedRef.current) {
        const capturedError = err instanceof Error ? err : new Error(String(err));
        setError(capturedError);
        console.error('Failed to load tenant property:', capturedError);
      }
    } finally {
      if (isMountedRef.current) {
        setIsLoading(false);
      }
    }
  }, [spHttpClient, key, discoverAppCatalogUrl, deserializeValue, isMountedRef]);
  
  // Auto-fetch on mount if enabled
  useEffect(() => {
    if (autoFetch && spHttpClient && key) {
      load().catch(() => {
        // Error already handled in load() function
      });
    }
  }, [autoFetch, spHttpClient, key, load]);
  
  // Computed state: ready when data loaded successfully
  const isReady = !isLoading && !error && data !== undefined;
  
  return {
    data,
    description,
    isLoading,
    error,
    load,
    isReady,
  };
}

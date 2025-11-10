// useSPFxTenantProperty.ts
// Hook to manage tenant-wide properties using SharePoint StorageEntity API

import { useState, useCallback, useEffect, useRef } from 'react';
import { useSPFxSPHttpClient } from './useSPFxSPHttpClient';
import { useSPFxPageContext } from './useSPFxPageContext';
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
   * Cleared on successful load or write.
   */
  readonly error: Error | undefined;
  
  /** 
   * Loading state for write operations.
   * True during write() calls.
   */
  readonly isWriting: boolean;
  
  /** 
   * Last error from write operations.
   * Cleared on successful write or load.
   */
  readonly writeError: Error | undefined;
  
  /** 
   * Whether the current user has permission to write tenant properties.
   * False if user lacks Manage Web permissions on tenant app catalog.
   */
  readonly canWrite: boolean;
  
  /** 
   * Manually load/reload the property from tenant app catalog.
   * Updates data, description, comment, isLoading, and error states.
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
   * Write property to tenant app catalog.
   * Creates property if it doesn't exist, updates if it does.
   * Updates isWriting and writeError states.
   * 
   * Smart serialization:
   * - Primitives (string, number, boolean, null, bigint) → String(content)
   * - Date objects → ISO string
   * - Objects/arrays → JSON.stringify(content)
   * 
   * @param content - Data to write
   * @param description - Optional description metadata (SharePoint only supports description, not comment)
   * @returns Promise that resolves when write completes
   * 
   * @example
   * ```tsx
   * const { write, isWriting } = useSPFxTenantProperty<string>('apiEndpoint');
   * 
   * const handleSave = async () => {
   *   await write('https://api.example.com', 'Production API endpoint');
   * };
   * ```
   */
  readonly write: (content: T, description?: string) => Promise<void>;
  
  /** 
   * Remove property from tenant app catalog.
   * Requires Manage Web permissions.
   * 
   * @returns Promise that resolves when removal completes
   * 
   * @example
   * ```tsx
   * const { remove } = useSPFxTenantProperty<string>('oldSetting');
   * 
   * const handleDelete = async () => {
   *   if (confirm('Delete this property?')) {
   *     await remove();
   *   }
   * };
   * ```
   */
  readonly remove: () => Promise<void>;
  
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
 * Note: SharePoint SetStorageEntity API only supports Value and Description
 */
interface IStorageEntity {
  Value: string;
  Description?: string;
}

/**
 * Tenant app catalog URL response
 */
interface ITenantAppCatalogResponse {
  CorporateCatalogUrl: string;
}

/**
 * Hook to manage tenant-wide properties using SharePoint StorageEntity API
 * 
 * Provides read/write operations for tenant-scoped properties stored in the
 * SharePoint tenant app catalog. Properties are accessible across all sites
 * in the tenant and support metadata (description only).
 * 
 * Features:
 * - Tenant-wide centralized storage (not site-specific)
 * - Smart serialization for primitives, Date, and complex objects
 * - Permission checking (canWrite flag)
 * - Optional metadata (description only - SharePoint limitation)
 * - Type-safe with TypeScript generics
 * - Memory leak safe with mounted state tracking
 * - Automatic app catalog URL discovery
 * - Remove operation for cleanup
 * 
 * Requirements:
 * - Tenant app catalog must be provisioned
 * - Read: Any authenticated user
 * - Write/Remove: Site Collection Administrator role on tenant app catalog site
 * 
 * Permission Notes:
 * - Being a Tenant Admin is NOT sufficient for write operations
 * - You must be explicitly added as Site Collection Administrator to the tenant app catalog
 * - Navigate to the app catalog site → Site Settings → Site Collection Administrators
 * - Add your user account if write operations fail with permission errors
 * 
 * Storage Format:
 * - Primitives (string, number, boolean, null) → stored as string
 * - Date → stored as ISO 8601 string
 * - Objects/arrays → stored as JSON string
 * 
 * @param key - Unique property key (e.g., 'appVersion', 'apiEndpoint', 'featureFlags')
 * @param autoFetch - Whether to automatically load property on mount. Default: true
 * 
 * @returns Object with data, metadata, loading states, error states, and CRUD functions
 * 
 * @example Basic usage - string property
 * ```tsx
 * function VersionDisplay() {
 *   const { data, isLoading, error, write, canWrite } = 
 *     useSPFxTenantProperty<string>('appVersion');
 *   
 *   if (isLoading) return <Spinner label="Loading version..." />;
 *   if (error) return <MessageBar messageBarType={MessageBarType.error}>
 *     Failed to load: {error.message}
 *   </MessageBar>;
 *   
 *   const handleUpdate = async () => {
 *     if (!canWrite) {
 *       alert('Insufficient permissions');
 *       return;
 *     }
 *     
 *     try {
 *       await write('2.0.1', 'Current application version');
 *       console.log('Version updated!');
 *     } catch (err) {
 *       console.error('Update failed:', err);
 *     }
 *   };
 *   
 *   return (
 *     <div>
 *       <Text>Current Version: {data ?? 'Not Set'}</Text>
 *       {canWrite && <PrimaryButton onClick={handleUpdate}>Update Version</PrimaryButton>}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Number property
 * ```tsx
 * const { data, write } = useSPFxTenantProperty<number>('maxUploadSize');
 * 
 * // Stored as "10485760" (10MB in bytes)
 * await write(10485760, 'Maximum file upload size in bytes');
 * 
 * // Read returns: 10485760 (number)
 * console.log(typeof data); // "number"
 * ```
 * 
 * @example Boolean flag
 * ```tsx
 * const { data: maintenanceMode, write } = useSPFxTenantProperty<boolean>('maintenanceMode');
 * 
 * // Stored as "true" or "false"
 * await write(true, 'Maintenance mode enabled');
 * 
 * if (maintenanceMode) {
 *   return <MessageBar>System is under maintenance</MessageBar>;
 * }
 * ```
 * 
 * @example Date property
 * ```tsx
 * const { data, write } = useSPFxTenantProperty<string>('lastDeployment');
 * 
 * // Write date as ISO string
 * await write(new Date().toISOString(), 'Last deployment timestamp');
 * 
 * // Read and convert back to Date
 * const lastDeploy = data ? new Date(data) : undefined;
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
 * const { data, write, isLoading } = useSPFxTenantProperty<FeatureFlags>('featureFlags');
 * 
 * // Stored as JSON string
 * await write({
 *   enableChat: true,
 *   enableAnalytics: false,
 *   maxUsers: 1000
 * }, 'Global feature flags configuration');
 * 
 * // Read returns parsed object
 * if (data?.enableChat) {
 *   return <ChatPanel />;
 * }
 * ```
 * 
 * @example Permission-aware UI
 * ```tsx
 * function TenantConfigPanel() {
 *   const { data, canWrite, write, isWriting, error, writeError } = 
 *     useSPFxTenantProperty<string>('apiEndpoint');
 *   
 *   const [editValue, setEditValue] = React.useState(data ?? '');
 *   
 *   React.useEffect(() => {
 *     setEditValue(data ?? '');
 *   }, [data]);
 *   
 *   if (!canWrite) {
 *     return (
 *       <MessageBar messageBarType={MessageBarType.info}>
 *         You don't have permission to edit tenant properties.
 *         Contact your SharePoint administrator.
 *       </MessageBar>
 *     );
 *   }
 *   
 *   const handleSave = async () => {
 *     try {
 *       await write(editValue, 'Production API endpoint URL');
 *     } catch (err) {
 *       console.error('Save failed:', err);
 *     }
 *   };
 *   
 *   return (
 *     <Stack tokens={{ childrenGap: 10 }}>
 *       <TextField 
 *         label="API Endpoint"
 *         value={editValue}
 *         onChange={(_, val) => setEditValue(val ?? '')}
 *         disabled={isWriting}
 *       />
 *       <PrimaryButton 
 *         onClick={handleSave} 
 *         disabled={isWriting || editValue === data}
 *       >
 *         {isWriting ? 'Saving...' : 'Save'}
 *       </PrimaryButton>
 *       {writeError && (
 *         <MessageBar messageBarType={MessageBarType.error}>
 *           {writeError.message}
 *         </MessageBar>
 *       )}
 *     </Stack>
 *   );
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
 * const { data, load, isLoading, write } = useSPFxTenantProperty<Config>(
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
 * @example Property removal
 * ```tsx
 * function PropertyManager() {
 *   const { data, remove, canWrite } = useSPFxTenantProperty<string>('deprecatedSetting');
 *   
 *   const handleDelete = async () => {
 *     if (!confirm('Delete this property? This cannot be undone.')) return;
 *     
 *     try {
 *       await remove();
 *       console.log('Property deleted');
 *     } catch (err) {
 *       console.error('Delete failed:', err);
 *     }
 *   };
 *   
 *   if (!data) return <Text>Property not found</Text>;
 *   
 *   return (
 *     <Stack tokens={{ childrenGap: 10 }}>
 *       <Text>Value: {data}</Text>
 *       {canWrite && (
 *         <DefaultButton onClick={handleDelete} text="Delete Property" />
 *       )}
 *     </Stack>
 *   );
 * }
 * ```
 * 
 * @example Multi-property dashboard
 * ```tsx
 * function TenantDashboard() {
 *   const version = useSPFxTenantProperty<string>('appVersion');
 *   const maintenance = useSPFxTenantProperty<boolean>('maintenanceMode');
 *   const lastUpdate = useSPFxTenantProperty<string>('lastUpdate');
 *   const config = useSPFxTenantProperty<AppConfig>('appConfig');
 *   
 *   const isLoading = version.isLoading || maintenance.isLoading || 
 *                     lastUpdate.isLoading || config.isLoading;
 *   
 *   if (isLoading) return <Spinner label="Loading dashboard..." />;
 *   
 *   return (
 *     <Stack tokens={{ childrenGap: 20 }}>
 *       <Text variant="xxLarge">Tenant Configuration</Text>
 *       
 *       <Stack tokens={{ childrenGap: 10 }}>
 *         <Label>App Version: {version.data ?? 'Not Set'}</Label>
 *         <Label>Maintenance Mode: {maintenance.data ? 'ON' : 'OFF'}</Label>
 *         <Label>Last Update: {lastUpdate.data ? new Date(lastUpdate.data).toLocaleString() : 'Never'}</Label>
 *       </Stack>
 *       
 *       {config.data && (
 *         <Stack tokens={{ childrenGap: 5 }}>
 *           <Text variant="large">Configuration</Text>
 *           <pre>{JSON.stringify(config.data, null, 2)}</pre>
 *         </Stack>
 *       )}
 *     </Stack>
 *   );
 * }
 * ```
 */
export function useSPFxTenantProperty<T = unknown>(
  key: string,
  autoFetch: boolean = true
): SPFxTenantPropertyResult<T> {
  const { client: spHttpClient } = useSPFxSPHttpClient();
  const pageContext = useSPFxPageContext();
  
  // State management
  const [data, setData] = useState<T | undefined>(undefined);
  const [description, setDescription] = useState<string | undefined>(undefined);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  const [isWriting, setIsWriting] = useState<boolean>(false);
  const [writeError, setWriteError] = useState<Error | undefined>(undefined);
  const [canWrite, setCanWrite] = useState<boolean>(false);
  const [appCatalogUrl, setAppCatalogUrl] = useState<string | undefined>(undefined);
  
  // Track component mounted state to prevent memory leaks
  const isMounted = useRef<boolean>(true);
  
  useEffect(() => {
    isMounted.current = true;
    return () => {
      isMounted.current = false;
    };
  }, []);
  
  /**
   * Check if value is a primitive type (including null, Date)
   */
  const isPrimitive = useCallback((val: unknown): boolean => {
    if (val === null) return true;
    const type = typeof val;
    // ES5-compatible array check
    const primitiveTypes = ['string', 'number', 'boolean', 'bigint'];
    if (primitiveTypes.indexOf(type) !== -1) return true;
    // Date is special: treat as primitive and convert to ISO string
    if (val instanceof Date) return true;
    return false;
  }, []);
  
  /**
   * Serialize content for storage
   * - Primitives → String(content)
   * - Date → ISO string
   * - Objects/arrays → JSON.stringify()
   */
  const serializeValue = useCallback((content: T): string => {
    if (isPrimitive(content)) {
      return content instanceof Date ? content.toISOString() : String(content);
    }
    return JSON.stringify(content);
  }, [isPrimitive]);
  
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
   * Discover tenant app catalog URL
   */
  const discoverAppCatalogUrl = useCallback(async (): Promise<string> => {
    if (!spHttpClient || !pageContext) {
      throw new Error('SPHttpClient or PageContext not available');
    }
    
    // Check cache first
    if (appCatalogUrl) {
      return appCatalogUrl;
    }
    
    try {
      const response: SPHttpClientResponse = await spHttpClient.get(
        `${pageContext.web.absoluteUrl}/_api/SP_TenantSettings_Current`,
        SPHttpClient.configurations.v1
      );
      
      if (!response.ok) {
        throw new Error(`Failed to discover app catalog: ${response.statusText}`);
      }
      
      const data: ITenantAppCatalogResponse = await response.json();
      
      if (!data.CorporateCatalogUrl) {
        throw new Error('Tenant app catalog is not provisioned. Please provision the app catalog first.');
      }
      
      if (isMounted.current) {
        setAppCatalogUrl(data.CorporateCatalogUrl);
      }
      
      return data.CorporateCatalogUrl;
    } catch (err) {
      throw new Error(`App catalog discovery failed: ${err instanceof Error ? err.message : String(err)}`);
    }
  }, [spHttpClient, pageContext, appCatalogUrl]);
  
  /**
   * Check write permissions on tenant app catalog
   * User must be Site Collection Administrator of the tenant app catalog
   */
  const checkWritePermission = useCallback(async (catalogUrl: string): Promise<boolean> => {
    if (!spHttpClient) return false;
    
    try {
      // Check if user is site collection admin
      const response: SPHttpClientResponse = await spHttpClient.get(
        `${catalogUrl}/_api/web/currentuser?$select=IsSiteAdmin`,
        SPHttpClient.configurations.v1
      );
      
      if (!response.ok) return false;
      
      const user = await response.json();
      return user.IsSiteAdmin === true;
    } catch {
      return false;
    }
  }, [spHttpClient]);
  
  /**
   * Load property from tenant app catalog
   */
  const load = useCallback(async (): Promise<void> => {
    if (!spHttpClient || !pageContext) {
      console.warn('SPHttpClient or PageContext not available yet. Skipping load.');
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
      
      // Check write permissions (async, non-blocking)
      checkWritePermission(catalogUrl).then(hasPermission => {
        if (isMounted.current) {
          setCanWrite(hasPermission);
        }
      }).catch(() => {
        if (isMounted.current) {
          setCanWrite(false);
        }
      });
      
      // Read property
      const response: SPHttpClientResponse = await spHttpClient.get(
        `${catalogUrl}/_api/web/GetStorageEntity('${encodeURIComponent(key)}')`,
        SPHttpClient.configurations.v1
      );
      
      if (!response.ok) {
        throw new Error(`Failed to read property: ${response.statusText}`);
      }
      
      const entity: IStorageEntity = await response.json();
      
      if (isMounted.current) {
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
      if (isMounted.current) {
        const error = err instanceof Error ? err : new Error(String(err));
        setError(error);
        console.error('Failed to load tenant property:', error);
      }
    } finally {
      if (isMounted.current) {
        setIsLoading(false);
      }
    }
  }, [spHttpClient, pageContext, key, discoverAppCatalogUrl, checkWritePermission, deserializeValue]);
  
  /**
   * Write property to tenant app catalog
   */
  const write = useCallback(async (
    content: T,
    desc?: string
  ): Promise<void> => {
    if (!spHttpClient || !pageContext) {
      throw new Error('SPHttpClient or PageContext not available. Cannot write property.');
    }
    
    if (!key) {
      throw new Error('key is required. Cannot write property.');
    }
    
    setIsWriting(true);
    setWriteError(undefined);
    
    try {
      // Discover app catalog URL
      const catalogUrl = await discoverAppCatalogUrl();
      
      // Serialize value
      const serializedValue = serializeValue(content);
      
      // Build request body - NOTE: SharePoint SetStorageEntity only supports key, value, and description
      const body = {
        key: key,
        value: serializedValue,
        description: desc ?? ''
        // comment is NOT supported by SharePoint SetStorageEntity API
      };
      
      const response: SPHttpClientResponse = await spHttpClient.post(
        `${catalogUrl}/_api/web/SetStorageEntity`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Content-Type': 'application/json;odata=verbose',
            'Accept': 'application/json;odata=verbose'
          },
          body: JSON.stringify(body)
        }
      );
      
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to write property: ${response.statusText}. ${errorText}`);
      }
      
      if (isMounted.current) {
        // Update local state to reflect successful write
        setData(content);
        setDescription(desc);
        // Clear read error if write succeeds
        setError(undefined);
      }
    } catch (err) {
      if (isMounted.current) {
        const error = err instanceof Error ? err : new Error(String(err));
        setWriteError(error);
        console.error('Failed to write tenant property:', error);
      }
      // Re-throw to allow caller to handle
      throw err;
    } finally {
      if (isMounted.current) {
        setIsWriting(false);
      }
    }
  }, [spHttpClient, pageContext, key, discoverAppCatalogUrl, serializeValue]);
  
  /**
   * Remove property from tenant app catalog
   */
  const remove = useCallback(async (): Promise<void> => {
    if (!spHttpClient || !pageContext) {
      throw new Error('SPHttpClient or PageContext not available. Cannot remove property.');
    }
    
    if (!key) {
      throw new Error('key is required. Cannot remove property.');
    }
    
    setIsWriting(true);
    setWriteError(undefined);
    
    try {
      // Discover app catalog URL
      const catalogUrl = await discoverAppCatalogUrl();
      
      const response: SPHttpClientResponse = await spHttpClient.post(
        `${catalogUrl}/_api/web/RemoveStorageEntity('${encodeURIComponent(key)}')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose'
          }
        }
      );
      
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to remove property: ${response.statusText}. ${errorText}`);
      }
      
      if (isMounted.current) {
        // Clear local state
        setData(undefined);
        setDescription(undefined);
        setError(undefined);
      }
    } catch (err) {
      if (isMounted.current) {
        const error = err instanceof Error ? err : new Error(String(err));
        setWriteError(error);
        console.error('Failed to remove tenant property:', error);
      }
      // Re-throw to allow caller to handle
      throw err;
    } finally {
      if (isMounted.current) {
        setIsWriting(false);
      }
    }
  }, [spHttpClient, pageContext, key, discoverAppCatalogUrl]);
  
  // Auto-fetch on mount if enabled
  useEffect(() => {
    if (autoFetch && spHttpClient && pageContext && key) {
      load().catch(() => {
        // Error already handled in load() function
      });
    }
  }, [autoFetch, spHttpClient, pageContext, key, load]);
  
  // Computed state: ready when data loaded successfully
  const isReady = !isLoading && !error && data !== undefined;
  
  return {
    data,
    description,
    isLoading,
    error,
    isWriting,
    writeError,
    canWrite,
    load,
    write,
    remove,
    isReady,
  };
}

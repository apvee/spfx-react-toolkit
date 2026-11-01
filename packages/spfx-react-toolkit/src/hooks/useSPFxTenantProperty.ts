/**
 * Tenant storage-entity reads retain data on same-identity reload failure.
 *
 * @packageDocumentation
 * @docPage hooks/storage/tenant-property
 */
// useSPFxTenantProperty.ts
// Hook to read tenant-wide properties using SharePoint StorageEntity API (read-only)

import { useState, useCallback, useEffect, useMemo, useRef } from 'react';
import { useAppCatalogUrl } from './useAppCatalogUrl.internal';
import { createSPFxTenantPropertyService } from '../services/spfx-tenant-property.service';

/**
 * Read-only tenant storage-entity value/description and reload state; failures preserve earlier data.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxTenantPropertyResult, useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const property: SPFxTenantPropertyResult<string> = useSPFxTenantProperty('appVersion');
 *   console.log(property.data, property.error);
 *   return <span>{property.data ?? 'Unavailable'}</span>;
 * }
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxTenantPropertyResult<T> {
  /**
   * Current property value, initially undefined. A failed reload preserves previously loaded data; inspect error/isReady before using it.
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
   * Error from the latest-started load; cleared when a new load starts.
   */
  readonly error: Error | undefined;

  /** 
   * Loads the property from the tenant app catalog. Missing client/key logs and resolves without dispatch. Read failures publish error and resolve rather than rejecting.
   * Updates data, description, isLoading, and error states.
   * 
   * @returns Resolves after the attempt; inspect the current render error and data.
   * 
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';
   *
   * // Render beneath the initialized host's matching SPFx provider.
   * function TenantVersion() {
   *   const { data, load, isLoading, error } = useSPFxTenantProperty<string>('appVersion', false);
   *   return (
   *     <div>
   *       <button onClick={load} disabled={isLoading}>Refresh</button>
   *       <span>{error?.message ?? data ?? 'Unavailable'}</span>
   *     </div>
   *   );
   * }
   * ```
   */
  readonly load: () => Promise<void>;

  /** 
   * Computed state: true if data is loaded successfully.
   * Equivalent to: !isLoading && !error && data !== undefined
   * 
   * Useful for conditional rendering:
   * ```tsx
   * import * as React from 'react';
   *
   * if (!isReady) return <span>Loading</span>;
   * return <div>{data}</div>;
   * ```
   */
  readonly isReady: boolean;
}

/**
 * Hook to read tenant-wide properties using SharePoint StorageEntity API (read-only)
 *
 * Provides read access to tenant-scoped properties stored in the SharePoint tenant
 * app catalog. Properties are accessible across all sites in the tenant.
 *
 * This hook exposes reads only; it has no write/remove operations. Manage tenant storage entities through your administrator's supported tooling.
 * For a read/write key-value store at tenant level, use `useSPFxTenantKeyValueStore`.
 *
 * Features:
 * - Tenant-wide centralized storage (not site-specific)
 * - Smart deserialization for primitives and complex objects
 * - Optional metadata (description only - SharePoint limitation)
 * - Type-safe with TypeScript generics
 * - Obsolete/unmounted completions cannot publish state
 * - Automatic app catalog URL discovery
 *
 * Requirements:
 * - Tenant app catalog must be provisioned
 * - SharePoint must allow the current user to read the selected storage entity
 *
 * Key, service or catalog-discovery changes clear prior data/description/error and invalidate old request ownership. Only the latest active-identity load publishes state; dispatched requests are not cancelled. Same-identity reloads preserve previous data/description, including after a failed load, so render error before treating retained data as current. Disabling autoFetch alone retains state and does not cancel pending work. T is a cast of JSON-first decoded storage data without validation.
 *
 * @param key - Unique property key (e.g., 'appVersion', 'apiEndpoint', 'featureFlags')
 * @param autoFetch - Whether to load on mount; omitted or undefined uses true.
 *
 * @returns Object with data, metadata, loading state, error state, and load function
 *
 * @example Basic usage - string property
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';
 *
 * function VersionDisplay() {
 *   const { data, isLoading, error } = useSPFxTenantProperty<string>('appVersion');
 *
 *   if (isLoading) return <span>Loading version...</span>;
 *   if (error) return <div role="alert">
 *     Failed to load: {error.message}
 *   </div>;
 *
 *   return <span>Current Version: {data ?? 'Not Set'}</span>;
 * }
 * ```
 *
 * @example Complex object with JSON
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   interface FeatureFlags {
 *     enableChat: boolean;
 *     enableAnalytics: boolean;
 *     maxUsers: number;
 *   }
 *
 *   const { data, isLoading } = useSPFxTenantProperty<FeatureFlags>('featureFlags');
 *
 *   // Returns parsed object (stored as JSON string via PowerShell)
 *   if (data?.enableChat) {
 *     return <section>Chat enabled</section>;
 *   }
 *   return <span>{data?.enableChat ? 'Chat enabled' : 'Chat disabled'}</span>;
 * }
 * ```
 *
 * @example With metadata viewing
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';
 *
 * function PropertyViewer() {
 *   const { data, description, isLoading } =
 *     useSPFxTenantProperty<string>('appConfig');
 *
 *   if (isLoading) return <span>Loading</span>;
 *
 *   return (
 *     <section>
 *       <p>Value: {data}</p>
 *       {description && <p>Description: {description}</p>}
 *     </section>
 *   );
 * }
 * ```
 *
 * @example Lazy loading with manual trigger
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const { data, load, isLoading } = useSPFxTenantProperty<{ title: string }> (
 *     'appConfig',
 *     false  // Don't auto-fetch
 *   );
 *
 *   return (
 *     <div>
 *       <button onClick={load} disabled={isLoading}>
 *         {isLoading ? 'Loading...' : 'Load Config'}
 *       </button>
 *       {data && <pre>{JSON.stringify(data)}</pre>}
 *     </div>
 *   );
 * }
 * ```
 *
 * @example Multi-property dashboard
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';
 *
 * function TenantDashboard() {
 *   const version = useSPFxTenantProperty<string>('appVersion');
 *   const maintenance = useSPFxTenantProperty<boolean>('maintenanceMode');
 *   const config = useSPFxTenantProperty<{ title: string }> ('appConfig');
 *
 *   const isLoading = version.isLoading || maintenance.isLoading || config.isLoading;
 *
 *   if (isLoading) return <span>Loading dashboard...</span>;
 *
 *   return (
 *     <section>
 *       <p>Tenant Configuration</p>
 *       <p>App Version: {version.data ?? 'Not Set'}</p>
 *       <p>Maintenance Mode: {maintenance.data ? 'ON' : 'OFF'}</p>
 *       {config.data && <pre>{JSON.stringify(config.data, undefined, 2)}</pre>}
 *     </section>
 *   );
 * }
 * ```
 * @throws When called outside an SPFx provider.
 * @template T - Expected application data shape; no runtime validation is performed.
 *
 * @useWhen read-only configuration is stored as a tenant app-catalog storage entity.
 *
 *
 * @example Compose feature flags endpoint configuration and branding
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';
 * function TenantConfiguredWidget({ request }: { request: (endpoint: string, apiKey: string) => Promise<unknown> }) {
 *   const flags = useSPFxTenantProperty<{ enableNewUI?: boolean; enableBetaFeatures?: boolean }>('FeatureFlags');
 *   const endpoint = useSPFxTenantProperty<string>('CustomApiEndpoint');
 *   const apiKey = useSPFxTenantProperty<string>('CustomApiKey');
 *   const logo = useSPFxTenantProperty<string>('CompanyLogo');
 *   const theme = useSPFxTenantProperty<{ primary: string }>('CompanyTheme');
 *   const help = useSPFxTenantProperty<string>('HelpDeskUrl');
 *   const [result, setResult] = React.useState<unknown>();
 *   const [failure, setFailure] = React.useState('');
 *   const load = async () => {
 *     if (!endpoint.data || !apiKey.data) return;
 *     try { setFailure(''); setResult(await request(endpoint.data, apiKey.data)); }
 *     catch (error) { setFailure(String(error)); }
 *   };
 *   const error = flags.error ?? endpoint.error ?? apiKey.error ?? logo.error ?? theme.error ?? help.error;
 *   return <div>{error && <p role="alert">{error.message}</p>}
 *     {logo.data && <img src={logo.data} alt="Company logo" />}{help.data && <a href={help.data}>Help</a>}
 *     <p>Configured primary color: {theme.data?.primary ?? '#0078d4'}</p>
 *     {flags.data?.enableNewUI && <p>New UI enabled</p>}{flags.data?.enableBetaFeatures && <p>Beta features enabled</p>}
 *     {(!endpoint.data || !apiKey.data) && <p>Tenant properties not configured. Contact your administrator.</p>}
 *     <button disabled={endpoint.isLoading || apiKey.isLoading || !endpoint.data || !apiKey.data} onClick={() => void load()}>Call configured API</button>
 *     {failure && <p role="alert">{failure}</p>}<pre>{JSON.stringify(result, undefined, 2)}</pre>
 *   </div>;
 * }
 * ```
 */
export function useSPFxTenantProperty<T = unknown>(
  key: string,
  autoFetch: boolean = true
): SPFxTenantPropertyResult<T> {
  const { spHttpClient, discoverAppCatalogUrl, isMountedRef } = useAppCatalogUrl();
  const tenantPropertyService = useMemo(
    () => spHttpClient ? createSPFxTenantPropertyService(spHttpClient) : undefined,
    [spHttpClient]
  );

  // State management
  const [data, setData] = useState<T | undefined>(undefined);
  const [description, setDescription] = useState<string | undefined>(undefined);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<Error | undefined>(undefined);

  const requestRef = useRef({ key, service: tenantPropertyService, discover: discoverAppCatalogUrl, latest: 0 });
  if (requestRef.current.key !== key || requestRef.current.service !== tenantPropertyService ||
      requestRef.current.discover !== discoverAppCatalogUrl) {
    requestRef.current = { key, service: tenantPropertyService, discover: discoverAppCatalogUrl, latest: 0 };
  }
  useEffect(() => {
    setData(undefined);
    setDescription(undefined);
    setError(undefined);
    setIsLoading(false);
  }, [key, tenantPropertyService, discoverAppCatalogUrl]);

  const currentIdentity = requestRef.current;

  /**
   * Load property from tenant app catalog
   */
  const load = useCallback(async (): Promise<void> => {
    if (!tenantPropertyService) {
      console.warn('SPHttpClient not available yet. Skipping load.');
      return;
    }

    if (!key) {
      console.warn('key is required. Skipping load.');
      return;
    }

    const identity = currentIdentity;
    const request = ++identity.latest;
    const isCurrent = (): boolean => isMountedRef.current && requestRef.current === identity &&
      identity.key === key && identity.service === tenantPropertyService && request === identity.latest;
    if (!isCurrent()) return;
    setIsLoading(true);
    setError(undefined);

    try {
      // Discover app catalog URL
      const catalogUrl = await discoverAppCatalogUrl();

      // Read property
      const property = await tenantPropertyService.get<T>(key, catalogUrl);

      if (isCurrent()) {
        setData(property.value);
        setDescription(property.description);
      }
    } catch (err) {
      if (isCurrent()) {
        const capturedError = err instanceof Error ? err : new Error(String(err));
        setError(capturedError);
        console.error('Failed to load tenant property:', capturedError);
      }
    } finally {
      if (isCurrent()) {
        setIsLoading(false);
      }
    }
  }, [tenantPropertyService, key, discoverAppCatalogUrl, isMountedRef, currentIdentity]);

  // Auto-fetch on mount if enabled
  useEffect(() => {
    if (autoFetch && tenantPropertyService && key) {
      load().catch(() => {
        // Error already handled in load() function
      });
    }
  }, [autoFetch, tenantPropertyService, key, load]);

  // Computed state: ready when data loaded successfully
  const isReady = !isLoading && !error && data !== undefined;

  return useMemo(() => ({
    data,
    description,
    isLoading,
    error,
    load,
    isReady,
  }), [data, description, isLoading, error, load, isReady]);
}

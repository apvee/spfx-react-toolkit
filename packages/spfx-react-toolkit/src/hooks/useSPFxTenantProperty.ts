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
 *   if (error) return <div role="alert">Failed to load: {error.message}</div>;
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
 *
 * @scenario Read tenant version and counter with explicit activation
 * ```tsx
 * import * as React from 'react';
 * import { makeStyles, Accordion, AccordionHeader, AccordionItem, AccordionPanel, Button, Text } from '@fluentui/react-components';
 * import { useSPFxInstanceInfo, useSPFxPageContext, useSPFxSPHttpClient, useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Presenter = (regions: DemoRegions) => React.ReactElement;
 * const useStyles = makeStyles({ payload: { minWidth: 0, maxWidth: '100%', overflowX: 'auto' } });
 * function present(regions: DemoRegions, render?: Presenter): React.ReactElement {
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 *
 * // Reset ends this resource session; the outer Demo and activation button remain mounted.
 * function TenantProperties({ current, render }: { current: () => boolean; render?: Presenter }): React.ReactElement {
 *   const version = useSPFxTenantProperty<string>('spfx-toolkit-test-version', false);
 *   const counter = useSPFxTenantProperty<number>('spfx-toolkit-test-counter', false);
 *   const styles = useStyles();
 *   const identity = React.useRef({ version: version.load, counter: counter.load, token: {} });
 *   if (identity.current.version !== version.load || identity.current.counter !== counter.load) identity.current = { version: version.load, counter: counter.load, token: {} };
 *   const token = identity.current.token;
 *   const mounted = React.useRef(true);
 *   const latest = React.useRef(0);
 *   const [execution, setExecution] = React.useState<{ token: object; operation: number; completed: boolean }>();
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
 *   const valid = (): boolean => mounted.current && current() && identity.current.token === token;
 *   const run = async (): Promise<void> => {
 *     if (!valid()) return;
 *     const operation = ++latest.current;
 *     const snapshot = { version: version.load, counter: counter.load };
 *     setExecution({ token, operation, completed: false });
 *     // These public loads resolve on failure. Inspect their CURRENT rendered errors below.
 *     await Promise.all([snapshot.version(), snapshot.counter()]);
 *     if (valid() && operation === latest.current && identity.current.version === snapshot.version && identity.current.counter === snapshot.counter) setExecution({ token, operation, completed: true });
 *   };
 *   const visible = valid() && execution?.token === token && execution.operation === latest.current ? execution : undefined;
 *   const failure = visible ? version.error ?? counter.error : undefined;
 *   const completed = visible?.completed && !failure;
 *   const regions: DemoRegions = {
 *     commands: <Button appearance="primary" disabled={version.isLoading || counter.isLoading} onClick={run}>Run</Button>,
 *     status: failure ? <Text role="alert">{failure.message}</Text> : <Text role="status">{visible && !visible.completed ? 'Reading tenant properties...' : completed ? 'Tenant property reads completed.' : 'Run to read the properties.'}</Text>,
 *     children: completed ? <dl aria-label="Tenant properties">
 *       <dt>spfx-toolkit-test-version</dt><dd>{version.data === undefined ? '(not set)' : String(version.data)}</dd>
 *       <dt>spfx-toolkit-test-counter</dt><dd>{counter.data === undefined ? '(not set)' : String(counter.data)}</dd>
 *       <dt>Version description</dt><dd>{version.description === undefined ? '(not set)' : typeof version.description === 'string' ? JSON.stringify(version.description) : '(legacy description unavailable)'}</dd>
 *       <dt>Counter description</dt><dd>{counter.description === undefined ? '(not set)' : typeof counter.description === 'string' ? JSON.stringify(counter.description) : '(legacy description unavailable)'}</dd>
 *       <dt>Version ready</dt><dd>{String(version.isReady)}</dd><dt>Counter ready</dt><dd>{String(counter.isReady)}</dd>
 *     </dl> : undefined,
 *     diagnostics: <>
 *       <Text as="p">Both hooks use manual loading. Run starts two independent catalog discoveries on their first read; later reads reuse each session's successful cache. Properties are read-only. Hook generics do not validate stored data. Current errors take precedence over retained data; failed reload values are hidden. Each Ready flag requires defined data and no loading/error; zero is valid.</Text>
 *       <Text as="p">Reset or host changes end this session and require explicit activation. Already-started logical reads are not cancelled: discovery can still continue to a property GET after departure, while obsolete results stay hidden. The wrapper observes web/root URL scalars on render because the public discovery hook does not independently track every in-place URL mutation.</Text>
 *       {completed && <Accordion collapsible><AccordionItem value="payload"><AccordionHeader>Full tenant property results</AccordionHeader><AccordionPanel><div className={styles.payload} role="region" aria-label="Full tenant property results" tabIndex={0}><pre>{JSON.stringify({ version: { value: version.data, description: version.description }, counter: { value: counter.data, description: counter.description } }, undefined, 2)}</pre></div></AccordionPanel></AccordionItem></Accordion>}
 *     </>,
 *   };
 *   return present(regions, render);
 * }
 *
 * export function Demo({ render, resetRevision = 0 }: { render?: Presenter; resetRevision?: number } = {}): React.ReactElement {
 *   const page = useSPFxPageContext();
 *   const { client } = useSPFxSPHttpClient();
 *   const instance = useSPFxInstanceInfo();
 *   const web = page.web.absoluteUrl.trim().replace(/\/+$/, '');
 *   const root = page.site.absoluteUrl.trim().replace(/\/+$/, '');
 *   const boundary = React.useRef({ page, client, web, root, instance: instance.id, revision: resetRevision, token: {} });
 *   const previous = boundary.current;
 *   if (previous.page !== page || previous.client !== client || previous.web !== web || previous.root !== root || previous.instance !== instance.id || previous.revision !== resetRevision) boundary.current = { page, client, web, root, instance: instance.id, revision: resetRevision, token: {} };
 *   const token = boundary.current.token;
 *   const [session, setSession] = React.useState<object>();
 *   const mounted = React.useRef(true);
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
 *   const current = (): boolean => mounted.current && boundary.current.token === token && page.web.absoluteUrl.trim().replace(/\/+$/, '') === web && page.site.absoluteUrl.trim().replace(/\/+$/, '') === root;
 *   const active = session === token;
 *   const activate = (): void => { if (current() && client && web && root && !active) setSession(token); };
 *   return <>
 *     <Button disabled={!client || !web || !root || active} onClick={activate}>Activate tenant properties</Button>
 *     {active ? <TenantProperties current={current} render={render}/> : present({ children: <Text>Activate tenant properties, then Run to read the version and counter. Selection, activation and Reset issue no HTTP requests.</Text> }, render)}
 *   </>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Run both tenant property loads explicitly and observe current values, descriptions, missing values and read failures.' };
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

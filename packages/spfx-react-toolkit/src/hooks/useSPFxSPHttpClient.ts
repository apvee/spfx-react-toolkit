/**
 * SharePoint REST calls use the host client; base URL state is caller-owned.
 *
 * @packageDocumentation
 * @docPage hooks/http-clients/sharepoint-http-client
 */
// useSPFxSPHttpClient.ts
// Hook to access SharePoint REST APIs with state management

import { useMemo, useState, useCallback } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';
import { useSPFxPageContext } from './useSPFxPageContext';
import { SPHttpClient } from '@microsoft/sp-http';
import { useAsyncInvoke } from './useAsyncInvoke.internal';

/**
 * Native SharePoint REST client, caller-managed base URL and counted invocation state local to this hook.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxSPHttpClientInfo, useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const http: SPFxSPHttpClientInfo = useSPFxSPHttpClient();
 *   console.log(http.baseUrl, http.isReady);
 *   return <span>{http.isReady ? 'Ready' : 'Unavailable'}</span>;
 * }
 * ```
 */
export interface SPFxSPHttpClientInfo {
  /**
   * Native SPFx client, undefined until available or when service consumption/initialization fails. Direct client requests do not update hook loading/error state.
   */
  readonly client: SPHttpClient | undefined;
  
  /**
   * Invoke SharePoint REST API call with automatic state management.
   * Tracks loading state and captures errors automatically.
   *
   * @typeParam T - Result produced by the supplied operation.
   * @param fn - Function that receives SPHttpClient and returns a promise
   * @returns Promise with the result
   * @throws Error if client is not available
   *
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
   * import { SPHttpClient } from '@microsoft/sp-http';
   *
   * function Example() {
   *   const { invoke, baseUrl, isReady, isLoading } = useSPFxSPHttpClient();
   *   const [result, setResult] = React.useState<unknown>();
   *   const [failure, setFailure] = React.useState<string>();
   *   const run = async () => {
   *     setFailure(undefined);
   *     try {
   *       const web = await invoke(client =>
   *         client.get(`${baseUrl}/_api/web`, SPHttpClient.configurations.v1)
   *           .then(res => res.json())
   *       );
   *       setResult(web);
   *     } catch (failure) {
   *       setFailure(failure instanceof Error ? failure.message : String(failure));
   *     }
   *   };
   *   return <>
   *     <button disabled={!isReady || isLoading} onClick={() => void run()}>Load</button>
   *     {failure && <span role="alert">{failure}</span>}
   *     <pre>{JSON.stringify(result)}</pre>
   *   </>;
   * }
   * ```
   * @throws Callback/request failures are normalized to Error, published when current, and rethrown. HTTP response.ok is not checked automatically.
   */
  readonly invoke: <T>(fn: /**
   * Operation receiving the initialized native client. Non-success HTTP responses require explicit checking.
   * @param client - Initialized client for this invocation.
   * @returns Promise for the operation result; rejection flows through invoke.
   */ (client: SPHttpClient) => Promise<T>) => Promise<T>;
  
  /**
   * True while any invoke call for the current client remains pending; direct client usage is not tracked.
   */
  readonly isLoading: boolean;
  
  /**
   * Error from the latest-started invoke call; each new call clears it. Older/client-obsolete completions cannot replace it. Direct client usage is not tracked.
   */
  readonly error: Error | undefined;
  
  /**
   * Clear the current error
   * @returns No value; initialization/context errors remain separate.
   */
  readonly clearError: () => void;
  
  /**
   * Changes the URL exposed by the hook for constructing cross-site requests; does not modify the native client or invoke callback URLs.
   * @param url - Base web URL used by the consumer.
   * @returns No value.
   */
  readonly setBaseUrl: (url: string) => void;
  
  /** Consumer base web URL; initially the supplied URL or current web, with one trailing slash removed. */
  readonly baseUrl: string;

  /**
   * True when the native client exists. Check readiness before invoking a request.
   *
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
   *
   * import { SPHttpClient } from '@microsoft/sp-http';
   *
   * function LoadButton() {
   *   const { isReady, invoke, baseUrl } = useSPFxSPHttpClient();
   *   const load = async () => {
   *     try {
   *       const data = await invoke(client => client.get(`${baseUrl}/_api/web`, SPHttpClient.configurations.v1).then(response => response.json()));
   *       console.log(data);
   *     } catch (error) {
   *       console.error(error);
   *     }
   *   };
   *   return <button disabled={!isReady} onClick={() => void load()}>Load</button>;
   * }
   * ```
   */
  readonly isReady: boolean;
}

/**
 * Accesses the SPFx SPHttpClient from ServiceScope. Direct client calls do not update hook state. invoke tracks all pending calls for its current client and publishes errors only from the latest-started call; obsolete completions still settle their promises. HTTP non-success responses must be checked by the callback; the hook does not inspect response.ok.
 * @param initialBaseUrl - Initial URL, trimmed and captured once; omitted/blank uses the current web URL. Use setBaseUrl for subsequent changes.
 * @returns Native client, invocation state and readiness helpers.
 * @throws Missing SPFx provider.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
 *
 * function ClientStatus() {
 *   const info = useSPFxSPHttpClient();
 *   return <span>{info.isReady ? 'Ready' : 'Unavailable'}</span>;
 * }
 * ```
 *
 * @useWhen an event or effect must call SharePoint REST using the host-authenticated client.
 *
 *
 * @example List items CRUD with visible actions
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
 * import { SPHttpClient } from '@microsoft/sp-http';
 *
 * interface ITask {
 *   Id: number;
 *   Title: string;
 *   Status: string;
 * }
 *
 * function TaskList() {
 *   const { invoke, isLoading, error, baseUrl, isReady } = useSPFxSPHttpClient();
 *   const [tasks, setTasks] = React.useState<ITask[]>([]);
 *
 *   // Load tasks
 *   const loadTasks = () => {
 *     invoke(client =>
 *       client.get(
 *         `${baseUrl}/_api/web/lists/getbytitle('Tasks')/items`,
 *         SPHttpClient.configurations.v1
 *       ).then(res => {
 *         if (!res.ok) throw new Error(`Request failed: ${res.status}`);
 *         return res.json();
 *       })
 *     ).then(result => setTasks(result.value)).catch(() => undefined);
 *   };
 *
 *   // Create task
 *   const createTask = (title: string) => {
 *     invoke(client =>
 *       client.post(
 *         `${baseUrl}/_api/web/lists/getbytitle('Tasks')/items`,
 *         SPHttpClient.configurations.v1,
 *         { body: JSON.stringify({ Title: title, Status: 'New' }) }
 *       ).then(res => {
 *         if (!res.ok) throw new Error(`Request failed: ${res.status}`);
 *         return res.json();
 *       })
 *     ).then(loadTasks).catch(() => undefined);
 *   };
 *
 *   // Update task
 *   const updateTask = (id: number, status: string) => {
 *     invoke(client =>
 *       client.post(
 *         `${baseUrl}/_api/web/lists/getbytitle('Tasks')/items(${id})`,
 *         SPHttpClient.configurations.v1,
 *         {
 *           headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
 *           body: JSON.stringify({ Status: status })
 *         }
 *       ).then(response => {
 *         if (!response.ok) throw new Error(`Request failed: ${response.status}`);
 *         return response;
 *       })
 *     ).then(loadTasks).catch(() => undefined);
 *   };
 *
 *   // Delete task
 *   const deleteTask = (id: number) => {
 *     invoke(client =>
 *       client.post(
 *         `${baseUrl}/_api/web/lists/getbytitle('Tasks')/items(${id})`,
 *         SPHttpClient.configurations.v1,
 *         { headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' } }
 *       ).then(response => {
 *         if (!response.ok) throw new Error(`Request failed: ${response.status}`);
 *         return response;
 *       })
 *     ).then(loadTasks).catch(() => undefined);
 *   };
 *
 *   React.useEffect(() => {
 *     if (isReady) loadTasks();
 *   }, [isReady]);
 *
 *   if (!isReady) return <p role="status">Initializing...</p>;
 *   if (isLoading) return <p role="status">Loading tasks...</p>;
 *   if (error) return <p role="alert">{error.message}</p>;
 *
 *   return (
 *     <div><button disabled={isLoading} onClick={() => createTask('New task')}>Create task</button><ul>
 *       {tasks.map(task => (
 *         <li key={task.Id}>
 *           {task.Title} - {task.Status}
 *           <button onClick={() => updateTask(task.Id, 'Done')}>Complete</button>
 *           <button onClick={() => deleteTask(task.Id)}>Delete</button>
 *         </li>
 *       ))}
 *     </ul></div>
 *   );
 * }
 * ```
 *
 *
 * @example Cross-site request without stale base state
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
 * import { SPHttpClient } from '@microsoft/sp-http';
 *
 * interface DocumentItem {
 *   Id: number;
 *   Title: string;
 * }
 *
 * function CrossSiteData() {
 *   const { invoke, setBaseUrl, isLoading, isReady, error } = useSPFxSPHttpClient();
 *   const [otherSiteData, setOtherSiteData] = React.useState<DocumentItem[]>([]);
 *
 *   const loadFromOtherSite = async () => {
 *     const targetUrl = 'https://tenant.sharepoint.com/sites/OtherSite';
 *     setBaseUrl(targetUrl);
 *     try {
 *       const result = await invoke(client =>
 *         client.get(
 *           `${targetUrl}/_api/web/lists/getbytitle('Documents')/items`,
 *           SPHttpClient.configurations.v1
 *         ).then(res => {
 *         if (!res.ok) throw new Error(`Request failed: ${res.status}`);
 *         return res.json();
 *       })
 *       );
 *       setOtherSiteData(result.value);
 *     } catch {
 *       // The hook exposes the invocation error for the UI below.
 *     }
 *   };
 *
 *   return <div>
 *     <button onClick={loadFromOtherSite} disabled={!isReady || isLoading}>
 *       Load from Other Site
 *     </button>
 *     {error && <p>{error.message}</p>}
 *     {otherSiteData.map(item => <p key={item.Id}>{item.Title}</p>)}
 *   </div>;
 * }
 * ```
 *
 * @scenario Read the current SharePoint web title and URL
 * ```tsx
 * import * as React from 'react';
 * import { Button, Accordion, AccordionItem, AccordionHeader, AccordionPanel, makeStyles } from '@fluentui/react-components';
 * import { useSPFxSPHttpClient, useSPFxContext, useSPFxServiceScope, useSPFxPageContext } from '@apvee/spfx-react-toolkit';
 * import { SPHttpClient } from '@microsoft/sp-http';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * const useStyles = makeStyles({ diagnostics: { minWidth: 0, maxWidth: '100%', overflowX: 'auto' } });
 * // JSON payloads may contain legacy null values; scenario-owned absence uses undefined.
 * function payloadText(value: unknown): string {
 *   const seen = new WeakSet<object>();
 *   return JSON.stringify(value, (key: string, item: unknown): unknown => {
 *     if (key.startsWith('_') || key === 'serviceScope' || key === 'service') return '[Private]';
 *     if (typeof item === 'object' && item !== undefined && item !== null) {
 *       if (seen.has(item)) return '[Circular]';
 *       seen.add(item);
 *     }
 *     return item;
 *   }, 2) ?? 'No response';
 * }
 * function responseRows(value: unknown): React.ReactNode {
 *   const entries = value !== undefined && typeof value === 'object' && value !== null
 *     ? Object.entries(value).filter(([, field]) => ['string', 'number', 'boolean'].includes(typeof field)).slice(0, 8)
 *     : value === undefined ? [] : [['Response', value]];
 *   return entries.map(([key, field]) => <React.Fragment key={String(key)}><dt>{String(key)}</dt><dd>{String(field)}</dd></React.Fragment>);
 * }
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const styles = useStyles();
 *   const { spfxContext } = useSPFxContext();
 *   const { serviceScope } = useSPFxServiceScope();
 *   const page = useSPFxPageContext();
 *   const info = useSPFxSPHttpClient();
 *   const currentWeb = page.web.absoluteUrl.trim().replace(/\/$/, '');
 *   React.useEffect(() => { if (info.baseUrl !== currentWeb) info.setBaseUrl(currentWeb); }, [currentWeb, info.baseUrl, info.setBaseUrl]);
 *   const [observedScope, setObservedScope] = React.useState(serviceScope);
 *
 *   React.useEffect(() => { setObservedScope(serviceScope);  }, [serviceScope ]);
 *   const initialized = observedScope === serviceScope ;
 *   const values: readonly unknown[] = [spfxContext, serviceScope, serviceScope?.consume, page, page.web.absoluteUrl, info.invoke, resetRevision , currentWeb, info.baseUrl];
 *   // A transition gets a new boundary even if a later draft/revision repeats an old value.
 *   const boundary = React.useRef(values);
 *   if (!values.every((value, index) => boundary.current[index] === value)) boundary.current = values;
 *   const identity = boundary.current;
 *   const live = React.useRef(identity);
 *   live.current = identity;
 *   const mounted = React.useRef(true);
 *   React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
 *   const same = (capture: readonly unknown[]): boolean => mounted.current && capture === live.current && capture[1] === spfxContext.serviceScope && capture[2] === serviceScope?.consume && capture[4] === page.web.absoluteUrl;
 *   const sequence = React.useRef(0);
 *   // Hook loading counts current-client calls; this view also remembers obsolete native work.
 *   const active = React.useRef(0);
 *   const [activeActions, setActiveActions] = React.useState(0);
 *   const [outcome, setOutcome] = React.useState<{ identity: readonly unknown[]; data?: unknown; failure?: string }>();
 *   const visible = outcome && same(outcome.identity) ? outcome : undefined;
 *   const execute = async (): Promise<void> => {
 *     if (!same(identity) || !initialized || !info.isReady || info.baseUrl !== currentWeb) return;
 *     const request = ++sequence.current;
 *     setOutcome(undefined);
 *     active.current++;
 *     setActiveActions(active.current);
 *     try {
 *       const data = await info.invoke(async client => {
 *         // invoke calls this synchronously today; keep the native I/O boundary guarded too.
 *         if (!same(identity)) return undefined;
 *         return client.get(`${info.baseUrl}/_api/web?$select=Title,Url`, SPHttpClient.configurations.v1).then(response => response.json());
 *       });
 *       if (same(identity) && sequence.current === request) setOutcome({ identity, data });
 *     } catch (failure) {
 *       if (same(identity) && sequence.current === request) setOutcome({ identity, failure: failure instanceof Error ? failure.message : String(failure) });
 *     } finally {
 *       active.current--;
 *       // Cleanup may release obsolete work, but never republishes its result or error.
 *       if (mounted.current) setActiveActions(active.current);
 *     }
 *   };
 *
 *   const regions: DemoRegions = {
 *     controls: undefined,
 *     commands: <><Button disabled={!initialized || !info.isReady || info.isLoading || activeActions > 0 || info.baseUrl !== currentWeb} onClick={execute}>Load current web</Button></>,
 *     status: <><p role="status">Ready: {initialized && info.isReady && info.baseUrl === currentWeb ? 'Yes' : 'No'}; {info.isLoading || activeActions > 0 ? 'Loading: underlying requests are still completing' : 'Idle'}</p>{visible?.failure && <p role="alert">{visible.failure}</p>}</>,
 *     children: <dl aria-label="SPHttpClient results"><dt>Response</dt><dd>{visible?.data === undefined ? 'Not loaded' : 'Loaded'}</dd><dt>Current request base</dt><dd>{currentWeb}</dd>{responseRows(visible?.data)}</dl>,
 *     diagnostics: <><p>Mount below the matching SPFx provider. Reset invalidates this view; it does not cancel requests. These JSON callbacks do not inspect HTTP response.ok.</p>{visible?.data !== undefined && <Accordion collapsible><AccordionItem value="response"><AccordionHeader>Raw response</AccordionHeader><AccordionPanel><div className={styles.diagnostics} role="region" aria-label="Raw response JSON"><pre>{payloadText(visible.data)}</pre></div></AccordionPanel></AccordionItem></Accordion>}</>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Read the current SharePoint web Title and Url explicitly; the source synchronizes its base URL without issuing a request.' };
 * ```
 */
export function useSPFxSPHttpClient(initialBaseUrl?: string): SPFxSPHttpClientInfo {
  const { consume } = useSPFxServiceScope();
  const pageContext = useSPFxPageContext();
  
  // Lazy consume SPHttpClient from ServiceScope (cached by useMemo)
  // Returns undefined if ServiceScope is not available
  const client = useMemo((): SPHttpClient | undefined => {
    try {
      return consume<SPHttpClient>(SPHttpClient.serviceKey);
    } catch (err) {
      console.error('Failed to consume SPHttpClient from ServiceScope:', err);
      return undefined;
    }
  }, [consume]);
  
  // Default to current site if no baseUrl provided
  const defaultBaseUrl = initialBaseUrl?.trim() || pageContext.web.absoluteUrl || '';
  
  // Normalize: remove trailing slash for consistency (ES5-compatible)
  const normalizedBaseUrl = defaultBaseUrl.charAt(defaultBaseUrl.length - 1) === '/'
    ? defaultBaseUrl.slice(0, -1) 
    : defaultBaseUrl;
  
  // BaseUrl state management (specific to this hook)
  const [baseUrl, setBaseUrlState] = useState<string>(normalizedBaseUrl);
  
  // Public setter for baseUrl with normalization
  const setBaseUrl = useCallback((url: string) => {
    const trimmed = url.trim();
    const normalized = trimmed.charAt(trimmed.length - 1) === '/' 
      ? trimmed.slice(0, -1) 
      : trimmed;
    setBaseUrlState(normalized);
  }, []);
  
  // Use shared async invocation pattern
  const { invoke, isLoading, error, clearError } = useAsyncInvoke(
    client,
    'SPHttpClient not available. Check SPFx context and ServiceScope.'
  );
  
  // Computed: ready when client is available
  const isReady = client !== undefined;
  
  return useMemo(() => ({
    client,
    invoke,
    isLoading,
    error,
    clearError,
    setBaseUrl,
    baseUrl,
    isReady,
  }), [client, invoke, isLoading, error, clearError, setBaseUrl, baseUrl, isReady]);
}

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

/**
 * Generic HTTP requests use the native host client with counted invocation state.
 *
 * @packageDocumentation
 * @docPage hooks/http-clients/http-client
 */
// useSPFxHttpClient.ts
// Hook to access generic HTTP client with state management

import { useMemo } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';
import { HttpClient } from '@microsoft/sp-http';
import { useAsyncInvoke } from './useAsyncInvoke.internal';

/**
 * Native generic HTTP client and counted invocation/loading/error controls local to this hook.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxHttpClientInfo, useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const http: SPFxHttpClientInfo = useSPFxHttpClient();
 *   console.log(http.isReady, http.error);
 *   return <span>{http.isReady ? 'Ready' : 'Unavailable'}</span>;
 * }
 * ```
 */
export interface SPFxHttpClientInfo {
  /**
   * Native SPFx client, undefined until available or when service consumption/initialization fails. Direct client requests do not update hook loading/error state.
   */
  readonly client: HttpClient | undefined;
  
  /**
   * Invoke HTTP API call with automatic state management.
   * Tracks loading state and captures errors automatically.
   *
   * @typeParam T - Result produced by the supplied operation.
   * @param fn - Function that receives HttpClient and returns a promise
   * @returns Promise with the result
   * @throws Error if client is not available
   *
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
   * import { HttpClient } from '@microsoft/sp-http';
   *
   * function Example() {
   *   const { invoke, isReady, isLoading } = useSPFxHttpClient();
   *   const [result, setResult] = React.useState<unknown>();
   *   const [failure, setFailure] = React.useState<string>();
   *   const run = async () => {
   *     setFailure(undefined);
   *     try {
   *       const data = await invoke(client =>
   *         client.get('https://api.example.com/data', HttpClient.configurations.v1)
   *           .then(res => res.json())
   *       );
   *       setResult(data);
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
   */ (client: HttpClient) => Promise<T>) => Promise<T>;
  
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
   * True when the native client exists. Check readiness before invoking a request.
   *
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
   *
   * import { HttpClient } from '@microsoft/sp-http';
   *
   * function LoadButton() {
   *   const { isReady, invoke } = useSPFxHttpClient();
   *   const load = async () => {
   *     try {
   *       const data = await invoke(client => client.get('https://api.example.com/data', HttpClient.configurations.v1).then(response => response.json()));
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
 * Accesses the SPFx HttpClient from ServiceScope. Direct client calls do not update hook state. invoke tracks all pending calls for its current client and publishes errors only from the latest-started call; obsolete completions still settle their promises. HTTP non-success responses must be checked by the callback; the hook does not inspect response.ok.
 * @returns Native client, invocation state and readiness helpers.
 * @throws Missing SPFx provider.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
 *
 * function ClientStatus() {
 *   const info = useSPFxHttpClient();
 *   return <span>{info.isReady ? 'Ready' : 'Unavailable'}</span>;
 * }
 * ```
 *
 * @useWhen an event or effect must issue a generic HTTP request with tracked async state.
 *
 *
 * @example Weather with readiness failure and retry
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
 * import { HttpClient } from '@microsoft/sp-http';
 *
 * function WeatherWidget() {
 *   const { invoke, isLoading, error, clearError, isReady } = useSPFxHttpClient();
 *   const [weather, setWeather] = React.useState<{ main?: { temp?: number } }>();
 *
 *   const loadWeather = () => {
 *     invoke(client =>
 *       client.get(
 *         'https://api.openweathermap.org/data/2.5/weather?q=London&appid=YOUR_KEY',
 *         HttpClient.configurations.v1
 *       ).then(res => {
 *         if (!res.ok) throw new Error(`Request failed: ${res.status}`);
 *         return res.json();
 *       })
 *     ).then(data => setWeather(data)).catch(() => undefined);
 *   };
 *
 *   React.useEffect(() => {
 *     if (isReady) loadWeather();
 *   }, [isReady]);
 *
 *   if (!isReady) return <p role="status">Initializing...</p>;
 *   if (isLoading) return <p role="status">Loading weather...</p>;
 *   if (error) return <div><p role="alert">{error.message}</p><button onClick={() => { clearError(); loadWeather(); }}>Retry</button></div>;
 *
 *   return <div>Temperature: {weather?.main?.temp}°K</div>;
 * }
 * ```
 *
 *
 * @example Post a webhook with invocation state
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
 * import { HttpClient } from '@microsoft/sp-http';
 *
 * function SlackNotifier() {
 *   const { invoke, isLoading, isReady } = useSPFxHttpClient();
 *
 *   const sendNotification = (message: string) => {
 *     invoke(client =>
 *       client.post(
 *         'https://hooks.slack.com/services/YOUR/WEBHOOK/URL',
 *         HttpClient.configurations.v1,
 *         {
 *           headers: { 'Content-Type': 'application/json' },
 *           body: JSON.stringify({ text: message })
 *         }
 *       ).then(response => {
 *         if (!response.ok) throw new Error(`Request failed: ${response.status}`);
 *         return response;
 *       })
 *     ).then(() => console.log('Notification sent')).catch(console.error);
 *   };
 *
 *   return (
 *     <button onClick={() => sendNotification('Hello!')} disabled={!isReady || isLoading}>
 *       Send to Slack
 *     </button>
 *   );
 * }
 * ```
 *
 * @example Load an external todo with tracked invocation state
 * ```tsx
 * import * as React from 'react';
 * import { Button, Accordion, AccordionItem, AccordionHeader, AccordionPanel, makeStyles } from '@fluentui/react-components';
 * import { useSPFxHttpClient, useSPFxContext, useSPFxServiceScope, useSPFxPageContext } from '@apvee/spfx-react-toolkit';
 * import { HttpClient } from '@microsoft/sp-http';
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
 *   const info = useSPFxHttpClient();
 *
 *   const [observedScope, setObservedScope] = React.useState(serviceScope);
 *
 *   React.useEffect(() => { setObservedScope(serviceScope);  }, [serviceScope ]);
 *   const initialized = observedScope === serviceScope ;
 *   const values: readonly unknown[] = [spfxContext, serviceScope, serviceScope?.consume, page, page.web.absoluteUrl, info.invoke, resetRevision ];
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
 *     if (!same(identity) || !initialized || !info.isReady ) return;
 *     const request = ++sequence.current;
 *     setOutcome(undefined);
 *     active.current++;
 *     setActiveActions(active.current);
 *     try {
 *       const data = await info.invoke(async client => {
 *         // invoke calls this synchronously today; keep the native I/O boundary guarded too.
 *         if (!same(identity)) return undefined;
 *         return client.get('https://jsonplaceholder.typicode.com/todos/1', HttpClient.configurations.v1).then(response => response.json());
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
 *     commands: <><Button disabled={!initialized || !info.isReady || info.isLoading || activeActions > 0 } onClick={execute}>Load external todo</Button></>,
 *     status: <><p role="status">Ready: {initialized && info.isReady  ? 'Yes' : 'No'}; {info.isLoading || activeActions > 0 ? 'Loading: underlying requests are still completing' : 'Idle'}</p>{visible?.failure && <p role="alert">{visible.failure}</p>}</>,
 *     children: <dl aria-label="HttpClient results"><dt>Response</dt><dd>{visible?.data === undefined ? 'Not loaded' : 'Loaded'}</dd>{responseRows(visible?.data)}</dl>,
 *     diagnostics: <><p>Mount below the matching SPFx provider. Reset invalidates this view; it does not cancel requests. These JSON callbacks do not inspect HTTP response.ok.</p>{visible?.data !== undefined && <Accordion collapsible><AccordionItem value="response"><AccordionHeader>Raw response</AccordionHeader><AccordionPanel><div className={styles.diagnostics} role="region" aria-label="Raw response JSON"><pre>{payloadText(visible.data)}</pre></div></AccordionPanel></AccordionItem></Accordion>}</>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Load the external todo explicitly; observe primitive response values, pending requests, caught failures and Reset invalidation.' };
 * ```
 */
export function useSPFxHttpClient(): SPFxHttpClientInfo {
  const { consume } = useSPFxServiceScope();
  
  // Lazy consume HttpClient from ServiceScope (cached by useMemo)
  // Returns undefined if ServiceScope is not available
  const client = useMemo((): HttpClient | undefined => {
    try {
      return consume<HttpClient>(HttpClient.serviceKey);
    } catch (err) {
      console.error('Failed to consume HttpClient from ServiceScope:', err);
      return undefined;
    }
  }, [consume]);
  
  // Use shared async invocation pattern
  const { invoke, isLoading, error, clearError } = useAsyncInvoke(
    client,
    'HttpClient not available. Check SPFx context and ServiceScope.'
  );
  
  // Computed: ready when client is available
  const isReady = client !== undefined;
  
  return useMemo(() => ({
    client,
    invoke,
    isLoading,
    error,
    clearError,
    isReady,
  }), [client, invoke, isLoading, error, clearError, isReady]);
}

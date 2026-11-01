/**
 * SPFx Graph client initialization is separate from tracked request state.
 *
 * @packageDocumentation
 * @docPage hooks/http-clients/graph-client
 */
// useSPFxMSGraphClient.ts
// Hook to access Microsoft Graph client with state management

import { useMemo, useState, useEffect, useRef } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';
import { MSGraphClientV3, MSGraphClientFactory } from '@microsoft/sp-http';
import { useAsyncInvoke } from './useAsyncInvoke.internal';

/**
 * Native SPFx v3 Graph client, counted invocation state and factory initialization/readiness state.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxMSGraphClientInfo, useSPFxMSGraphClient } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const graph: SPFxMSGraphClientInfo = useSPFxMSGraphClient();
 *   console.log(graph.isReady, graph.initError);
 *   return <span>{graph.initError?.message ?? (graph.isReady ? 'Ready' : 'Unavailable')}</span>;
 * }
 * ```
 */
export interface SPFxMSGraphClientInfo {
  /**
   * Native MSGraphClientV3 from SPFx.
   * Provides access to Microsoft Graph API with built-in authentication.
   * Will be undefined until initialization completes.
   */
  readonly client: MSGraphClientV3 | undefined;
  
  /**
   * Invoke Graph API call with automatic state management.
   * Tracks loading state and captures errors automatically.
   *
   * @typeParam T - Result produced by the supplied operation.
   * @param fn - Function that receives Graph client and returns a promise
   * @returns Promise with the result
   * @throws Error if client is not initialized yet
   *
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxMSGraphClient } from '@apvee/spfx-react-toolkit';
   *
   * function Example() {
   *   const { invoke, isReady, isLoading, initError } = useSPFxMSGraphClient();
   *   const [result, setResult] = React.useState<unknown>();
   *   const [failure, setFailure] = React.useState<string>();
   *   const run = async () => {
   *     setFailure(undefined);
   *     try {
   *       const user = await invoke(client =>
   *         client.api('/me').select('displayName,mail').get()
   *       );
   *       setResult(user);
   *     } catch (failure) {
   *       setFailure(failure instanceof Error ? failure.message : String(failure));
   *     }
   *   };
   *   return <>
   *     <button disabled={!isReady || isLoading} onClick={() => void run()}>Load</button>
   *     {(failure ?? initError?.message) && <span role="alert">{failure ?? initError?.message}</span>}
   *     <pre>{JSON.stringify(result)}</pre>
   *   </>;
   * }
   * ```
   * @throws Callback/request failures are normalized to Error, published when current, and rethrown. Successful payloads are returned without additional inspection.
   */
  readonly invoke: <T>(fn: /**
   * Operation receiving the initialized native Graph client; native request errors flow through invoke.
   * @param client - Initialized client for this invocation.
   * @returns Promise for the operation result; rejection flows through invoke.
   */ (client: MSGraphClientV3) => Promise<T>) => Promise<T>;
  
  /**
   * True while any invoke call for the current client remains pending; direct client usage is not tracked.
   */
  readonly isLoading: boolean;
  
  /**
   * Error from the latest-started invoke call; each new call clears it. Older/client-obsolete completions cannot replace it. Direct client usage is not tracked.
   */
  readonly error: Error | undefined;
  
  /**
   * Clear the current error from invoke() calls
   * @returns No value; initialization/context errors remain separate.
   */
  readonly clearError: () => void;

  /**
   * True while Graph client initialization is pending. False alone does not guarantee readiness.
   *
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxMSGraphClient } from '@apvee/spfx-react-toolkit';
   *
   * function Example() {
   *   const { isInitializing, isReady } = useSPFxMSGraphClient();
   *   if (isInitializing) return <span>Initializing Graph</span>;
   *   if (!isReady) return <span>Graph unavailable</span>;
   *   return <span>Ready</span>;
   * }
   * ```
   */
  readonly isInitializing: boolean;

  /**
   * Graph client initialization error, or undefined when none is published. Invocation errors are exposed separately by error.
   *
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxMSGraphClient } from '@apvee/spfx-react-toolkit';
   *
   * function Example() {
   *   const { initError } = useSPFxMSGraphClient();
   *   if (initError) return <span>{initError.message}</span>;
   *   return <span>No initialization error</span>;
   * }
   * ```
   */
  readonly initError: Error | undefined;

  /**
   * Computed state: true when client is ready for use.
   * Equivalent to: client !== undefined && !isInitializing && !initError
   *
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxMSGraphClient } from '@apvee/spfx-react-toolkit';
   *
   * function Example() {
   *   const { invoke, isReady, isLoading, initError } = useSPFxMSGraphClient();
   *   const [result, setResult] = React.useState<unknown>();
   *   const [failure, setFailure] = React.useState<string>();
   *   const run = async () => {
   *     setFailure(undefined);
   *     try {
   *       const data = await invoke(c => c.api('/me').get());
   *       setResult(data);
   *     } catch (failure) {
   *       setFailure(failure instanceof Error ? failure.message : String(failure));
   *     }
   *   };
   *   return <>
   *     <button disabled={!isReady || isLoading} onClick={() => void run()}>Load</button>
   *     {(failure ?? initError?.message) && <span role="alert">{failure ?? initError?.message}</span>}
   *     <pre>{JSON.stringify(result)}</pre>
   *   </>;
   * }
   * ```
   */
  readonly isReady: boolean;
}

/**
 * Accesses the SPFx MSGraphClientV3 from ServiceScope with asynchronous initialization. A replacement factory hides the former client immediately and initializes its own client; obsolete/unmounted results are ignored. Direct client calls do not update hook state. invoke tracks all pending calls for its current client and publishes errors only from the latest-started call; obsolete completions still settle their promises. Graph SDK/callback rejections are normalized to Error; the hook does not inspect successful response payloads.
 * @returns Native client, invocation state and initialization/readiness state.
 * @throws Missing SPFx provider or synchronous native factory-consumption failures.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxMSGraphClient } from '@apvee/spfx-react-toolkit';
 *
 * function ClientStatus() {
 *   const info = useSPFxMSGraphClient();
 *   return <span>{info.isReady ? 'Ready' : 'Unavailable'}</span>;
 * }
 * ```
 *
 * @useWhen an event or effect must call Graph with the SPFx-provided v3 client.
 *
 *
 * @example User profile with initialization request and retry
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxMSGraphClient } from '@apvee/spfx-react-toolkit';
 * // Mount under the initialized matching provider. These shapes express expected Graph responses.
 * function UserProfile() {
 *   const { invoke, isLoading, error, clearError, isInitializing, initError, isReady } = useSPFxMSGraphClient();
 *   const [data, setData] = React.useState<{ displayName?: string; jobTitle?: string; department?: string }>();
 *   const load = React.useCallback(() => {
 *     void invoke(client => client.api('/me').select('displayName,mail,jobTitle,department').get()).then(setData).catch(() => undefined);
 *   }, [invoke]);
 *   React.useEffect(() => { if (isReady) load(); }, [isReady, load]);
 *   if (isInitializing) return <p role="status">Initializing Graph</p>;
 *   if (initError) return <p role="alert">{initError.message}</p>;
 *   return <div><button disabled={!isReady || isLoading} onClick={() => { clearError(); load(); }}>Reload</button>
 *     {error && <p role="alert">{error.message}</p>}
 *     <p>{data?.displayName} {data?.jobTitle} {data?.department}</p>
 *   </div>;
 * }
 * ```
 *
 *
 * @example Read recent high-importance mail
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxMSGraphClient } from '@apvee/spfx-react-toolkit';
 * // Mount under the initialized matching provider. These shapes express expected Graph responses.
 * function RecentMail() {
 *   const { invoke, isLoading, error, clearError, isInitializing, initError, isReady } = useSPFxMSGraphClient();
 *   const [data, setData] = React.useState<{ value: { id: string; subject?: string; from?: { emailAddress?: { name?: string } } }[] }>();
 *   const load = React.useCallback(() => {
 *     // The last 30 days, newest first. Graph requires ordered properties first in the filter:
 *     // https://learn.microsoft.com/en-us/graph/api/user-list-messages#using-filter-and-orderby-in-the-same-query
 *     const windowEnd = new Date();
 *     const windowStart = new Date(windowEnd.getTime() - 30 * 24 * 60 * 60 * 1000);
 *     void invoke(client => client.api('/me/messages').version('v1.0')
 *       .select('subject,from,receivedDateTime')
 *       .filter(`receivedDateTime ge ${windowStart.toISOString()} and receivedDateTime le ${windowEnd.toISOString()} and importance eq 'high'`)
 *       .orderby('receivedDateTime DESC').top(10).get()).then(setData).catch(() => undefined);
 *   }, [invoke]);
 *   React.useEffect(() => { if (isReady) load(); }, [isReady, load]);
 *   if (isInitializing) return <p role="status">Initializing Graph</p>;
 *   if (initError) return <p role="alert">{initError.message}</p>;
 *   return <div><button disabled={!isReady || isLoading} onClick={() => { clearError(); load(); }}>Reload</button>
 *     {error && <p role="alert">{error.message}</p>}
 *     <ul>{data?.value.map(message => <li key={message.id}><strong>{message.subject}</strong><span>{message.from?.emailAddress?.name}</span></li>)}</ul>
 *   </div>;
 * }
 * ```
 *
 *
 * @example Read the coming week calendar view
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxMSGraphClient } from '@apvee/spfx-react-toolkit';
 * type GraphUtcStart = { dateTime?: string; timeZone?: string };
 * // This recipe requests UTC response values, not arbitrary Windows/IANA zone conversion.
 * // Graph separates offset-free dateTime from timeZone; preserve and check both.
 * // https://learn.microsoft.com/en-us/graph/api/user-list-calendarview#request-headers
 * function parseUtcStart(start: GraphUtcStart | undefined): Date | undefined {
 *   if (start?.timeZone !== 'UTC' || typeof start.dateTime !== 'string') return undefined;
 *   const match = /^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})(?:\.(\d{1,7}))?$/.exec(start.dateTime);
 *   if (!match) return undefined;
 *   // Graph fractions can have seven digits; Date represents milliseconds. Appending Z is essential.
 *   const iso = `${match[1]}.${(match[2] ?? '').padEnd(3, '0').slice(0, 3)}Z`;
 *   const instant = new Date(iso);
 *   return Number.isFinite(instant.getTime()) && instant.toISOString() === iso ? instant : undefined;
 * }
 * // Mount under the initialized matching provider. The shape expresses an expected Graph response.
 * function UpcomingEvents() {
 *   const { invoke, isLoading, error, clearError, isInitializing, initError, isReady } = useSPFxMSGraphClient();
 *   const [data, setData] = React.useState<{ value: { id: string; subject?: string; start?: GraphUtcStart }[] }>();
 *   const load = React.useCallback(() => {
 *     const windowStart = new Date();
 *     const windowEnd = new Date(windowStart.getTime() + 7 * 24 * 60 * 60 * 1000);
 *     void invoke(client => client.api('/me/calendarView').header('Prefer', 'outlook.timezone="UTC"')
 *       .query({ startDateTime: windowStart.toISOString(), endDateTime: windowEnd.toISOString() })
 *       .select('subject,start,end,location').orderby('start/dateTime').top(5).get()).then(setData).catch(() => undefined);
 *   }, [invoke]);
 *   React.useEffect(() => { if (isReady) load(); }, [isReady, load]);
 *   if (isInitializing) return <p role="status">Initializing Graph</p>;
 *   if (initError) return <p role="alert">{initError.message}</p>;
 *   return <div><button disabled={!isReady || isLoading} onClick={() => { clearError(); load(); }}>Reload</button>
 *     {error && <p role="alert">{error.message}</p>}
 *     <ul>{data?.value.map(event => {
 *       const start = parseUtcStart(event.start);
 *       return <li key={event.id}>{event.subject} - {start ? start.toLocaleString() : 'Start time unavailable'}</li>;
 *     })}</ul>
 *   </div>;
 * }
 * ```
 */
export function useSPFxMSGraphClient(): SPFxMSGraphClientInfo {
  const { consume } = useSPFxServiceScope();
  
  const factory = useMemo(() => consume<MSGraphClientFactory>(MSGraphClientFactory.serviceKey), [consume]);
  const factoryIdentity = useRef({ factory });
  if (factoryIdentity.current.factory !== factory) factoryIdentity.current = { factory };
  const current = factoryIdentity.current;
  const [initialization, setInitialization] = useState<{
    identity: typeof current;
    client?: MSGraphClientV3;
    error?: Error;
  }>();
  useEffect(() => {
    let disposed = false;
    Promise.resolve().then(() => factory.getClient('3')).then(graphClient => {
      if (!disposed && factoryIdentity.current === current) setInitialization({ identity: current, client: graphClient });
    }, failure => {
      if (!disposed && factoryIdentity.current === current) {
        setInitialization({ identity: current, error: failure instanceof Error ? failure : new Error(String(failure)) });
      }
    });
    return () => { disposed = true; };
  }, [factory, current]);
  const visible = initialization?.identity === current ? initialization : undefined;
  const client = visible?.client;
  const initError = visible?.error;
  const isInitializing = visible === undefined;

  // Use shared async invocation pattern
  const invocation = useAsyncInvoke(
    client,
    'Graph client not initialized. Wait for client to be available or check initError.'
  );
  
  // ═══════════════════════════════════════════════════════════════════════════
  // COMPUTED STATE & RETURN
  // ═══════════════════════════════════════════════════════════════════════════
  
  const { invoke, clearError } = invocation;
  const isLoading = client !== undefined && invocation.isLoading;
  const error = client === undefined ? undefined : invocation.error;

  // Computed: ready when client is available and no errors
  const isReady = client !== undefined && !isInitializing && !initError;
  
  return useMemo(() => ({
    client,
    invoke,
    isLoading,
    error,
    clearError,
    isInitializing,
    initError,
    isReady,
  }), [client, invoke, isLoading, error, clearError, isInitializing, initError, isReady]);
}

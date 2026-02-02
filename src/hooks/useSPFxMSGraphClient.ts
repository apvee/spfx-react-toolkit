// useSPFxMSGraphClient.ts
// Hook to access Microsoft Graph client with state management

import { useMemo, useState, useEffect, useRef } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';
import { MSGraphClientV3, MSGraphClientFactory } from '@microsoft/sp-http';
import { useAsyncInvoke } from './useAsyncInvoke.internal';

/**
 * Return type for useSPFxMSGraphClient hook
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
   * @param fn - Function that receives Graph client and returns a promise
   * @returns Promise with the result
   * @throws Error if client is not initialized yet
   * 
   * @example
   * ```tsx
   * const { invoke } = useSPFxMSGraphClient();
   * 
   * const user = await invoke(client => 
   *   client.api('/me').select('displayName,mail').get()
   * );
   * ```
   */
  readonly invoke: <T>(fn: (client: MSGraphClientV3) => Promise<T>) => Promise<T>;
  
  /** 
   * Loading state - true during invoke() calls.
   * Does not track direct client usage or initialization.
   */
  readonly isLoading: boolean;
  
  /** 
   * Last error from invoke() calls.
   * Does not capture errors from direct client usage or initialization.
   * @see initError for initialization errors
   */
  readonly error: Error | undefined;
  
  /** Clear the current error from invoke() calls */
  readonly clearError: () => void;

  /**
   * True while the Graph client is being initialized.
   * Use this to show a loading indicator during startup.
   * 
   * @example
   * ```tsx
   * const { client, isInitializing } = useSPFxMSGraphClient();
   * 
   * if (isInitializing) return <Spinner label="Initializing Graph..." />;
   * if (!client) return <Error message="Graph client unavailable" />;
   * ```
   */
  readonly isInitializing: boolean;

  /**
   * Error that occurred during client initialization.
   * If set, the client will remain undefined.
   * 
   * @example
   * ```tsx
   * const { initError } = useSPFxMSGraphClient();
   * 
   * if (initError) {
   *   return <MessageBar messageBarType={MessageBarType.error}>
   *     Failed to initialize Graph: {initError.message}
   *   </MessageBar>;
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
   * const { isReady, client, invoke } = useSPFxMSGraphClient();
   * 
   * if (!isReady) return <Spinner />;
   * 
   * // Safe to use client or invoke
   * const data = await invoke(c => c.api('/me').get());
   * ```
   */
  readonly isReady: boolean;
}

/**
 * Hook to access Microsoft Graph client with built-in state management
 * 
 * Provides native MSGraphClientV3 for authenticated Microsoft Graph API access.
 * Offers two usage patterns:
 * 
 * 1. **invoke()** - Automatic state management (loading + error tracking)
 * 2. **client** - Direct access for full control
 * 
 * For type safety, install @microsoft/microsoft-graph-types:
 * ```bash
 * npm install @microsoft/microsoft-graph-types --save-dev
 * ```
 * 
 * And import SPFx types:
 * ```typescript
 * import { MSGraphClientV3 } from '@microsoft/sp-http';
 * import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
 * ```
 * 
 * Requirements:
 * - SPFx ServiceScope with MSGraphClientFactory service (v1.15.0+)
 * - Appropriate Microsoft Graph permissions granted by admin
 * 
 * @remarks
 * This hook consumes MSGraphClientFactory from SPFx ServiceScope using dependency injection.
 * The factory is consumed lazily and cached. The factory.getClient() method is then called
 * asynchronously to obtain the MSGraphClientV3 instance, which may be undefined until initialized.
 * 
 * @example Using invoke with state management
 * ```tsx
 * import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
 * 
 * function UserProfile() {
 *   const { invoke, isLoading, error, clearError } = useSPFxMSGraphClient();
 *   const [user, setUser] = useState<MicrosoftGraph.User>();
 *   
 *   const loadUser = () => {
 *     invoke(client => 
 *       client.api('/me')
 *         .select('displayName,mail,jobTitle')
 *         .get()
 *     ).then(setUser);
 *   };
 *   
 *   useEffect(() => { loadUser(); }, []);
 *   
 *   if (isLoading) return <Spinner />;
 *   if (error) return (
 *     <MessageBar messageBarType={MessageBarType.error}>
 *       {error.message}
 *       <Link onClick={() => { clearError(); loadUser(); }}>Retry</Link>
 *     </MessageBar>
 *   );
 *   
 *   return <div>{user?.displayName}</div>;
 * }
 * ```
 * 
 * @example Using client directly for advanced control
 * ```tsx
 * import { MSGraphClientV3 } from '@microsoft/sp-http';
 * import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
 * 
 * function MessagesList() {
 *   const { client } = useSPFxMSGraphClient();
 *   const [messages, setMessages] = useState<MicrosoftGraph.Message[]>([]);
 *   const [loading, setLoading] = useState(false);
 *   
 *   if (!client) return <Spinner label="Initializing Graph client..." />;
 *   
 *   const loadMessages = async () => {
 *     setLoading(true);
 *     try {
 *       const result = await client.api('/me/messages')
 *         .version('v1.0')
 *         .select('subject,from,receivedDateTime')
 *         .filter("importance eq 'high'")
 *         .orderBy('receivedDateTime DESC')
 *         .top(20)
 *         .get();
 *       setMessages(result.value);
 *     } catch (err) {
 *       console.error(err);
 *     } finally {
 *       setLoading(false);
 *     }
 *   };
 *   
 *   return (
 *     <>
 *       <button onClick={loadMessages} disabled={loading}>Load</button>
 *       {loading && <Spinner />}
 *       <MessageList items={messages} />
 *     </>
 *   );
 * }
 * ```
 * 
 * @example CRUD operations with invoke
 * ```tsx
 * function ContactsManager() {
 *   const { invoke, isLoading, error } = useSPFxMSGraphClient();
 *   const [contacts, setContacts] = useState([]);
 *   
 *   const loadContacts = () => {
 *     invoke(client => 
 *       client.api('/me/contacts').get()
 *     ).then(result => setContacts(result.value));
 *   };
 *   
 *   const createContact = (contact: any) => {
 *     invoke(client => 
 *       client.api('/me/contacts').post(contact)
 *     ).then(newContact => setContacts([...contacts, newContact]));
 *   };
 *   
 *   const updateContact = (id: string, changes: any) => {
 *     invoke(client => 
 *       client.api(`/me/contacts/${id}`).patch(changes)
 *     ).then(loadContacts);
 *   };
 *   
 *   const deleteContact = (id: string) => {
 *     invoke(client => 
 *       client.api(`/me/contacts/${id}`).delete()
 *     ).then(() => setContacts(contacts.filter(c => c.id !== id)));
 *   };
 *   
 *   return (
 *     <ContactsUI 
 *       contacts={contacts} 
 *       loading={isLoading}
 *       error={error}
 *       onCreate={createContact}
 *       onUpdate={updateContact}
 *       onDelete={deleteContact}
 *     />
 *   );
 * }
 * ```
 * 
 * @example Batch operations with client
 * ```tsx
 * function BatchOperations() {
 *   const { client } = useSPFxMSGraphClient();
 *   
 *   if (!client) return <Spinner />;
 *   
 *   const batchRequest = async () => {
 *     const batch = {
 *       requests: [
 *         { id: '1', method: 'GET', url: '/me' },
 *         { id: '2', method: 'GET', url: '/me/messages?$top=5' }
 *       ]
 *     };
 *     
 *     const result = await client.api('/$batch').post(batch);
 *     console.log(result);
 *   };
 * }
 * ```
 */
export function useSPFxMSGraphClient(): SPFxMSGraphClientInfo {
  const { consume } = useSPFxServiceScope();
  
  // ═══════════════════════════════════════════════════════════════════════════
  // STATE
  // ═══════════════════════════════════════════════════════════════════════════
  
  const [client, setClient] = useState<MSGraphClientV3 | undefined>(undefined);
  const [isInitializing, setIsInitializing] = useState<boolean>(true);
  const [initError, setInitError] = useState<Error | undefined>(undefined);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // REFS (for cleanup and preventing double initialization)
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Track component mounted state to prevent memory leaks
  const isMountedRef = useRef<boolean>(true);
  
  // Track if initialization has been attempted (prevent double init)
  const initAttemptedRef = useRef<boolean>(false);
  
  // Cleanup on unmount
  useEffect(() => {
    return () => {
      isMountedRef.current = false;
    };
  }, []);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // FACTORY (lazy consume from ServiceScope)
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Lazy consume MSGraphClientFactory from ServiceScope (cached by useMemo)
  const factory = useMemo(() => {
    return consume<MSGraphClientFactory>(MSGraphClientFactory.serviceKey);
  }, [consume]);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // INITIALIZATION EFFECT
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Initialize Graph client (factory.getClient is async)
  useEffect(() => {
    // Prevent double initialization
    if (initAttemptedRef.current) {
      return;
    }
    initAttemptedRef.current = true;
    
    // Reset state for new initialization
    setIsInitializing(true);
    setInitError(undefined);
    
    // Get MSGraphClientV3 (version 3 of Microsoft Graph JavaScript Client Library)
    factory
      .getClient('3')
      .then((graphClient: MSGraphClientV3) => {
        // Only update state if still mounted
        if (isMountedRef.current) {
          setClient(graphClient);
          setIsInitializing(false);
        }
      })
      .catch((err: unknown) => {
        // Only update state if still mounted
        if (isMountedRef.current) {
          const error = err instanceof Error ? err : new Error(String(err));
          setInitError(error);
          setIsInitializing(false);
          console.error('Failed to initialize MSGraphClient:', error);
        }
      });
  }, [factory]);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // ASYNC INVOKE PATTERN
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Use shared async invocation pattern
  const { invoke, isLoading, error, clearError } = useAsyncInvoke(
    client,
    'Graph client not initialized. Wait for client to be available or check initError.'
  );
  
  // ═══════════════════════════════════════════════════════════════════════════
  // COMPUTED STATE & RETURN
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Computed: ready when client is available and no errors
  const isReady = client !== undefined && !isInitializing && !initError;
  
  return {
    client,
    invoke,
    isLoading,
    error,
    clearError,
    isInitializing,
    initError,
    isReady,
  };
}

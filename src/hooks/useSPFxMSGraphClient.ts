// useSPFxMSGraphClient.ts
// Hook to access Microsoft Graph client with state management

import { useMemo, useState, useCallback, useEffect } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';
import { MSGraphClientV3, MSGraphClientFactory } from '@microsoft/sp-http';

/**
 * Return type for useSPFxMSGraphClient hook
 */
export interface SPFxMSGraphClientInfo {
  /** 
   * Native MSGraphClientV3 from SPFx.
   * Provides access to Microsoft Graph API with built-in authentication.
   */
  readonly client: MSGraphClientV3 | undefined;
  
  /** 
   * Invoke Graph API call with automatic state management.
   * Tracks loading state and captures errors automatically.
   * 
   * @param fn - Function that receives Graph client and returns a promise
   * @returns Promise with the result
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
   * Does not track direct client usage.
   */
  readonly isLoading: boolean;
  
  /** 
   * Last error from invoke() calls.
   * Does not capture errors from direct client usage.
   */
  readonly error: Error | undefined;
  
  /** Clear the current error */
  readonly clearError: () => void;
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
  
  // Lazy consume MSGraphClientFactory from ServiceScope (cached by useMemo)
  const factory = useMemo(() => {
    return consume<MSGraphClientFactory>(MSGraphClientFactory.serviceKey);
  }, [consume]);
  
  const [client, setClient] = useState<MSGraphClientV3 | undefined>(undefined);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  
  // Initialize Graph client (factory.getClient is async)
  useEffect(() => {
    // Get MSGraphClientV3 (version 3 of Microsoft Graph JavaScript Client Library)
    factory
      .getClient('3')
      .then((graphClient: MSGraphClientV3) => {
        setClient(graphClient);
      })
      .catch((err: Error) => {
        console.error('Failed to initialize MSGraphClient:', err);
      });
  }, [factory]);
  
  // Invoke with automatic state management
  const invoke = useCallback(
    async <T>(fn: (client: MSGraphClientV3) => Promise<T>): Promise<T> => {
      if (!client) {
        throw new Error('Graph client not initialized. Wait for client to be available.');
      }
      
      setIsLoading(true);
      setError(undefined);
      
      try {
        const result = await fn(client);
        return result;
      } catch (err) {
        const error = err instanceof Error ? err : new Error(String(err));
        setError(error);
        throw error;
      } finally {
        setIsLoading(false);
      }
    },
    [client]
  );
  
  // Clear error helper
  const clearError = useCallback(() => {
    setError(undefined);
  }, []);
  
  return {
    client,
    invoke,
    isLoading,
    error,
    clearError,
  };
}

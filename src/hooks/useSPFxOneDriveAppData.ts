// useSPFxOneDriveAppData.ts
// Hook to manage JSON files in OneDrive appRoot folder with state management

import { useState, useCallback, useEffect, useRef } from 'react';
import { useSPFxMSGraphClient } from './useSPFxMSGraphClient';

/**
 * Return type for useSPFxOneDriveAppData hook
 */
export interface SPFxOneDriveAppDataResult<T> {
  /** 
   * The loaded data from OneDrive.
   * Undefined if not loaded yet or on error.
   */
  readonly data: T | undefined;
  
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
   * Manually load/reload the file from OneDrive.
   * Updates data, isLoading, and error states.
   * 
   * @returns Promise that resolves when load completes
   * 
   * @example
   * ```tsx
   * const { data, load } = useSPFxOneDriveAppData<Config>('config.json', undefined, false);
   * 
   * // Load on button click
   * <button onClick={load}>Refresh</button>
   * ```
   */
  readonly load: () => Promise<void>;
  
  /** 
   * Write data to OneDrive file.
   * Creates file if it doesn't exist, updates if it does.
   * Updates isWriting and writeError states.
   * 
   * @param content - Data to write (will be JSON.stringify'd)
   * @returns Promise that resolves when write completes
   * 
   * @example
   * ```tsx
   * const { write, isWriting } = useSPFxOneDriveAppData<Config>('config.json');
   * 
   * const handleSave = async () => {
   *   await write({ theme: 'dark', language: 'en' });
   * };
   * ```
   */
  readonly write: (content: T) => Promise<void>;
  
  /** 
   * Computed state: true if data is loaded successfully.
   * Equivalent to: !isLoading && !error && data !== undefined
   * 
   * Useful for conditional rendering:
   * ```tsx
   * if (!isReady) return <Spinner />;
   * return <div>{data.title}</div>;
   * ```
   */
  readonly isReady: boolean;
}

/**
 * Hook to manage JSON files in user's OneDrive appRoot folder
 * 
 * Provides unified read/write operations for JSON data stored in OneDrive's special
 * appRoot folder (accessible per-app, user-scoped storage).
 * 
 * Features:
 * - Automatic JSON serialization/deserialization
 * - Separate loading states for read/write operations
 * - Optional auto-fetch on mount
 * - Folder/namespace support for file organization
 * - Type-safe with TypeScript generics
 * - Memory leak safe with mounted state tracking
 * - Error handling for both read and write operations
 * 
 * Requirements:
 * - Microsoft Graph permissions: Files.ReadWrite or Files.ReadWrite.AppFolder
 * - User must be authenticated
 * 
 * @param fileName - Name of the JSON file (e.g., 'config.json', 'settings.json')
 * @param folder - Optional folder/namespace identifier for file organization.
 *                 Will be sanitized to prevent path traversal.
 *                 Examples: 'my-app', instanceId (GUID), 'config-v2'
 * @param autoFetch - Whether to automatically load file on mount. Default: true
 * 
 * @returns Object with data, loading states, error states, and read/write functions
 * 
 * @example Basic usage - auto-fetch from root
 * ```tsx
 * import type { MyConfig } from './types';
 * 
 * function ConfigPanel() {
 *   const { data, isLoading, error, write, isWriting } = 
 *     useSPFxOneDriveAppData<MyConfig>('config.json');
 *   
 *   if (isLoading) return <Spinner label="Loading configuration..." />;
 *   if (error) return <MessageBar messageBarType={MessageBarType.error}>
 *     Failed to load: {error.message}
 *   </MessageBar>;
 *   
 *   const handleSave = async (newConfig: MyConfig) => {
 *     try {
 *       await write(newConfig);
 *       console.log('Saved successfully!');
 *     } catch (err) {
 *       console.error('Save failed:', err);
 *     }
 *   };
 *   
 *   return (
 *     <div>
 *       <TextField 
 *         value={data?.title} 
 *         onChange={(_, val) => handleSave({ ...data, title: val })}
 *         disabled={isWriting}
 *       />
 *       {isWriting && <Spinner label="Saving..." />}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example With folder namespace
 * ```tsx
 * // Store files in a dedicated folder
 * const { data, write } = useSPFxOneDriveAppData<State>(
 *   'state.json',
 *   'my-app-v2'  // Files stored in appRoot:/my-app-v2/state.json
 * );
 * ```
 * 
 * @example Per-instance storage (multi-instance support)
 * ```tsx
 * // Each WebPart instance has its own data
 * const { id } = useSPFxInstanceInfo();
 * const { data, write } = useSPFxOneDriveAppData<Settings>(
 *   'settings.json',
 *   id  // Files stored in appRoot:/abc-123-guid/settings.json
 * );
 * ```
 * 
 * @example Lazy loading (manual load)
 * ```tsx
 * const { data, load, isLoading, write } = useSPFxOneDriveAppData<Cache>(
 *   'cache.json',
 *   'my-app',
 *   false  // Don't auto-fetch
 * );
 * 
 * return (
 *   <div>
 *     <button onClick={load} disabled={isLoading}>
 *       {isLoading ? 'Loading...' : 'Load Cache'}
 *     </button>
 *     {data && <pre>{JSON.stringify(data, null, 2)}</pre>}
 *   </div>
 * );
 * ```
 * 
 * @example Multiple files in same namespace
 * ```tsx
 * function MyApp() {
 *   const config = useSPFxOneDriveAppData<Config>('config.json', 'myapp');
 *   const state = useSPFxOneDriveAppData<State>('state.json', 'myapp');
 *   const cache = useSPFxOneDriveAppData<Cache>('cache.json', 'myapp');
 *   
 *   // All files stored in appRoot:/myapp/
 *   // Easy to manage and clean up as a group
 * }
 * ```
 * 
 * @example Error handling and retry
 * ```tsx
 * function DataManager() {
 *   const { data, error, load, writeError, write, isReady } = 
 *     useSPFxOneDriveAppData<MyData>('data.json');
 *   
 *   if (error) {
 *     return (
 *       <MessageBar 
 *         messageBarType={MessageBarType.error}
 *         actions={<button onClick={load}>Retry</button>}
 *       >
 *         Load failed: {error.message}
 *       </MessageBar>
 *     );
 *   }
 *   
 *   if (writeError) {
 *     return (
 *       <MessageBar messageBarType={MessageBarType.warning}>
 *         Save failed: {writeError.message}
 *       </MessageBar>
 *     );
 *   }
 *   
 *   if (!isReady) return <Spinner />;
 *   
 *   return <DataDisplay data={data} onSave={write} />;
 * }
 * ```
 * 
 * @example CRUD-like operations
 * ```tsx
 * interface TodoList {
 *   items: Array<{ id: string; text: string; done: boolean }>;
 * }
 * 
 * function TodoApp() {
 *   const { data, write, isLoading, isWriting } = 
 *     useSPFxOneDriveAppData<TodoList>('todos.json', 'todo-app');
 *   
 *   const addTodo = async (text: string) => {
 *     const newItem = { id: crypto.randomUUID(), text, done: false };
 *     await write({
 *       items: [...(data?.items ?? []), newItem]
 *     });
 *   };
 *   
 *   const toggleTodo = async (id: string) => {
 *     await write({
 *       items: data?.items.map(item => 
 *         item.id === id ? { ...item, done: !item.done } : item
 *       ) ?? []
 *     });
 *   };
 *   
 *   const deleteTodo = async (id: string) => {
 *     await write({
 *       items: data?.items.filter(item => item.id !== id) ?? []
 *     });
 *   };
 *   
 *   if (isLoading) return <Spinner />;
 *   
 *   return (
 *     <div>
 *       <TodoList 
 *         items={data?.items ?? []} 
 *         onToggle={toggleTodo}
 *         onDelete={deleteTodo}
 *       />
 *       <AddTodoForm onAdd={addTodo} disabled={isWriting} />
 *     </div>
 *   );
 * }
 * ```
 */
export function useSPFxOneDriveAppData<T = unknown>(
  fileName: string,
  folder?: string,
  autoFetch: boolean = true
): SPFxOneDriveAppDataResult<T> {
  const { client } = useSPFxMSGraphClient();
  
  // State management
  const [data, setData] = useState<T | undefined>(undefined);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  const [isWriting, setIsWriting] = useState<boolean>(false);
  const [writeError, setWriteError] = useState<Error | undefined>(undefined);
  
  // Track component mounted state to prevent memory leaks
  const isMounted = useRef<boolean>(true);
  
  useEffect(() => {
    isMounted.current = true;
    return () => {
      isMounted.current = false;
    };
  }, []);
  
  /**
   * Build Graph API path with optional folder namespace
   * Sanitizes folder name to prevent path traversal attacks
   */
  const buildApiPath = useCallback((file: string, folderName?: string): string => {
    const basePath = '/me/drive/special/appRoot:';
    
    if (folderName) {
      // Sanitize folder name: only allow alphanumeric, hyphens, underscores
      // This prevents path traversal (../) and other injection attacks
      const safeFolderName = folderName.replace(/[^a-zA-Z0-9-_]/g, '-');
      return `${basePath}/${safeFolderName}/${file}:/content`;
    }
    
    return `${basePath}/${file}:/content`;
  }, []);
  
  /**
   * Load file from OneDrive
   * Updates data, isLoading, and error states
   */
  const load = useCallback(async (): Promise<void> => {
    if (!client) {
      console.warn('Graph client not available yet. Skipping load.');
      return;
    }
    
    if (!fileName) {
      console.warn('fileName is required. Skipping load.');
      return;
    }
    
    setIsLoading(true);
    setError(undefined);
    
    try {
      const apiPath = buildApiPath(fileName, folder);
      const fileContent = await client.api(apiPath).get();
      
      if (isMounted.current) {
        // Parse JSON if response is string, otherwise use as-is
        if (typeof fileContent === 'string') {
          try {
            setData(JSON.parse(fileContent) as T);
          } catch (parseError) {
            throw new Error(`Failed to parse JSON: ${parseError instanceof Error ? parseError.message : 'Unknown error'}`);
          }
        } else {
          setData(fileContent as T);
        }
      }
    } catch (err) {
      if (isMounted.current) {
        const error = err instanceof Error ? err : new Error(String(err));
        setError(error);
        // Don't throw - allow component to handle error via state
        console.error('Failed to load file from OneDrive:', error);
      }
    } finally {
      if (isMounted.current) {
        setIsLoading(false);
      }
    }
  }, [client, fileName, folder, buildApiPath]);
  
  /**
   * Write data to OneDrive file
   * Creates file if it doesn't exist, updates if it does (upsert)
   * Updates isWriting and writeError states
   */
  const write = useCallback(async (content: T): Promise<void> => {
    if (!client) {
      throw new Error('Graph client not available. Cannot write file.');
    }
    
    if (!fileName) {
      throw new Error('fileName is required. Cannot write file.');
    }
    
    setIsWriting(true);
    setWriteError(undefined);
    
    try {
      const apiPath = buildApiPath(fileName, folder);
      
      // Always stringify to ensure valid JSON
      const jsonContent = JSON.stringify(content);
      
      await client
        .api(apiPath)
        .header('Content-Type', 'application/json')
        .put(jsonContent);
      
      if (isMounted.current) {
        // Update local data to reflect successful write
        setData(content);
        // Clear read error if write succeeds (fresh state)
        setError(undefined);
      }
    } catch (err) {
      if (isMounted.current) {
        const error = err instanceof Error ? err : new Error(String(err));
        setWriteError(error);
        console.error('Failed to write file to OneDrive:', error);
      }
      // Re-throw to allow caller to handle
      throw err;
    } finally {
      if (isMounted.current) {
        setIsWriting(false);
      }
    }
  }, [client, fileName, folder, buildApiPath]);
  
  // Auto-fetch on mount if enabled
  useEffect(() => {
    if (autoFetch && client && fileName) {
      load().catch(() => {
        // Error already handled in load() function
      });
    }
  }, [autoFetch, client, fileName, load]);
  
  // Computed state: ready when data loaded successfully
  const isReady = !isLoading && !error && data !== undefined;
  
  return {
    data,
    isLoading,
    error,
    isWriting,
    writeError,
    load,
    write,
    isReady,
  };
}

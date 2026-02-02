// useSPFxOneDriveAppData.ts
// Hook to manage JSON files in OneDrive appRoot folder with state management

import { useState, useCallback, useEffect, useRef, useMemo } from 'react';
import { useSPFxMSGraphClient } from './useSPFxMSGraphClient';

// ═══════════════════════════════════════════════════════════════════════════
// PURE FUNCTIONS (extracted for stability and testability)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Build Graph API path with optional folder namespace.
 * Sanitizes folder name to prevent path traversal attacks.
 * 
 * @param fileName - Name of the file
 * @param folderName - Optional folder namespace
 * @returns Full Graph API path for file content
 */
function buildApiPath(fileName: string, folderName?: string): string {
  const basePath = '/me/drive/special/appRoot:';

  if (folderName) {
    // Sanitize folder name: only allow alphanumeric, hyphens, underscores
    // This prevents path traversal (../) and other injection attacks
    const safeFolderName = folderName.replace(/[^a-zA-Z0-9-_]/g, '-');
    return `${basePath}/${safeFolderName}/${fileName}:/content`;
  }

  return `${basePath}/${fileName}:/content`;
}

/**
 * Check if an error indicates a 404 / itemNotFound response from Graph API.
 * 
 * @param err - The error to check
 * @returns True if the error indicates file not found
 */
function isNotFoundError(err: unknown): boolean {
  const anyErr = err as {
    statusCode?: number;
    status?: number;
    code?: string;
    message?: string;
    body?: { error?: { code?: string; message?: string } };
  };

  // Check status codes
  if (anyErr?.statusCode === 404 || anyErr?.status === 404) return true;

  // Check error codes
  const code = anyErr?.code ?? anyErr?.body?.error?.code;
  if (code && /itemnotfound/i.test(code)) return true;

  // Check error messages as fallback
  const message = anyErr?.message ?? anyErr?.body?.error?.message;
  if (message && /(\b404\b|not found|itemnotfound)/i.test(message)) return true;

  return false;
}

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
   * True if the file does not exist in OneDrive (404 / itemNotFound).
   * In the legacy signature this is treated as a non-error.
   */
  readonly isNotFound: boolean;

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
 * Optional configuration for useSPFxOneDriveAppData.
 */
export interface SPFxOneDriveAppDataOptions<T> {
  /** Optional folder/namespace identifier for file organization. */
  folder?: string;

  /** Whether to automatically load file on mount. Default: true */
  autoFetch?: boolean;

  /**
   * Default value to use when the file is missing (404).
   * If provided, load() will set data to this value and will not set error.
   */
  defaultValue?: T;

  /**
   * If true, when the file is missing (404) and defaultValue is provided,
   * the hook will attempt to create the file by writing defaultValue.
   */
  createIfMissing?: boolean;
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
  options?: SPFxOneDriveAppDataOptions<T>
): SPFxOneDriveAppDataResult<T> {
  const {
    client,
    isReady: isClientReady,
    isInitializing: isClientInitializing,
    initError: clientInitError
  } = useSPFxMSGraphClient();

  // ═══════════════════════════════════════════════════════════════════════════
  // OPTIONS (stabilized with useMemo to prevent unnecessary re-renders)
  // ═══════════════════════════════════════════════════════════════════════════

  const resolvedOptions = useMemo(() => options ?? {}, [options]);

  // Extract stable primitive values for dependency arrays
  const folder = resolvedOptions.folder;
  const shouldAutoFetch = resolvedOptions.autoFetch ?? true;
  const defaultValue = resolvedOptions.defaultValue;
  const createIfMissing = resolvedOptions.createIfMissing ?? false;

  // ═══════════════════════════════════════════════════════════════════════════
  // STATE MANAGEMENT
  // ═══════════════════════════════════════════════════════════════════════════

  const [data, setData] = useState<T | undefined>(defaultValue);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  const [isWriting, setIsWriting] = useState<boolean>(false);
  const [writeError, setWriteError] = useState<Error | undefined>(undefined);
  const [isNotFound, setIsNotFound] = useState<boolean>(false);

  // ═══════════════════════════════════════════════════════════════════════════
  // REFS (for cleanup and stable references)
  // ═══════════════════════════════════════════════════════════════════════════

  // Track component mounted state to prevent memory leaks
  const isMountedRef = useRef<boolean>(true);
  useEffect(() => {
    return () => {
      isMountedRef.current = false;
    };
  }, []);

  // Track if createIfMissing write has been attempted (to prevent multiple writes)
  const createAttemptedRef = useRef<boolean>(false);

  // ═══════════════════════════════════════════════════════════════════════════
  // WRITE CALLBACK (defined first, no dependency on load)
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Write data to OneDrive file
   * Creates file if it doesn't exist, updates if it does (upsert)
   * Updates isWriting and writeError states
   */
  const write = useCallback(async (content: T): Promise<void> => {
    if (!client) {
      if (isClientInitializing) {
        throw new Error('Graph client is still initializing. Please wait and try again.');
      }
      if (clientInitError) {
        throw new Error(`Graph client initialization failed: ${clientInitError.message}`);
      }
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

      if (isMountedRef.current) {
        // Update local data to reflect successful write
        setData(content);
        setIsNotFound(false);
        // Clear read error if write succeeds (fresh state)
        setError(undefined);
      }
    } catch (err) {
      if (isMountedRef.current) {
        const writeErr = err instanceof Error ? err : new Error(String(err));
        setWriteError(writeErr);
        console.error('Failed to write file to OneDrive:', writeErr);
      }
      // Re-throw to allow caller to handle
      throw err;
    } finally {
      if (isMountedRef.current) {
        setIsWriting(false);
      }
    }
  }, [client, fileName, folder, isClientInitializing, clientInitError]);

  // ═══════════════════════════════════════════════════════════════════════════
  // LOAD CALLBACK (NO dependency on write - uses effect for createIfMissing)
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Load file from OneDrive
   * Updates data, isLoading, and error states
   * Does NOT call write directly - createIfMissing is handled by separate effect
   */
  const load = useCallback(async (): Promise<void> => {
    if (!client) {
      if (isClientInitializing) {
        console.info('Graph client is still initializing. Skipping load - will auto-retry when ready.');
        return;
      }
      if (clientInitError) {
        console.error('Graph client initialization failed:', clientInitError.message);
        return;
      }
      console.warn('Graph client not available. Skipping load.');
      return;
    }

    if (!fileName) {
      console.warn('fileName is required. Skipping load.');
      return;
    }

    // Reset createAttempted flag when load is called (fresh attempt)
    createAttemptedRef.current = false;

    setIsLoading(true);
    setError(undefined);
    setIsNotFound(false);

    try {
      const apiPath = buildApiPath(fileName, folder);
      const fileContent = await client.api(apiPath).get();

      if (isMountedRef.current) {
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
        setIsNotFound(false);
      }
    } catch (err) {
      if (isMountedRef.current) {
        const notFound = isNotFoundError(err);
        setIsNotFound(notFound);

        if (notFound) {
          // Missing file is treated as a non-error.
          // Set data to defaultValue if provided, otherwise undefined
          if (defaultValue !== undefined) {
            setData(defaultValue);
          } else {
            setData(undefined);
          }
          setError(undefined);
          console.info('OneDrive file not found. isNotFound=true');
          // NOTE: createIfMissing is handled by separate useEffect
          return;
        }

        const loadError = err instanceof Error ? err : new Error(String(err));
        setError(loadError);
        console.error('Failed to load file from OneDrive:', loadError);
      }
    } finally {
      if (isMountedRef.current) {
        setIsLoading(false);
      }
    }
  }, [client, fileName, folder, defaultValue, isClientInitializing, clientInitError]); // ← NO write, NO createIfMissing

  // ═══════════════════════════════════════════════════════════════════════════
  // EFFECTS
  // ═══════════════════════════════════════════════════════════════════════════

  // Auto-fetch on mount if enabled (wait for client to be ready)
  useEffect(() => {
    if (shouldAutoFetch && isClientReady && fileName) {
      load().catch(() => {
        // Error already handled in load() function
      });
    }
  }, [shouldAutoFetch, isClientReady, fileName, load]);

  // Separate effect for createIfMissing - reacts to isNotFound state
  // This breaks the circular dependency: load → write
  useEffect(() => {
    // Guard conditions:
    // 1. File must be not found
    // 2. createIfMissing must be enabled
    // 3. defaultValue must be provided
    // 4. Must not be currently writing (prevent double-write)
    // 5. Must not have already attempted create (prevent infinite loop)
    // 6. Must not be currently loading (wait for load to complete)
    if (
      isNotFound &&
      createIfMissing &&
      defaultValue !== undefined &&
      !isWriting &&
      !createAttemptedRef.current &&
      !isLoading
    ) {
      createAttemptedRef.current = true;
      write(defaultValue).catch((writeErr) => {
        // write() already updates writeError state
        console.error('Failed to create missing file in OneDrive:', writeErr);
      });
    }
  }, [isNotFound, createIfMissing, defaultValue, isWriting, isLoading, write]);

  // ═══════════════════════════════════════════════════════════════════════════
  // COMPUTED STATE & RETURN
  // ═══════════════════════════════════════════════════════════════════════════

  // Computed state: ready when data loaded successfully
  const isReady = !isLoading && !error && data !== undefined;

  return {
    data,
    isLoading,
    error,
    isWriting,
    writeError,
    isNotFound,
    load,
    write,
    isReady,
  };
}

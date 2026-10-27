// useSPFxOneDriveAppData.ts
// Hook to manage JSON files in OneDrive approot folder with state management

import { useState, useCallback, useEffect, useRef, useMemo } from 'react';
import { useSPFxMSGraphClient } from './useSPFxMSGraphClient';
import { createSPFxOneDriveAppDataService } from '../services/spfx-onedrive-app-data.service';

/**
 * Return type for useSPFxOneDriveAppData hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const file: SPFxOneDriveAppDataResult<{ theme: string }> = useSPFxOneDriveAppData('settings.json');
 * console.log(file.data?.theme, file.isNotFound);
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxOneDriveAppDataResult<T> {
  /**
   * Current parsed data or fallback, initially defaultValue. Failed reloads preserve previous data; inspect error/isReady.
   */
  readonly data: T | undefined;

  /** 
   * Loading state for read operations.
   * True during initial load or manual load() calls.
   */
  readonly isLoading: boolean;

  /**
   * Error from the latest-started read; cleared when a read starts or the latest write succeeds.
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
   * Error from the latest operation when it was a write; cleared when a new write starts or the target changes. Reads do not clear it.
   */
  readonly writeError: Error | undefined;

  /**
   * Reads Graph metadata and downloads JSON via native fetch without authorization headers/cookies. Missing metadata uses defaultValue; failed reads publish error and resolve. Unavailable clients or empty file names skip the read.
   * @returns Resolves after the read attempt; inspect error and isNotFound.
   * @example
   * ```tsx
   * // Inside a React component using the host Fluent renderer:
   * const file = useSPFxOneDriveAppData<{ theme: string }>('settings.json', { autoFetch: false });
   * return <button onClick={file.load}>Refresh</button>;
   * ```
   */
  readonly load: () => Promise<void>;

  /**
   * Creates or replaces JSON file content and publishes successful data when the request is current.
   * @param content - JSON-serializable data to write.
   * @returns Resolves after the write succeeds.
   * @throws Missing/initializing client, empty fileName, serialization or Graph failures.
   */
  readonly write: (content: T) => Promise<void>;

  /**
   * True when !isLoading, no read error and data is defined; fallback/default data can satisfy this without a successful server read, and write status is independent.
   */
  readonly isReady: boolean;
}

/**
 * Optional configuration for useSPFxOneDriveAppData.
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const options: SPFxOneDriveAppDataOptions<{ theme: string }> = { folder: 'dashboard', defaultValue: { theme: 'dark' }, createIfMissing: true };
 * const file = useSPFxOneDriveAppData('settings.json', options);
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxOneDriveAppDataOptions<T> {
  /**
   * Optional folder namespace; characters other than letters, digits, hyphens and underscores become hyphens. fileName is not sanitized.
   */
  folder?: string;

  /** Whether to automatically load file on mount. Default: true */
  autoFetch?: boolean;

  /**
   * Initial/fallback data when a file is confirmed missing. Changing it alone does not reset current data; future missing reads use the latest value.
   */
  defaultValue?: T;

  /**
   * Defaults to false. After a current confirmed missing read, writes a defined defaultValue once per load attempt; failed downloads do not trigger creation.
   */
  createIfMissing?: boolean;
}

/**
 * Manages JSON content in the authenticated user's OneDrive app root. Requires Graph permissions for the requested read/write operations and native fetch for preauthenticated downloads. T is an expected shape without runtime validation. Auto-load defaults to true; target changes reset state and obsolete completions cannot publish it. Read failures publish error and resolve; writes reject and publish writeError. Recognized missing metadata returns fallback data without an error.
 * @typeParam T - Expected JSON data shape.
 * @param fileName - File name, passed into the Graph path without sanitization.
 * @param options - Folder, auto-load, fallback and missing-file creation settings.
 * @returns Read/write state and explicit load/write methods.
 * @throws When called outside an SPFx provider.
 *
 * @example
 * ```tsx
 * function Settings() {
 *   const file = useSPFxOneDriveAppData<{ theme: string }>('settings.json', {
 *     folder: 'dashboard', autoFetch: false, defaultValue: { theme: 'dark' }
 *   });
 *   return <button onClick={file.load} disabled={file.isLoading}>Load settings</button>;
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
  const defaultValueRef = useRef(defaultValue);
  defaultValueRef.current = defaultValue;
  const createIfMissing = resolvedOptions.createIfMissing ?? false;

  const oneDriveAppDataService = useMemo(() => (
    client ? createSPFxOneDriveAppDataService(client) : undefined
  ), [client]);

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

  const createAttemptedRef = useRef<boolean>(false);
  const identityRef = useRef({ service: oneDriveAppDataService, fileName, folder,
    latest: 0, reads: 0, writes: 0, missingRequest: 0 });
  if (identityRef.current.service !== oneDriveAppDataService ||
      identityRef.current.fileName !== fileName || identityRef.current.folder !== folder) {
    identityRef.current = { service: oneDriveAppDataService, fileName, folder,
      latest: 0, reads: 0, writes: 0, missingRequest: 0 };
  }
  useEffect(() => {
    setData(defaultValueRef.current);
    setError(undefined);
    setWriteError(undefined);
    setIsLoading(false);
    setIsWriting(false);
    setIsNotFound(false);
    createAttemptedRef.current = false;
  }, [oneDriveAppDataService, fileName, folder]);

  const currentIdentity = identityRef.current;

  // ═══════════════════════════════════════════════════════════════════════════
  // WRITE CALLBACK (defined first, no dependency on load)
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Write data to OneDrive file
   * Creates file if it doesn't exist, updates if it does (upsert)
   * Updates isWriting and writeError states
   */
  const write = useCallback(async (content: T): Promise<void> => {
    const service = oneDriveAppDataService;

    if (!client || !service) {
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

    const identity = currentIdentity;
    const request = ++identity.latest;
    const isIdentityCurrent = (): boolean => isMountedRef.current && identityRef.current === identity &&
      identity.service === service && identity.fileName === fileName && identity.folder === folder;
    identity.writes++;
    identity.missingRequest = 0;
    if (isIdentityCurrent()) { setIsWriting(true); setWriteError(undefined); }

    try {
      await service.write<T>(fileName, content, folder);

      if (isIdentityCurrent() && request === identity.latest) {
        // Update local data to reflect successful write
        setData(content);
        setIsNotFound(false);
        // Clear read error if write succeeds (fresh state)
        setError(undefined);
      }
    } catch (err) {
      if (isIdentityCurrent() && request === identity.latest) {
        const writeErr = err instanceof Error ? err : new Error(String(err));
        setWriteError(writeErr);
        console.error('Failed to write file to OneDrive:', writeErr);
      }
      // Re-throw to allow caller to handle
      throw err;
    } finally {
      identity.writes--;
      if (isIdentityCurrent()) setIsWriting(identity.writes > 0);
    }
  }, [client, oneDriveAppDataService, fileName, folder, isClientInitializing, clientInitError, currentIdentity]);

  // ═══════════════════════════════════════════════════════════════════════════
  // LOAD CALLBACK (NO dependency on write - uses effect for createIfMissing)
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Load file from OneDrive
   * Updates data, isLoading, and error states
   * Does NOT call write directly - createIfMissing is handled by separate effect
   */
  const load = useCallback(async (): Promise<void> => {
    const service = oneDriveAppDataService;

    if (!client || !service) {
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
    const identity = currentIdentity;
    if (identityRef.current === identity) createAttemptedRef.current = false;
    const request = ++identity.latest;
    const fallbackValue = defaultValueRef.current;
    const isIdentityCurrent = (): boolean => isMountedRef.current && identityRef.current === identity &&
      identity.service === service && identity.fileName === fileName && identity.folder === folder;
    identity.reads++;
    identity.missingRequest = 0;
    if (isIdentityCurrent()) {
      setIsLoading(true);
      setError(undefined);
      setIsNotFound(false);
    }

    try {
      const readResult = await service.read<T>(fileName, folder);

      if (isIdentityCurrent() && request === identity.latest) {
        if (readResult.isNotFound) {
          identity.missingRequest = request;
          // Missing file is treated as a non-error.
          // Set data to defaultValue if provided, otherwise undefined
          if (fallbackValue !== undefined) {
            setData(fallbackValue);
          } else {
            setData(undefined);
          }
          setIsNotFound(true);
          setError(undefined);
          console.info('OneDrive file not found. isNotFound=true');
          // NOTE: createIfMissing is handled by separate useEffect
          return;
        }

        setData(readResult.data);
        setIsNotFound(false);
      }
    } catch (err) {
      if (isIdentityCurrent() && request === identity.latest) {
        setIsNotFound(false);
        const loadError = err instanceof Error ? err : new Error(String(err));
        setError(loadError);
        console.error('Failed to load file from OneDrive:', loadError);
      }
    } finally {
      identity.reads--;
      if (isIdentityCurrent()) setIsLoading(identity.reads > 0);
    }
  }, [client, oneDriveAppDataService, fileName, folder, isClientInitializing, clientInitError, currentIdentity]); // ← NO write, NO createIfMissing

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
      identityRef.current.missingRequest > 0 &&
      identityRef.current.missingRequest === identityRef.current.latest &&
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

  return useMemo(() => ({
    data,
    isLoading,
    error,
    isWriting,
    writeError,
    isNotFound,
    load,
    write,
    isReady,
  }), [data, isLoading, error, isWriting, writeError, isNotFound, load, write, isReady]);
}

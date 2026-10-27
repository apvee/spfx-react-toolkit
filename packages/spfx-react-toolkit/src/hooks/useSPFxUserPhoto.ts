// useSPFxUserPhoto.ts
// Hook to load user photos from Microsoft Graph API

import { useState, useCallback, useEffect, useRef, useMemo } from 'react';
import { useSPFxMSGraphClient } from './useSPFxMSGraphClient';
import { createSPFxUserPhotoService } from '../services/spfx-user-photo.service';

/**
 * Available photo sizes from Microsoft Graph
 *
 * @example
 * ```tsx
 * const size: SPFxUserPhotoSize = '96x96';
 * ```
 */
export type SPFxUserPhotoSize = 
  | '48x48' 
  | '64x64' 
  | '96x96' 
  | '120x120' 
  | '240x240' 
  | '360x360' 
  | '432x432' 
  | '504x504' 
  | '648x648';

/**
 * Options for loading a specific user's photo
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const options: SPFxUserPhotoOptions = { email: 'user@contoso.com', size: '96x96', autoFetch: false };
 * const photo = useSPFxUserPhoto(options);
 * ```
 */
export interface SPFxUserPhotoOptions {
  /**
   * Graph user/object ID; takes precedence over email when nonempty.
   */
  userId?: string;
  
  /**
   * User email/UPN; used when userId is absent/empty, otherwise ignored. Without either identifier, targets /me.
   */
  email?: string;
  
  /**
   * Photo size to retrieve.
   * Default: '240x240'
   * 
   * Available sizes:
   * - 48x48: Tiny avatar
   * - 64x64: Small avatar
   * - 96x96: Standard avatar
   * - 120x120: Medium avatar
   * - 240x240: Large avatar (default)
   * - 360x360+: Extra large profiles
   */
  size?: SPFxUserPhotoSize;
  
  /**
   * Whether to automatically fetch photo on mount.
   * Default: true
   */
  autoFetch?: boolean;
}

/**
 * Return type for useSPFxUserPhoto hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const photo: SPFxUserPhotoResult = useSPFxUserPhoto();
 * return photo.photoUrl ? <img src={photo.photoUrl} alt="Profile" /> : <span>No photo</span>;
 * ```
 */
export interface SPFxUserPhotoResult {
  /**
   * Object URL created from the photo Blob. Revoked when replaced, target changes, on load failure or unmount; do not retain it beyond this lifecycle.
   */
  readonly photoUrl: string | undefined;
  
  /**
   * Raw photo blob data.
   * Useful if you need to process the image (upload, transform, etc.)
   */
  readonly photoBlob: Blob | undefined;
  
  /**
   * Loading state.
   * True while fetching photo from Microsoft Graph.
   */
  readonly isLoading: boolean;
  
  /**
   * Last error encountered.
   * Common errors:
   * - 404: Photo not found (user has no photo)
   * - 403: Insufficient permissions
   * - 401: Authentication failed
   */
  readonly error: Error | undefined;
  
  /**
   * Reloads the current photo; failures publish error and resolve rather than rejecting. Obsolete/unmounted callbacks skip work.
   * @returns Resolves after the attempt; inspect error/photoUrl.
   */
  readonly reload: () => Promise<void>;
  
  /**
   * Computed ready state.
   * True when photo is loaded successfully (photoUrl is available).
   * False while loading, on error, or if photo doesn't exist.
   * 
   * Useful for conditional rendering:
   * ```tsx
   * if (!isReady) return <Spinner />;
   * return <img src={photoUrl} />;
   * ```
   */
  readonly isReady: boolean;
}

interface GraphPhotoErrorShape {
  readonly statusCode?: number;
  readonly status?: number;
  readonly message?: string;
  readonly body?: {
    readonly error?: {
      readonly message?: string;
    };
  };
}

function getPhotoErrorStatus(err: unknown): number | undefined {
  const graphError = (typeof err === 'object' && err !== null)
    ? err as GraphPhotoErrorShape
    : {};

  return graphError.statusCode ?? graphError.status;
}

function getPhotoErrorMessage(err: unknown, error: Error): string {
  const graphError = (typeof err === 'object' && err !== null)
    ? err as GraphPhotoErrorShape
    : {};

  return graphError.message ?? graphError.body?.error?.message ?? error.message;
}

function createEnhancedPhotoError(err: unknown): Error {
  const error = err instanceof Error ? err : new Error(String(err));
  const status = getPhotoErrorStatus(err);
  const message = getPhotoErrorMessage(err, error);

  if (status === 404 || message.indexOf('404') !== -1) {
    error.message = 'Photo not found. User may not have a profile photo.';
  } else if (status === 403 || message.indexOf('403') !== -1) {
    error.message = 'Insufficient permissions to access photo. Check Graph API permissions.';
  } else if (status === 401 || message.indexOf('401') !== -1) {
    error.message = 'Authentication failed. User may not be signed in.';
  }

  return error;
}

/**
 * Loads a Graph profile photo and manages its Blob/object URL lifecycle. Defaults to /me, 240x240 and automatic fetching. Identifier/size changes clear prior state; latest requests publish results, and obsolete/unmounted completions cannot publish them. URLs are revoked on replacement, target change, failure and unmount. Failures are captured in error and reload resolves; Graph permissions are enforced by the server.
 * @param options - User ID (preferred over email), email/UPN, size and auto-fetch settings.
 * @returns Photo data, loading/error/readiness state and a reload method.
 * @throws Missing SPFx provider or synchronous Graph factory-consumption failures.
 *
 * @example
 * ```tsx
 * function ProfilePhoto() {
 *   const photo = useSPFxUserPhoto({ email: 'user@contoso.com', size: '96x96' });
 *   return photo.photoUrl ? <img src={photo.photoUrl} alt="Profile" /> : <span>{photo.error?.message}</span>;
 * }
 * ```
 */
export function useSPFxUserPhoto(
  options?: SPFxUserPhotoOptions
): SPFxUserPhotoResult {
  const { client: graphClient } = useSPFxMSGraphClient();
  const userPhotoService = useMemo(() => (
    graphClient ? createSPFxUserPhotoService(graphClient) : undefined
  ), [graphClient]);
  
  // Destructure options with defaults
  const {
    userId,
    email,
    size = '240x240',
    autoFetch = true
  } = options || {};
  
  // State management
  const [photoUrl, setPhotoUrl] = useState<string | undefined>(undefined);
  const [photoBlob, setPhotoBlob] = useState<Blob | undefined>(undefined);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  
  // Track component mounted state and current blob URL for cleanup
  const isMounted = useRef<boolean>(true);
  const currentBlobUrl = useRef<string | undefined>(undefined);
  const requestIdRef = useRef(0);
  const identityRef = useRef({ userPhotoService, userId, email, size });
  if (identityRef.current.userPhotoService !== userPhotoService ||
      identityRef.current.userId !== userId || identityRef.current.email !== email ||
      identityRef.current.size !== size) {
    identityRef.current = { userPhotoService, userId, email, size };
    requestIdRef.current += 1;
  }
  const identity = identityRef.current;
  
  useEffect(() => {
    isMounted.current = true;
    return () => {
      isMounted.current = false;
      requestIdRef.current += 1;
      // Cleanup: revoke blob URL to prevent memory leaks
      if (currentBlobUrl.current) {
        URL.revokeObjectURL(currentBlobUrl.current);
        currentBlobUrl.current = undefined;
      }
    };
  }, []);
  
  useEffect(() => {
    if (currentBlobUrl.current) {
      URL.revokeObjectURL(currentBlobUrl.current);
      currentBlobUrl.current = undefined;
    }
    setPhotoBlob(undefined);
    setPhotoUrl(undefined);
    setError(undefined);
    setIsLoading(false);
  }, [identity]);

  /**
   * Load photo from Microsoft Graph
   */
  const load = useCallback(async (): Promise<void> => {
    if (!isMounted.current || identity !== identityRef.current) {
      return;
    }
    const requestId = ++requestIdRef.current;
    const isCurrentRequest = (): boolean => isMounted.current &&
      identity === identityRef.current && requestId === requestIdRef.current;
    const service = userPhotoService;

    if (!graphClient || !service) {
      const err = new Error('MSGraphClient not available. Cannot load photo.');
      console.error('[useSPFxUserPhoto]', err);
      
      if (isCurrentRequest()) {
        setError(err);
      }
      return;
    }
    
    setIsLoading(true);
    setError(undefined);
    
    try {
      // Fetch photo blob from Graph API
      const blob = await service.getPhotoBlob({ userId, email, size });
      
      if (isCurrentRequest()) {
        // Revoke previous blob URL if exists
        if (currentBlobUrl.current) {
          URL.revokeObjectURL(currentBlobUrl.current);
        }
        
        // Create new blob URL
        const blobUrl = URL.createObjectURL(blob);
        currentBlobUrl.current = blobUrl;
        
        setPhotoBlob(blob);
        setPhotoUrl(blobUrl);
      }
    } catch (err) {
      if (isCurrentRequest()) {
        const error = createEnhancedPhotoError(err);

        if (currentBlobUrl.current) {
          URL.revokeObjectURL(currentBlobUrl.current);
          currentBlobUrl.current = undefined;
        }
        
        setError(error);
        setPhotoUrl(undefined);
        setPhotoBlob(undefined);
        console.error('[useSPFxUserPhoto] Failed to load photo:', error);
      }
    } finally {
      if (isCurrentRequest()) {
        setIsLoading(false);
      }
    }
  }, [graphClient, userPhotoService, userId, email, size, identity]);
  
  // Auto-fetch on mount if enabled
  useEffect(() => {
    if (autoFetch && userPhotoService) {
      load().catch(() => {
        // Error already handled in load() function
      });
    }
  }, [autoFetch, userPhotoService, load]);
  
  // Computed state: ready when photo loaded successfully
  const isReady = !isLoading && !error && photoUrl !== undefined;
  
  return useMemo(() => ({
    photoUrl,
    photoBlob,
    isLoading,
    error,
    reload: load,
    isReady
  }), [photoUrl, photoBlob, isLoading, error, load, isReady]);
}

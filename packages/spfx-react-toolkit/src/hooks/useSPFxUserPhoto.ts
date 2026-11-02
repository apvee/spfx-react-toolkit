/**
 * Graph profile photos use managed Blob URLs with explicit cleanup.
 *
 * @packageDocumentation
 * @docPage hooks/user-site
 */
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
 * import * as React from 'react';
 * import { SPFxUserPhotoSize } from '@apvee/spfx-react-toolkit';
 *
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
 * import * as React from 'react';
 * import { SPFxUserPhotoOptions, useSPFxUserPhoto } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const options: SPFxUserPhotoOptions = { email: 'user@contoso.com', size: '96x96', autoFetch: false };
 *   const photo = useSPFxUserPhoto(options);
 *   return photo.photoUrl ? <img src={photo.photoUrl} alt="Profile" /> : <span>Photo unavailable</span>;
 * }
 * ```
 */
export interface SPFxUserPhotoOptions {
  /**
   * Graph user/object ID; omitted, undefined or empty falls back to email, then /me. Nonempty values take precedence over email.
   */
  userId?: string;
  
  /**
   * User email/UPN; used when userId is absent/empty, otherwise ignored. Omitted, undefined or empty targets /me when no ID is provided.
   */
  email?: string;
  
  /**
   * Photo size to retrieve.
   * Omitted or undefined uses '240x240'.
   * @defaultValue '240x240'
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
   * Omitted or undefined uses true.
   * @defaultValue true
   */
  autoFetch?: boolean;
}

/**
 * Graph photo Blob/object URL with owned cleanup, reload control and current-request state.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxUserPhotoResult, useSPFxUserPhoto } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const photo: SPFxUserPhotoResult = useSPFxUserPhoto();
 *   return photo.photoUrl ? <img src={photo.photoUrl} alt="Profile" /> : <span>No photo</span>;
 * }
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
   * import * as React from 'react';
   *
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
 * @param options - User ID, email/UPN, size and auto-fetch settings; omitted or undefined uses /me, 240x240 and auto-fetch true.
 * @returns Photo data, loading/error/readiness state and a reload method.
 * @throws Missing SPFx provider or synchronous Graph factory-consumption failures.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxUserPhoto } from '@apvee/spfx-react-toolkit';
 *
 * function ProfilePhoto() {
 *   const photo = useSPFxUserPhoto({ email: 'user@contoso.com', size: '96x96' });
 *   return photo.photoUrl ? <img src={photo.photoUrl} alt="Profile" /> : <span>{photo.error?.message}</span>;
 * }
 * ```
 *
 * @useWhen an avatar must load a Graph photo with hook-owned object URL cleanup.
 *
 * @remarks
 * Mount below the matching SPFx provider with Fluent peers. Start gates Graph initialization; Run alone requests /me at the default 240x240 and displays it at 96px. Reset preserves the operational child and pending request: a late current success may create a hidden hook-owned URL. Only normal replacement, target change, failure or unmount revokes it. Reset does not clear hook state or cancel I/O.
 *
 * @example Read the current user photo with managed URL cleanup
 * ```tsx
 * import * as React from 'react';
 * import { Avatar, Button, Image, Label, MessageBar, MessageBarBody, makeStyles } from '@fluentui/react-components';
 * import { useSPFxMSGraphClient, useSPFxUserInfo, useSPFxUserPhoto } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type DemoProps = { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number };
 * const useStyles = makeStyles({ profile: { display: 'flex', alignItems: 'center', gap: '12px' } });
 * function ActivePhoto({ render, resetRevision = 0 }: DemoProps): React.ReactElement {
 *   const classes = useStyles();
 *   const user = useSPFxUserInfo();
 *   const graph = useSPFxMSGraphClient();
 *   const photo = useSPFxUserPhoto({ autoFetch: false });
 *   const live = React.useRef({ graph, photo });
 *   live.current = { graph, photo };
 *   const identity = React.useRef({ client: graph.client, reload: photo.reload, revision: resetRevision });
 *   if (identity.current.client !== graph.client || identity.current.reload !== photo.reload || identity.current.revision !== resetRevision) {
 *     identity.current = { client: graph.client, reload: photo.reload, revision: resetRevision };
 *   }
 *   const boundary = identity.current;
 *   const mounted = React.useRef(true);
 *   React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
 *   const sequence = React.useRef(0);
 *   const pending = React.useRef<{ reload: typeof photo.reload; request: number }>();
 *   const [executed, setExecuted] = React.useState<{ boundary: typeof boundary; request: number; completed: boolean }>();
 *   const run = React.useCallback(async (): Promise<void> => {
 *     if (!mounted.current || identity.current !== boundary || !live.current.graph.isReady || !live.current.graph.client || live.current.photo.isLoading || pending.current?.reload === boundary.reload) return;
 *     const request = ++sequence.current;
 *     pending.current = { reload: boundary.reload, request };
 *     setExecuted({ boundary, request, completed: false });
 *     try {
 *       await boundary.reload();
 *       if (mounted.current && identity.current === boundary && sequence.current === request) setExecuted({ boundary, request, completed: true });
 *     } finally {
 *       if (pending.current?.request === request) pending.current = undefined;
 *     }
 *   }, [boundary]);
 *   const current = executed?.boundary === boundary && executed.request === sequence.current && executed.completed;
 *   const error = current ? photo.error : undefined;
 *   const regions: DemoRegions = {
 *     commands: <><Button appearance="primary" disabled={!graph.isReady || !graph.client || photo.isLoading} onClick={() => { run().catch(() => undefined); }}>Run</Button></>,
 *     status: graph.initError ? <MessageBar intent="error"><MessageBarBody>{graph.initError.message}</MessageBarBody></MessageBar> : error ? <MessageBar intent="error"><MessageBarBody>{error.message}</MessageBarBody></MessageBar> : undefined,
 *     children: <>{!graph.isReady ? <Label>Microsoft Graph unavailable. Wait for client initialization.</Label> : photo.isLoading ? <Label>Loading...</Label> : <div className={classes.profile}>{current && photo.photoUrl ? <Image fit="cover" shape="circular" src={photo.photoUrl} alt={user.displayName} width={96} height={96} /> : <Avatar name={user.displayName} size={96} />}<Label>{user.displayName}</Label></div>}</>,
 *     diagnostics: <><dl><dt>Ready</dt><dd>{current ? String(photo.isReady) : 'Not requested'}</dd><dt>Blob</dt><dd>{current && photo.photoBlob ? `${photo.photoBlob.size} bytes` : 'Not requested'}</dd></dl><p>Reset hides local output and invalidates publication. It preserves the active hook, pending work and hook-owned URL; replacement, failure, target change and unmount release that URL.</p></>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export function Demo({ render, resetRevision = 0 }: DemoProps = {}): React.ReactElement {
 *   const [started, setStarted] = React.useState(false);
 *   if (started) return <ActivePhoto render={render} resetRevision={resetRevision} />;
 *   const regions: DemoRegions = { commands: <><Button appearance="primary" onClick={() => setStarted(true)}>Start</Button></>, children: <p>Start initializes Microsoft Graph under the matching SPFx provider. Run then reads the current user's photo.</p> };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Start initializes Graph without a photo request; Run reads the current user at default 240x240, displays a 96px photo or named fallback, loading, error and blob bytes. Reset hides output while preserving the hook, pending work and owned URL until normal cleanup.' };
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

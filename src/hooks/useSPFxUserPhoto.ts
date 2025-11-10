// useSPFxUserPhoto.ts
// Hook to load user photos from Microsoft Graph API

import { useState, useCallback, useEffect, useRef } from 'react';
import { useSPFxMSGraphClient } from './useSPFxMSGraphClient';

/**
 * Available photo sizes from Microsoft Graph
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
 */
export interface SPFxUserPhotoOptions {
  /**
   * User ID (Graph ID or Azure AD Object ID).
   * Example: 'abc-123-def-456'
   */
  userId?: string;
  
  /**
   * User email/UPN (User Principal Name).
   * Example: 'user@contoso.com'
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
 */
export interface SPFxUserPhotoResult {
  /**
   * Photo URL ready for use in <img src={photoUrl} />
   * Generated using URL.createObjectURL() from the blob.
   * Undefined if not loaded yet, on error, or if photo doesn't exist.
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
   * Reload the photo.
   * Useful for refresh buttons or retry logic.
   * 
   * @returns Promise that resolves when reload completes
   * 
   * @example
   * ```tsx
   * const { photoUrl, reload, isLoading } = useSPFxUserPhoto();
   * 
   * <button onClick={reload} disabled={isLoading}>
   *   Refresh Photo
   * </button>
   * ```
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

/**
 * Hook to load user photos from Microsoft Graph API
 * 
 * Provides easy access to user profile photos with automatic blob URL management.
 * Supports loading current user's photo or any user's photo by ID or email.
 * 
 * Features:
 * - Current user or specific user by ID/email
 * - Multiple photo sizes (48x48 to 648x648)
 * - Automatic blob URL creation and cleanup
 * - Memory leak prevention (revokes URLs on unmount)
 * - Type-safe with TypeScript
 * - Auto-fetch on mount (configurable)
 * - Manual reload function
 * 
 * Microsoft Graph Permissions Required:
 * - **User.Read**: Required for current user's photo (/me/photo)
 * - **User.ReadBasic.All**: Required for other users' photos (/users/{id}/photo)
 * - **User.Read.All**: Alternative permission for other users (more privileged)
 * 
 * Permission Notes:
 * - Application must have appropriate Graph API permissions configured in Azure AD
 * - Permissions must be consented by admin or user (depending on permission type)
 * - 404 errors typically mean the user has no photo (not a permission issue)
 * - 403 errors indicate insufficient permissions
 * 
 * Graph API Endpoints Used:
 * - Current user: GET /me/photo/{size}/$value
 * - By ID: GET /users/{id}/photo/{size}/$value
 * - By email: GET /users/{email}/photo/{size}/$value
 * 
 * @param options - Optional. Configuration for loading specific user's photo
 * 
 * @example Current user photo
 * ```tsx
 * function UserAvatar() {
 *   const { photoUrl, isLoading, error } = useSPFxUserPhoto();
 *   
 *   if (isLoading) return <Spinner />;
 *   if (error) return <DefaultAvatar />;
 *   
 *   return <img src={photoUrl} alt="User" style={{ width: 240, height: 240 }} />;
 * }
 * ```
 * 
 * @example Specific user by email
 * ```tsx
 * function TeamMemberAvatar({ email }: { email: string }) {
 *   const { photoUrl, isLoading } = useSPFxUserPhoto({ 
 *     email,
 *     size: '96x96'
 *   });
 *   
 *   return (
 *     <div>
 *       {isLoading ? (
 *         <Spinner />
 *       ) : (
 *         <img src={photoUrl || '/default-avatar.png'} alt={email} />
 *       )}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Specific user by ID with reload
 * ```tsx
 * function ProfileCard({ userId }: { userId: string }) {
 *   const { photoUrl, reload, isLoading, error } = useSPFxUserPhoto({ 
 *     userId,
 *     size: '360x360'
 *   });
 *   
 *   return (
 *     <Stack>
 *       {error ? (
 *         <MessageBar messageBarType={MessageBarType.error}>
 *           Failed to load photo: {error.message}
 *         </MessageBar>
 *       ) : (
 *         <img src={photoUrl} alt="Profile" />
 *       )}
 *       
 *       <PrimaryButton 
 *         onClick={reload} 
 *         disabled={isLoading}
 *         text="Refresh Photo"
 *       />
 *     </Stack>
 *   );
 * }
 * ```
 * 
 * @example Multiple sizes for responsive images
 * ```tsx
 * function ResponsiveAvatar() {
 *   const small = useSPFxUserPhoto({ size: '96x96' });
 *   const large = useSPFxUserPhoto({ size: '240x240' });
 *   
 *   return (
 *     <picture>
 *       <source media="(min-width: 768px)" srcSet={large.photoUrl} />
 *       <img src={small.photoUrl} alt="User Avatar" />
 *     </picture>
 *   );
 * }
 * ```
 * 
 * @example Lazy loading with manual fetch
 * ```tsx
 * function LazyAvatar({ email }: { email: string }) {
 *   const { photoUrl, reload, isLoading } = useSPFxUserPhoto({ 
 *     email,
 *     autoFetch: false  // Don't load on mount
 *   });
 *   
 *   return (
 *     <div>
 *       {photoUrl ? (
 *         <img src={photoUrl} alt="Avatar" />
 *       ) : (
 *         <button onClick={reload} disabled={isLoading}>
 *           {isLoading ? 'Loading...' : 'Load Photo'}
 *         </button>
 *       )}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example With error fallback
 * ```tsx
 * function SafeAvatar() {
 *   const { photoUrl, error, isReady } = useSPFxUserPhoto();
 *   
 *   if (!isReady) {
 *     return <Persona text="Loading..." size={PersonaSize.size72} />;
 *   }
 *   
 *   if (error || !photoUrl) {
 *     // Fallback to Fluent UI Persona with initials
 *     return <Persona text="John Doe" size={PersonaSize.size72} />;
 *   }
 *   
 *   return <img src={photoUrl} alt="User" className="avatar-round" />;
 * }
 * ```
 * 
 * @example Access raw blob for processing
 * ```tsx
 * function PhotoUploader() {
 *   const { photoBlob, photoUrl } = useSPFxUserPhoto();
 *   
 *   const handleUploadToAzure = async () => {
 *     if (!photoBlob) return;
 *     
 *     const formData = new FormData();
 *     formData.append('photo', photoBlob, 'profile.jpg');
 *     
 *     await fetch('/api/upload', {
 *       method: 'POST',
 *       body: formData
 *     });
 *   };
 *   
 *   return (
 *     <div>
 *       <img src={photoUrl} alt="Preview" />
 *       <button onClick={handleUploadToAzure}>Upload to Azure</button>
 *     </div>
 *   );
 * }
 * ```
 */
export function useSPFxUserPhoto(
  options?: SPFxUserPhotoOptions
): SPFxUserPhotoResult {
  const { client: graphClient } = useSPFxMSGraphClient();
  
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
  
  useEffect(() => {
    isMounted.current = true;
    return () => {
      isMounted.current = false;
      // Cleanup: revoke blob URL to prevent memory leaks
      if (currentBlobUrl.current) {
        URL.revokeObjectURL(currentBlobUrl.current);
        currentBlobUrl.current = undefined;
      }
    };
  }, []);
  
  /**
   * Build Graph API endpoint based on user identifier
   */
  const buildPhotoEndpoint = useCallback((): string => {
    // Determine base path
    let basePath: string;
    
    if (userId) {
      // Specific user by ID
      basePath = `/users/${userId}`;
    } else if (email) {
      // Specific user by email
      basePath = `/users/${email}`;
    } else {
      // Current user
      basePath = '/me';
    }
    
    // Append photo size endpoint
    return `${basePath}/photos/${size}/$value`;
  }, [userId, email, size]);
  
  /**
   * Load photo from Microsoft Graph
   */
  const load = useCallback(async (): Promise<void> => {
    if (!graphClient) {
      const err = new Error('MSGraphClient not available. Cannot load photo.');
      console.error('[useSPFxUserPhoto]', err);
      
      if (isMounted.current) {
        setError(err);
      }
      return;
    }
    
    setIsLoading(true);
    setError(undefined);
    
    try {
      const endpoint = buildPhotoEndpoint();
      
      // Fetch photo blob from Graph API
      const blob: Blob = await graphClient
        .api(endpoint)
        .get();
      
      if (isMounted.current) {
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
      if (isMounted.current) {
        const error = err instanceof Error ? err : new Error(String(err));
        
        // Enhanced error messages (ES5-compatible)
        if (error.message.indexOf('404') !== -1) {
          error.message = 'Photo not found. User may not have a profile photo.';
        } else if (error.message.indexOf('403') !== -1) {
          error.message = 'Insufficient permissions to access photo. Check Graph API permissions.';
        } else if (error.message.indexOf('401') !== -1) {
          error.message = 'Authentication failed. User may not be signed in.';
        }
        
        setError(error);
        setPhotoUrl(undefined);
        setPhotoBlob(undefined);
        console.error('[useSPFxUserPhoto] Failed to load photo:', error);
      }
    } finally {
      if (isMounted.current) {
        setIsLoading(false);
      }
    }
  }, [graphClient, buildPhotoEndpoint]);
  
  // Auto-fetch on mount if enabled
  useEffect(() => {
    if (autoFetch && graphClient) {
      load().catch(() => {
        // Error already handled in load() function
      });
    }
  }, [autoFetch, graphClient, load]);
  
  // Computed state: ready when photo loaded successfully
  const isReady = !isLoading && !error && photoUrl !== undefined;
  
  return {
    photoUrl,
    photoBlob,
    isLoading,
    error,
    reload: load,
    isReady
  };
}

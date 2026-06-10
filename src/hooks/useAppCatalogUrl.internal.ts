// useAppCatalogUrl.internal.ts
// Internal hook to discover and cache the tenant app catalog URL

import { useState, useCallback, useMemo, useRef, useEffect } from 'react';
import { useSPFxSPHttpClient } from './useSPFxSPHttpClient';
import { useSPFxPageContext } from './useSPFxPageContext';
import type { SPHttpClient } from '@microsoft/sp-http';
import { createSPFxAppCatalogService } from '../services/spfx-app-catalog.service';

/**
 * Return type for useAppCatalogUrl internal hook
 */
export interface AppCatalogUrlInfo {
  /** Resolved tenant app catalog URL, or undefined if not yet discovered */
  readonly appCatalogUrl: string | undefined;
  /** SPHttpClient instance from context */
  readonly spHttpClient: SPHttpClient | undefined;
  /** Discover and return the tenant app catalog URL (cached after first call) */
  readonly discoverAppCatalogUrl: () => Promise<string>;
  /** Check if current user is Site Collection Admin on the app catalog */
  readonly checkWritePermission: (catalogUrl: string) => Promise<boolean>;
  /** Ref tracking component mounted state */
  readonly isMountedRef: React.MutableRefObject<boolean>;
}

/**
 * Internal hook to discover and cache the tenant app catalog URL.
 * Shared by useSPFxTenantProperty and useSPFxTenantKeyValueStore.
 *
 * @returns Object with app catalog URL, discovery function, and permission check
 * @internal
 */
export function useAppCatalogUrl(): AppCatalogUrlInfo {
  const { client: spHttpClient } = useSPFxSPHttpClient();
  const pageContext = useSPFxPageContext();
  const appCatalogService = useMemo(
    () => createSPFxAppCatalogService(spHttpClient, pageContext),
    [spHttpClient, pageContext]
  );

  const [appCatalogUrl, setAppCatalogUrl] = useState<string | undefined>(undefined);
  const appCatalogUrlRef = useRef<string | undefined>(undefined);

  const isMountedRef = useRef<boolean>(true);

  useEffect(() => {
    isMountedRef.current = true;
    return () => {
      isMountedRef.current = false;
    };
  }, []);

  const discoverAppCatalogUrl = useCallback(async (): Promise<string> => {
    if (!spHttpClient || !pageContext) {
      throw new Error('SPHttpClient or PageContext not available');
    }

    if (appCatalogUrlRef.current) {
      return appCatalogUrlRef.current;
    }

    try {
      const discoveredCatalogUrl = await appCatalogService.discoverUrl();

      // eslint-disable-next-line require-atomic-updates -- Idempotent: always sets the same discovered URL
      appCatalogUrlRef.current = discoveredCatalogUrl;
      if (isMountedRef.current) {
        setAppCatalogUrl(discoveredCatalogUrl);
      }

      return discoveredCatalogUrl;
    } catch (err) {
      throw new Error(`App catalog discovery failed: ${err instanceof Error ? err.message : String(err)}`);
    }
  }, [spHttpClient, pageContext, appCatalogService]);

  const checkWritePermission = useCallback(async (catalogUrl: string): Promise<boolean> => {
    return appCatalogService.canCurrentUserWrite(catalogUrl);
  }, [appCatalogService]);

  return useMemo(() => ({
    appCatalogUrl,
    spHttpClient,
    discoverAppCatalogUrl,
    checkWritePermission,
    isMountedRef,
  }), [appCatalogUrl, spHttpClient, discoverAppCatalogUrl, checkWritePermission]);
}

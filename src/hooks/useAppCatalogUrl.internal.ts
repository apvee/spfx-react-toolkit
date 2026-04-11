// useAppCatalogUrl.internal.ts
// Internal hook to discover and cache the tenant app catalog URL

import { useState, useCallback, useRef, useEffect } from 'react';
import { useSPFxSPHttpClient } from './useSPFxSPHttpClient';
import { useSPFxPageContext } from './useSPFxPageContext';
import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * Tenant app catalog URL response
 */
interface ITenantAppCatalogResponse {
  CorporateCatalogUrl: string;
}

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

  const [appCatalogUrl, setAppCatalogUrl] = useState<string | undefined>(undefined);

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

    if (appCatalogUrl) {
      return appCatalogUrl;
    }

    try {
      const response: SPHttpClientResponse = await spHttpClient.get(
        `${pageContext.web.absoluteUrl}/_api/SP_TenantSettings_Current`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to discover app catalog: ${response.statusText}`);
      }

      const data: ITenantAppCatalogResponse = await response.json();

      if (!data.CorporateCatalogUrl) {
        throw new Error('Tenant app catalog is not provisioned. Please provision the app catalog first.');
      }

      if (isMountedRef.current) {
        setAppCatalogUrl(data.CorporateCatalogUrl);
      }

      return data.CorporateCatalogUrl;
    } catch (err) {
      throw new Error(`App catalog discovery failed: ${err instanceof Error ? err.message : String(err)}`);
    }
  }, [spHttpClient, pageContext, appCatalogUrl]);

  const checkWritePermission = useCallback(async (catalogUrl: string): Promise<boolean> => {
    if (!spHttpClient) return false;

    try {
      const response: SPHttpClientResponse = await spHttpClient.get(
        `${catalogUrl}/_api/web/currentuser?$select=IsSiteAdmin`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) return false;

      const user = await response.json();
      return user.IsSiteAdmin === true;
    } catch {
      return false;
    }
  }, [spHttpClient]);

  return {
    appCatalogUrl,
    spHttpClient,
    discoverAppCatalogUrl,
    checkWritePermission,
    isMountedRef,
  };
}

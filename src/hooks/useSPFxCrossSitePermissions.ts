import { useState, useEffect, useCallback } from 'react';
import { SPPermission } from '@microsoft/sp-page-context';
import { SPHttpClient } from '@microsoft/sp-http';
import { useSPFxSPHttpClient } from './useSPFxSPHttpClient';

/**
 * Options for cross-site permissions retrieval
 */
export interface SPFxCrossSitePermissionsOptions {
  /**
   * Optional web URL within the site (e.g., '/sites/mysite/subweb')
   */
  webUrl?: string;

  /**
   * Optional list ID to retrieve list-level permissions
   */
  listId?: string;
}

/**
 * Information about cross-site permissions
 */
export interface SPFxCrossSitePermissionsInfo {
  /**
   * Site-level permissions
   */
  sitePermissions?: SPPermission;

  /**
   * Web-level permissions
   */
  webPermissions?: SPPermission;

  /**
   * List-level permissions (if listId provided)
   */
  listPermissions?: SPPermission;

  /**
   * Check if user has specific web permission
   */
  hasWebPermission: (permission: SPPermission) => boolean;

  /**
   * Check if user has specific site permission
   */
  hasSitePermission: (permission: SPPermission) => boolean;

  /**
   * Check if user has specific list permission
   */
  hasListPermission: (permission: SPPermission) => boolean;

  /**
   * Loading state
   */
  isLoading: boolean;

  /**
   * Error state
   */
  error?: Error;
}

/**
 * Hook to retrieve permissions for a different site/web/list
 * 
 * @param siteUrl - The target site URL (optional - no fetch if undefined/empty)
 * @param options - Optional configuration (webUrl, listId)
 * @returns Permissions information with helper methods
 * 
 * @example
 * ```tsx
 * // Lazy loading - no fetch until URL is set
 * const [targetUrl, setTargetUrl] = useState<string | undefined>(undefined);
 * const { 
 *   sitePermissions, 
 *   webPermissions, 
 *   hasWebPermission,
 *   isLoading 
 * } = useSPFxCrossSitePermissions(targetUrl, {
 *   webUrl: 'https://contoso.sharepoint.com/sites/target/subweb'
 * });
 * 
 * // Set URL when ready - triggers fetch
 * setTargetUrl('https://contoso.sharepoint.com/sites/target');
 * 
 * if (!isLoading && hasWebPermission(SPPermission.addListItems)) {
 *   // User can add items
 * }
 * ```
 */
export function useSPFxCrossSitePermissions(
  siteUrl?: string,
  options?: SPFxCrossSitePermissionsOptions
): SPFxCrossSitePermissionsInfo {
  const { invoke } = useSPFxSPHttpClient();

  const [sitePermissions, setSitePermissions] = useState<SPPermission | undefined>();
  const [webPermissions, setWebPermissions] = useState<SPPermission | undefined>();
  const [listPermissions, setListPermissions] = useState<SPPermission | undefined>();
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<Error | undefined>();

  // Helper: Convert EffectiveBasePermissions to SPPermission
  const convertToSPPermission = useCallback((data: { High: number; Low: number }): SPPermission => {
    return new SPPermission(data);
  }, []);

  // Fetch permissions when inputs change
  useEffect(() => {
    // Skip fetch if siteUrl is not provided or empty
    if (!siteUrl || siteUrl.trim() === '') {
      // Reset to idle state (no loading, no error)
      setIsLoading(false);
      setError(undefined);
      setSitePermissions(undefined);
      setWebPermissions(undefined);
      setListPermissions(undefined);
      return;
    }

    setIsLoading(true);
    setError(undefined);

    const targetWebUrl = options?.webUrl || siteUrl;

    // Fetch site permissions
    const fetchSitePermissions = invoke(client =>
      client.get(
        `${siteUrl}/_api/site/effectivebasepermissions`,
        SPHttpClient.configurations.v1
      )
      .then(res => res.json())
      .then((data: { EffectiveBasePermissions?: { High: number; Low: number } }) => {
        if (data.EffectiveBasePermissions) {
          return convertToSPPermission(data.EffectiveBasePermissions);
        }
        return undefined;
      })
    );

    // Fetch web permissions
    const fetchWebPermissions = invoke(client =>
      client.get(
        `${targetWebUrl}/_api/web/effectivebasepermissions`,
        SPHttpClient.configurations.v1
      )
      .then(res => res.json())
      .then((data: { EffectiveBasePermissions?: { High: number; Low: number } }) => {
        if (data.EffectiveBasePermissions) {
          return convertToSPPermission(data.EffectiveBasePermissions);
        }
        return undefined;
      })
    );

    // Fetch list permissions if listId provided
    const fetchListPermissions = options?.listId
      ? invoke(client =>
          client.get(
            `${targetWebUrl}/_api/web/lists(guid'${options.listId}')/effectivebasepermissions`,
            SPHttpClient.configurations.v1
          )
          .then(res => res.json())
          .then((data: { EffectiveBasePermissions?: { High: number; Low: number } }) => {
            if (data.EffectiveBasePermissions) {
              return convertToSPPermission(data.EffectiveBasePermissions);
            }
            return undefined;
          })
        )
      : Promise.resolve(undefined);

    // Execute all fetches in parallel
    Promise.all([fetchSitePermissions, fetchWebPermissions, fetchListPermissions])
      .then(([site, web, list]) => {
        setSitePermissions(site);
        setWebPermissions(web);
        setListPermissions(list);
        setIsLoading(false);
      })
      .catch(err => {
        setError(err instanceof Error ? err : new Error(String(err)));
        setIsLoading(false);
      });
  }, [siteUrl, options?.webUrl, options?.listId, invoke, convertToSPPermission]);

  // Helper methods
  const hasWebPermission = useCallback(
    (permission: SPPermission): boolean => {
      return webPermissions?.hasPermission(permission) ?? false;
    },
    [webPermissions]
  );

  const hasSitePermission = useCallback(
    (permission: SPPermission): boolean => {
      return sitePermissions?.hasPermission(permission) ?? false;
    },
    [sitePermissions]
  );

  const hasListPermission = useCallback(
    (permission: SPPermission): boolean => {
      return listPermissions?.hasPermission(permission) ?? false;
    },
    [listPermissions]
  );

  return {
    sitePermissions,
    webPermissions,
    listPermissions,
    hasWebPermission,
    hasSitePermission,
    hasListPermission,
    isLoading,
    error,
  };
}

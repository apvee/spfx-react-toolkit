/**
 * Discover a tenant app catalog lazily and inspect its current-user site-administrator flag.
 *
 * @packageDocumentation
 * @docPage services/app-catalog
 */
import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * Current web URL required by the corresponding service factory.
 *
 * @example
 * ```ts
 * import { SPFxAppCatalogPageContextLike } from '@apvee/spfx-react-toolkit';
 *
 * const page: SPFxAppCatalogPageContextLike = { web: { absoluteUrl: 'https://contoso.sharepoint.com/sites/team' } };
 * ```
 */
export interface SPFxAppCatalogPageContextLike {
  /**
   * Current web metadata used to resolve SharePoint endpoints.
   */
  web: {
    /**
     * Absolute HTTP(S) URL of the current web.
     */
    absoluteUrl: string;
  };
}

/**
 * Tenant app catalog discovery and site-administrator indicator. Discovery caches a successful URL within this service instance. Concurrent first discovery calls are not deduplicated; failures are not cached. No catalog is provisioned and tenant settings are never mutated.
 *
 * @example
 * ```ts
 * import type { PageContext } from '@microsoft/sp-page-context';
 * import type { SPHttpClient } from '@microsoft/sp-http';
 * import { SPFxAppCatalogService, createSPFxAppCatalogService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(pageContext: PageContext, client: SPHttpClient) {
 *   const catalog: SPFxAppCatalogService = createSPFxAppCatalogService(client, pageContext);
 *   const url = await catalog.discoverUrl();
 * }
 * ```
 */
export interface SPFxAppCatalogService {
  /**
   * Discovers the tenant app catalog.
   * @returns Its URL, cached after successful discovery.
   * @throws Missing dependencies, failed requests, or an unprovisioned catalog.
   */
  discoverUrl: () => Promise<string>;
  /**
   * Checks the catalog current user's IsSiteAdmin flag, not granular list permissions.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns False if the client is missing or lookup fails; otherwise the administrator flag.
   */
  canCurrentUserWrite: (catalogUrl: string) => Promise<boolean>;
}

interface TenantAppCatalogResponse {
  CorporateCatalogUrl: string;
}

/**
 * Creates an isolated catalog discovery cache; no request is made at construction.
 * @param spHttpClient - Caller-owned SharePoint client, possibly unavailable.
 * @param pageContext - Current web metadata, possibly unavailable.
 * @returns Explicit discovery and administrator-check operations.
 *
 * @example Basic usage
 * ```ts
 * import type { WebPartContext } from '@microsoft/sp-webpart-base';
 * import { createSPFxAppCatalogService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(context: WebPartContext) {
 *   const catalog = createSPFxAppCatalogService(context.spHttpClient, context.pageContext);
 *   const url = await catalog.discoverUrl();
 * }
 * ```
 * @useWhen You need lazy tenant app-catalog discovery or its site-administrator indicator outside React.
 *
 *
 * @example Discover the catalog and compare its admin indicator
 * ```ts
 * import type { SPHttpClient } from '@microsoft/sp-http';
 * import type { PageContext } from '@microsoft/sp-page-context';
 *
 * import { createSPFxAppCatalogService } from '@apvee/spfx-react-toolkit';
 * async function demonstrate(context: { spHttpClient: SPHttpClient; pageContext: PageContext }) {
 *   const catalog = createSPFxAppCatalogService(context.spHttpClient, context.pageContext);
 *   const url = await catalog.discoverUrl();
 *   const isCatalogAdmin = await catalog.canCurrentUserWrite(url);
 * }
 * ```
 */
export function createSPFxAppCatalogService(
  spHttpClient: SPHttpClient | undefined,
  pageContext: SPFxAppCatalogPageContextLike | undefined
): SPFxAppCatalogService {
  let cachedCatalogUrl: string | undefined;

  const discoverUrl = async (): Promise<string> => {
    if (!spHttpClient || !pageContext) {
      throw new Error('SPHttpClient or PageContext not available');
    }

    if (cachedCatalogUrl) {
      return cachedCatalogUrl;
    }

    const response: SPHttpClientResponse = await spHttpClient.get(
      `${pageContext.web.absoluteUrl}/_api/SP_TenantSettings_Current`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to discover app catalog: ${response.statusText}`);
    }

    const data: TenantAppCatalogResponse = await response.json();

    if (!data.CorporateCatalogUrl) {
      throw new Error('Tenant app catalog is not provisioned. Please provision the app catalog first.');
    }

    // eslint-disable-next-line require-atomic-updates -- Idempotent cache of the discovered tenant app catalog URL.
    cachedCatalogUrl = data.CorporateCatalogUrl;

    return data.CorporateCatalogUrl;
  };

  const canCurrentUserWrite = async (catalogUrl: string): Promise<boolean> => {
    if (!spHttpClient) {
      return false;
    }

    try {
      const response: SPHttpClientResponse = await spHttpClient.get(
        `${catalogUrl}/_api/web/currentuser?$select=IsSiteAdmin`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return false;
      }

      const user = await response.json() as { IsSiteAdmin?: boolean };

      return user.IsSiteAdmin === true;
    } catch {
      return false;
    }
  };

  return {
    discoverUrl,
    canCurrentUserWrite
  };
}

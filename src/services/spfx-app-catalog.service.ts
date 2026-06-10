import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';

export interface SPFxAppCatalogPageContextLike {
  web: {
    absoluteUrl: string;
  };
}

export interface SPFxAppCatalogService {
  discoverUrl: () => Promise<string>;
  canCurrentUserWrite: (catalogUrl: string) => Promise<boolean>;
}

interface TenantAppCatalogResponse {
  CorporateCatalogUrl: string;
}

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

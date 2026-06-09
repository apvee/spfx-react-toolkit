import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';

import {
  deserializeTenantValue,
  escapeODataValue
} from '../helpers/spfx-tenant-value.helpers';

export interface SPFxTenantProperty<T = unknown> {
  readonly value: T | undefined;
  readonly description: string | undefined;
}

export interface SPFxTenantPropertyService {
  get: <T = unknown>(key: string, catalogUrl: string) => Promise<SPFxTenantProperty<T>>;
}

interface StorageEntityResponse {
  Value?: string;
  Description?: string;
}

function encodeODataPathValue(value: string): string {
  return encodeURIComponent(escapeODataValue(value)).replace(/'/g, '%27');
}

export function createSPFxTenantPropertyService(
  spHttpClient: SPHttpClient
): SPFxTenantPropertyService {
  const get = async <T = unknown>(
    key: string,
    catalogUrl: string
  ): Promise<SPFxTenantProperty<T>> => {
    const safeKey = encodeODataPathValue(key);
    const response: SPHttpClientResponse = await spHttpClient.get(
      `${catalogUrl}/_api/web/GetStorageEntity('${safeKey}')`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to read property: ${response.statusText}`);
    }

    const entity: StorageEntityResponse = await response.json();

    if (entity.Value !== undefined && entity.Value !== null) {
      return {
        value: deserializeTenantValue<T>(entity.Value),
        description: entity.Description
      };
    }

    return {
      value: undefined,
      description: undefined
    };
  };

  return {
    get
  };
}

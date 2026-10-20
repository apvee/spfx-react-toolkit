import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';

import {
  deserializeTenantValue,
  escapeODataValue
} from '../helpers/spfx-tenant-value.helpers';

/**
 * Deserialized StorageEntity value and description. T is a compile-time expectation, not runtime validation; legacy serialization may coerce JSON-like strings.
 *
 * @example
 * ```ts
 * const property: SPFxTenantProperty<boolean> = await service.get('featureEnabled', catalogUrl);
 * console.log(property.value);
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxTenantProperty<T = unknown> {
  /**
   * Deserialized value, or undefined when SharePoint returns no Value.
   */
  readonly value: T | undefined;
  /**
   * Stored description, or undefined when no value/description is returned.
   */
  readonly description: string | undefined;
}

/**
 * Reads StorageEntity properties from an explicitly supplied tenant app catalog.
 *
 * @example
 * ```ts
 * const service: SPFxTenantPropertyService = createSPFxTenantPropertyService(client);
 * const setting = await service.get<boolean>('featureEnabled', catalogUrl);
 * ```
 */
export interface SPFxTenantPropertyService {
  /**
   * Reads and deserializes one StorageEntity.
   * @param key - Property key; escaped for the OData path.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Value and description, possibly undefined.
   * @throws Failed HTTP requests and response/serialization failures.
   */
  get: <T = unknown>(key: string, catalogUrl: string) => Promise<SPFxTenantProperty<T>>;
}

interface StorageEntityResponse {
  Value?: string;
  Description?: string;
}

function encodeODataPathValue(value: string): string {
  return encodeURIComponent(escapeODataValue(value)).replace(/'/g, '%27');
}

/**
 * Creates explicit StorageEntity reads without discovering the catalog.
 * @param spHttpClient - Caller-owned SharePoint client.
 * @returns A property-reading service; construction performs no request.
 *
 * @example
 * ```ts
 * const service = createSPFxTenantPropertyService(context.spHttpClient);
 * const setting = await service.get<string>('appVersion', catalogUrl);
 * ```
 */
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

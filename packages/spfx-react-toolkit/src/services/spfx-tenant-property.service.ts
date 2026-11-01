/**
 * Read legacy StorageEntity values from a caller-selected tenant app catalog.
 *
 * @packageDocumentation
 * @docPage services/tenant-property
 */
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
 * import { SPFxTenantProperty, SPFxTenantPropertyService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(catalogUrl: string, service: SPFxTenantPropertyService) {
 *   const property: SPFxTenantProperty<boolean> = await service.get('featureEnabled', catalogUrl);
 *   console.log(property.value);
 * }
 * ```
 * @template T - Expected application data shape; defaults to unknown and performs no runtime validation.
 */
export interface SPFxTenantProperty<T = unknown> {
  /**
   * Deserialized value, or undefined when SharePoint returns no Value or a legacy REST-null Value. A stored string is parsed JSON-first without schema validation or Date reconstruction.
   */
  readonly value: T | undefined;
  /**
   * Stored description passed through unchanged when Value exists; undefined when Value is missing/legacy REST null, even if a description was supplied.
   */
  readonly description: string | undefined;
}

/**
 * Reads StorageEntity properties from an explicitly supplied tenant app catalog. No discovery, cache, provisioning or write method is added.
 *
 * @example
 * ```ts
 * import type { SPHttpClient } from '@microsoft/sp-http';
 * import { SPFxTenantPropertyService, createSPFxTenantPropertyService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(client: SPHttpClient, catalogUrl: string) {
 *   const service: SPFxTenantPropertyService = createSPFxTenantPropertyService(client);
 *   const setting = await service.get<boolean>('featureEnabled', catalogUrl);
 * }
 * ```
 */
export interface SPFxTenantPropertyService {
  /**
   * Reads and deserializes one StorageEntity.
   * @typeParam T - Expected deserialized value; defaults to unknown and performs no runtime validation.
   * @param key - Required property key; OData-escaped and URL-encoded for the path.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Value and description, possibly undefined.
   * @throws Every unsuccessful HTTP response, including 404/403, and response parsing failures; HTTP failure is never an absent-property success.
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
 * @example Basic usage
 * ```ts
 * import type { WebPartContext } from '@microsoft/sp-webpart-base';
 * import { createSPFxTenantPropertyService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(context: WebPartContext, catalogUrl: string) {
 *   const service = createSPFxTenantPropertyService(context.spHttpClient);
 *   const setting = await service.get<string>('appVersion', catalogUrl);
 * }
 * ```
 * @useWhen You read existing StorageEntity values from an already known tenant catalog.
 *
 *
 * @example Read and validate an unknown StorageEntity
 * ```ts
 * import type { SPHttpClient } from '@microsoft/sp-http';
 *
 * import { createSPFxTenantPropertyService } from '@apvee/spfx-react-toolkit';
 * async function demonstrate(context: { spHttpClient: SPHttpClient }, catalogUrl: string) {
 *   const service = createSPFxTenantPropertyService(context.spHttpClient);
 *   const property = await service.get<unknown>('my-app:featureEnabled', catalogUrl);
 *   if (typeof property.value === 'boolean') console.log(property.value);
 * }
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

import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';

import {
  deserializeTenantValue,
  escapeODataValue,
  serializeTenantValue
} from '../helpers/spfx-tenant-value.helpers';

const LIST_TITLE = 'TenantKeyValueStore';

export interface SPFxTenantKeyValueStoreServiceItem<T = unknown> {
  readonly key: string;
  readonly value: T;
  readonly description: string | undefined;
  readonly id: number;
}

export interface SPFxTenantKeyValueStoreService {
  ensureListReady: (catalogUrl: string) => Promise<void>;
  get: <T = unknown>(key: string, catalogUrl: string) => Promise<SPFxTenantKeyValueStoreServiceItem<T> | undefined>;
  list: (catalogUrl: string) => Promise<SPFxTenantKeyValueStoreServiceItem<unknown>[]>;
  save: <T = unknown>(key: string, value: T, catalogUrl: string, description?: string) => Promise<void>;
  remove: (key: string, catalogUrl: string) => Promise<void>;
}

interface ListItemResponse {
  Id: number;
  Title: string;
  Value: string;
  Description?: string;
}

interface ListItemsResponse {
  value: ListItemResponse[];
}

interface FieldsCheckResponse {
  value: Array<{ InternalName: string }>;
}

function getListApiUrl(catalogUrl: string): string {
  return `${catalogUrl}/_api/web/lists/getByTitle('${LIST_TITLE}')`;
}

function encodeQueryValue(value: string): string {
  return encodeURIComponent(value).replace(/'/g, '%27');
}

async function createField(
  spHttpClient: SPHttpClient,
  listApiUrl: string,
  fieldTitle: string
): Promise<void> {
  const response: SPHttpClientResponse = await spHttpClient.post(
    `${listApiUrl}/fields`,
    SPHttpClient.configurations.v1,
    {
      body: JSON.stringify({
        FieldTypeKind: 3,
        Title: fieldTitle
      })
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to create ${fieldTitle} field: ${response.statusText}. ${errorText}`);
  }
}

function mapItem<T>(raw: ListItemResponse): SPFxTenantKeyValueStoreServiceItem<T> {
  return {
    key: raw.Title,
    value: deserializeTenantValue<T>(raw.Value),
    description: raw.Description || undefined,
    id: raw.Id
  };
}

export function createSPFxTenantKeyValueStoreService(
  spHttpClient: SPHttpClient
): SPFxTenantKeyValueStoreService {
  const listProvisioned = new Map<string, boolean>();
  const provisioningPromises = new Map<string, Promise<void>>();

  const ensureListReady = (catalogUrl: string): Promise<void> => {
    if (listProvisioned.get(catalogUrl)) {
      return Promise.resolve();
    }

    const provisioningPromise = provisioningPromises.get(catalogUrl);
    if (provisioningPromise) {
      return provisioningPromise;
    }

    const doProvision = async (): Promise<void> => {
      const listApiUrl = getListApiUrl(catalogUrl);

      const listResponse: SPHttpClientResponse = await spHttpClient.get(
        `${listApiUrl}?$select=Id`,
        SPHttpClient.configurations.v1
      );

      const listExists = listResponse.status !== 404;

      if (listExists && !listResponse.ok) {
        throw new Error(`Failed to check list existence: ${listResponse.statusText}`);
      }

      if (listExists) {
        const fieldsResponse: SPHttpClientResponse = await spHttpClient.get(
          `${listApiUrl}/fields?$filter=InternalName eq 'Value' or InternalName eq 'Description'&$select=InternalName`,
          SPHttpClient.configurations.v1
        );

        if (!fieldsResponse.ok) {
          throw new Error(`Failed to check list fields: ${fieldsResponse.statusText}`);
        }

        const fields: FieldsCheckResponse = await fieldsResponse.json();
        const existingFields = fields.value.map(function(field): string {
          return field.InternalName;
        });
        const hasValue = existingFields.includes('Value');
        const hasDescription = existingFields.includes('Description');

        if (hasValue && hasDescription) {
          return;
        }

        if (!hasValue) {
          await createField(spHttpClient, listApiUrl, 'Value');
        }

        if (!hasDescription) {
          await createField(spHttpClient, listApiUrl, 'Description');
        }

        return;
      }

      const createListResponse: SPHttpClientResponse = await spHttpClient.post(
        `${catalogUrl}/_api/web/lists`,
        SPHttpClient.configurations.v1,
        {
          body: JSON.stringify({
            BaseTemplate: 100,
            Title: LIST_TITLE,
            Hidden: true,
            NoCrawl: true
          })
        }
      );

      if (!createListResponse.ok) {
        const errorText = await createListResponse.text();
        throw new Error(`Failed to create list: ${createListResponse.statusText}. ${errorText}`);
      }

      const createdListApiUrl = getListApiUrl(catalogUrl);

      await createField(spHttpClient, createdListApiUrl, 'Value');
      await createField(spHttpClient, createdListApiUrl, 'Description');

      const titleResponse: SPHttpClientResponse = await spHttpClient.post(
        `${createdListApiUrl}/fields/getByInternalNameOrTitle('Title')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'X-HTTP-Method': 'MERGE',
            'If-Match': '*'
          },
          body: JSON.stringify({
            Indexed: true,
            EnforceUniqueValues: true
          })
        }
      );

      if (!titleResponse.ok) {
        console.warn('Failed to set Title uniqueness constraint. It may already be configured.');
      }
    };

    const cleanupMutex = (): void => {
      provisioningPromises.delete(catalogUrl);
    };

    const nextProvisioningPromise = doProvision()
      .then(function(): void {
        listProvisioned.set(catalogUrl, true);
        cleanupMutex();
      })
      .catch(function(error: Error): never {
        cleanupMutex();
        throw error;
      });

    provisioningPromises.set(catalogUrl, nextProvisioningPromise);

    return nextProvisioningPromise;
  };

  const findItemByKey = async (
    catalogUrl: string,
    key: string
  ): Promise<ListItemResponse | undefined> => {
    const filter = `Title eq '${escapeODataValue(key)}'`;
    const query = [
      `$filter=${encodeQueryValue(filter)}`,
      `$select=${encodeQueryValue('Id,Title,Value,Description')}`,
      '$top=1'
    ].join('&');
    const response: SPHttpClientResponse = await spHttpClient.get(
      `${getListApiUrl(catalogUrl)}/items?${query}`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to find item: ${response.statusText}`);
    }

    const data: ListItemsResponse = await response.json();

    return data.value.length > 0 ? data.value[0] : undefined;
  };

  const get = async <T = unknown>(
    key: string,
    catalogUrl: string
  ): Promise<SPFxTenantKeyValueStoreServiceItem<T> | undefined> => {
    const raw = await findItemByKey(catalogUrl, key);

    return raw ? mapItem<T>(raw) : undefined;
  };

  const list = async (catalogUrl: string): Promise<SPFxTenantKeyValueStoreServiceItem<unknown>[]> => {
    const response: SPHttpClientResponse = await spHttpClient.get(
      `${getListApiUrl(catalogUrl)}/items?$select=Id,Title,Value,Description&$orderby=Title`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to list items: ${response.statusText}`);
    }

    const data: ListItemsResponse = await response.json();

    return data.value.map(function(raw): SPFxTenantKeyValueStoreServiceItem<unknown> {
      return mapItem<unknown>(raw);
    });
  };

  const save = async <T = unknown>(
    key: string,
    value: T,
    catalogUrl: string,
    description?: string
  ): Promise<void> => {
    await ensureListReady(catalogUrl);

    const serializedValue = serializeTenantValue(value);
    const existing = await findItemByKey(catalogUrl, key);

    if (existing) {
      const updateResponse: SPHttpClientResponse = await spHttpClient.post(
        `${getListApiUrl(catalogUrl)}/items(${existing.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'X-HTTP-Method': 'MERGE',
            'If-Match': '*'
          },
          body: JSON.stringify({
            Value: serializedValue,
            Description: description ?? existing.Description ?? ''
          })
        }
      );

      if (!updateResponse.ok) {
        const errorText = await updateResponse.text();
        throw new Error(`Failed to update item: ${updateResponse.statusText}. ${errorText}`);
      }

      return;
    }

    const createResponse: SPHttpClientResponse = await spHttpClient.post(
      `${getListApiUrl(catalogUrl)}/items`,
      SPHttpClient.configurations.v1,
      {
        body: JSON.stringify({
          Title: key,
          Value: serializedValue,
          Description: description ?? ''
        })
      }
    );

    if (!createResponse.ok) {
      const errorText = await createResponse.text();
      throw new Error(`Failed to create item: ${createResponse.statusText}. ${errorText}`);
    }
  };

  const remove = async (key: string, catalogUrl: string): Promise<void> => {
    await ensureListReady(catalogUrl);

    const existing = await findItemByKey(catalogUrl, key);

    if (!existing) {
      return;
    }

    const deleteResponse: SPHttpClientResponse = await spHttpClient.post(
      `${getListApiUrl(catalogUrl)}/items(${existing.Id})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*'
        }
      }
    );

    if (!deleteResponse.ok) {
      const errorText = await deleteResponse.text();
      throw new Error(`Failed to remove item: ${deleteResponse.statusText}. ${errorText}`);
    }
  };

  return {
    ensureListReady,
    get,
    list,
    save,
    remove
  };
}

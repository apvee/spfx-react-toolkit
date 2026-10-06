import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';

import {
  deserializeTenantValue,
  escapeODataValue,
  serializeTenantValue
} from '../helpers/spfx-tenant-value.helpers';

import {
  createSiteStoreSetup, mayRecoverSiteStoreCreation, normalizeSiteStoreUrl,
  readSiteStorePage, resolveSiteStorePageUrl, siteStoreListUrl
} from './spfx-site-key-value-store.internal';

const LIST_TITLE = 'TenantKeyValueStore';

export interface SPFxListKeyValueStoreItem<T = unknown> {
  readonly key: string;
  readonly value: T;
  readonly description: string | undefined;
  readonly id: number;
}

export interface SPFxListKeyValueStoreService {
  ensureListReady: (webUrl: string) => Promise<void>;
  get: <T = unknown>(key: string, webUrl: string) => Promise<SPFxListKeyValueStoreItem<T> | undefined>;
  list: (webUrl: string) => Promise<SPFxListKeyValueStoreItem<unknown>[]>;
  save: <T = unknown>(key: string, value: T, webUrl: string, description?: string) => Promise<void>;
  remove: (key: string, webUrl: string) => Promise<void>;
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

function getListApiUrl(webUrl: string): string {
  return `${webUrl}/_api/web/lists/getByTitle('${LIST_TITLE}')`;
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

function mapItem<T>(raw: ListItemResponse): SPFxListKeyValueStoreItem<T> {
  return {
    key: raw.Title,
    value: deserializeTenantValue<T>(raw.Value),
    description: raw.Description || undefined,
    id: raw.Id
  };
}

export function createSPFxListKeyValueStoreService(
  spHttpClient: SPHttpClient,
  profile: 'tenant' | 'site'
): SPFxListKeyValueStoreService {
  const site = profile === 'site' ? createSiteStoreSetup(spHttpClient) : undefined;
  const normalize = (url: string): string => site ? normalizeSiteStoreUrl(url) : url;
  const listApi = (url: string): string => site ? siteStoreListUrl(url) : getListApiUrl(url);

  const listProvisioned = new Map<string, boolean>();
  const provisioningPromises = new Map<string, Promise<void>>();

  const ensureListReady = (webUrl: string): Promise<void> => {
    if (site) return site.ensure(webUrl);
    if (listProvisioned.get(webUrl)) {
      return Promise.resolve();
    }

    const provisioningPromise = provisioningPromises.get(webUrl);
    if (provisioningPromise) {
      return provisioningPromise;
    }

    const doProvision = async (): Promise<void> => {
      const listApiUrl = listApi(webUrl);

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
        `${webUrl}/_api/web/lists`,
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

      const createdListApiUrl = listApi(webUrl);

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
      provisioningPromises.delete(webUrl);
    };

    const nextProvisioningPromise = doProvision()
      .then(function(): void {
        listProvisioned.set(webUrl, true);
        cleanupMutex();
      })
      .catch(function(error: Error): never {
        cleanupMutex();
        throw error;
      });

    provisioningPromises.set(webUrl, nextProvisioningPromise);

    return nextProvisioningPromise;
  };

  const findItemByKey = async (
    webUrl: string,
    key: string
  ): Promise<ListItemResponse | undefined> => {
    const filter = `Title eq '${escapeODataValue(key)}'`;
    const query = [
      `$filter=${encodeQueryValue(filter)}`,
      `$select=${encodeQueryValue('Id,Title,Value,Description')}`,
      site ? '$top=2' : '$top=1'
    ].join('&');
    const response: SPHttpClientResponse = await spHttpClient.get(
      `${listApi(webUrl)}/items?${query}`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to find item: ${response.statusText}`);
    }

    const data: ListItemsResponse = await response.json();
    if (site) {
      const page = readSiteStorePage(data);
      if (page.next || page.rows.length > 1) {
        throw new Error('Site store key lookup requires one exact unique key.');
      }
      return page.rows[0];
    }
    return data.value.length > 0 ? data.value[0] : undefined;
  };

  const get = async <T = unknown>(
    key: string,
    webUrl: string
  ): Promise<SPFxListKeyValueStoreItem<T> | undefined> => {
    webUrl = normalize(webUrl);
    if (site && !await site.probe(webUrl)) return undefined;
    const raw = await findItemByKey(webUrl, key);

    return raw ? mapItem<T>(raw) : undefined;
  };

  const list = async (webUrl: string): Promise<SPFxListKeyValueStoreItem<unknown>[]> => {
    webUrl = normalize(webUrl);
    if (site && !await site.probe(webUrl)) return [];
    const initialUrl = `${listApi(webUrl)}/items?$select=Id,Title,Value,Description&$orderby=Title`;
    const response: SPHttpClientResponse = await spHttpClient.get(
      `${listApi(webUrl)}/items?$select=Id,Title,Value,Description&$orderby=Title`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to list items: ${response.statusText}`);
    }

    const data: ListItemsResponse = await response.json();
    if (site) {
      const rows: ListItemResponse[] = [];
      let page = readSiteStorePage(data);
      let currentUrl = initialUrl;
      const seen = new Set<string>([new URL(initialUrl).href]);
      rows.push(...page.rows);
      while (page.next) {
        currentUrl = resolveSiteStorePageUrl(page.next, currentUrl, initialUrl, seen);
        seen.add(currentUrl);
        const nextResponse = await spHttpClient.get(currentUrl, SPHttpClient.configurations.v1);
        if (!nextResponse.ok) throw new Error(`Failed to list items: ${nextResponse.statusText}`);
        page = readSiteStorePage(await nextResponse.json());
        rows.push(...page.rows);
      }
      return rows.map(raw => mapItem<unknown>(raw));
    }

    return data.value.map(function(raw): SPFxListKeyValueStoreItem<unknown> {
      return mapItem<unknown>(raw);
    });
  };

  const save = async <T = unknown>(
    key: string,
    value: T,
    webUrl: string,
    description?: string
  ): Promise<void> => {
    webUrl = normalize(webUrl);
    await ensureListReady(webUrl);

    const serializedValue = serializeTenantValue(value);
    const existing = await findItemByKey(webUrl, key);

    const updateItem = async (item: ListItemResponse): Promise<void> => {
      const updateResponse: SPHttpClientResponse = await spHttpClient.post(
        `${listApi(webUrl)}/items(${item.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'X-HTTP-Method': 'MERGE',
            'If-Match': '*'
          },
          body: JSON.stringify({
            Value: serializedValue,
            Description: description ?? item.Description ?? ''
          })
        }
      );

      if (!updateResponse.ok) {
        const errorText = await updateResponse.text();
        throw new Error(`Failed to update item: ${updateResponse.statusText}. ${errorText}`);
      }
    };
    if (existing) {
      await updateItem(existing);
      return;
    }

    const createResponse: SPHttpClientResponse = await spHttpClient.post(
      `${listApi(webUrl)}/items`,
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
      const original = new Error(`Failed to create item: ${createResponse.statusText}. ${errorText}`);
      if (site && mayRecoverSiteStoreCreation(createResponse)) {
        try {
          await site.fields(webUrl, true);
          const concurrent = await findItemByKey(webUrl, key);
          if (!concurrent) throw original;
          await updateItem(concurrent);
          return;
        } catch { throw original; }
      }
      throw original;
    }
  };

  const remove = async (key: string, webUrl: string): Promise<void> => {
    webUrl = normalize(webUrl);
    if (site) {
      if (!await site.probe(webUrl)) return;
    } else {
      await ensureListReady(webUrl);
    }

    const existing = await findItemByKey(webUrl, key);

    if (!existing) {
      return;
    }

    const deleteResponse: SPHttpClientResponse = await spHttpClient.post(
      `${listApi(webUrl)}/items(${existing.Id})`,
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
      const original = new Error(`Failed to remove item: ${deleteResponse.statusText}. ${errorText}`);
      if (site && deleteResponse.status === 404) {
        try {
          if (!await site.probe(webUrl) || !await findItemByKey(webUrl, key)) return;
        } catch { throw original; }
      }
      throw original;
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

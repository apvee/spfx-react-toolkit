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
  etag?: string;
}

function encodeQueryValue(value: string): string {
  return encodeURIComponent(value).replace(/'/g, '%27');
}

function serializeStoreValue(value: unknown): string {
  if (typeof value === 'bigint') return JSON.stringify(String(value));
  if (typeof value === 'string') {
    // Empty Note text is returned as null by SharePoint; quote it to preserve a new empty string.
    if (value === '') return JSON.stringify(value);
    try {
      JSON.parse(value);
      return JSON.stringify(value);
    } catch {
      return value;
    }
  }
  return serializeTenantValue(value);
}

class MutationSnapshotError extends Error {
  public constructor(public readonly status: number, statusText: string) {
    super(`Failed to read item for mutation: ${statusText}`);
  }
}

function requireItemEtag(etag: string | undefined): string {
  if (!etag || etag === '*' || /[\r\n]/.test(etag)) {
    throw new Error('A valid item ETag is required for a key-value store mutation.');
  }
  return etag;
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
  const setup = createSiteStoreSetup(spHttpClient, profile);
  const normalize = (url: string): string => normalizeSiteStoreUrl(url, profile);
  const listApi = (url: string): string => siteStoreListUrl(url, profile);
  const ensureListReady = setup.ensure;

  const findItemByKey = async (
    webUrl: string,
    key: string
  ): Promise<ListItemResponse | undefined> => {
    const filter = `Title eq '${escapeODataValue(key)}'`;
    const query = [
      `$filter=${encodeQueryValue(filter)}`,
      `$select=${encodeQueryValue('Id,Title,Value,Description')}`,
      '$top=2'
    ].join('&');
    const response: SPHttpClientResponse = await spHttpClient.get(
      `${listApi(webUrl)}/items?${query}`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to find item: ${response.statusText}`);
    }

    const data: unknown = await response.json();
    const page = readSiteStorePage(data, profile);
    if (page.next || page.rows.length > 1) {
      throw new Error(`${profile === 'site' ? 'Site' : 'Tenant'} store key lookup requires one exact unique key.`);
    }
    return page.rows[0];
  };

  const mutationItem = async (item: ListItemResponse, webUrl: string): Promise<ListItemResponse & { etag: string }> => {
    if (item.etag !== undefined) return { ...item, etag: requireItemEtag(item.etag) };
    const response = await spHttpClient.get(
      `${listApi(webUrl)}/items(${item.Id})?$select=Id,Title,Value,Description`, SPHttpClient.configurations.v1
    );
    if (!response.ok) throw new MutationSnapshotError(response.status, response.statusText);
    const data: unknown = await response.json();
    const snapshot = readSiteStorePage({ value: [data] }, profile).rows[0];
    if (snapshot.Id !== item.Id || snapshot.Title !== item.Title) {
      throw new Error(`${profile === 'site' ? 'Site' : 'Tenant'} store key changed before mutation. Reload and retry.`);
    }
    const etag = response.headers?.get('ETag') ?? snapshot.etag;
    return { ...snapshot, etag: requireItemEtag(etag) };
  };

  const get = async <T = unknown>(
    key: string,
    webUrl: string
  ): Promise<SPFxListKeyValueStoreItem<T> | undefined> => {
    webUrl = normalize(webUrl);
    if (!await setup.probe(webUrl)) return undefined;
    const raw = await findItemByKey(webUrl, key);

    return raw ? mapItem<T>(raw) : undefined;
  };

  const list = async (webUrl: string): Promise<SPFxListKeyValueStoreItem<unknown>[]> => {
    webUrl = normalize(webUrl);
    if (!await setup.probe(webUrl)) return [];
    const initialUrl = `${listApi(webUrl)}/items?$select=Id,Title,Value,Description&$orderby=Title`;
    const response: SPHttpClientResponse = await spHttpClient.get(
      `${listApi(webUrl)}/items?$select=Id,Title,Value,Description&$orderby=Title`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to list items: ${response.statusText}`);
    }

    const data: unknown = await response.json();
    const rows: ListItemResponse[] = [];
    let page = readSiteStorePage(data, profile);
    let currentUrl = initialUrl;
    const seen = new Set<string>([new URL(initialUrl).href]);
    rows.push(...page.rows);
    while (page.next) {
      currentUrl = resolveSiteStorePageUrl(page.next, currentUrl, initialUrl, seen, profile);
      seen.add(currentUrl);
      const nextResponse = await spHttpClient.get(currentUrl, SPHttpClient.configurations.v1);
      if (!nextResponse.ok) throw new Error(`Failed to list items: ${nextResponse.statusText}`);
      page = readSiteStorePage(await nextResponse.json(), profile);
      rows.push(...page.rows);
    }
    return rows.map(raw => mapItem<unknown>(raw));
  };

  const save = async <T = unknown>(
    key: string,
    value: T,
    webUrl: string,
    description?: string
  ): Promise<void> => {
    webUrl = normalize(webUrl);
    const serializedValue = serializeStoreValue(value);
    await ensureListReady(webUrl);
    const existing = await findItemByKey(webUrl, key);

    const updateItem = async (item: ListItemResponse): Promise<void> => {
      const snapshot = await mutationItem(item, webUrl);
      const updateResponse: SPHttpClientResponse = await spHttpClient.post(
        `${listApi(webUrl)}/items(${item.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'X-HTTP-Method': 'MERGE',
            'If-Match': snapshot.etag
          },
          body: JSON.stringify({
            Value: serializedValue,
            Description: description ?? snapshot.Description ?? ''
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
      if (mayRecoverSiteStoreCreation(createResponse)) {
        let concurrent: ListItemResponse | undefined;
        try {
          await setup.fields(webUrl, true);
          concurrent = await findItemByKey(webUrl, key);
          if (!concurrent) throw original;
        } catch { throw original; }
        await updateItem(concurrent);
        return;
      }
      throw original;
    }
  };

  const remove = async (key: string, webUrl: string): Promise<void> => {
    webUrl = normalize(webUrl);
    if (!await setup.probe(webUrl)) return;

    const existing = await findItemByKey(webUrl, key);

    if (!existing) {
      return;
    }

    let snapshot: ListItemResponse & { etag: string };
    try {
      snapshot = await mutationItem(existing, webUrl);
    } catch (error) {
      if (error instanceof MutationSnapshotError && error.status === 404) {
        try {
          if (!await setup.probe(webUrl) || !await findItemByKey(webUrl, key)) return;
        } catch { throw error; }
      }
      throw error;
    }
    const deleteResponse: SPHttpClientResponse = await spHttpClient.post(
      `${listApi(webUrl)}/items(${existing.Id})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': snapshot.etag
        }
      }
    );

    if (!deleteResponse.ok) {
      const errorText = await deleteResponse.text();
      const original = new Error(`Failed to remove item: ${deleteResponse.statusText}. ${errorText}`);
      if (deleteResponse.status === 404) {
        try {
          if (!await setup.probe(webUrl) || !await findItemByKey(webUrl, key)) return;
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

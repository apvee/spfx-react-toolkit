import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';
import { deserializeValue, escapeODataValue } from './useSPFxTenantKeyValueStore.serialization.internal';
import type { SPFxTenantKeyValueStoreItem } from './useSPFxTenantKeyValueStore';

export const TENANT_KEY_VALUE_STORE_LIST_TITLE = 'TenantKeyValueStore';

export interface ListItemResponse {
    Id: number;
    Title: string;
    Value: string;
    Description?: string;
}

export interface ListItemsResponse {
    value: ListItemResponse[];
}

export interface FieldsCheckResponse {
    value: Array<{ InternalName: string }>;
}

export function getListApiUrl(catalogUrl: string): string {
    return `${catalogUrl}/_api/web/lists/getByTitle('${TENANT_KEY_VALUE_STORE_LIST_TITLE}')`;
}

export async function createField(
    client: SPHttpClient,
    listApiUrl: string,
    fieldTitle: string
): Promise<void> {
    const response: SPHttpClientResponse = await client.post(
        `${listApiUrl}/fields`,
        SPHttpClient.configurations.v1,
        {
            body: JSON.stringify({
                FieldTypeKind: 3,
                Title: fieldTitle,
            }),
        }
    );

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to create ${fieldTitle} field: ${response.statusText}. ${errorText}`);
    }
}

export async function ensureTenantKeyValueStoreList(
    client: SPHttpClient,
    catalogUrl: string
): Promise<void> {
    const listApiUrl = getListApiUrl(catalogUrl);

    const listResponse: SPHttpClientResponse = await client.get(
        `${listApiUrl}?$select=Id`,
        SPHttpClient.configurations.v1
    );

    const listExists = listResponse.status !== 404;

    if (listExists && !listResponse.ok) {
        throw new Error(`Failed to check list existence: ${listResponse.statusText}`);
    }

    if (listExists) {
        const fieldsResponse: SPHttpClientResponse = await client.get(
            `${listApiUrl}/fields?$filter=InternalName eq 'Value' or InternalName eq 'Description'&$select=InternalName`,
            SPHttpClient.configurations.v1
        );

        if (!fieldsResponse.ok) {
            throw new Error(`Failed to check list fields: ${fieldsResponse.statusText}`);
        }

        const fields: FieldsCheckResponse = await fieldsResponse.json();
        const existingFields = fields.value.map(f => f.InternalName);
        const hasValue = existingFields.includes('Value');
        const hasDescription = existingFields.includes('Description');

        if (hasValue && hasDescription) {
            return;
        }

        if (!hasValue) {
            await createField(client, listApiUrl, 'Value');
        }
        if (!hasDescription) {
            await createField(client, listApiUrl, 'Description');
        }

        return;
    }

    const createListResponse: SPHttpClientResponse = await client.post(
        `${catalogUrl}/_api/web/lists`,
        SPHttpClient.configurations.v1,
        {
            body: JSON.stringify({
                BaseTemplate: 100,
                Title: TENANT_KEY_VALUE_STORE_LIST_TITLE,
                Hidden: true,
                NoCrawl: true,
            }),
        }
    );

    if (!createListResponse.ok) {
        const errorText = await createListResponse.text();
        throw new Error(`Failed to create list: ${createListResponse.statusText}. ${errorText}`);
    }

    await createField(client, getListApiUrl(catalogUrl), 'Value');
    await createField(client, getListApiUrl(catalogUrl), 'Description');

    const titleResp: SPHttpClientResponse = await client.post(
        `${getListApiUrl(catalogUrl)}/fields/getByInternalNameOrTitle('Title')`,
        SPHttpClient.configurations.v1,
        {
            headers: {
                'X-HTTP-Method': 'MERGE',
                'If-Match': '*',
            },
            body: JSON.stringify({
                Indexed: true,
                EnforceUniqueValues: true,
            }),
        }
    );

    if (!titleResp.ok) {
        console.warn('Failed to set Title uniqueness constraint. It may already be configured.');
    }
}

export async function findTenantKeyValueStoreItemByKey(
    client: SPHttpClient,
    catalogUrl: string,
    key: string
): Promise<ListItemResponse | undefined> {
    const safeKey = escapeODataValue(key);
    const response: SPHttpClientResponse = await client.get(
        `${getListApiUrl(catalogUrl)}/items?$filter=Title eq '${safeKey}'&$select=Id,Title,Value,Description&$top=1`,
        SPHttpClient.configurations.v1
    );

    if (!response.ok) {
        throw new Error(`Failed to find item: ${response.statusText}`);
    }

    const data: ListItemsResponse = await response.json();
    return data.value.length > 0 ? data.value[0] : undefined;
}

export function mapTenantKeyValueStoreItem<T>(raw: ListItemResponse): SPFxTenantKeyValueStoreItem<T> {
    return {
        key: raw.Title,
        value: deserializeValue<T>(raw.Value),
        description: raw.Description || undefined,
        id: raw.Id,
    };
}

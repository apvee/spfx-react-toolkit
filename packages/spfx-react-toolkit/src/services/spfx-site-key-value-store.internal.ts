import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';

type StoreProfile = 'site' | 'tenant';

interface SiteField {
  InternalName: string;
  FieldTypeKind: number;
  Indexed?: boolean;
  EnforceUniqueValues?: boolean;
}

export interface SiteStoreRow {
  Id: number;
  Title: string;
  Value: string;
  Description?: string;
  etag?: string;
}

export function normalizeSiteStoreUrl(url: string, profile: StoreProfile = 'site'): string {
  const normalized = url.trim().replace(/\/+$/, '');
  if (!normalized) throw new Error(`A nonempty ${profile === 'site' ? 'site collection' : 'tenant app catalog'} URL is required.`);
  return normalized;
}

export function siteStoreListUrl(url: string, profile: StoreProfile = 'site'): string {
  return `${normalizeSiteStoreUrl(url, profile)}/_api/web/lists/getByTitle('${profile === 'site' ? 'Site' : 'Tenant'}KeyValueStore')`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

export function readSiteStorePage(data: unknown, profile: StoreProfile = 'site'): { rows: SiteStoreRow[]; next?: string } {
  if (!isRecord(data) || !Array.isArray(data.value)) {
    throw new Error(`Invalid ${profile} key-value store items response.`);
  }
  const rows = data.value.map((raw: unknown): SiteStoreRow => {
    if (!isRecord(raw) || !Number.isInteger(raw.Id) || Number(raw.Id) <= 0 ||
      typeof raw.Title !== 'string' || (raw.Value !== null && typeof raw.Value !== 'string') ||
      (raw.Description !== undefined && raw.Description !== null && typeof raw.Description !== 'string')) {
      throw new Error(`Invalid ${profile} key-value store item row.`);
    }
    const metadata = isRecord(raw.__metadata) ? raw.__metadata.etag : undefined;
    const etag = raw['@odata.etag'] ?? raw['odata.etag'] ?? metadata;
    // SharePoint returns null for an empty Note. Tenant's legacy JSON-first reader returned null; site readers historically returned ''.
    return { Id: Number(raw.Id), Title: raw.Title, Value: typeof raw.Value === 'string' ? raw.Value : profile === 'tenant' ? 'null' : '', Description: typeof raw.Description === 'string' ? raw.Description : undefined,
      etag: typeof etag === 'string' ? etag : undefined };
  });
  const next = data['@odata.nextLink'] ?? data['odata.nextLink'];
  if (next !== undefined && (typeof next !== 'string' || !next.trim())) {
    throw new Error(`Invalid ${profile} key-value store continuation link.`);
  }
  return { rows, next: typeof next === 'string' ? next : undefined };
}

export function resolveSiteStorePageUrl(link: string, currentUrl: string, itemsUrl: string, seen: Set<string>, profile: StoreProfile = 'site'): string {
  const next = new URL(link, currentUrl);
  const expected = new URL(itemsUrl);
  if (next.origin !== expected.origin || decodeURIComponent(next.pathname).toLowerCase() !== decodeURIComponent(expected.pathname).toLowerCase() ||
    next.hash || next.username || next.password || seen.has(next.href)) {
    throw new Error(`Foreign or repeated ${profile} key-value store continuation link.`);
  }
  return next.href;
}

export async function siteStoreResponseError(response: SPHttpClientResponse, operation: string): Promise<Error> {
  return new Error(`Failed to ${operation}: ${response.statusText}. ${await response.text()}`);
}

export function mayRecoverSiteStoreCreation(response: SPHttpClientResponse): boolean {
  return response.status !== 401 && response.status !== 403;
}

export function createSiteStoreSetup(spHttpClient: SPHttpClient, profile: StoreProfile = 'site'): {
  exists: (url: string) => Promise<boolean>;
  probe: (url: string) => Promise<boolean>;
  fields: (url: string, requireUnique?: boolean) => Promise<SiteField[]>;
  ensure: (url: string) => Promise<void>;
} {
  const ready = new Set<string>();
  const pending = new Map<string, Promise<void>>();
  const exists = async (url: string): Promise<boolean> => {
    const response = await spHttpClient.get(`${siteStoreListUrl(url, profile)}?$select=Id`, SPHttpClient.configurations.v1);
    if (response.status === 404) return false;
    if (!response.ok) throw new Error(`Failed to check list existence: ${response.statusText}`);
    return true;
  };
  const fields = async (url: string, requireUnique = false): Promise<SiteField[]> => {
    const response = await spHttpClient.get(
      `${siteStoreListUrl(url, profile)}/fields?$filter=InternalName eq 'Title' or InternalName eq 'Value' or InternalName eq 'Description'&$select=InternalName,FieldTypeKind,Indexed,EnforceUniqueValues`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) throw new Error(`Failed to check list fields: ${response.statusText}`);
    const data: unknown = await response.json();
    if (!isRecord(data) || !Array.isArray(data.value)) throw new Error(`Invalid ${profile} store field schema.`);
    const names = new Set<string>();
    const result = data.value.map((field: unknown): SiteField => {
      if (!isRecord(field) || typeof field.InternalName !== 'string' || typeof field.FieldTypeKind !== 'number' || names.has(field.InternalName)) {
        throw new Error(`Invalid ${profile} store field schema.`);
      }
      names.add(field.InternalName);
      const expected = field.InternalName === 'Title' ? 2 : 3;
      if (field.FieldTypeKind !== expected) throw new Error(`Incompatible ${field.InternalName} field type in ${profile} store schema.`);
      return { InternalName: field.InternalName, FieldTypeKind: field.FieldTypeKind, Indexed: field.Indexed === true, EnforceUniqueValues: field.EnforceUniqueValues === true };
    });
    if (requireUnique) validateComplete(result, true, profile);
    return result;
  };
  const probe = async (url: string): Promise<boolean> => {
    if (!await exists(url)) return false;
    validateComplete(await fields(url), false, profile);
    return true;
  };
  const configureSchema = async (url: string): Promise<void> => {
    let current = await fields(url);
    for (const name of ['Value', 'Description']) {
      if (current.some(field => field.InternalName === name)) continue;
      const response = await spHttpClient.post(`${siteStoreListUrl(url, profile)}/fields`, SPHttpClient.configurations.v1, {
        body: JSON.stringify({ FieldTypeKind: 3, Title: name })
      });
      if (!response.ok) {
        const original = await siteStoreResponseError(response, `create ${name} field`);
        if (!mayRecoverSiteStoreCreation(response)) throw original;
        try {
          current = await fields(url);
          if (!current.some(field => field.InternalName === name)) throw original;
        } catch { throw original; }
      } else {
        current = [...current, { InternalName: name, FieldTypeKind: 3 }];
      }
    }
    validateComplete(current, false, profile);
    const title = current.find(field => field.InternalName === 'Title');
    if (!title?.Indexed || !title.EnforceUniqueValues) {
      const response = await spHttpClient.post(`${siteStoreListUrl(url, profile)}/fields/getByInternalNameOrTitle('Title')`, SPHttpClient.configurations.v1, {
        headers: { 'X-HTTP-Method': 'MERGE', 'If-Match': '*' }, body: JSON.stringify({ Indexed: true, EnforceUniqueValues: true })
      });
      if (!response.ok) throw await siteStoreResponseError(response, 'set Title uniqueness constraint');
      await fields(url, true);
    }
  };
  const provision = async (url: string): Promise<void> => {
    let recoveredCreationError: Error | undefined;
    if (!await exists(url)) {
      const response = await spHttpClient.post(`${url}/_api/web/lists`, SPHttpClient.configurations.v1, {
        body: JSON.stringify({ BaseTemplate: 100, Title: profile === 'site' ? 'SiteKeyValueStore' : 'TenantKeyValueStore', Hidden: true, NoCrawl: true })
      });
      if (!response.ok) {
        const original = await siteStoreResponseError(response, 'create list');
        if (!mayRecoverSiteStoreCreation(response)) throw original;
        try { if (!await exists(url)) throw original; } catch { throw original; }
        recoveredCreationError = original;
      }
    }
    try {
      await configureSchema(url);
    } catch (error) {
      throw recoveredCreationError ?? error;
    }
  };
  const ensure = (input: string): Promise<void> => {
    let url: string;
    try { url = normalizeSiteStoreUrl(input, profile); } catch (error) { return Promise.reject(error); }
    if (ready.has(url)) return Promise.resolve();
    const active = pending.get(url);
    if (active) return active;
    const promise = provision(url).then(() => { ready.add(url); }).finally(() => { pending.delete(url); });
    pending.set(url, promise);
    return promise;
  };
  return { exists, probe, fields, ensure };
}

function validateComplete(fields: SiteField[], requireUnique: boolean, profile: StoreProfile): void {
  for (const name of ['Title', 'Value', 'Description']) {
    const field = fields.find(candidate => candidate.InternalName === name);
    if (!field) throw new Error(`Missing ${name} field in ${profile} store schema.`);
    if (requireUnique && name === 'Title' && (!field.Indexed || !field.EnforceUniqueValues)) {
      throw new Error(`${profile === 'site' ? 'Site' : 'Tenant'} store schema requires indexed unique Title keys.`);
    }
  }
}

function readPermissionMask(data: unknown): Record<string, unknown> | undefined {
  if (!isRecord(data)) return undefined;
  if ('d' in data) {
    return isRecord(data.d) && isRecord(data.d.EffectiveBasePermissions) ? data.d.EffectiveBasePermissions : undefined;
  }
  if ('EffectiveBasePermissions' in data) {
    return isRecord(data.EffectiveBasePermissions) ? data.EffectiveBasePermissions : undefined;
  }
  return data;
}

function permissionWord(value: unknown): number | undefined {
  if (typeof value !== 'number' && (typeof value !== 'string' || !/^\d+$/.test(value))) return undefined;
  const number = Number(value);
  return Number.isInteger(number) && number >= 0 && number <= 0xffffffff ? number : undefined;
}

export async function canWriteKeyValueStore(spHttpClient: SPHttpClient, inputUrl: string, profile: StoreProfile): Promise<boolean> {
  const setup = createSiteStoreSetup(spHttpClient, profile);
  try {
    const url = normalizeSiteStoreUrl(inputUrl, profile);
    const exists = await setup.exists(url);
    const response = await spHttpClient.get(
      `${exists ? siteStoreListUrl(url, profile) : `${url}/_api/web`}/effectiveBasePermissions`, SPHttpClient.configurations.v1
    );
    if (!response.ok) return false;
    const data: unknown = await response.json();
    const record = readPermissionMask(data);
    if (!record) return false;
    const High = permissionWord(record.High);
    const Low = permissionWord(record.Low);
    if (High === undefined || Low === undefined) return false;
    const permissions = new SPPermission({ High, Low });
    return permissions.hasAllPermissions(SPPermission.addListItems, SPPermission.editListItems, SPPermission.deleteListItems) &&
      (exists || permissions.hasPermission(SPPermission.manageLists));
  } catch { return false; }
}

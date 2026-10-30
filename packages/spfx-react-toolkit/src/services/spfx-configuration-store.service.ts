import type { SPHttpClient } from '@microsoft/sp-http';
import { createSPFxAppCatalogService } from './spfx-app-catalog.service';
import { createSPFxSiteKeyValueStoreService } from './spfx-site-key-value-store.service';
import { createSPFxTenantKeyValueStoreService } from './spfx-tenant-key-value-store.service';

/**
 * Physical configuration destination. Site means the collection root, shared by its subsites.
 * @example
 * const scope: SPFxConfigurationStoreScope = 'site';
 */
export type SPFxConfigurationStoreScope = 'tenant' | 'site';

/**
 * Allowed physical stores and effective-read policy. Layered reads site first, then tenant on confirmed absence.
 * @example
 * const mode: SPFxConfigurationStoreMode = 'layered';
 */
export type SPFxConfigurationStoreMode = SPFxConfigurationStoreScope | 'layered';

/**
 * Required mode and its addressable dependencies; no defaults. URLs are snapshotted and normalized at construction.
 * @example
 * const options: SPFxConfigurationStoreServiceOptions = {
 *   mode: 'layered', discoveryWebUrl: context.pageContext.web.absoluteUrl,
 *   siteCollectionUrl: context.pageContext.site.absoluteUrl,
 * };
 */
export type SPFxConfigurationStoreServiceOptions =
  | {
    /** Only the tenant store is available. */
    readonly mode: 'tenant';
    /** Current web URL used for lazy tenant catalog discovery; not the catalog URL. */
    readonly discoveryWebUrl: string;
    /** Unused in tenant mode. */
    readonly siteCollectionUrl?: never;
  }
  | {
    /** Only the site collection store is available. */
    readonly mode: 'site';
    /** Absolute collection-root URL, never a subsite URL. */
    readonly siteCollectionUrl: string;
    /** Unused in site mode. */
    readonly discoveryWebUrl?: never;
  }
  | {
    /** Site values override complete tenant values by key. */
    readonly mode: 'layered';
    /** Current web URL used for lazy tenant catalog discovery. */
    readonly discoveryWebUrl: string;
    /** Absolute collection-root URL. */
    readonly siteCollectionUrl: string;
  };

/**
 * Hook configuration. The matching SPFx provider supplies the client and mode-specific URLs.
 * @example
 * const options: SPFxConfigurationStoreOptions = { mode: 'site' };
 */
export interface SPFxConfigurationStoreOptions {
  /** Required read policy and allowed physical scopes; no default. */
  readonly mode: SPFxConfigurationStoreMode;
}

/**
 * Explicit mutation destination, independent of a previously read value's source.
 * @example
 * const target: SPFxConfigurationStoreTarget = { scope: 'site' };
 */
export interface SPFxConfigurationStoreTarget {
  /** Required physical scope; must be enabled by the configured mode. */
  readonly scope: SPFxConfigurationStoreScope;
}

/**
 * Save destination and destination-local metadata.
 * @example
 * const options: SPFxConfigurationStoreSaveOptions = { scope: 'site', description: '' };
 */
export interface SPFxConfigurationStoreSaveOptions extends SPFxConfigurationStoreTarget {
  /** Omitted preserves existing destination metadata; empty string clears it. Never inherits metadata from another scope. */
  readonly description?: string;
}

/**
 * One physical record and its provenance. T describes expected data; it does not validate stored content.
 * @example
 * const item: SPFxConfigurationStoreItem<number> | undefined = await store.get<number>('my-app:pageSize');
 * console.log(item?.value, item?.source);
 */
export interface SPFxConfigurationStoreItem<T = unknown> {
  /** Stored key with server-provided casing, without normalization. */
  readonly key: string;
  /** Whole deserialized value, retaining the source store's historical interpretation. */
  readonly value: T;
  /** Metadata from the same physical record; absent/empty descriptions read as undefined. */
  readonly description: string | undefined;
  /** Source-list item ID, not a globally unique identity. */
  readonly id: number;
  /** Physical store that supplied this record; does not route later writes. */
  readonly source: SPFxConfigurationStoreScope;
}

/**
 * Configuration operations without React. All failed operations reject; only confirmed absence is empty.
 * @example
 * await store.save('my-app:pageSize', 50, { scope: 'site' });
 * const item = await store.get<number>('my-app:pageSize');
 * await store.remove('my-app:pageSize', { scope: 'site' });
 */
export interface SPFxConfigurationStoreService {
  /**
   * Reads the mode's effective record without provisioning. Layered reads site first and consults tenant only after confirmed site absence.
   * @param key - String key forwarded unchanged; equality is delegated to SharePoint.
   * @returns Whole winning record with source, or undefined for confirmed absence in all required scopes.
   * @throws Discovery, access, schema or lookup failures; a failed site read never enables tenant fallback.
   */
  get: <T = unknown>(key: string) => Promise<SPFxConfigurationStoreItem<T> | undefined>;
  /**
   * Inventories one physical store, preserving paging order, casing and legacy duplicates; never merges scopes.
   * @param scope - Required enabled physical scope.
   * @returns Complete physical inventory, or [] for confirmed absent list. No partial results on page failures.
   * @throws Disabled/invalid scope, discovery, schema or read failures.
   */
  listScope: (scope: SPFxConfigurationStoreScope) => Promise<SPFxConfigurationStoreItem<unknown>[]>;
  /**
   * Replaces the complete value only at the explicit destination, provisioning that store if necessary. No JSON-property merge or automatic read follows.
   * @param key - String key forwarded unchanged.
   * @param value - Existing store encoding applies: new text is preserved, bigint reads as decimal string and Date as ISO text. T is not validation.
   * @param options - Required enabled scope and optional destination description.
   * @returns Resolves after successful destination mutation.
   * @throws Invalid target/input, serialization, discovery, setup or mutation failures, including observed-ETag conflicts. No failover or additional retry.
   */
  save: <T = unknown>(key: string, value: T, options: SPFxConfigurationStoreSaveOptions) => Promise<void>;
  /**
   * Removes only the selected physical record without provisioning. Removing a site override reveals tenant data on the next effective read.
   * @param key - String key forwarded unchanged.
   * @param target - Required enabled physical scope, even in single-store modes.
   * @returns Resolves after deletion or confirmed absence; never deletes both scopes.
   * @throws Invalid target, discovery or mutation failures/conflicts.
   */
  remove: (key: string, target: SPFxConfigurationStoreTarget) => Promise<void>;
}

/**
 * Creates an isolated facade over existing list stores. Construction performs no requests. Tenant discovery is lazy and shared while pending within this instance; successful URLs use the existing catalog cache. Site-only operations never need catalog access.
 * Existing APIs/data, nullable-Note legacy differences, serialization and ETag preconditions remain unchanged. No cross-store transaction, value cache or permission indicator is provided.
 * @param client - Caller-owned authenticated SharePoint client.
 * @param options - Required mode and relevant URLs; copied at construction, never read again from the caller's object.
 * @returns Effective reads, physical inventories and explicitly scoped mutations.
 * @throws Invalid mode or missing/empty required URLs synchronously. Operation failures reject their Promises. Tenant serialization failures can follow discovery, but never setup/item writes.
 * @example
 * const store = createSPFxConfigurationStoreService(context.spHttpClient, {
 *   mode: 'layered', discoveryWebUrl: context.pageContext.web.absoluteUrl,
 *   siteCollectionUrl: context.pageContext.site.absoluteUrl,
 * });
 * await store.save('my-app:pageSize', 20, { scope: 'tenant' });
 * await store.save('my-app:pageSize', 50, { scope: 'site' });
 * console.log((await store.get<number>('my-app:pageSize'))?.source); // site
 */
export function createSPFxConfigurationStoreService(
  client: SPHttpClient,
  options: SPFxConfigurationStoreServiceOptions
): SPFxConfigurationStoreService {
  const mode = options?.mode;
  if (mode !== 'tenant' && mode !== 'site' && mode !== 'layered') throw new Error('A valid configuration store mode is required.');
  const normalize = (url: string | undefined): string => {
    if (typeof url !== 'string') throw new Error('A nonempty configuration store URL is required.');
    const normalized = url.trim().replace(/\/+$/, '');
    if (!normalized) throw new Error('A nonempty configuration store URL is required.');
    return normalized;
  };
  const siteUrl = mode !== 'tenant' ? normalize(options.siteCollectionUrl) : undefined;
  const discoveryUrl = mode !== 'site' ? normalize(options.discoveryWebUrl) : undefined;
  const site = siteUrl ? createSPFxSiteKeyValueStoreService(client) : undefined;
  const tenant = discoveryUrl ? createSPFxTenantKeyValueStoreService(client) : undefined;
  const catalog = discoveryUrl ? createSPFxAppCatalogService(client, { web: { absoluteUrl: discoveryUrl } }) : undefined;
  let pendingCatalog: Promise<string> | undefined;
  const catalogUrl = (): Promise<string> => {
    if (!catalog) return Promise.reject(new Error('Tenant scope is not enabled.'));
    if (!pendingCatalog) {
      pendingCatalog = catalog.discoverUrl().finally(() => { pendingCatalog = undefined; });
    }
    return pendingCatalog;
  };
  const scopeOf = (scope: SPFxConfigurationStoreScope): SPFxConfigurationStoreScope => {
    if ((scope !== 'tenant' && scope !== 'site') || (mode !== 'layered' && scope !== mode)) throw new Error('An enabled tenant or site scope is required.');
    return scope;
  };
  const keyOf = (key: string): void => { if (typeof key !== 'string') throw new Error('A string configuration key is required.'); };
  const targetOf = (target: SPFxConfigurationStoreTarget): SPFxConfigurationStoreScope => scopeOf(target?.scope);
  const physicalGet = async <T,>(key: string, scope: SPFxConfigurationStoreScope): Promise<SPFxConfigurationStoreItem<T> | undefined> => {
    const item = scope === 'site' && site && siteUrl ? await site.get<T>(key, siteUrl)
      : tenant ? await tenant.get<T>(key, await catalogUrl()) : undefined;
    return item === undefined ? undefined : { ...item, source: scope };
  };
  return {
    get: async <T,>(key: string): Promise<SPFxConfigurationStoreItem<T> | undefined> => {
      keyOf(key);
      if (mode !== 'layered') return physicalGet<T>(key, mode);
      const local = await physicalGet<T>(key, 'site');
      return local !== undefined ? local : physicalGet<T>(key, 'tenant');
    },
    listScope: async (requestedScope): Promise<SPFxConfigurationStoreItem<unknown>[]> => {
      const scope = scopeOf(requestedScope);
      const items = scope === 'site' && site && siteUrl ? await site.list(siteUrl)
        : tenant ? await tenant.list(await catalogUrl()) : [];
      return items.map(item => ({ ...item, source: scope }));
    },
    save: async <T,>(key: string, value: T, saveOptions: SPFxConfigurationStoreSaveOptions): Promise<void> => {
      keyOf(key);
      const scope = targetOf(saveOptions);
      const description = saveOptions.description;
      if (description !== undefined && typeof description !== 'string') throw new Error('Configuration description must be a string.');
      if (scope === 'site' && site && siteUrl) await site.save(key, value, siteUrl, description);
      else if (tenant) await tenant.save(key, value, await catalogUrl(), description);
    },
    remove: async (key, target): Promise<void> => {
      keyOf(key);
      const scope = targetOf(target);
      if (scope === 'site' && site && siteUrl) await site.remove(key, siteUrl);
      else if (tenant) await tenant.remove(key, await catalogUrl());
    }
  };
}

import { SPHttpClient } from '@microsoft/sp-http';

import { createSPFxListKeyValueStoreService } from './spfx-list-key-value-store.internal';

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

export function createSPFxTenantKeyValueStoreService(
  spHttpClient: SPHttpClient
): SPFxTenantKeyValueStoreService {
  return createSPFxListKeyValueStoreService(spHttpClient, 'tenant');
}

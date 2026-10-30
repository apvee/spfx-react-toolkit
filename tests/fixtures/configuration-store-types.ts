import type { SPHttpClient } from '@microsoft/sp-http';
import { createSPFxConfigurationStoreService, useSPFxConfigurationStore } from '@apvee/spfx-react-toolkit';
import type { SPFxConfigurationStoreItem } from '@apvee/spfx-react-toolkit/services';
import type { SPFxConfigurationStoreItem as HookItem, SPFxConfigurationStoreOptions } from '@apvee/spfx-react-toolkit/hooks';

export function contracts(client: SPHttpClient): void {
  const service = createSPFxConfigurationStoreService(client, { mode: 'site', siteCollectionUrl: 'https://tenant.test/sites/root' });
  void service.save('my-app:pageSize', 50, { scope: 'site' });
  void service.remove('my-app:pageSize', { scope: 'site' });
  const item: Promise<SPFxConfigurationStoreItem<number> | undefined> = service.get<number>('my-app:pageSize');
  const equivalent: Promise<HookItem<number> | undefined> = item;
  void equivalent;
  // @ts-expect-error Every mutation requires an explicit scope.
  void service.save('key', 1);
  // @ts-expect-error Every removal requires an explicit scope.
  void service.remove('key');
  // @ts-expect-error Only physical scopes are accepted.
  void service.save('key', 1, { scope: 'layered' });
  // @ts-expect-error Description must be text.
  void service.save('key', 1, { scope: 'site', description: 7 });
  // @ts-expect-error Inventories are explicitly physical, never an effective merged list.
  void service.list();
  // @ts-expect-error Tenant discovery URL is required.
  createSPFxConfigurationStoreService(client, { mode: 'tenant' });
  // @ts-expect-error Mode-specific URLs cannot be supplied to site-only mode.
  createSPFxConfigurationStoreService(client, { mode: 'site', siteCollectionUrl: 'https://tenant.test', discoveryWebUrl: 'https://tenant.test' });
  // @ts-expect-error Both URL dependencies are required in layered mode.
  createSPFxConfigurationStoreService(client, { mode: 'layered', discoveryWebUrl: 'https://tenant.test' });
}
export function useContracts(options: SPFxConfigurationStoreOptions): void {
  const store = useSPFxConfigurationStore(options);
  const ready: boolean = store.isReady;
  void ready;
  void store.listScope('tenant');
  // @ts-expect-error Mode is required.
  useSPFxConfigurationStore({});
  // @ts-expect-error No implicit write destination option.
  useSPFxConfigurationStore({ mode: 'layered', writeTo: 'site' });
}

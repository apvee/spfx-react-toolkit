import type { SPHttpClient } from '@microsoft/sp-http';
import { createSPFxConfigurationStoreService, useSPFxConfigurationStore } from '@apvee/spfx-react-toolkit';
import { createSPFxConfigurationStoreService as domainFactory } from '@apvee/spfx-react-toolkit/services';
import { useSPFxConfigurationStore as domainHook } from '@apvee/spfx-react-toolkit/hooks';
import type { SPFxConfigurationStoreItem, SPFxConfigurationStoreServiceOptions } from '@apvee/spfx-react-toolkit/services';
import type { SPFxConfigurationStoreItem as HookItem, SPFxConfigurationStoreResult } from '@apvee/spfx-react-toolkit/hooks';

export const configurationContracts = { createSPFxConfigurationStoreService, domainFactory, useSPFxConfigurationStore, domainHook };
export async function serviceExample(client: SPHttpClient): Promise<HookItem<number> | undefined> {
  const options: SPFxConfigurationStoreServiceOptions = {
    mode: 'layered', discoveryWebUrl: 'https://tenant.test/sites/root/child', siteCollectionUrl: 'https://tenant.test/sites/root'
  };
  const store = domainFactory(client, options);
  await store.save('my-app:pageSize', 50, { scope: 'site' });
  const item: SPFxConfigurationStoreItem<number> | undefined = await store.get<number>('my-app:pageSize');
  await store.listScope('site');
  await store.remove('my-app:pageSize', { scope: 'site' });
  return item;
}
export function useConfigurationExample(): SPFxConfigurationStoreResult {
  return domainHook({ mode: 'layered' });
}

import { width as rootWidth, useSx as rootSx, useStableCallback as rootCallback } from '@apvee/spfx-react-toolkit';
import { SPFxWebPartProvider } from '@apvee/spfx-react-toolkit/core';
import { useSPFxProperties, useSPFxPnPListById } from '@apvee/spfx-react-toolkit/hooks';
import { createSPFxPnPListService } from '@apvee/spfx-react-toolkit/services';
import { width as helperWidth } from '@apvee/spfx-react-toolkit/helpers';
import { width, useSx } from '@apvee/spfx-react-toolkit/styles';
import type { SxDescriptor, SxFunction } from '@apvee/spfx-react-toolkit/styles';
import { useSx as leafSx } from '@apvee/spfx-react-toolkit/styles/useSx';
import { useStableCallback } from '@apvee/spfx-react-toolkit/hooks/useStableCallback';
import { width as directoryWidth } from '@apvee/spfx-react-toolkit/lib/helpers/styles';
import { useStableCallback as legacyCallback } from '@apvee/spfx-react-toolkit/lib/hooks/useStableCallback';
import { useSx as legacySx } from '@apvee/spfx-react-toolkit/lib/hooks/useSx.js';
import type { SPFxPnPListSelector } from '@apvee/spfx-react-toolkit/lib/services/spfx-pnp-list.service';
import { serializeValue } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxTenantKeyValueStore.serialization.internal';
import { useSPFxContext } from '@apvee/spfx-react-toolkit/lib';
// @ts-expect-error Private clean paths are deliberately unavailable.
import {} from '@apvee/spfx-react-toolkit/hooks/useAsyncInvoke.internal';
// @ts-expect-error Private descriptor implementation has no clean alias.
import {} from '@apvee/spfx-react-toolkit/styles/descriptor.internal';
// @ts-expect-error Types are public through styles, not the branded internal types module.
import {} from '@apvee/spfx-react-toolkit/styles/types';
// @ts-expect-error Private brands must not be re-exported by the public facade.
import { sxDescriptorBrand } from '@apvee/spfx-react-toolkit/styles';

export const contracts = { rootWidth, rootSx, rootCallback, SPFxWebPartProvider, createSPFxPnPListService,
  helperWidth, directoryWidth, legacyCallback, legacySx, serializeValue, useSPFxContext };
export const descriptor: SxDescriptor = width.full;
export const composer: typeof useSx = leafSx;
export function useContracts(): { title: string | undefined; text: string; pending: Promise<number>; sx: SxFunction } {
  const properties = useSPFxProperties<{ title: string }>();
  properties.setProperties({ title: 'Title' });
  // @ts-expect-error Property generic cannot become any.
  properties.setProperties({ title: 7 });
  const list = useSPFxPnPListById<{ Title: string }>('guid');
  // @ts-expect-error List item fields retain the generic shape.
  list.create({ Title: false });
  const sync = useStableCallback((id: number, label: string) => label + id);
  const asyncCallback = legacyCallback(async (id: number) => id);
  // @ts-expect-error Callback argument types remain inferred.
  sync('bad', 'label');
  // @ts-expect-error Callback return types remain inferred.
  const wrong: number = sync(1, 'label');
  void wrong;
  return { title: properties.properties?.title, text: sync(1, 'label'), pending: asyncCallback(1), sx: useSx() };
}
export function readonlySelector(selector: Extract<SPFxPnPListSelector, { kind: 'id' }>): void {
  // @ts-expect-error Selector discriminants remain readonly.
  selector.kind = 'id';
  // @ts-expect-error Selector values remain readonly.
  selector.id = 'changed';
}

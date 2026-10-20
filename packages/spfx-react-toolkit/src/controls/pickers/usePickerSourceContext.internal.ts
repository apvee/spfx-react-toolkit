import { useMemo } from 'react';
import { useSPFxContext } from '../../hooks/useSPFxContext';
import { useSPFxPageContext } from '../../hooks/useSPFxPageContext';
import { getSPFxCorrelationInfo } from '../../helpers/spfx-page-context.helpers';
import { useSPFxSPHttpClient } from '../../hooks/useSPFxSPHttpClient';
import { normalizePickerUrl } from './picker-identity.internal';
import type { PickerSourceContext } from './picker-source.types.internal';

/** Reuses native provider/client readiness; changing the source uses an explicit URL rather than initial hook baseUrl. */
export function usePickerSourceContext(): PickerSourceContext {
  const { spfxContext } = useSPFxContext(), page = useSPFxPageContext(), { client } = useSPFxSPHttpClient();
  const tenantId = getSPFxCorrelationInfo(page).tenantId ?? '';
  const userId = page.user.loginName;
  const webUrl = normalizePickerUrl(page.web.absoluteUrl), siteCollectionUrl = normalizePickerUrl(page.site.absoluteUrl);
  return useMemo(() => ({ owner: spfxContext, tenantId, userId, webUrl, siteCollectionUrl, spClient: client }), [spfxContext, tenantId, userId, webUrl, siteCollectionUrl, client]);
}

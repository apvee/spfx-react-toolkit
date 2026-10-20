import * as React from 'react';
import { useSPFxContext } from '../../hooks/useSPFxContext';
import { useSPFxPageContext } from '../../hooks/useSPFxPageContext';
import { useSPFxSPHttpClient } from '../../hooks/useSPFxSPHttpClient';
import { useAccessRequest } from '../../hooks/useAccessRequest.internal';
import { normalizeSharePointGroups, normalizeSharePointGroupWeb, readSharePointGroupAccess } from '../../services/spfx-sharepoint-groups.internal';
import { createAccessTrimBranch, readAccessTrimChildren, renderAccessTrim, validateTrimMatch } from '../access-trim-children.internal';
import type { SharePointGroupTrimProps, SharePointGroupTrimBranchProps, SharePointGroupSelector } from './sharepoint-group-trim.types';

const family = {};
/**
 * Mounts protected content after current-user direct membership matches all/any
 * selected SharePoint groups. Resolves selected-web user ID and every reference;
 * missing/ambiguous groups, malformed payloads and failed reads select ErrorFallback.
 * Does not expand Entra principals inside site groups. Requires the matching SPFx
 * provider/native SPHttpClient. Grants are advisory; server access still applies.
 * Requests share only while inflight within one provider; identity changes hide
 * former grants synchronously. No wrapper or inactive consumer mounting.
 * @param props - Nonempty ID/name selectors, optional web/match and structural children.
 * @returns Active branch/content in a React fragment; omitted content is empty.
 * @throws Missing matching SPFx provider.
 * @example
 * ```tsx
 * <SharePointGroupTrim groups={[{ name: 'Project Readers' }]}>
 *   <section>Project</section>
 *   <SharePointGroupTrim.Loading>Checking</SharePointGroupTrim.Loading>
 *   <SharePointGroupTrim.Fallback>No membership</SharePointGroupTrim.Fallback>
 *   <SharePointGroupTrim.ErrorFallback>Unavailable</SharePointGroupTrim.ErrorFallback>
 * </SharePointGroupTrim>
 * ```
 */
function GroupTrim(props: SharePointGroupTrimProps): React.ReactElement {
  const provider = useSPFxContext();
  const page = useSPFxPageContext();
  const { client } = useSPFxSPHttpClient();
  const children = readAccessTrimChildren(props.children, family);
  let error = children.error;
  let selectors: readonly SharePointGroupSelector[] = [];
  let webUrl = '';
  let match: 'all' | 'any' = 'all';
  try {
    selectors = normalizeSharePointGroups(props.groups);
    webUrl = normalizeSharePointGroupWeb(props.webUrl, page.web.absoluteUrl);
    match = validateTrimMatch(props.match);
    if (!client) throw new Error('SharePoint group client is unavailable.');
  } catch (failure) { error = failure instanceof Error ? failure : new Error(String(failure)); }
  const key = JSON.stringify([webUrl, selectors, match]);
  const state = useAccessRequest({
    provider,
    identity: [provider.spfxContext, client, page, page.user?.loginName, page.user?.email, page.aadInfo?.userId?.toString(), key],
    requestKey: `sp-groups:${key}`,
    request: client && !error ? () => readSharePointGroupAccess(client, webUrl, selectors, match) : undefined,
    error,
  });
  return renderAccessTrim(children, { granted: state.data === true, loading: state.isLoading, error: state.error });
}
/**
 * Structural pending-membership branch, unique and direct beneath SharePointGroupTrim.
 * @param props - Optional children only; omission renders empty.
 * @returns Active children without a wrapper.
 * @example <SharePointGroupTrim.Loading>Checking</SharePointGroupTrim.Loading>
 */
const Loading: React.ComponentType<SharePointGroupTrimBranchProps> = createAccessTrimBranch(family, 'loading');
/**
 * Structural denial branch after all selected references/reads succeed.
 * @param props - Optional children only; unique and direct beneath SharePointGroupTrim.
 * @returns Active children without a wrapper; omission renders empty.
 * @example <SharePointGroupTrim.Fallback>No membership</SharePointGroupTrim.Fallback>
 */
const Fallback: React.ComponentType<SharePointGroupTrimBranchProps> = createAccessTrimBranch(family, 'fallback');
/**
 * Structural malformed-input/composition or failed-read branch.
 * @param props - Optional children only; unique and direct beneath SharePointGroupTrim.
 * @returns Active children without a wrapper; omission renders empty.
 * @example <SharePointGroupTrim.ErrorFallback>Unavailable</SharePointGroupTrim.ErrorFallback>
 */
const ErrorFallback: React.ComponentType<SharePointGroupTrimBranchProps> = createAccessTrimBranch(family, 'error');
/**
 * Headless direct SharePoint-group trim using a matching SPFx provider/native
 * SPHttpClient. Resolves every selector and the selected web's current user ID;
 * Entra principals are not expanded. webUrl defaults to current web, match to all;
 * any accepts one group only after all reads succeed. Only protected or one unique
 * direct family branch mounts; omitted content is empty. No wrapper/styles/callbacks.
 * Requests deduplicate inflight per provider only; changed identity masks old grants.
 * Server authorization remains required.
 * @param props - Required nonempty group selectors, optional webUrl/match/children.
 * @returns Active content in a React fragment; failed reads select ErrorFallback.
 * @throws Missing matching SPFx provider.
 * @example
 * ```tsx
 * <SharePointGroupTrim groups={[{ name: 'Readers' }]}>
 *   Team content
 *   <SharePointGroupTrim.Fallback>No membership</SharePointGroupTrim.Fallback>
 * </SharePointGroupTrim>
 * ```
 */
export const SharePointGroupTrim = Object.assign(GroupTrim, {
  /** Pending membership branch, unique and direct beneath SharePointGroupTrim.
   * @param props - Optional children only; omission is empty.
   * @returns Active children without a wrapper.
   * @example <SharePointGroupTrim.Loading>Checking</SharePointGroupTrim.Loading>
   */
  Loading,
  /** Successful direct-membership denial branch; unique direct child of this family.
   * @param props - Optional children only; omission is empty.
   * @returns Active children without a wrapper.
   * @example <SharePointGroupTrim.Fallback>No membership</SharePointGroupTrim.Fallback>
   */
  Fallback,
  /** Invalid input/composition or failed-read branch; unique direct child of this family.
   * @param props - Optional children only; omission is empty.
   * @returns Active children without a wrapper.
   * @example <SharePointGroupTrim.ErrorFallback>Unavailable</SharePointGroupTrim.ErrorFallback>
   */
  ErrorFallback,
});

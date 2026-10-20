import * as React from 'react';
import { useSPFxContext } from '../../hooks/useSPFxContext';
import { useSPFxPageContext } from '../../hooks/useSPFxPageContext';
import { useSPFxMSGraphClient } from '../../hooks/useSPFxMSGraphClient';
import { useAccessRequest } from '../../hooks/useAccessRequest.internal';
import { normalizeEntraGroups, normalizeEntraGroupType, normalizeEntraRelationship, readEntraGroupAccess } from '../../services/spfx-entra-groups.internal';
import { createAccessTrimBranch, readAccessTrimChildren, renderAccessTrim, validateTrimMatch } from '../access-trim-children.internal';
import type { EntraGroupTrimProps, EntraGroupTrimBranchProps, EntraGroupSelector, EntraGroupType, EntraGroupRelationship } from './entra-group-trim.types';

const family = {};
/**
 * Mounts protected content after native Graph current-user membership/ownership
 * matches selected groups. Validates every selector and resolves names/metadata-required
 * IDs, applying groupType before named ambiguity. All-ID membership uses checkMemberGroups
 * in chunks of twenty without metadata; missing and nonmember IDs are indistinguishable. Owners read IDs across every owners page. Errors in any
 * selected reference/read select ErrorFallback even for any/memberOrOwner.
 * Requires a matching SPFx provider and sufficient Graph consent. Membership is
 * transitive (M365 direct); hidden membership omissions cannot prove real-world
 * absence. Server authorization remains required. Requests share inflight only
 * per provider; identity changes hide former grants during render.
 * @param props - Selectors, control-level type/relationship/match and structural children.
 * @returns Active content in a fragment with no DOM wrapper; omission is empty.
 * @throws Missing matching SPFx provider.
 * @example
 * ```tsx
 * <EntraGroupTrim groups={[{ name: 'Engineering' }]} groupType="M365" relationship="memberOrOwner">
 *   <section>Team</section>
 *   <EntraGroupTrim.Loading>Checking</EntraGroupTrim.Loading>
 *   <EntraGroupTrim.Fallback>No relationship</EntraGroupTrim.Fallback>
 *   <EntraGroupTrim.ErrorFallback>Unavailable</EntraGroupTrim.ErrorFallback>
 * </EntraGroupTrim>
 * ```
 */
function GroupTrim(props: EntraGroupTrimProps): React.ReactElement {
  const provider = useSPFxContext();
  const page = useSPFxPageContext();
  const graph = useSPFxMSGraphClient();
  const children = readAccessTrimChildren(props.children, family);
  let error = children.error ?? graph.initError;
  let selectors: readonly EntraGroupSelector[] = [];
  let type: EntraGroupType = 'All';
  let relationship: EntraGroupRelationship = 'member';
  let match: 'all' | 'any' = 'all';
  try {
    selectors = normalizeEntraGroups(props.groups);
    type = normalizeEntraGroupType(props.groupType);
    relationship = normalizeEntraRelationship(props.relationship);
    match = validateTrimMatch(props.match);
    if (!graph.client && !graph.isInitializing && !error) throw new Error('Entra group client is unavailable.');
  } catch (failure) { error = failure instanceof Error ? failure : new Error(String(failure)); }
  const key = JSON.stringify([selectors, type, relationship, match]);
  const client = graph.client;
  const state = useAccessRequest({
    provider,
    identity: [provider.spfxContext, client, page, page.user?.loginName, page.user?.email, page.aadInfo?.userId?.toString(), key],
    requestKey: `entra-groups:${key}`,
    request: client && !error ? () => readEntraGroupAccess(client, selectors, type, relationship, match) : undefined,
    error,
  });
  return renderAccessTrim(children, { granted: state.data === true, loading: graph.isInitializing || state.isLoading, error: state.error });
}
/**
 * Structural branch during Graph initialization or a pending relationship check.
 * @param props - Optional children only; unique and direct beneath EntraGroupTrim.
 * @returns Active children without a wrapper; omission renders empty.
 * @example <EntraGroupTrim.Loading>Checking</EntraGroupTrim.Loading>
 */
const Loading: React.ComponentType<EntraGroupTrimBranchProps> = createAccessTrimBranch(family, 'loading');
/**
 * Structural denial branch after every requested reference/read succeeds.
 * @param props - Optional children only; unique and direct beneath EntraGroupTrim.
 * @returns Active children without a wrapper; omission renders empty.
 * @example <EntraGroupTrim.Fallback>No relationship</EntraGroupTrim.Fallback>
 */
const Fallback: React.ComponentType<EntraGroupTrimBranchProps> = createAccessTrimBranch(family, 'fallback');
/**
 * Structural malformed-input/composition, initialization or failed-read branch.
 * @param props - Optional children only; unique and direct beneath EntraGroupTrim.
 * @returns Active children without a wrapper; omission renders empty.
 * @example <EntraGroupTrim.ErrorFallback>Unavailable</EntraGroupTrim.ErrorFallback>
 */
const ErrorFallback: React.ComponentType<EntraGroupTrimBranchProps> = createAccessTrimBranch(family, 'error');
/**
 * Headless Entra trim beneath a matching SPFx provider with native Graph consent.
 * Defaults to All/member/all. Control-level classification filters exact-name
 * candidates before ambiguity; All-ID members avoid metadata and use chunks of
 * twenty. Owners use every owners page and current /me ID. Errors in any requested
 * read fail the check even for any/memberOrOwner. Membership is transitive (M365
 * direct); hidden omissions cannot prove real-world absence. Only current active
 * content mounts, without wrappers/styles/callbacks; omitted content is empty.
 * Inflight sharing is provider-local; changed identity masks obsolete grants.
 * Server authorization remains required.
 * @param props - Required nonempty selectors and optional groupType/relationship/match/children.
 * @returns Active content in a React fragment; errors select ErrorFallback.
 * @throws Missing matching SPFx provider.
 * @example
 * ```tsx
 * <EntraGroupTrim groups={[{ name: 'Engineering' }]} groupType="M365">
 *   Team content
 *   <EntraGroupTrim.Fallback>No membership</EntraGroupTrim.Fallback>
 * </EntraGroupTrim>
 * ```
 */
export const EntraGroupTrim = Object.assign(GroupTrim, {
  /** Graph initialization or pending-check branch; unique direct child of EntraGroupTrim.
   * @param props - Optional children only; omission is empty.
   * @returns Active children without a wrapper.
   * @example <EntraGroupTrim.Loading>Checking</EntraGroupTrim.Loading>
   */
  Loading,
  /** Successful relationship-denial branch; unique direct child of this family.
   * @param props - Optional children only; omission is empty.
   * @returns Active children without a wrapper.
   * @example <EntraGroupTrim.Fallback>No relationship</EntraGroupTrim.Fallback>
   */
  Fallback,
  /** Invalid input/composition, initialization or failed-read branch; unique direct child.
   * @param props - Optional children only; omission is empty.
   * @returns Active children without a wrapper.
   * @example <EntraGroupTrim.ErrorFallback>Unavailable</EntraGroupTrim.ErrorFallback>
   */
  ErrorFallback,
});

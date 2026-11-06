/**
 * Advisory content branches selected by current-user Entra membership or ownership.
 *
 * @packageDocumentation
 * @docPage controls/entra-group-trim
 */
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
 * @returns Active content in a fragment with no DOM wrapper; when omitted, is empty.
 * @throws Missing matching SPFx provider.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { EntraGroupTrim } from '@apvee/spfx-react-toolkit';
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
 * @returns Active children without a wrapper; when omitted, renders empty.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { EntraGroupTrim } from '@apvee/spfx-react-toolkit';
 * <EntraGroupTrim groups={[{ name: 'Engineering' }]}>
 *   <EntraGroupTrim.Loading>Checking</EntraGroupTrim.Loading>
 * </EntraGroupTrim>
 * ```
 */
const Loading: React.ComponentType<EntraGroupTrimBranchProps> = createAccessTrimBranch(family, 'loading');
/**
 * Structural denial branch after every requested reference/read succeeds.
 * @param props - Optional children only; unique and direct beneath EntraGroupTrim.
 * @returns Active children without a wrapper; when omitted, renders empty.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { EntraGroupTrim } from '@apvee/spfx-react-toolkit';
 * <EntraGroupTrim groups={[{ name: 'Engineering' }]}>
 *   <EntraGroupTrim.Fallback>No relationship</EntraGroupTrim.Fallback>
 * </EntraGroupTrim>
 * ```
 */
const Fallback: React.ComponentType<EntraGroupTrimBranchProps> = createAccessTrimBranch(family, 'fallback');
/**
 * Structural malformed-input/composition, initialization or failed-read branch.
 * @param props - Optional children only; unique and direct beneath EntraGroupTrim.
 * @returns Active children without a wrapper; when omitted, renders empty.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { EntraGroupTrim } from '@apvee/spfx-react-toolkit';
 * <EntraGroupTrim groups={[{ name: 'Engineering' }]}>
 *   <EntraGroupTrim.ErrorFallback>Unavailable</EntraGroupTrim.ErrorFallback>
 * </EntraGroupTrim>
 * ```
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
 * @example Entra membership branches
 * ```tsx
 * import * as React from 'react';
 * import { EntraGroupTrim } from '@apvee/spfx-react-toolkit';
 *
 * <EntraGroupTrim groups={[{ name: 'Engineering' }]} groupType="M365">
 *   Team content
 *   <EntraGroupTrim.Fallback>No membership</EntraGroupTrim.Fallback>
 * </EntraGroupTrim>;
 * ```
 *
 * @scenario Stage Entra relationships and observe protected lifetimes
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Select, makeStyles } from '@fluentui/react-components';
 * import { EntraGroupTrim, type EntraGroupSelector, type EntraGroupType, type EntraGroupRelationship, useSPFxContext, useSPFxPageContext } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; children?: React.ReactNode; diagnostics?: React.ReactNode };
 * type Draft = { selector: 'name' | 'id'; values: string; groupType: EntraGroupType; relationship: EntraGroupRelationship; match: 'all' | 'any' };
 * type Attempt = Readonly<{ id: number; groups: readonly EntraGroupSelector[]; groupType: EntraGroupType; relationship: EntraGroupRelationship; match: 'all' | 'any' }>;
 * type Observation = { mounts: number; unmounts: number; active: boolean; counter: number; attached: boolean };
 * type Lifetime = { active: boolean; retired: boolean };
 * const defaults: Draft = { selector: 'name', values: 'Project Readers', groupType: 'All', relationship: 'member', match: 'all' };
 * const empty: Observation = { mounts: 0, unmounts: 0, active: false, counter: 0, attached: false };
 *
 * function ProtectedContent({ current, observe, className }: { current: () => boolean; observe: (event: 'mount' | 'unmount' | 'attach' | 'detach' | 'counter', lifetime: Lifetime, count?: number) => void; className: string }): React.ReactElement {
 *   const [count, setCount] = React.useState(0);
 *   const counter = React.useRef(0);
 *   const live = React.useRef(false);
 *   const lifetime = React.useRef<Lifetime>({ active: false, retired: false });
 *   const node = React.useRef<HTMLDivElement>();
 *   const callerRef = React.useCallback((element: HTMLDivElement | null) => {
 *     if (lifetime.current.retired) return;
 *     // React 17 ref detachment uses the legacy null sentinel; local absence is undefined.
 *     node.current = element ?? undefined;
 *     observe(element ? 'attach' : 'detach', lifetime.current);
 *   }, [observe]);
 *   React.useEffect(() => {
 *     live.current = true;
 *     observe('mount', lifetime.current);
 *     // The real ref is attached before passive cleanup of a previous attempt.
 *     if (node.current) observe('attach', lifetime.current);
 *     return () => { live.current = false; observe('unmount', lifetime.current); lifetime.current.retired = true; };
 *   }, [observe]);
 *   const increment = (): void => {
 *     if (!live.current || !current()) return;
 *     counter.current++; setCount(counter.current); observe('counter', lifetime.current, counter.current);
 *   };
 *   return <div ref={callerRef} className={className} data-trim-demo="protected"><strong>Allowed: protected content is mounted.</strong><Button data-trim-demo="increment" onClick={increment}>Protected counter: {count}</Button></div>;
 * }
 * function Branch({ level, state }: { level: 'outer' | 'inner'; state: 'loading' | 'denied' | 'error' }): React.ReactElement {
 *   return <p data-trim-demo="state" data-trim-level={level} data-trim-state={state}>{level}: {state === 'error' ? 'Error: check unavailable; inspect directory references, consent, payloads and transport.' : state === 'loading' ? 'Loading: checking current access.' : 'Denied: requested access was not observed.'}</p>;
 * }
 * const useStyles = makeStyles({ controls: { display: 'grid', gap: '12px' }, protected: { padding: '12px', border: '1px solid currentColor' }, diagnostics: { whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' } });
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const classes = useStyles();
 *   const owner = useSPFxContext(), page = useSPFxPageContext();
 *   // Observe actual public host identities only; the trims remain decision authority.
 *   const identity = [owner, owner.spfxContext, page, page.user.loginName, page.user.email, page.aadInfo?.userId?.toString(), page.web.absoluteUrl];
 *   const boundary = React.useRef({ revision: resetRevision, identity, epoch: 0, runEpoch: 0, attempt: 0, mounted: true });
 *   const [draft, setDraft] = React.useState<Draft>(defaults);
 *   const [applied, setApplied] = React.useState<Attempt>();
 *   const [observation, setObservation] = React.useState<Observation>(empty);
 *   if (boundary.current.revision !== resetRevision || !identity.every((value, index) => Object.is(value, boundary.current.identity[index]))) {
 *     const reset = boundary.current.revision !== resetRevision;
 *     boundary.current = { revision: resetRevision, identity, epoch: boundary.current.epoch + (reset ? 1 : 0), runEpoch: boundary.current.runEpoch + 1, attempt: reset ? 0 : boundary.current.attempt, mounted: true };
 *     if (reset) { setDraft(defaults); setApplied(undefined); setObservation(empty); }
 *   }
 *   const epoch = boundary.current.epoch, runEpoch = boundary.current.runEpoch, attempt = boundary.current.attempt;
 *   React.useEffect(() => { boundary.current.mounted = true; return () => { boundary.current.mounted = false; boundary.current.epoch++; }; }, []);
 *   const current = React.useCallback((): boolean => boundary.current.mounted && boundary.current.epoch === epoch && boundary.current.attempt === applied?.id, [epoch, applied]);
 *   const observe = React.useCallback((event: 'mount' | 'unmount' | 'attach' | 'detach' | 'counter', lifetime: Lifetime, count = 0): void => {
 *     if (!boundary.current.mounted || boundary.current.epoch !== epoch || lifetime.retired) return;
 *     if (event === 'unmount') {
 *       if (!lifetime.active) return;
 *       lifetime.active = false;
 *       setObservation(value => ({ ...value, unmounts: value.unmounts + 1, active: false, attached: false }));
 *     } else if (current()) {
 *       if (event === 'mount') { if (lifetime.active) return; lifetime.active = true; setObservation(value => ({ ...value, mounts: value.mounts + 1, active: true, counter: 0 })); }
 *       else if (event === 'counter' && lifetime.active) setObservation(value => ({ ...value, counter: count }));
 *       else if (event === 'attach' || event === 'detach') setObservation(value => ({ ...value, attached: event === 'attach' }));
 *     }
 *   }, [epoch, current]);
 *   const run = (): void => {
 *     if (!boundary.current.mounted || boundary.current.epoch !== epoch || boundary.current.runEpoch !== runEpoch || boundary.current.attempt !== attempt) return;
 *     // Map every entry, including empty pieces, so the public validator owns errors.
 *     const groups: readonly EntraGroupSelector[] = Object.freeze(draft.values.split(';').map(value => Object.freeze(draft.selector === 'id' ? { id: value.trim() } : { name: value.trim() })));
 *     const id = boundary.current.attempt + 1;
 *     boundary.current.attempt = id;
 *     setApplied(Object.freeze({ id, groups, groupType: draft.groupType, relationship: draft.relationship, match: draft.match }));
 *   };
 *   const controls = <div className={classes.controls}>
 *     <Field label="Group selector"><Select aria-label="Group selector" value={draft.selector} onChange={(_, data) => { const selector = data.value; if (selector === 'name' || selector === 'id') setDraft({ ...draft, selector }); }}><option value="name">Exact names</option><option value="id">GUIDs</option></Select></Field>
 *     <Field label="Group values (semicolon-separated)"><Input aria-label="Group values" value={draft.values} onChange={(_, data) => setDraft({ ...draft, values: data.value })} /></Field>
 *     <Field label="Group type"><Select aria-label="Group type" value={draft.groupType} onChange={(_, data) => { const groupType = data.value; if (groupType === 'All' || groupType === 'M365' || groupType === 'Security') setDraft({ ...draft, groupType }); }}><option>All</option><option>M365</option><option>Security</option></Select></Field>
 *     <Field label="Relationship"><Select aria-label="Relationship" value={draft.relationship} onChange={(_, data) => { const relationship = data.value; if (relationship === 'member' || relationship === 'owner' || relationship === 'memberOrOwner') setDraft({ ...draft, relationship }); }}><option>member</option><option>owner</option><option>memberOrOwner</option></Select></Field>
 *     <Field label="Match"><Select aria-label="Match" value={draft.match} onChange={(_, data) => { const match = data.value; if (match === 'all' || match === 'any') setDraft({ ...draft, match }); }}><option value="all">all</option><option value="any">any</option></Select></Field>
 *   </div>;
 *   const protectedContent = <ProtectedContent current={current} observe={observe} className={classes.protected} />;
 *   const regions: DemoRegions = {
 *     controls,
 *     commands: <div><Button data-trim-demo="run" onClick={run}>Run</Button></div>,
 *     status: <><p>Requires a matching SPFx provider and native Graph client. The sample requests User.Read and GroupMember.Read.All; tenant approval is separate. Security membership is transitive; M365 membership is direct. Hidden membership omissions do not prove absence. All + ID + member bypasses metadata: an unreturned GUID is Denied and cannot distinguish a missing group from nonmembership. Exact names resolve every page with type filtering before ambiguity; typed IDs require metadata. Ownership uses the Graph /me ID and every direct owners page. Synchronized groups, non-Unified mail-enabled groups and distribution groups do not support ownership; missing/ambiguous references, type mismatches and failed reads select Error even with any/memberOrOwner. Continuations support graph.microsoft.com only, not national-cloud hosts. Server authorization still governs protected operations. Host identity changes require a rerender with fresh context; Run repeats a settled check, without a settled cache or transport cancellation promise.</p><p data-trim-demo="mount-gauge">Protected mounts: {observation.mounts}; unmounts: {observation.unmounts}; active: {String(observation.active)}; ref attached: {String(observation.attached)}</p></>,
 *     children: applied ? <EntraGroupTrim key={applied.id} groups={applied.groups} groupType={applied.groupType} relationship={applied.relationship} match={applied.match}>
 *       {protectedContent}
 *       <EntraGroupTrim.Loading><Branch level="outer" state="loading" /></EntraGroupTrim.Loading>
 *       <EntraGroupTrim.Fallback><Branch level="outer" state="denied" /></EntraGroupTrim.Fallback>
 *       <EntraGroupTrim.ErrorFallback><Branch level="outer" state="error" /></EntraGroupTrim.ErrorFallback>
 *     </EntraGroupTrim> : <p data-trim-demo="idle">Idle: stage a target, then Run.</p>,
 *     diagnostics: <pre className={classes.diagnostics} data-trim-demo="applied-target">{JSON.stringify({ attempt: applied?.id, groups: applied?.groups, groupType: applied?.groupType, relationship: applied?.relationship, match: applied?.match, ...observation }, undefined, 2)}</pre>,
 *   };
 *   return render ? render(regions) : <section>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</section>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Run staged Entra group checks, then observe protected content only after the selected membership or ownership relationship allows it; drafts and Reset do not start checks.' };
 * ```
 *
 * @useWhen Current-user Entra membership or ownership should select advisory content branches beneath an SPFx provider.
 * @avoidWhen Do not use a rendering decision as authorization for server operations.
 */
export const EntraGroupTrim = Object.assign(GroupTrim, {
  /**
   * Graph initialization or pending-check branch; unique direct child of EntraGroupTrim.
   * Accepts Optional children only; when omitted, is empty.
   * @returns Active children without a wrapper.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { EntraGroupTrim } from '@apvee/spfx-react-toolkit';
   * <EntraGroupTrim groups={[{ name: 'Engineering' }]}>
   *   <EntraGroupTrim.Loading>Checking</EntraGroupTrim.Loading>
   * </EntraGroupTrim>
   * ```
   */
  Loading,
  /**
   * Successful relationship-denial branch; unique direct child of this family.
   * Accepts Optional children only; when omitted, is empty.
   * @returns Active children without a wrapper.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { EntraGroupTrim } from '@apvee/spfx-react-toolkit';
   * <EntraGroupTrim groups={[{ name: 'Engineering' }]}>
   *   <EntraGroupTrim.Fallback>No relationship</EntraGroupTrim.Fallback>
   * </EntraGroupTrim>
   * ```
   */
  Fallback,
  /**
   * Invalid input/composition, initialization or failed-read branch; unique direct child.
   * Accepts Optional children only; when omitted, is empty.
   * @returns Active children without a wrapper.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { EntraGroupTrim } from '@apvee/spfx-react-toolkit';
   * <EntraGroupTrim groups={[{ name: 'Engineering' }]}>
   *   <EntraGroupTrim.ErrorFallback>Unavailable</EntraGroupTrim.ErrorFallback>
   * </EntraGroupTrim>
   * ```
   */
  ErrorFallback,
});

/**
 * Advisory content branches selected by direct SharePoint site-group membership.
 *
 * @packageDocumentation
 * @docPage controls/sharepoint-group-trim
 */
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
 * import * as React from 'react';
 * import { SharePointGroupTrim } from '@apvee/spfx-react-toolkit';
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
 * @param props - Optional children only; when omitted, renders empty.
 * @returns Active children without a wrapper.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SharePointGroupTrim } from '@apvee/spfx-react-toolkit';
 * <SharePointGroupTrim groups={[{ name: 'Project Readers' }]}>
 *   <SharePointGroupTrim.Loading>Checking</SharePointGroupTrim.Loading>
 * </SharePointGroupTrim>
 * ```
 */
const Loading: React.ComponentType<SharePointGroupTrimBranchProps> = createAccessTrimBranch(family, 'loading');
/**
 * Structural denial branch after all selected references/reads succeed.
 * @param props - Optional children only; unique and direct beneath SharePointGroupTrim.
 * @returns Active children without a wrapper; when omitted, renders empty.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SharePointGroupTrim } from '@apvee/spfx-react-toolkit';
 * <SharePointGroupTrim groups={[{ name: 'Project Readers' }]}>
 *   <SharePointGroupTrim.Fallback>No membership</SharePointGroupTrim.Fallback>
 * </SharePointGroupTrim>
 * ```
 */
const Fallback: React.ComponentType<SharePointGroupTrimBranchProps> = createAccessTrimBranch(family, 'fallback');
/**
 * Structural malformed-input/composition or failed-read branch.
 * @param props - Optional children only; unique and direct beneath SharePointGroupTrim.
 * @returns Active children without a wrapper; when omitted, renders empty.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SharePointGroupTrim } from '@apvee/spfx-react-toolkit';
 * <SharePointGroupTrim groups={[{ name: 'Project Readers' }]}>
 *   <SharePointGroupTrim.ErrorFallback>Unavailable</SharePointGroupTrim.ErrorFallback>
 * </SharePointGroupTrim>
 * ```
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
 * @example Direct SharePoint membership branches
 * ```tsx
 * import * as React from 'react';
 * import { SharePointGroupTrim } from '@apvee/spfx-react-toolkit';
 *
 * <SharePointGroupTrim groups={[{ name: 'Readers' }]}>
 *   Team content
 *   <SharePointGroupTrim.Fallback>No membership</SharePointGroupTrim.Fallback>
 * </SharePointGroupTrim>;
 * ```
 *
 * @scenario Stage direct SharePoint group membership and observe protected lifetimes
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Select, makeStyles } from '@fluentui/react-components';
 * import { SharePointGroupTrim, type SharePointGroupSelector, useSPFxContext, useSPFxPageContext, useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; children?: React.ReactNode; diagnostics?: React.ReactNode };
 * type Draft = { selector: 'name' | 'id'; values: string; webUrl: string; match: 'all' | 'any' };
 * type Attempt = Readonly<{ id: number; groups: readonly SharePointGroupSelector[]; webUrl?: string; match: 'all' | 'any' }>;
 * type Observation = { mounts: number; unmounts: number; active: boolean; counter: number; attached: boolean };
 * type Lifetime = { active: boolean; retired: boolean };
 * const defaults: Draft = { selector: 'name', values: 'Project Readers', webUrl: '', match: 'all' };
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
 *   return <p data-trim-demo="state" data-trim-level={level} data-trim-state={state}>{level}: {state === 'error' ? 'Error: check unavailable; inspect selected-web references, payloads and transport.' : state === 'loading' ? 'Loading: checking current access.' : 'Denied: requested access was not observed.'}</p>;
 * }
 * const useStyles = makeStyles({ controls: { display: 'grid', gap: '12px' }, protected: { padding: '12px', border: '1px solid currentColor' }, diagnostics: { whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' } });
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const classes = useStyles();
 *   const owner = useSPFxContext(), page = useSPFxPageContext(), { client } = useSPFxSPHttpClient();
 *   // Observe actual public host identities only; the trims remain decision authority.
 *   const identity = [owner, owner.spfxContext, client, page, page.user.loginName, page.user.email, page.aadInfo?.userId?.toString(), page.web.absoluteUrl];
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
 *     const webUrl = draft.webUrl.trim() || undefined;
 *     // Map every entry, including empty pieces, so the public validator owns errors.
 *     const groups: readonly SharePointGroupSelector[] = Object.freeze(draft.values.split(';').map(value => Object.freeze(draft.selector === 'id' ? { id: Number(value.trim()) } : { name: value.trim() })));
 *     const id = boundary.current.attempt + 1;
 *     boundary.current.attempt = id;
 *     setApplied(Object.freeze({ id, groups, webUrl, match: draft.match }));
 *   };
 *   const controls = <div className={classes.controls}>
 *     <Field label="Group selector"><Select aria-label="Group selector" value={draft.selector} onChange={(_, data) => { const selector = data.value; if (selector === 'name' || selector === 'id') setDraft({ ...draft, selector }); }}><option value="name">Exact names</option><option value="id">Numeric IDs</option></Select></Field>
 *     <Field label="Group values (semicolon-separated)"><Input aria-label="Group values" value={draft.values} onChange={(_, data) => setDraft({ ...draft, values: data.value })} /></Field>
 *     <Field label="Web URL (empty uses current web)"><Input aria-label="Web URL" value={draft.webUrl} onChange={(_, data) => setDraft({ ...draft, webUrl: data.value })} /></Field>
 *     <Field label="Match"><Select aria-label="Match" value={draft.match} onChange={(_, data) => { const match = data.value; if (match === 'all' || match === 'any') setDraft({ ...draft, match }); }}><option value="all">all</option><option value="any">any</option></Select></Field>
 *   </div>;
 *   const protectedContent = <ProtectedContent current={current} observe={observe} className={classes.protected} />;
 *   const regions: DemoRegions = {
 *     controls,
 *     commands: <div><Button data-trim-demo="run" onClick={run}>Run</Button></div>,
 *     status: <><p>Requires a matching SPFx provider and native SPHttpClient. Checks direct membership in the selected web using that web's own current-user ID and every selected group and user page. Entra principals are not expanded; membership does not prove resource permissions. Server authorization still governs protected operations. Host identity changes require a rerender with fresh context; Run repeats a settled check, without a settled cache or transport cancellation promise.</p><p data-trim-demo="mount-gauge">Protected mounts: {observation.mounts}; unmounts: {observation.unmounts}; active: {String(observation.active)}; ref attached: {String(observation.attached)}</p></>,
 *     children: applied ? <SharePointGroupTrim key={applied.id} groups={applied.groups} webUrl={applied.webUrl} match={applied.match}>
 *       {protectedContent}
 *       <SharePointGroupTrim.Loading><Branch level="outer" state="loading" /></SharePointGroupTrim.Loading>
 *       <SharePointGroupTrim.Fallback><Branch level="outer" state="denied" /></SharePointGroupTrim.Fallback>
 *       <SharePointGroupTrim.ErrorFallback><Branch level="outer" state="error" /></SharePointGroupTrim.ErrorFallback>
 *     </SharePointGroupTrim> : <p data-trim-demo="idle">Idle: stage a target, then Run.</p>,
 *     diagnostics: <pre className={classes.diagnostics} data-trim-demo="applied-target">{JSON.stringify({ attempt: applied?.id, groups: applied?.groups, webUrl: applied?.webUrl, match: applied?.match, ...observation }, undefined, 2)}</pre>,
 *   };
 *   return render ? render(regions) : <section>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</section>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Run staged direct SharePoint group checks for the selected web, then observe protected content only when the current membership decision allows it; drafts and Reset do not start checks.' };
 * ```
 *
 * @useWhen Direct membership in specific SharePoint site groups should select advisory content branches.
 * @avoidWhen Use EntraGroupTrim for directory groups; SharePoint groups do not expand contained Entra principals.
 */
export const SharePointGroupTrim = Object.assign(GroupTrim, {
  /**
   * Pending membership branch, unique and direct beneath SharePointGroupTrim.
   * Accepts Optional children only; when omitted, is empty.
   * @returns Active children without a wrapper.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { SharePointGroupTrim } from '@apvee/spfx-react-toolkit';
   * <SharePointGroupTrim groups={[{ name: 'Project Readers' }]}>
   *   <SharePointGroupTrim.Loading>Checking</SharePointGroupTrim.Loading>
   * </SharePointGroupTrim>
   * ```
   */
  Loading,
  /**
   * Successful direct-membership denial branch; unique direct child of this family.
   * Accepts Optional children only; when omitted, is empty.
   * @returns Active children without a wrapper.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { SharePointGroupTrim } from '@apvee/spfx-react-toolkit';
   * <SharePointGroupTrim groups={[{ name: 'Project Readers' }]}>
   *   <SharePointGroupTrim.Fallback>No membership</SharePointGroupTrim.Fallback>
   * </SharePointGroupTrim>
   * ```
   */
  Fallback,
  /**
   * Invalid input/composition or failed-read branch; unique direct child of this family.
   * Accepts Optional children only; when omitted, is empty.
   * @returns Active children without a wrapper.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { SharePointGroupTrim } from '@apvee/spfx-react-toolkit';
   * <SharePointGroupTrim groups={[{ name: 'Project Readers' }]}>
   *   <SharePointGroupTrim.ErrorFallback>Unavailable</SharePointGroupTrim.ErrorFallback>
   * </SharePointGroupTrim>
   * ```
   */
  ErrorFallback,
});

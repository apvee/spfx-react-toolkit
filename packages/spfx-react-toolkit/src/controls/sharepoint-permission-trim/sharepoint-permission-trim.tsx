/**
 * Advisory content branches selected by effective SharePoint resource permissions.
 *
 * @packageDocumentation
 * @docPage controls/sharepoint-permission-trim
 */
import * as React from 'react';
import type { SPPermission } from '@microsoft/sp-page-context';
import { useSPFxCrossSitePermissions } from '../../hooks/useSPFxCrossSitePermissions';
import type { SPFxScopedPermissionsOptions } from '../../services/spfx-permission-target.types';
import { createAccessTrimBranch, readAccessTrimChildren, renderAccessTrim, validateTrimMatch } from '../access-trim-children.internal';
import type { SharePointPermissionTrimProps, SharePointPermissionTrimBranchProps } from './sharepoint-permission-trim.types';

const family = {};
function isRequestedPermission(permission: unknown): permission is SPPermission {
  if (!permission || typeof permission !== 'object') return false;
  const value = (permission as { readonly value?: unknown }).value;
  if (!value || typeof value !== 'object') return false;
  const mask = value as { readonly High?: unknown; readonly Low?: unknown };
  return [mask.High, mask.Low].every(part => typeof part === 'number' && Number.isInteger(part) && part >= 0 && part <= 4294967295);
}
/**
 * Mounts protected children only when current effective permissions match. Uses
 * useSPFxCrossSitePermissions for the selected resource; a current native web
 * mask avoids HTTP. Denials, loading and validation/transport errors select the
 * respective structural branch. Identity changes hide previous grants during
 * render. This is advisory presentation: server operations still enforce access.
 * @param props - Scoped target, nonempty permissions, match and structural children.
 * @returns Active content in a React fragment, without a DOM wrapper.
 * @throws Missing matching SPFx provider.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SharePointPermissionTrim } from '@apvee/spfx-react-toolkit';
 * import { SPPermission } from '@microsoft/sp-page-context';
 * <SharePointPermissionTrim permissions={[SPPermission.addListItems]} scope="list" listTitle="Tasks">
 *   <button>Create task</button>
 *   <SharePointPermissionTrim.Loading>Checking</SharePointPermissionTrim.Loading>
 *   <SharePointPermissionTrim.Fallback>Read only</SharePointPermissionTrim.Fallback>
 *   <SharePointPermissionTrim.ErrorFallback>Check unavailable</SharePointPermissionTrim.ErrorFallback>
 * </SharePointPermissionTrim>
 * ```
 */
function PermissionTrim(props: SharePointPermissionTrimProps): React.ReactElement {
  const children = readAccessTrimChildren(props.children, family);
  let validationError = children.error;
  let match: 'all' | 'any' = 'all';
  try {
    match = validateTrimMatch(props.match);
    if (!Array.isArray(props.permissions) || props.permissions.length === 0 ||
        !Array.from(props.permissions).every(isRequestedPermission)) {
      throw new Error('SharePointPermissionTrim requires a nonempty array of SPPermission values.');
    }
  } catch (failure) {
    validationError = failure instanceof Error ? failure : new Error(String(failure));
  }
  // A deliberately invalid scoped target disables all permission transport while
  // retaining an unconditional hook call. It cannot select any native mask.
  const disabledTarget = { scope: 'invalid' } as unknown as SPFxScopedPermissionsOptions;
  const info = useSPFxCrossSitePermissions(validationError ? disabledTarget : props);
  const granted = !validationError && !info.error && info.permissions !== undefined &&
    (match === 'all' ? props.permissions.every(info.hasPermission) : props.permissions.some(info.hasPermission));
  return renderAccessTrim(children, { granted, loading: info.isLoading, error: validationError ?? info.error });
}
/**
 * Structural loading branch, shown while the selected permission check is pending.
 * @param props - Optional children only; must be a unique direct child of SharePointPermissionTrim.
 * @returns Its active children without a DOM wrapper; omitted children render empty.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SharePointPermissionTrim } from '@apvee/spfx-react-toolkit';
 * import { SPPermission } from '@microsoft/sp-page-context';
 * <SharePointPermissionTrim permissions={[SPPermission.addListItems]}>
 *   <SharePointPermissionTrim.Loading>Checking</SharePointPermissionTrim.Loading>
 * </SharePointPermissionTrim>
 * ```
 */
const Loading: React.ComponentType<SharePointPermissionTrimBranchProps> = createAccessTrimBranch(family, 'loading');
/**
 * Structural denial branch, shown after a successful check without a matching grant.
 * @param props - Optional children only; must be a unique direct child of SharePointPermissionTrim.
 * @returns Its active children without a DOM wrapper; omitted children render empty.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SharePointPermissionTrim } from '@apvee/spfx-react-toolkit';
 * import { SPPermission } from '@microsoft/sp-page-context';
 * <SharePointPermissionTrim permissions={[SPPermission.addListItems]}>
 *   <SharePointPermissionTrim.Fallback>Read only</SharePointPermissionTrim.Fallback>
 * </SharePointPermissionTrim>
 * ```
 */
const Fallback: React.ComponentType<SharePointPermissionTrimBranchProps> = createAccessTrimBranch(family, 'fallback');
/**
 * Structural error branch for malformed input/composition and failed permission reads.
 * @param props - Optional children only; must be a unique direct child of SharePointPermissionTrim.
 * @returns Its active children without a DOM wrapper; omitted children render empty.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SharePointPermissionTrim } from '@apvee/spfx-react-toolkit';
 * import { SPPermission } from '@microsoft/sp-page-context';
 * <SharePointPermissionTrim permissions={[SPPermission.addListItems]}>
 *   <SharePointPermissionTrim.ErrorFallback>Unavailable</SharePointPermissionTrim.ErrorFallback>
 * </SharePointPermissionTrim>
 * ```
 */
const ErrorFallback: React.ComponentType<SharePointPermissionTrimBranchProps> = createAccessTrimBranch(family, 'error');
/**
 * Headless current-user effective permission trim beneath a matching SPFx provider.
 * Uses the scoped permissions hook: web is default, list/item/folder require their
 * selectors, all is default and any accepts one requested native mask. Only the
 * current positive decision mounts protected children. Loading/denial/error select
 * their unique direct family branches; omitted active content is empty. No DOM
 * wrapper, styling or callbacks. Server authorization remains required.
 * @param props - Scoped target, required nonempty permissions and optional match/children.
 * @returns Active content in a React fragment. Identity changes hide obsolete grants.
 * @throws Missing matching SPFx provider.
 * @example Current web permission branches
 * ```tsx
 * import * as React from 'react';
 * import { SharePointPermissionTrim } from '@apvee/spfx-react-toolkit';
 * import { SPPermission } from '@microsoft/sp-page-context';
 * <SharePointPermissionTrim permissions={[SPPermission.addListItems]}>
 *   <button>Create</button>
 *   <SharePointPermissionTrim.Fallback>Read only</SharePointPermissionTrim.Fallback>
 * </SharePointPermissionTrim>
 * ```
 *
 * @example Stage resource permissions and observe protected lifetimes
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Select, makeStyles } from '@fluentui/react-components';
 * import { SharePointPermissionTrim, type SPFxScopedPermissionsOptions, useSPFxContext, useSPFxPageContext, useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
 * import { SPPermission } from '@microsoft/sp-page-context';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; children?: React.ReactNode; diagnostics?: React.ReactNode };
 * type Draft = { scope: 'web' | 'list' | 'item' | 'folder'; webUrl: string; selector: 'title' | 'id' | 'path'; list: string; itemId: string; folderPath: string; permission: 'view' | 'edit' | 'both'; match: 'all' | 'any';  };
 * type Attempt = Readonly<{ id: number; target: SPFxScopedPermissionsOptions; permissions: readonly SPPermission[]; match: 'all' | 'any' }>;
 * type Observation = { mounts: number; unmounts: number; active: boolean; counter: number; attached: boolean };
 * type Lifetime = { active: boolean; retired: boolean };
 * const defaults: Draft = { scope: 'web', webUrl: '', selector: 'title', list: 'Tasks', itemId: '1', folderPath: '/sites/team/Shared Documents/Reports', permission: 'view', match: 'all',  };
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
 *   return <p data-trim-demo="state" data-trim-level={level} data-trim-state={state}>{level}: {state === 'error' ? 'Error: check unavailable; inspect selectors, consent and transport.' : state === 'loading' ? 'Loading: checking current access.' : 'Denied: requested access was not observed.'}</p>;
 * }
 * const useStyles = makeStyles({ controls: { display: 'grid', gap: '12px' }, protected: { padding: '12px', border: '1px solid currentColor' }, diagnostics: { whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' } });
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const classes = useStyles();
 *   const owner = useSPFxContext(), page = useSPFxPageContext(), { client } = useSPFxSPHttpClient();
 *   // Observe actual public host identities only; the trims remain decision authority.
 *   const identity = [owner, owner.spfxContext, client, page, page.user.loginName, page.user.email, page.web.absoluteUrl];
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
 *     let target: SPFxScopedPermissionsOptions;
 *     if (draft.scope === 'web') target = { scope: 'web', webUrl };
 *     else {
 *       const selector = draft.selector === 'id' ? { listId: draft.list } : draft.selector === 'path' ? { listPath: draft.list } : { listTitle: draft.list };
 *       if (draft.scope === 'item') target = { ...selector, scope: 'item', webUrl, itemId: Number(draft.itemId) };
 *       else if (draft.scope === 'folder') target = { ...selector, scope: 'folder', webUrl, folderPath: draft.folderPath };
 *       else target = { ...selector, scope: 'list', webUrl };
 *     }
 *     const permissions = draft.permission === 'both' ? [SPPermission.viewListItems, SPPermission.editListItems] : [draft.permission === 'edit' ? SPPermission.editListItems : SPPermission.viewListItems];
 *     const id = boundary.current.attempt + 1;
 *     boundary.current.attempt = id;
 *     setApplied(Object.freeze({ id, target: Object.freeze(target), permissions: Object.freeze(permissions), match: draft.match }));
 *   };
 *   const controls = <div className={classes.controls}>
 *     <Field label="Resource scope"><Select aria-label="Resource scope" value={draft.scope} onChange={(_, data) => { const scope = data.value; if (scope === 'web' || scope === 'list' || scope === 'item' || scope === 'folder') setDraft({ ...draft, scope }); }}>{['web', 'list', 'item', 'folder'].map(value => <option key={value}>{value}</option>)}</Select></Field>
 *     <Field label="Web URL (empty uses current web)"><Input aria-label="Web URL" value={draft.webUrl} onChange={(_, data) => setDraft({ ...draft, webUrl: data.value })} /></Field>
 *     {draft.scope !== 'web' && <><Field label="List selector"><Select aria-label="List selector" value={draft.selector} onChange={(_, data) => { const selector = data.value; if (selector === 'title' || selector === 'id' || selector === 'path') setDraft({ ...draft, selector }); }}><option value="title">Exact title</option><option value="id">GUID</option><option value="path">Decoded web-relative path</option></Select></Field><Field label="List value"><Input aria-label="List value" value={draft.list} onChange={(_, data) => setDraft({ ...draft, list: data.value })} /></Field></>}
 *     {draft.scope === 'item' && <Field label="Item ID"><Input aria-label="Item ID" value={draft.itemId} onChange={(_, data) => setDraft({ ...draft, itemId: data.value })} /></Field>}
 *     {draft.scope === 'folder' && <Field label="Decoded server-relative folder path"><Input aria-label="Folder path" value={draft.folderPath} onChange={(_, data) => setDraft({ ...draft, folderPath: data.value })} /></Field>}
 *     <Field label="Permissions"><Select aria-label="Permissions" value={draft.permission} onChange={(_, data) => { const permission = data.value; if (permission === 'view' || permission === 'edit' || permission === 'both') setDraft({ ...draft, permission }); }}><option value="view">viewListItems</option><option value="edit">editListItems</option><option value="both">Both</option></Select></Field>
 *     <Field label="Match"><Select aria-label="Match" value={draft.match} onChange={(_, data) => { const match = data.value; if (match === 'all' || match === 'any') setDraft({ ...draft, match }); }}><option value="all">all</option><option value="any">any</option></Select></Field>
 *
 *   </div>;
 *   const protectedContent = <ProtectedContent current={current} observe={observe} className={classes.protected} />;
 *   const regions: DemoRegions = {
 *     controls,
 *     commands: <div><Button data-trim-demo="run" onClick={run}>Run</Button></div>,
 *     status: <><p>Requires a matching SPFx provider. Server authorization still governs every protected operation. Current native PageContext masks refresh only when the host supplies fresh context and rerenders; a current mask can decide synchronously without HTTP or Loading.</p><p data-trim-demo="mount-gauge">Protected mounts: {observation.mounts}; unmounts: {observation.unmounts}; active: {String(observation.active)}; ref attached: {String(observation.attached)}</p></>,
 *     children: applied ? <SharePointPermissionTrim key={applied.id} {...applied.target} permissions={applied.permissions} match={applied.match}>
 *       {protectedContent}
 *       <SharePointPermissionTrim.Loading><Branch level="outer" state="loading" /></SharePointPermissionTrim.Loading>
 *       <SharePointPermissionTrim.Fallback><Branch level="outer" state="denied" /></SharePointPermissionTrim.Fallback>
 *       <SharePointPermissionTrim.ErrorFallback><Branch level="outer" state="error" /></SharePointPermissionTrim.ErrorFallback>
 *     </SharePointPermissionTrim> : <p data-trim-demo="idle">Idle: stage a target, then Run.</p>,
 *     diagnostics: <pre className={classes.diagnostics} data-trim-demo="applied-target">{JSON.stringify({ attempt: applied?.id, target: applied?.target, match: applied?.match, permissions: applied?.permissions.map(value => value.value), ...observation }, undefined, 2)}</pre>,
 *   };
 *   return render ? render(regions) : <section>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</section>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Run a staged resource permission check and observe that protected content mounts only after a real grant; draft edits preserve the applied check and Reset returns to idle.' };
 * ```
 *
 * @example Compose permission and Entra checks with nested AND
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Select, makeStyles } from '@fluentui/react-components';
 * import { SharePointPermissionTrim, EntraGroupTrim, type SPFxScopedPermissionsOptions, useSPFxContext, useSPFxPageContext, useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
 * import { SPPermission } from '@microsoft/sp-page-context';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; children?: React.ReactNode; diagnostics?: React.ReactNode };
 * type Draft = { scope: 'web' | 'list' | 'item' | 'folder'; webUrl: string; selector: 'title' | 'id' | 'path'; list: string; itemId: string; folderPath: string; permission: 'view' | 'edit' | 'both'; match: 'all' | 'any'; group: string; };
 * type Attempt = Readonly<{ id: number; target: SPFxScopedPermissionsOptions; permissions: readonly SPPermission[]; match: 'all' | 'any'; group: string }>;
 * type Observation = { mounts: number; unmounts: number; active: boolean; counter: number; attached: boolean };
 * type Lifetime = { active: boolean; retired: boolean };
 * const defaults: Draft = { scope: 'web', webUrl: '', selector: 'title', list: 'Tasks', itemId: '1', folderPath: '/sites/team/Shared Documents/Reports', permission: 'view', match: 'all', group: 'Project Readers', };
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
 *   return <p data-trim-demo="state" data-trim-level={level} data-trim-state={state}>{level}: {state === 'error' ? 'Error: check unavailable; inspect selectors, consent and transport.' : state === 'loading' ? 'Loading: checking current access.' : 'Denied: requested access was not observed.'}</p>;
 * }
 * const useStyles = makeStyles({ controls: { display: 'grid', gap: '12px' }, protected: { padding: '12px', border: '1px solid currentColor' }, diagnostics: { whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' } });
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const classes = useStyles();
 *   const owner = useSPFxContext(), page = useSPFxPageContext(), { client } = useSPFxSPHttpClient();
 *   // Observe actual public host identities only; the trims remain decision authority.
 *   const identity = [owner, owner.spfxContext, client, page, page.user.loginName, page.user.email, page.web.absoluteUrl];
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
 *     let target: SPFxScopedPermissionsOptions;
 *     if (draft.scope === 'web') target = { scope: 'web', webUrl };
 *     else {
 *       const selector = draft.selector === 'id' ? { listId: draft.list } : draft.selector === 'path' ? { listPath: draft.list } : { listTitle: draft.list };
 *       if (draft.scope === 'item') target = { ...selector, scope: 'item', webUrl, itemId: Number(draft.itemId) };
 *       else if (draft.scope === 'folder') target = { ...selector, scope: 'folder', webUrl, folderPath: draft.folderPath };
 *       else target = { ...selector, scope: 'list', webUrl };
 *     }
 *     const permissions = draft.permission === 'both' ? [SPPermission.viewListItems, SPPermission.editListItems] : [draft.permission === 'edit' ? SPPermission.editListItems : SPPermission.viewListItems];
 *     const id = boundary.current.attempt + 1;
 *     boundary.current.attempt = id;
 *     setApplied(Object.freeze({ id, target: Object.freeze(target), permissions: Object.freeze(permissions), match: draft.match, group: draft.group }));
 *   };
 *   const controls = <div className={classes.controls}>
 *     <Field label="Resource scope"><Select aria-label="Resource scope" value={draft.scope} onChange={(_, data) => { const scope = data.value; if (scope === 'web' || scope === 'list' || scope === 'item' || scope === 'folder') setDraft({ ...draft, scope }); }}>{['web', 'list', 'item', 'folder'].map(value => <option key={value}>{value}</option>)}</Select></Field>
 *     <Field label="Web URL (empty uses current web)"><Input aria-label="Web URL" value={draft.webUrl} onChange={(_, data) => setDraft({ ...draft, webUrl: data.value })} /></Field>
 *     {draft.scope !== 'web' && <><Field label="List selector"><Select aria-label="List selector" value={draft.selector} onChange={(_, data) => { const selector = data.value; if (selector === 'title' || selector === 'id' || selector === 'path') setDraft({ ...draft, selector }); }}><option value="title">Exact title</option><option value="id">GUID</option><option value="path">Decoded web-relative path</option></Select></Field><Field label="List value"><Input aria-label="List value" value={draft.list} onChange={(_, data) => setDraft({ ...draft, list: data.value })} /></Field></>}
 *     {draft.scope === 'item' && <Field label="Item ID"><Input aria-label="Item ID" value={draft.itemId} onChange={(_, data) => setDraft({ ...draft, itemId: data.value })} /></Field>}
 *     {draft.scope === 'folder' && <Field label="Decoded server-relative folder path"><Input aria-label="Folder path" value={draft.folderPath} onChange={(_, data) => setDraft({ ...draft, folderPath: data.value })} /></Field>}
 *     <Field label="Permissions"><Select aria-label="Permissions" value={draft.permission} onChange={(_, data) => { const permission = data.value; if (permission === 'view' || permission === 'edit' || permission === 'both') setDraft({ ...draft, permission }); }}><option value="view">viewListItems</option><option value="edit">editListItems</option><option value="both">Both</option></Select></Field>
 *     <Field label="Match"><Select aria-label="Match" value={draft.match} onChange={(_, data) => { const match = data.value; if (match === 'all' || match === 'any') setDraft({ ...draft, match }); }}><option value="all">all</option><option value="any">any</option></Select></Field>
 *     <Field label="Nested Entra group name"><Input aria-label="Nested Entra group name" value={draft.group} onChange={(_, data) => setDraft({ ...draft, group: data.value })} /></Field>
 *   </div>;
 *   const protectedContent = <ProtectedContent current={current} observe={observe} className={classes.protected} />;
 *   const regions: DemoRegions = {
 *     controls,
 *     commands: <div><Button data-trim-demo="run" onClick={run}>Run</Button></div>,
 *     status: <><p>Requires a matching SPFx provider. Server authorization still governs every protected operation. Current native PageContext masks refresh only when the host supplies fresh context and rerenders; a current mask can decide synchronously without HTTP or Loading.</p><p>The sample already requests Graph User.Read and GroupMember.Read.All. Consent and hidden membership limitations can prevent a result; no grant is requested here.</p><p data-trim-demo="mount-gauge">Protected mounts: {observation.mounts}; unmounts: {observation.unmounts}; active: {String(observation.active)}; ref attached: {String(observation.attached)}</p></>,
 *     children: applied ? <SharePointPermissionTrim key={applied.id} {...applied.target} permissions={applied.permissions} match={applied.match}>
 *       <EntraGroupTrim groups={[{ name: applied.group }]} groupType="All" relationship="member" match="all">
 *         {protectedContent}
 *         <EntraGroupTrim.Loading><Branch level="inner" state="loading" /></EntraGroupTrim.Loading>
 *         <EntraGroupTrim.Fallback><Branch level="inner" state="denied" /></EntraGroupTrim.Fallback>
 *         <EntraGroupTrim.ErrorFallback><Branch level="inner" state="error" /></EntraGroupTrim.ErrorFallback>
 *       </EntraGroupTrim>
 *       <SharePointPermissionTrim.Loading><Branch level="outer" state="loading" /></SharePointPermissionTrim.Loading>
 *       <SharePointPermissionTrim.Fallback><Branch level="outer" state="denied" /></SharePointPermissionTrim.Fallback>
 *       <SharePointPermissionTrim.ErrorFallback><Branch level="outer" state="error" /></SharePointPermissionTrim.ErrorFallback>
 *     </SharePointPermissionTrim> : <p data-trim-demo="idle">Idle: stage a target, then Run.</p>,
 *     diagnostics: <pre className={classes.diagnostics} data-trim-demo="applied-target">{JSON.stringify({ attempt: applied?.id, target: applied?.target, match: applied?.match, permissions: applied?.permissions.map(value => value.value), group: applied?.group, ...observation }, undefined, 2)}</pre>,
 *   };
 *   return render ? render(regions) : <section>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</section>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Run a staged permission check, then observe that the inner Entra check starts only after the outer grant and protected content requires both decisions.' };
 * ```
 *
 * @useWhen Effective current-user SharePoint permission masks should select resource-specific content branches.
 * @avoidWhen Keep server authorization on every protected operation; the trim only controls rendering.
 */
export const SharePointPermissionTrim = Object.assign(PermissionTrim, {
  /**
   * Pending permission branch, unique and direct beneath SharePointPermissionTrim.
   * Accepts Optional children only; when omitted, is empty.
   * @returns Active children without a wrapper.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { SharePointPermissionTrim } from '@apvee/spfx-react-toolkit';
   * import { SPPermission } from '@microsoft/sp-page-context';
   * <SharePointPermissionTrim permissions={[SPPermission.addListItems]}>
   *   <SharePointPermissionTrim.Loading>Checking</SharePointPermissionTrim.Loading>
   * </SharePointPermissionTrim>
   * ```
   */
  Loading,
  /**
   * Successful denial branch, unique and direct beneath SharePointPermissionTrim.
   * Accepts Optional children only; when omitted, is empty.
   * @returns Active children without a wrapper.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { SharePointPermissionTrim } from '@apvee/spfx-react-toolkit';
   * import { SPPermission } from '@microsoft/sp-page-context';
   * <SharePointPermissionTrim permissions={[SPPermission.addListItems]}>
   *   <SharePointPermissionTrim.Fallback>Read only</SharePointPermissionTrim.Fallback>
   * </SharePointPermissionTrim>
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
   * import { SharePointPermissionTrim } from '@apvee/spfx-react-toolkit';
   * import { SPPermission } from '@microsoft/sp-page-context';
   * <SharePointPermissionTrim permissions={[SPPermission.addListItems]}>
   *   <SharePointPermissionTrim.ErrorFallback>Unavailable</SharePointPermissionTrim.ErrorFallback>
   * </SharePointPermissionTrim>
   * ```
   */
  ErrorFallback,
});

/**
 * @packageDocumentation
 * @docPage controls/pickers
 */
import { PickerMessageError } from './picker-message.internal';
import * as React from 'react';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { useSPFxServiceScope } from '../../hooks/useSPFxServiceScope';
import { useSPFxMSGraphClient } from '../../hooks/useSPFxMSGraphClient';
import type { PrincipalPickerProps, PrincipalPickerOption } from './picker.types';
import type { PrincipalPickerValue } from './picker-values.types';
import { createPrincipalPickerSource } from './principal-picker.source.internal';
import { usePickerSourceContext } from './usePickerSourceContext.internal';
import { usePickerRuntime } from './usePickerRuntime.internal';
import { PickerRenderer } from './picker-renderer.internal';

interface PrincipalGraphAccess {
  readonly getClient: () => Promise<MSGraphClientV3>;
  publish(client: MSGraphClientV3 | undefined, initializing: boolean, error: Error | undefined): void;
  dispose(): void;
}

/**
 * Per-provider pending dependency; SharePoint-only operations never wait for Graph.
 */
function createPrincipalGraphAccess(): PrincipalGraphAccess {
  let current: { readonly client: MSGraphClientV3 } | { readonly error: Error } | undefined;
  const waiters = new Set<{ readonly resolve: (client: MSGraphClientV3) => void; readonly reject: (error: Error) => void }>();
  const settle = (): void => {
    if (current === undefined) return;
    for (const waiter of waiters) {
      if ('client' in current) waiter.resolve(current.client);
      else waiter.reject(current.error);
    }
    waiters.clear();
  };
  return {
    getClient() {
      if (current !== undefined) return 'client' in current ? Promise.resolve(current.client) : Promise.reject(current.error);
      return new Promise<MSGraphClientV3>((resolve, reject) => { waiters.add({ resolve, reject }); });
    },
    publish(client, initializing, error) {
      current = initializing ? undefined : client !== undefined ? { client } : { error: error ?? new PickerMessageError({ group: 'errors', key: 'graphUnavailable' }) };
      settle();
    },
    dispose() { current = { error: new Error('Principal Graph dependency was disposed.') }; settle(); },
  };
}

/**
 * Unified user/SharePoint/Entra group picker beneath an SPFx provider.
 * Defaults: TagPicker, multiple users, available IDs, no personal/current suggestions and no provisioning.
 * Delegated People.Read/User.Read and directory/membership permissions depend on enabled sources. Graph-dependent operations wait for native client initialization; terminal initialization errors are reported without blocking SharePoint-only operations.
 * @param props - Principal categories/scopes, independent consumer/seed models and explicit ID resolution policy.
 * @returns A native Fluent picker. ensureUser only runs on accepted user actions or explicit ID resolution.
 * @example Select recipients with explicit identity modes
 * ```tsx
 * import * as React from 'react';
 * import { PrincipalPicker } from '@apvee/spfx-react-toolkit';
 * <PrincipalPicker label="Destinatari" principalTypes={['user','microsoft365Group']} userIdentity={{mode:'both',ensureUser:false}} />
 * ```
 *
 * @useWhen A user must select people or configured SharePoint/Entra group categories with verified scoped identities.
 * @avoidWhen Use advisory access trims for membership decisions; a picker selection is not an authorization check.
 *
 * @example Explore principal discovery and explicit identity resolution
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Select, Switch, makeStyles } from '@fluentui/react-components';
 * import { PrincipalPicker, useSPFxContext, type PrincipalPickerOption, type PickerSelectionData, type PrincipalType } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Configuration = {
 *   presentation: 'combobox' | 'tagPicker'; multi: boolean; searchable: boolean;
 *   colorful: boolean; customStrings: boolean; customStatus: boolean; header: string;  webUrl: string; types: 'users' | 'groups' | 'all'; idMode: 'available' | 'entra' | 'sharePoint' | 'both'; ensureUser: boolean; personal: boolean; currentUser: boolean;
 * };
 * const baseline: Configuration = { presentation: 'combobox', multi: false, searchable: true, colorful: false, customStrings: false, customStatus: false, header: '', webUrl: '', types: 'users', idMode: 'available', ensureUser: false, personal: false, currentUser: false };
 * const useStyles = makeStyles({ controls: { display: 'grid', gap: '12px' }, commands: { display: 'flex', flexWrap: 'wrap', gap: '8px' }, diagnostics: { whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' } });
 *
 * // Mount beneath the existing SPFxWebPartProvider/WebPartShell, never another provider.
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const classes = useStyles();
 *   const { spfxContext } = useSPFxContext();
 *   const scope = spfxContext.serviceScope;
 *   const boundary = React.useRef({ resetRevision, owner: spfxContext, scope });
 *   const session = boundary.current;
 *   const [config, setConfig] = React.useState<Configuration>(baseline);
 *   const [started, setStarted] = React.useState(false);
 *   const [selection, setSelection] = React.useState<PickerSelectionData<PrincipalPickerOption> | undefined>();
 *   const [available, setAvailable] = React.useState<readonly PrincipalPickerOption[] | undefined>();
 *   const [error, setError] = React.useState('');
 *   const [refresh, setRefresh] = React.useState(0);
 *   const current = (): boolean => boundary.current === session;
 *   const active = started && session.resetRevision === resetRevision && session.owner === spfxContext && session.scope === scope;
 *   React.useLayoutEffect(() => {
 *     const observed = boundary.current;
 *     if (observed.resetRevision === resetRevision && observed.owner === spfxContext && observed.scope === scope) return;
 *     boundary.current = { resetRevision, owner: spfxContext, scope };
 *     setStarted(false); setConfig(baseline); setSelection(undefined); setAvailable(undefined); setError(''); setRefresh(0);
 *   }, [resetRevision, spfxContext, scope]);
 *   const update = (patch: Partial<Configuration>): void => {
 *     if (!current()) return;
 *     const next = { ...config, ...patch };
 *     const normalizedWeb = (draft: string): string => {
 *       const web = draft.trim() || spfxContext.pageContext.web.absoluteUrl;
 *       try { return new URL(web).href.replace(/\/$/, '').toLowerCase(); }
 *       catch { return web; } // Keep unarmed partial URL drafts editable; the picker validates its actual source.
 *     };
 *     // The actual Principal source serializes its identity policy verbatim.
 *     // SharePoint/both target spelling is therefore part of that binding; no extra normalization is promised.
 *     const sourceKey = (value: Configuration): string => JSON.stringify([
 *       normalizedWeb(value.webUrl), value.types, value.idMode === 'sharePoint' || value.idMode === 'both'
 *         ? { mode: value.idMode, webUrl: value.webUrl.trim() || undefined, ensureUser: value.ensureUser }
 *         : { mode: value.idMode }, value.personal, value.currentUser,
 *     ]);
 *     // Retire old commands synchronously without remounting the operational picker.
 *     if (sourceKey(config) !== sourceKey(next)) boundary.current = { ...session };
 *     setConfig(previous => ({ ...previous, ...patch }));
 *   };
 *   const target = config.webUrl.trim() || undefined;
 *   const principalTypes: readonly PrincipalType[] = config.types === 'users' ? ['user'] : config.types === 'groups'
 *     ? ['sharePointGroup', 'microsoft365Group', 'securityGroup', 'distributionList']
 *     : ['user', 'sharePointGroup', 'microsoft365Group', 'securityGroup', 'distributionList'];
 *   const userIdentity = config.idMode === 'sharePoint' || config.idMode === 'both'
 *     ? { mode: config.idMode, webUrl: target, ensureUser: config.ensureUser } : { mode: config.idMode };
 *   const common = {
 *     label: 'PrincipalPicker', presentation: config.presentation, multiselect: config.multi, searchable: config.searchable,
 *     colorfulTagAvatars: config.colorful, refreshKey: refresh,
 *     selectedOptions: selection?.selectedOptions ?? [], initialOptions: selection?.selectedItems, availableOptions: available,
 *     suggestionsHeader: config.header.trim() || undefined,
 *     emptyContent: config.customStatus ? 'No candidates: type at least two characters.' : undefined,
 *     loadingContent: config.customStatus ? <span>Fetching live suggestions…</span> : undefined,
 *     errorContent: config.customStatus ? { children: 'Source unavailable: inspect diagnostics and retry.' } : undefined,
 *     strings: config.customStrings ? { labels: { selectOne: 'Choose source content', selectMany: 'Choose source content', search: 'Find a source item', retry: 'Try this source again' }, status: { noOptions: 'No matching source items' } } : undefined,
 *     onOptionSelect: (_: unknown, data: PickerSelectionData<PrincipalPickerOption>): void => { if (current()) { setSelection(data); setError(''); } },
 *     onError: (failure: Error): void => { if (current()) setError(failure.message); },
 *   };
 *   const controls = <div className={classes.controls}>
 *     <Field label="Presentation"><Select aria-label="Presentation" value={config.presentation} onChange={(_, data) => update({ presentation: data.value === 'tagPicker' ? 'tagPicker' : 'combobox' })}><option value="combobox">Dropdown</option><option value="tagPicker">TagPicker</option></Select></Field>
 *     <Switch label="Multiple selection" checked={config.multi} onChange={(_, data) => update({ multi: data.checked })} />
 *     <Switch label="Enable search" checked={config.searchable} onChange={(_, data) => update({ searchable: data.checked })} />
 *     <Switch label="Colorful tag avatars" checked={config.colorful} disabled={config.presentation !== 'tagPicker'} onChange={(_, data) => update({ colorful: data.checked })} />
 *     <Field label="Suggestion header (empty: default)"><Input aria-label="Suggestion header" value={config.header} onChange={(_, data) => update({ header: data.value })} /></Field>
 *     <Switch label="Customize status content" checked={config.customStatus} onChange={(_, data) => update({ customStatus: data.checked })} />
 *     <Switch label="Custom search and status text" checked={config.customStrings} onChange={(_, data) => update({ customStrings: data.checked })} />
 *     <Field label="Web URL (empty: current web)"><Input aria-label="Web URL" value={config.webUrl} onChange={(_, data) => update({ webUrl: data.value })} /></Field>
 *     <Field label="Principal types"><Select aria-label="Principal types" value={config.types} onChange={(_, data) => update({ types: data.value === 'groups' ? 'groups' : data.value === 'all' ? 'all' : 'users' })}><option value="users">Users</option><option value="groups">Groups</option><option value="all">Users and groups</option></Select></Field>
 *     <Field label="User IDs"><Select aria-label="User IDs" value={config.idMode} onChange={(_, data) => update({ idMode: data.value === 'both' ? 'both' : data.value === 'sharePoint' ? 'sharePoint' : data.value === 'entra' ? 'entra' : 'available' })}><option>available</option><option>entra</option><option>sharePoint</option><option>both</option></Select></Field>
 *     <Switch label="Allow ensureUser on explicit confirmation" checked={config.ensureUser} disabled={config.idMode !== 'sharePoint' && config.idMode !== 'both'} onChange={(_, data) => update({ ensureUser: data.checked })} />
 *     <Switch label="Load Graph personal suggestions" checked={config.personal} onChange={(_, data) => update({ personal: data.checked })} />
 *     <Switch label="Offer current user" checked={config.currentUser} onChange={(_, data) => update({ currentUser: data.checked })} />
 *   </div>;
 *   const commands = <div className={classes.commands}>
 *     <Button onClick={() => { if (current()) setStarted(true); }}>Start</Button>
 *     <Button onClick={() => { if (current()) setAvailable(selection?.selectedItems ?? []); }}>Use confirmed models as consumer options</Button>
 *     <Button onClick={() => { if (current()) setAvailable([]); }}>Empty consumer options</Button>
 *     <Button onClick={() => { if (current()) setSelection(undefined); }}>Clear controlled selection</Button>
 *     <Button onClick={() => { if (current()) { setError(''); setRefresh(value => value + 1); } }}>Refresh / retry</Button>
 *   </div>;
 *   const values = selection?.selectedValues ?? [];
 *   const regions: DemoRegions = {
 *     controls, commands,
 *     status: <p role="status">{error || (active ? 'Started: use the picker below.' : 'Not started. Start explicitly to read the source.')}</p>,
 *     children: <><p>Personal and current-user suggestions start off. Graph-dependent discovery and identity resolution wait for client initialization; SharePoint-only choices remain usable after a Graph initialization failure. Only explicit confirmation or ID resolution can use opt-in ensureUser in sharePoint/both modes. Source confirmation and any ensureUser request precede validateSelection: a validator cannot prevent an already dispatched write. Reset retires this operational picker and prevents obsolete local publication; an already sent server write can still finish. Retirement does not abort, undo or serialize writes across sessions.</p><p>Sample defaults explicitly select Dropdown, single selection, search enabled and colorful avatars off. Copy confirmed models, clear selection, then reopen to compare independent consumer choices. Empty consumer options retains selection. No history is saved. Presentation, labels, strings and equivalent models retain the live picker. Popup search keeps the selected label in the trigger. Clearing query only clears search; selection clear removes its value. Multiple selected rows and tags can be removed without searching. Escape returns focus; Tab leaves without accepting a highlighted row.</p>{active ? <PrincipalPicker {...common} webUrl={target} principalTypes={principalTypes} userIdentity={userIdentity} loadPersonalSuggestions={config.personal} initialSuggestions={{ includeCurrentUser: config.currentUser }} /> : <p>Start is required before the operational picker mounts.</p>}</>,
 *     diagnostics: <pre className={classes.diagnostics}>{JSON.stringify({ started: active, selectedOptions: selection?.selectedOptions ?? [], selectedItems: selection?.selectedItems ?? [], selectedValues: values, unresolvedOptions: selection?.unresolvedOptions ?? [], reason: selection?.reason ?? 'No action', optionValue: selection?.optionValue, optionText: selection?.optionText, roundTrip: JSON.stringify(JSON.parse(JSON.stringify(values))) === JSON.stringify(values), consumerOptions: available, error }, undefined, 2)}</pre>,
 *   };
 *   return render ? render(regions) : <section>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</section>;
 * }
 * export const demo = { host: 'webPart', kind: 'write', contentPlacement: 'configuration', observe: 'Start principal discovery, compare user and group sources and ID modes, and observe that only explicit confirmation can use the optional ensureUser policy.' };
 * ```
 */
export function PrincipalPicker(props: PrincipalPickerProps): React.ReactElement {
  const base = usePickerSourceContext(), graph = useSPFxMSGraphClient(), { serviceScope } = useSPFxServiceScope();
  const access = React.useMemo(() => createPrincipalGraphAccess(), [base.owner, serviceScope]);
  React.useLayoutEffect(() => { access.publish(graph.client, graph.isInitializing, graph.initError); }, [access, graph.client, graph.isInitializing, graph.initError]);
  React.useEffect(() => () => access.dispose(), [access]);
  // Readiness preserves in-flight actions; replacing the owner/factory still creates a new binding.
  const context = React.useMemo(() => ({ ...base, graphClientIdentity: access, getGraphClient: access.getClient }), [base, access]);
  const source = createPrincipalPickerSource(props, context);
  const runtime = usePickerRuntime<PrincipalPickerValue, PrincipalPickerOption>(props, source, { multiselect: true });
  return <PickerRenderer<PrincipalPickerValue, PrincipalPickerOption> props={props} runtime={runtime} defaultPresentation="tagPicker" defaultMultiselect />;
}

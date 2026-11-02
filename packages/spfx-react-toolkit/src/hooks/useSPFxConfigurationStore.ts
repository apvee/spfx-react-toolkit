/**
 * Explicit configuration scopes support site overrides over tenant defaults.
 *
 * @packageDocumentation
 * @docPage hooks/storage/configuration-store
 */
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { useSPFxSPHttpClient } from './useSPFxSPHttpClient';
import { useSPFxPageContext } from './useSPFxPageContext';
import { createSPFxConfigurationStoreService } from '../services/spfx-configuration-store.service';
import type {
  SPFxConfigurationStoreItem, SPFxConfigurationStoreOptions, SPFxConfigurationStoreService
} from '../services/spfx-configuration-store.service';

export type {
  SPFxConfigurationStoreScope, SPFxConfigurationStoreMode, SPFxConfigurationStoreServiceOptions,
  SPFxConfigurationStoreOptions, SPFxConfigurationStoreTarget, SPFxConfigurationStoreSaveOptions,
  SPFxConfigurationStoreItem, SPFxConfigurationStoreService
} from '../services/spfx-configuration-store.service';

/**
 * Configuration operations plus independent read/write state. Inherits all methods from SPFxConfigurationStoreService; unlike the historical storage hooks, failed reads reject as well as publishing error. No data or permission indicator is cached.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxConfigurationStoreResult, useSPFxConfigurationStore } from '@apvee/spfx-react-toolkit';
 *
 * function LoadSetting() {
 *   const store: SPFxConfigurationStoreResult = useSPFxConfigurationStore({ mode: 'layered' });
 *   const load = async () => { try { console.log(await store.get('my-app:pageSize')); } catch (error) { console.error(error); } };
 *   return <button disabled={!store.isReady} onClick={load}>Load</button>;
 * }
 * ```
 */
export interface SPFxConfigurationStoreResult extends SPFxConfigurationStoreService {
  /** True when client and required local URLs are available; does not establish discovery, list existence or access. */
  readonly isReady: boolean;
  /** True while any current-identity get/listScope is pending. */
  readonly isLoading: boolean;
  /** Latest-started read failure, cleared on another read; undefined initially. Non-Error rejections are normalized for state only. */
  readonly error: Error | undefined;
  /** True while any current-identity save/remove is pending. */
  readonly isWriting: boolean;
  /** Latest-started mutation failure, cleared on another mutation; undefined initially. */
  readonly writeError: Error | undefined;
}

/**
 * Uses the current provider's client and site collection to resolve configuration. Construction/mount never read, provision or check permissions. Layered reads use the entire site record when present, including false/zero/empty and legacy null values, and consult tenant only after confirmed absence. All mutations require an enabled explicit scope.
 * Read and write promises preserve original rejections while publishing their own error channels. Current-identity pending operations are counted independently; old callbacks retain original destinations, but obsolete/unmounted completions cannot publish state. No automatic refetch or request cancellation is provided.
 * @param options - Required tenant/site/layered mode; equivalent inline objects preserve callbacks. Site uses pageContext.site.absoluteUrl; tenant discovery uses a snapshot of pageContext.web.absoluteUrl.
 * @returns Explicit service operations and local readiness/loading/errors. Method parameters and return values inherit SPFxConfigurationStoreService.
 * @throws Outside a matching host-specific SPFx provider, or for an invalid mode during rendering. Not-ready operations and service failures reject their Promises.
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxConfigurationStore } from '@apvee/spfx-react-toolkit';
 *
 * function PageSize() {
 *   const store = useSPFxConfigurationStore({ mode: 'layered' });
 *   const save = async () => {
 *     try { await store.save('my-app:pageSize', 50, { scope: 'site' }); }
 *     catch (error) { console.error(error); }
 *   };
 *   return <button disabled={!store.isReady || store.isWriting} onClick={save}>Save site override</button>;
 * }
 * ```
 *
 * @useWhen configuration needs explicit tenant/site scopes or site overrides over tenant defaults.
 *
 *
 * @example Read a site override then reveal its tenant default
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxConfigurationStore } from '@apvee/spfx-react-toolkit/hooks';
 * import type { SPFxConfigurationStoreItem } from '@apvee/spfx-react-toolkit/hooks';
 *
 * function ConfigurationExample() {
 *   const store = useSPFxConfigurationStore({ mode: 'layered' });
 *   const [item, setItem] = React.useState<SPFxConfigurationStoreItem<number>>();
 *   const run = async (): Promise<void> => {
 *     try {
 *       await store.save('my-app:pageSize', 20, { scope: 'tenant' });
 *       await store.save('my-app:pageSize', 50, { scope: 'site' });
 *       setItem(await store.get<number>('my-app:pageSize')); // 50, source site
 *       await store.remove('my-app:pageSize', { scope: 'site' });
 *       setItem(await store.get<number>('my-app:pageSize')); // 20, source tenant
 *     } catch {
 *       setItem(undefined); // Render the current hook error below.
 *     }
 *   };
 *   return <>
 *     <button disabled={!store.isReady || store.isLoading || store.isWriting} onClick={run}>Run configuration example</button>
 *     <p>{item ? `${item.source}: ${item.value}` : 'No result'}</p>
 *     {(store.error ?? store.writeError) && <p role="alert">{(store.error ?? store.writeError)?.message}</p>}
 *   </>;
 * }
 * ```
 *

 *

 *

 *

 *

 *

 *

 *

 *

 *
 * @example Compare layered configuration and physical inventories
 * ```tsx
 * import * as React from 'react';
 * import { Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, Select, Text, makeStyles } from '@fluentui/react-components';
 * import { useSPFxConfigurationStore, useSPFxInstanceInfo, useSPFxPageContext, useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
 * import type { SPFxConfigurationStoreItem, SPFxConfigurationStoreMode, SPFxConfigurationStoreScope } from '@apvee/spfx-react-toolkit';
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Action = 'get' | 'list' | 'save' | 'remove';
 * type Format = 'text' | 'json' | 'bigint' | 'date';
 * type Ticket = { promise?: Promise<void> };
 * type Outcome = { token: object; steps: string[]; message?: string; failure?: string; item?: SPFxConfigurationStoreItem; items?: SPFxConfigurationStoreItem[] };
 * const useStyles = makeStyles({ payload: { overflowX: 'auto', maxWidth: '100%', minWidth: 0 } });
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const styles = useStyles();
 *   const page = useSPFxPageContext();
 *   const { client } = useSPFxSPHttpClient();
 *   const instance = useSPFxInstanceInfo();
 *   const [mode, setMode] = React.useState<SPFxConfigurationStoreMode>('layered');
 *   const [selectedScope, setScope] = React.useState<SPFxConfigurationStoreScope>('site');
 *   const scope = mode === 'layered' ? selectedScope : mode;
 *   const store = useSPFxConfigurationStore({ mode });
 *   const root = mode !== 'tenant' ? page.site.absoluteUrl.trim().replace(/\/+$/, '') : undefined;
 *   const web = mode !== 'site' ? page.web.absoluteUrl.trim().replace(/\/+$/, '') : undefined;
 *   const [action, setAction] = React.useState<Action>('get');
 *   const [key, setKey] = React.useState('spfx-toolkit-config-demo-disposable');
 *   const [value, setValue] = React.useState('');
 *   const [format, setFormat] = React.useState<Format>('text');
 *   const [replaceDescription, setReplaceDescription] = React.useState(false);
 *   const [description, setDescription] = React.useState('Disposable configuration store demo item');
 *   const [consent, setConsent] = React.useState<object>();
 *   const approval = React.useRef<object>();
 *   const approve = (next: object | undefined): void => { approval.current = next; setConsent(next); };
 *   const fingerprint = JSON.stringify([mode, scope, action, key, value, format, replaceDescription, description]);
 *   const boundary = React.useRef({ resetRevision, client, page, instance: instance.id, root, web, fingerprint, token: {} });
 *   const previous = boundary.current;
 *   if (previous.resetRevision !== resetRevision || previous.client !== client || previous.page !== page || previous.instance !== instance.id || previous.root !== root || previous.web !== web || previous.fingerprint !== fingerprint) boundary.current = { resetRevision, client, page, instance: instance.id, root, web, fingerprint, token: {} };
 *   const token = boundary.current.token;
 *   const mounted = React.useRef(true);
 *   const ledger = React.useRef(new Set<Ticket>());
 *   const sequence = React.useRef<object>();
 *   const latest = React.useRef(0);
 *   const [, refresh] = React.useState(0);
 *   const [outcome, setOutcome] = React.useState<Outcome>();
 *   const [confirmation, setConfirmation] = React.useState<{ confirm: () => void; cancel: () => void; key: string; scope: SPFxConfigurationStoreScope }>();
 *   const confirmationRef = React.useRef<{ token: object; owner: object; settle: (value: boolean) => void }>();
 *   const focus = React.useRef<{ trigger: HTMLButtonElement; owns: () => boolean }>();
 *   const current = (): boolean => mounted.current && boundary.current.token === token && (root === undefined || page.site.absoluteUrl.trim().replace(/\/+$/, '') === root) && (web === undefined || page.web.absoluteUrl.trim().replace(/\/+$/, '') === web);
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; confirmationRef.current?.settle(false); confirmationRef.current = undefined; }; }, []);
 *   React.useLayoutEffect(() => {
 *     if (confirmationRef.current && confirmationRef.current.owner !== token) { confirmationRef.current.settle(false); confirmationRef.current = undefined; setConfirmation(undefined); }
 *   }, [token]);
 *   React.useLayoutEffect(() => { const request = focus.current; if (!confirmation && request && sequence.current === undefined && ledger.current.size === 0 && !request.trigger.disabled) { focus.current = undefined; if (request.owns() && request.trigger.isConnected) request.trigger.focus(); } });
 *   const resetSeen = React.useRef(resetRevision);
 *   React.useLayoutEffect(() => {
 *     if (resetSeen.current !== resetRevision) {
 *       resetSeen.current = resetRevision; approve(undefined); setOutcome(undefined); setMode('layered'); setScope('site'); setAction('get'); setKey('spfx-toolkit-config-demo-disposable'); setValue(''); setFormat('text'); setReplaceDescription(false); setDescription('Disposable configuration store demo item');
 *     }
 *   }, [resetRevision]);
 *   const busy = (): boolean => ledger.current.size > 0 || sequence.current !== undefined || confirmationRef.current !== undefined;
 *   const locked = busy();
 *   const updateDraft = (changed: boolean, update: () => void): void => {
 *     if (!changed || !current() || busy()) return;
 *     // Retire old admission before React can batch the changed draft with old consent/Run.
 *     boundary.current = { ...boundary.current, token: {} }; approve(undefined); update();
 *   };
 *   const disposable = key.startsWith('spfx-toolkit-config-demo-') && key.length > 'spfx-toolkit-config-demo-'.length;
 *   const visible = current() && outcome?.token === token ? outcome : undefined;
 *   const askRemoval = (destination: SPFxConfigurationStoreScope, trigger: HTMLButtonElement, owns: () => boolean): Promise<boolean> => new Promise(resolve => {
 *     const confirmationToken = {};
 *     confirmationRef.current = { token: confirmationToken, owner: token, settle: resolve };
 *     const close = (accepted: boolean): void => {
 *       if (!owns() || confirmationRef.current?.token !== confirmationToken || ledger.current.size > 0) return;
 *       confirmationRef.current = undefined; focus.current = { trigger, owns }; setConfirmation(undefined); resolve(accepted);
 *     };
 *     setConfirmation({ key, scope: destination, confirm: () => close(true), cancel: () => close(false) });
 *   });
 *   const run = async (trigger: HTMLButtonElement, recipe = false): Promise<void> => {
 *     if (!current() || busy() || !store.isReady || (action !== 'list' && !key)) return;
 *     const mutating = recipe || action === 'save' || action === 'remove';
 *     if (mutating && (!disposable || approval.current !== token || (recipe && mode !== 'layered'))) return;
 *     const operation = ++latest.current, sequenceToken = {};
 *     const owns = (): boolean => current() && latest.current === operation;
 *     const steps: string[] = [];
 *     if (mutating) { sequence.current = sequenceToken; approve(undefined); refresh(count => count + 1); }
 *     setOutcome({ token, steps: [], message: recipe ? 'Running recipe...' : action === 'get' ? 'Reading...' : action === 'list' ? 'Loading physical inventory...' : 'Preparing mutation...' });
 *     const mark = (message: string, data: { item?: SPFxConfigurationStoreItem; items?: SPFxConfigurationStoreItem[] } = {}): void => {
 *       if (!owns()) return;
 *       steps.push(message); setOutcome(previousOutcome => ({ ...(previousOutcome?.token === token ? previousOutcome : { token }), ...data, token, steps: [...steps], message }));
 *     };
 *     const write = async (dispatch: () => Promise<void>): Promise<boolean> => {
 *       if (!owns() || ledger.current.size > 0) return false;
 *       const ticket: Ticket = {}; ledger.current.add(ticket); refresh(count => count + 1);
 *       const release = (): void => { ledger.current.delete(ticket); if (mounted.current) refresh(count => count + 1); };
 *       let promise: Promise<void>;
 *       try { promise = dispatch(); ticket.promise = promise; } catch (error) { release(); throw error; }
 *       // The ORIGINAL public promise owns every internal RPC, including lazy discovery and setup.
 *       await promise.finally(release); return owns();
 *     };
 *     try {
 *       if (recipe) {
 *         if (!await write(() => store.save(key, 20, { scope: 'tenant' }))) return;
 *         mark('Saved tenant default 20.');
 *         if (!owns() || !await write(() => store.save(key, 50, { scope: 'site', description: 'Local override' }))) return;
 *         mark('Saved site override 50.');
 *         if (!owns()) return; const item = await store.get<unknown>(key); if (!owns()) return;
 *         mark('Read the effective site override.', { item });
 *         const removeSite = (): Promise<void> => store.remove(key, { scope: 'site' });
 *         if (!owns() || !await askRemoval('site', trigger, owns) || !owns()) { mark('Site removal cancelled; completed writes remain.'); return; }
 *         if (!await write(removeSite)) return;
 *         mark('Removed only the site override.');
 *         if (!owns()) return; const inherited = await store.get<unknown>(key); if (!owns()) return;
 *         mark('Read the tenant default after removal.', { item: inherited });
 *       } else if (action === 'get') {
 *         const item = await store.get<unknown>(key); if (owns()) mark(item ? `Found ${key}.` : `${key} not found.`, { item });
 *       } else if (action === 'list') {
 *         const items = await store.listScope(scope); if (owns()) mark(`Loaded ${items.length} ${scope} physical item(s).`, { items });
 *       } else if (action === 'save') {
 *         const input: unknown = format === 'json' ? JSON.parse(value) : format === 'bigint' ? BigInt(value) : format === 'date' ? new Date(value) : value;
 *         if (input instanceof Date && Number.isNaN(input.getTime())) throw new Error('Enter a valid Date.');
 *         if (await write(() => store.save(key, input, replaceDescription ? { scope, description } : { scope }))) mark(`Saved ${key} in ${scope}.`);
 *       } else {
 *         const removeCaptured = (): Promise<void> => store.remove(key, { scope });
 *         if (await askRemoval(scope, trigger, owns) && owns() && await write(removeCaptured)) mark(`Removed ${key} from ${scope}. Read again to resolve inheritance.`);
 *       }
 *     } catch (error) { if (owns()) setOutcome(previousOutcome => ({ ...(previousOutcome?.token === token ? previousOutcome : { token, steps: [] }), failure: error instanceof Error ? error.message : String(error) })); }
 *     finally { if (sequence.current === sequenceToken) sequence.current = undefined; if (mounted.current) refresh(count => count + 1); }
 *   };
 *   // Tie Fluent disabledFocusable to the exact disabled flag: return focus after consumed consent without enabling invocation.
 *   const shown = (value: unknown): string => (JSON.stringify(value) ?? String(value)) + (value === '' ? ' (empty string)' : '');
 *   const runDisabled = !store.isReady || locked || (action !== 'list' && !key) || (action === 'save' || action === 'remove') && (!disposable || consent !== token);
 *   const recipeDisabled = !store.isReady || locked || mode !== 'layered' || !disposable || consent !== token;
 *   const regions: DemoRegions = {
 *     controls: <>
 *       <Field label="Mode"><Select aria-label="Mode" disabled={locked} value={mode} onChange={(_, data) => updateDraft(data.value !== mode, () => setMode(data.value as SPFxConfigurationStoreMode))}><option value="layered">Tenant defaults with site overrides</option><option value="site">Site collection only</option><option value="tenant">Tenant only</option></Select></Field>
 *       <Field label="Operation"><Select aria-label="Operation" disabled={locked} value={action} onChange={(_, data) => updateDraft(data.value !== action, () => setAction(data.value as Action))}><option value="get">Read effective value</option><option value="list">Physical inventory</option><option value="save">Save</option><option value="remove">Remove</option></Select></Field>
 *       <Field label="Physical scope"><Select aria-label="Physical scope" disabled={locked || mode !== 'layered'} value={scope} onChange={(_, data) => { if (data.value !== 'site' && data.value !== 'tenant' || mode !== 'layered' && data.value !== mode) return; updateDraft(data.value !== scope, () => setScope(data.value as SPFxConfigurationStoreScope)); }}><option value="site" disabled={mode === 'tenant'}>Site collection</option><option value="tenant" disabled={mode === 'site'}>Tenant app catalog</option></Select></Field>
 *       <Field label="Key"><Input aria-label="Key" disabled={locked} value={key} onChange={(_, data) => updateDraft(data.value !== key, () => setKey(data.value))}/></Field>
 *       {action === 'save' && <><Field label="Value format"><Select aria-label="Value format" disabled={locked} value={format} onChange={(_, data) => updateDraft(data.value !== format, () => setFormat(data.value as Format))}><option value="text">Text</option><option value="json">JSON</option><option value="bigint">Big integer</option><option value="date">Date (ISO)</option></Select></Field>
 *         <Field label="Value"><Input aria-label="Value" disabled={locked} value={value} onChange={(_, data) => updateDraft(data.value !== value, () => setValue(data.value))}/></Field>
 *         <Checkbox label="Replace/clear destination description" disabled={locked} checked={replaceDescription} onChange={(_, data) => updateDraft((data.checked === true) !== replaceDescription, () => setReplaceDescription(data.checked === true))}/>
 *         <Field label="Description"><Input aria-label="Description" disabled={locked || !replaceDescription} value={description} onChange={(_, data) => updateDraft(data.value !== description, () => setDescription(data.value))}/></Field></>}
 *       <Text as="p">Manual {action}: scope {scope}, key {key}, format {format}, value {JSON.stringify(value)}, {replaceDescription ? `replace description with ${JSON.stringify(description)}` : 'preserve destination description'}. Recipe: tenant 20 (preserve description), site 50 (Local override), then remove only site, at this same exact key. Tenant means this captured discovery session, not the current web as catalog.</Text>
 *       <Checkbox label="I consent to these exact disposable mutations" disabled={locked || !disposable} checked={consent === token} onChange={(_, data) => { if (current() && !busy()) approve(data.checked === true ? token : undefined); }}/>
 *     </>,
 *     commands: <><Button appearance="primary" disabled={runDisabled} disabledFocusable={runDisabled} onClick={event => run(event.currentTarget)}>Run</Button><Button disabled={recipeDisabled} disabledFocusable={recipeDisabled} onClick={event => run(event.currentTarget, true)}>Run default and override recipe</Button></>,
 *     status: <>{ledger.current.size > 0 && <Text role="status">Previous write still pending</Text>}{visible?.failure ? <Text role="alert">{visible.failure}</Text> : <Text role="status">{visible?.message ?? 'Not run'}</Text>}{visible && <ol aria-label="Completed configuration steps">{visible.steps.map((step, index) => <li key={index}>{step}</li>)}</ol>}</>,
 *     children: <><dl><dt>Mode (required)</dt><dd>{mode}</dd><dt>Physical scope</dt><dd>{scope}</dd><dt>Collection root URL</dt><dd>{page.site.absoluteUrl}</dd><dt>Current web URL</dt><dd>{page.web.absoluteUrl}</dd><dt>Discovery web URL</dt><dd>{web ?? '(not used in site mode)'}</dd><dt>Provider instance</dt><dd>{instance.id}</dd><dt>Client ready</dt><dd>{String(store.isReady)}</dd><dt>Current read pending</dt><dd>{String(visible !== undefined && store.isLoading)}</dd><dt>Current write pending</dt><dd>{String(visible !== undefined && store.isWriting)}</dd>
 *       {visible?.item && <><dt>Source</dt><dd>{visible.item.source}</dd><dt>Selected key</dt><dd>{visible.item.key}</dd><dt>Selected value</dt><dd>{shown(visible.item.value)}</dd><dt>Value type</dt><dd>{typeof visible.item.value}</dd><dt>Description</dt><dd>{visible.item.description ?? '(not set)'}</dd><dt>Physical item ID</dt><dd>{visible.item.id}</dd></>}</dl>
 *       {visible?.items && <><Text>Inventory count: {visible.items.length}</Text><dl aria-label="First ten configuration records">{visible.items.slice(0, 10).map((item, index) => <React.Fragment key={index}><dt>{item.key}</dt><dd>{item.source} / {item.id}: {shown(item.value)}</dd></React.Fragment>)}</dl></>}
 *       <Dialog open={confirmation !== undefined} onOpenChange={(_, data) => { if (!data.open) confirmation?.cancel(); }}><DialogSurface><DialogBody><DialogTitle>Confirm physical removal</DialogTitle><DialogContent>Remove exact key {confirmation?.key} only from {confirmation?.scope}, collection root {root ?? '(not used)'}, captured tenant discovery web {web ?? '(not used)'}. No other scope is removed.</DialogContent><DialogActions><Button onClick={confirmation?.cancel}>Cancel</Button><Button appearance="primary" onClick={confirmation?.confirm}>Confirm removal</Button></DialogActions></DialogBody></DialogSurface></Dialog>
 *     </>,
 *     diagnostics: <><Text as="p">Read is effective; inventory and changes target one physical store. Only Save may initialize that store. Objects replace whole records; false, zero, empty and legacy null remain overrides. Physical catalog URL is unavailable through this facade. No canWrite is supplied. Manual Save/Remove never automatically reads. Remove a site override then Read again to see tenant defaults. Tenant cleanup requires a separate explicit consented tenant Remove. Reset restores local drafts and hides outcomes; the original public operation may continue. No cancellation, rollback, transaction or lock across scenarios is provided.</Text><div role="region" aria-label="Configuration payload" tabIndex={0} className={styles.payload}><pre>{JSON.stringify(visible, undefined, 2)}</pre></div></>,
 *   };
 *   return render ? render(regions) : <section aria-label="Configuration exploration">{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</section>;
 * }
 * export const demo = { host: 'webPart', kind: 'write', contentPlacement: 'configuration', observe: 'Compare tenant, site and layered reads with physical inventories and explicitly scoped consented changes.' };
 * ```
 */
export function useSPFxConfigurationStore(options: SPFxConfigurationStoreOptions): SPFxConfigurationStoreResult {
  const { client } = useSPFxSPHttpClient();
  const pageContext = useSPFxPageContext();
  const mode = options?.mode;
  if (mode !== 'tenant' && mode !== 'site' && mode !== 'layered') throw new Error('A valid configuration store mode is required.');
  const siteCollectionUrl = mode !== 'tenant' ? (pageContext.site.absoluteUrl || '').trim().replace(/\/+$/, '') : undefined;
  const discoveryWebUrl = mode !== 'site' ? (pageContext.web.absoluteUrl || '').trim().replace(/\/+$/, '') : undefined;
  const service = useMemo(() => {
    if (!client) return undefined;
    if (mode === 'site') return siteCollectionUrl ? createSPFxConfigurationStoreService(client, { mode, siteCollectionUrl }) : undefined;
    if (mode === 'tenant') return discoveryWebUrl ? createSPFxConfigurationStoreService(client, { mode, discoveryWebUrl }) : undefined;
    return siteCollectionUrl && discoveryWebUrl ? createSPFxConfigurationStoreService(client, { mode, siteCollectionUrl, discoveryWebUrl }) : undefined;
  }, [client, mode, siteCollectionUrl, discoveryWebUrl]);
  const identity = useMemo(() => ({ service, mode, siteCollectionUrl, discoveryWebUrl, reads: 0, writes: 0, readId: 0, writeId: 0 }), [service, mode, siteCollectionUrl, discoveryWebUrl]);
  const currentRef = useRef(identity);
  currentRef.current = identity;
  const mounted = useRef(true);
  const initial = useMemo(() => ({ identity, isLoading: false, isWriting: false,
    error: undefined as Error | undefined, writeError: undefined as Error | undefined }), [identity]);
  const [state, setState] = useState(initial);
  const isCurrent = useCallback(() => mounted.current && currentRef.current === identity, [identity]);
  useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
  useEffect(() => { setState(previous => previous.identity === identity ? previous : initial); }, [identity, initial]);

  const run = useCallback(async <T,>(write: boolean, operation: (store: SPFxConfigurationStoreService) => Promise<T>): Promise<T> => {
    const request = write ? ++identity.writeId : ++identity.readId;
    if (write) identity.writes++; else identity.reads++;
    if (isCurrent()) setState(previous => isCurrent() ? {
      ...(previous.identity === identity ? previous : initial),
      ...(write ? { isWriting: true, writeError: undefined } : { isLoading: true, error: undefined })
    } : previous);
    try {
      if (!service) throw new Error('SPHttpClient and the mode-specific configuration URLs are required.');
      return await operation(service);
    } catch (failure) {
      if (isCurrent() && request === (write ? identity.writeId : identity.readId)) {
        const error = failure instanceof Error ? failure : new Error(String(failure));
        setState(previous => isCurrent() ? { ...previous, ...(write ? { writeError: error } : { error }) } : previous);
      }
      throw failure;
    } finally {
      if (write) identity.writes--; else identity.reads--;
      if (isCurrent()) setState(previous => isCurrent() ? {
        ...previous, isLoading: identity.reads > 0, isWriting: identity.writes > 0
      } : previous);
    }
  }, [service, identity, initial, isCurrent]);
  const get = useCallback(<T = unknown>(key: string): Promise<SPFxConfigurationStoreItem<T> | undefined> => run(false, store => store.get<T>(key)), [run]);
  const listScope = useCallback<SPFxConfigurationStoreService['listScope']>(scope => run(false, store => store.listScope(scope)), [run]);
  const save = useCallback<SPFxConfigurationStoreService['save']>((key, value, target) => run(true, store => store.save(key, value, target)), [run]);
  const remove = useCallback<SPFxConfigurationStoreService['remove']>((key, target) => run(true, store => store.remove(key, target)), [run]);
  const visible = state.identity === identity ? state : initial;
  return useMemo(() => ({ get, listScope, save, remove, isReady: service !== undefined,
    isLoading: visible.isLoading, isWriting: visible.isWriting, error: visible.error, writeError: visible.writeError }),
  [get, listScope, save, remove, service, visible.isLoading, visible.isWriting, visible.error, visible.writeError]);
}

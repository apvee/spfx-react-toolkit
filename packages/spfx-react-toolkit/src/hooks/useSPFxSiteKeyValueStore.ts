/**
 * The legacy collection-root list store is shared by all subsites.
 *
 * @packageDocumentation
 * @docPage hooks/storage/site-key-value-store
 */
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { useSPFxSPHttpClient } from './useSPFxSPHttpClient';
import { useSPFxPageContext } from './useSPFxPageContext';
import { createSPFxSiteKeyValueStoreService } from '../services/spfx-site-key-value-store.service';

/**
 * A key and its deserialized value from the collection's root-web store.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxSiteKeyValueStoreItem, useSPFxSiteKeyValueStore } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const store = useSPFxSiteKeyValueStore();
 *   const [result, setResult] = React.useState<unknown>();
 *   const [failure, setFailure] = React.useState<string>();
 *   const run = async () => {
 *     setFailure(undefined);
 *     try {
 *       const item: SPFxSiteKeyValueStoreItem<boolean> | undefined = await store.get('featureEnabled');
 *       setResult(item?.value);
 *     } catch (failure) {
 *       setFailure(failure instanceof Error ? failure.message : String(failure));
 *     }
 *   };
 *   return <>
 *     <button disabled={!store.isReady || store.isLoading} onClick={() => void run()}>Load setting</button>
 *     {(failure ?? store.error?.message) && <span role="alert">{failure ?? store.error?.message}</span>}
 *     <pre>{JSON.stringify(result)}</pre>
 *   </>;
 * }
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxSiteKeyValueStoreItem<T = unknown> {
  /** Property key, stored in the Title column. */
  readonly key: string;
  /** Deserialized value; the generic does not perform runtime validation. */
  readonly value: T;
  /** Optional description metadata. */
  readonly description: string | undefined;
  /** SharePoint list item ID. */
  readonly id: number;
}

/**
 * State and operations for a site collection key-value store.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxSiteKeyValueStoreResult, useSPFxSiteKeyValueStore } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const store: SPFxSiteKeyValueStoreResult = useSPFxSiteKeyValueStore();
 *   const [result, setResult] = React.useState<unknown>();
 *   const [failure, setFailure] = React.useState<string>();
 *   const run = async () => {
 *     setFailure(undefined);
 *     try {
 *       const item = await store.get<boolean>('featureEnabled');
 *       setResult(item?.value);
 *     } catch (failure) {
 *       setFailure(failure instanceof Error ? failure.message : String(failure));
 *     }
 *   };
 *   return <>
 *     <button disabled={!store.isReady || store.isLoading} onClick={() => void run()}>Load setting</button>
 *     {(failure ?? store.error?.message) && <span role="alert">{failure ?? store.error?.message}</span>}
 *     <pre>{JSON.stringify(result)}</pre>
 *   </>;
 * }
 * ```
 */
export interface SPFxSiteKeyValueStoreResult {
  /** True while any read operation is pending. */
  readonly isLoading: boolean;
  /** Error from the latest-started read; cleared when another read starts. */
  readonly error: Error | undefined;
  /** True while any write operation is pending. */
  readonly isWriting: boolean;
  /** Error from the latest-started write; cleared when another write starts. */
  readonly writeError: Error | undefined;
  /** Effective write permission indicator, initially false; refreshed on mount and after each current-identity read/write succeeds or fails. Failed checks return false and only the latest check may publish. SharePoint still enforces grants and schema repair may need additional permissions. */
  readonly canWrite: boolean;
  /** True when the client and a nonempty collection root URL are available. */
  readonly isReady: boolean;
  /**
   * Reads without provisioning; failed requests publish error and return undefined.
   * @typeParam T - Expected deserialized value shape; no runtime validation.
   * @param key - Key to read.
   * @returns An item, or undefined for absent keys/lists or caught read failures.
   * @throws Not-ready operations reject; writes also reject request/setup failures.
   */
  readonly get: <T = unknown>(key: string) => Promise<SPFxSiteKeyValueStoreItem<T> | undefined>;
  /**
   * Reads all pages without provisioning; failed requests publish error and return [].
   * @returns All items, or [] for absent lists or caught read failures.
   * @throws Not-ready operations reject; writes also reject request/setup failures.
   */
  readonly list: () => Promise<SPFxSiteKeyValueStoreItem<unknown>[]>;
  /**
   * Ensures setup, then writes with an observed item ETag; missing ETags and HTTP 412 conflicts publish writeError and reject without a blind overwrite. This is not compare-and-swap against a previous get().
   * @param key - Key to create/update.
   * @typeParam T - Supplied value type; serialization may change its read-back representation.
   * @param value - Value to encode. Empty/JSON-like strings preserve their text, bigint reads as exact decimal string and Date as ISO string; T does not validate runtime content. Historical values keep their previous interpretation.
   * @param description - Omitted preserves existing description; empty string clears it.
   * @returns Resolves after successful write.
   * @throws Not-ready operations reject; writes also reject request/setup failures.
   */
  readonly save: <T = unknown>(key: string, value: T, description?: string) => Promise<void>;
  /**
   * Deletes without provisioning using an observed ETag; failures/conflicts publish writeError and reject. A missing list/key is a no-op only after confirmed absence.
   * @param key - Key to delete.
   * @returns Resolves after deletion or confirmed absence.
   * @throws Not-ready operations reject; writes also reject request/setup failures.
   */
  readonly remove: (key: string) => Promise<void>;
}

/**
 * Manages SiteKeyValueStore in pageContext.site.absoluteUrl, shared by all subsites. Mount checks permissions without provisioning; every settled current-identity operation refreshes them. New writes preserve text and bigint precision without rewriting historical data; Date and bigint read as strings. Mutations use item ETags and reject conflicts rather than blindly overwriting. Legacy empty Note values still read as empty strings. canWrite is an indicator; SharePoint enforces grants. Read failures return fallbacks and publish error, writes reject and publish writeError. Not-ready operations reject. Old collection/client completions cannot publish current state.
 * @returns Store readiness, read/write state and explicit operations.
 * @throws When called outside an SPFx provider.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxSiteKeyValueStore } from '@apvee/spfx-react-toolkit';
 *
 * function SiteSetting() {
 *   const store = useSPFxSiteKeyValueStore();
 *   const [result, setResult] = React.useState<unknown>();
 *   const [failure, setFailure] = React.useState<string>();
 *   const run = async () => {
 *     setFailure(undefined);
 *     try {
 *       const item = await store.get<boolean>('featureEnabled');
 *       setResult(item?.value);
 *     } catch (failure) {
 *       setFailure(failure instanceof Error ? failure.message : String(failure));
 *     }
 *   };
 *   return <>
 *     <button disabled={!store.isReady || store.isLoading} onClick={() => void run()}>Load setting</button>
 *     {(failure ?? store.error?.message) && <span role="alert">{failure ?? store.error?.message}</span>}
 *     <pre>{JSON.stringify(result)}</pre>
 *   </>;
 * }
 * ```
 *
 * @useWhen legacy shared configuration belongs to the current site collection.
 *
 * @avoidWhen new configuration needs layered tenant defaults or explicit mutation scopes; use useSPFxConfigurationStore.
 *
 *
 * @example Read and save with separate hook errors
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxSiteKeyValueStore, SPFxSiteKeyValueStoreItem } from '@apvee/spfx-react-toolkit';
 *
 * function SiteSettings() {
 *   const store = useSPFxSiteKeyValueStore();
 *   const [item, setItem] = React.useState<SPFxSiteKeyValueStoreItem<unknown>>();
 *   const [hasRead, setHasRead] = React.useState(false);
 *   const [message, setMessage] = React.useState('');
 *
 *   const load = async () => {
 *     setHasRead(false);
 *     setMessage('');
 *     try {
 *       setItem(await store.get('toolkit-site-demo-setting'));
 *       setHasRead(true);
 *     } catch (failure) {
 *       setMessage(String(failure)); // For example, an operation before readiness.
 *     }
 *   };
 *
 *   const save = async () => {
 *     setMessage('');
 *     try {
 *       await store.save('toolkit-site-demo-setting', { enabled: true }, 'Collection setting');
 *       setMessage('Saved');
 *     } catch (failure) {
 *       setMessage(String(failure));
 *     }
 *   };
 *
 *   return (
 *     <div>
 *       <button onClick={load} disabled={!store.isReady || store.isLoading}>Read</button>
 *       <button onClick={save} disabled={!store.isReady || !store.canWrite || store.isWriting}>Save</button>
 *       {store.error ? <p role="alert">{store.error.message}</p> :
 *         hasRead && !store.isLoading && <p>{item ? JSON.stringify(item.value) : 'Key not found'}</p>}
 *       {store.writeError && <p role="alert">{store.writeError.message}</p>}
 *       {message && <p>{message}</p>}
 *     </div>
 *   );
 * }
 * ```
 *
 * @example Explore collection root records with guarded mutations
 * ```tsx
 * import * as React from 'react';
 * import { makeStyles, Accordion, AccordionHeader, AccordionItem, AccordionPanel, Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, Select, Text } from '@fluentui/react-components';
 * import { useSPFxInstanceInfo, useSPFxPageContext, useSPFxSiteKeyValueStore, useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
 * import type { SPFxSiteKeyValueStoreItem } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Presenter = (regions: DemoRegions) => React.ReactElement;
 * type Session = { generation: object; root: string; web: string };
 * type Action = 'get' | 'list' | 'save' | 'remove';
 * type Format = 'text' | 'json' | 'bigint' | 'date';
 * const useStyles = makeStyles({ payload: { overflowX: 'auto', maxWidth: '100%', minWidth: 0 } });
 * type Ticket = object;
 *
 * function present(regions: DemoRegions, render?: Presenter): React.ReactElement {
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 *
 * // This bounded child owns the hook session. Reset deliberately unmounts it.
 * function SiteSession({ session, current, write, pending, isLocked, render }: {
 *   session: Session; current: () => boolean; write: (operation: () => Promise<void>) => Promise<void> | undefined;
 *   pending: boolean; isLocked: () => boolean; render?: Presenter;
 * }): React.ReactElement {
 *   const store = useSPFxSiteKeyValueStore();
 *   const styles = useStyles();
 *   const [action, setAction] = React.useState<Action>('get');
 *   const [key, setKey] = React.useState('spfx-toolkit-site-demo-disposable');
 *   const [value, setValue] = React.useState('');
 *   const [format, setFormat] = React.useState<Format>('text');
 *   const [replaceDescription, setReplaceDescription] = React.useState(false);
 *   const [description, setDescription] = React.useState('Disposable SPFx React Toolkit site demo item');
 *   const [consent, setConsent] = React.useState<object>();
 *   const approval = React.useRef<object>();
 *   approval.current = consent;
 *   const approve = (next: object | undefined): void => { approval.current = next; setConsent(next); };
 *   const [outcome, setOutcome] = React.useState<{ boundary: object; action: Action; message?: string; failure?: string; item?: SPFxSiteKeyValueStoreItem; items?: SPFxSiteKeyValueStoreItem[] }>();
 *   const boundary = React.useRef({ fingerprint: '', token: {} });
 *   const fingerprint = JSON.stringify([action, key, value, format, replaceDescription, description]);
 *   if (boundary.current.fingerprint !== fingerprint) boundary.current = { fingerprint, token: {} };
 *   const target = boundary.current.token;
 *   const latest = React.useRef(0);
 *   const mounted = React.useRef(true);
 *   const confirmationRef = React.useRef<object>();
 *   const restoreFocus = React.useRef<{ trigger: HTMLButtonElement; current: () => boolean }>();
 *   const [confirmation, setConfirmation] = React.useState<{ confirm: () => void; cancel: () => void; key: string }>();
 *   React.useLayoutEffect(() => {
 *     const request = restoreFocus.current;
 *     if (!confirmation && request) { restoreFocus.current = undefined; if (request.current() && request.trigger.isConnected) request.trigger.focus(); }
 *   }, [confirmation]);
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; confirmationRef.current = undefined; }; }, []);
 *   const valid = (): boolean => mounted.current && current() && boundary.current.token === target;
 *   const disposable = key.startsWith('spfx-toolkit-site-demo-') && key.length > 'spfx-toolkit-site-demo-'.length;
 *   const visible = valid() && outcome?.boundary === target ? outcome : undefined;
 *   const reading = visible?.action === 'get' || visible?.action === 'list';
 *   // The public read resolves a fallback. Its CURRENT rendered error takes precedence.
 *   const failure = visible ? visible.failure ?? (reading ? store.error?.message : store.writeError?.message) : undefined;
 *   const locked = pending || confirmation !== undefined;
 *   const updateDraft = (changed: boolean, update: () => void): void => {
 *     if (!changed || !valid() || locked || isLocked() || confirmationRef.current) return;
 *     // Retire captured admission before React can batch the draft state update.
 *     boundary.current = { ...boundary.current, token: {} };
 *     approve(undefined);
 *     update();
 *   };
 *
 *   const run = async (trigger: HTMLButtonElement): Promise<void> => {
 *     if (!valid() || isLocked() || pending || confirmationRef.current || !store.isReady || store.isLoading || store.isWriting) return;
 *     if (action !== 'list' && !key) return;
 *     const mutating = action === 'save' || action === 'remove';
 *     if (mutating && (!disposable || approval.current !== target)) return;
 *     const operation = ++latest.current;
 *     const owns = (): boolean => valid() && latest.current === operation;
 *     const execute = async (remove?: () => Promise<void>): Promise<void> => {
 *       if (!owns()) return;
 *       setOutcome({ boundary: target, action });
 *       try {
 *         if (action === 'get') {
 *           const item = await store.get(key);
 *           if (owns()) setOutcome({ boundary: target, action, item, message: item ? `Found ${key}.` : `${key} not found.` });
 *         } else if (action === 'list') {
 *           const items = await store.list();
 *           if (owns()) setOutcome({ boundary: target, action, items, message: `Loaded ${items.length} site key-value item(s).` });
 *         } else {
 *           let input: unknown = value;
 *           if (action === 'save') {
 *             if (format === 'json') input = JSON.parse(value);
 *             else if (format === 'bigint') input = BigInt(value);
 *             else if (format === 'date') { const date = new Date(value); if (!Number.isFinite(date.getTime())) throw new Error('Enter a valid Date.'); input = date; }
 *           }
 *           // The parent retains the ORIGINAL public promise through Reset/child unmount.
 *           const promise = write(remove ?? (() => replaceDescription ? store.save(key, input, description) : store.save(key, input)));
 *           if (!promise) return;
 *           await promise;
 *           if (owns()) { approve(undefined); setOutcome({ boundary: target, action, message: `${action === 'save' ? 'Saved' : 'Removed'} ${key}.` }); }
 *         }
 *       } catch (error) {
 *         if (owns()) setOutcome({ boundary: target, action, failure: error instanceof Error ? error.message : String(error) });
 *       }
 *     };
 *     if (action !== 'remove') { await execute(); return; }
 *     const token = {};
 *     const remove = store.remove;
 *     const capturedKey = key;
 *     confirmationRef.current = token;
 *     const cancel = (): void => {
 *       if (!owns() || confirmationRef.current !== token) return;
 *       restoreFocus.current = { trigger, current: owns };
 *       confirmationRef.current = undefined; setConfirmation(undefined);
 *     };
 *     const confirm = (): Promise<void> | undefined => {
 *       if (!owns() || confirmationRef.current !== token || approval.current !== target) return;
 *       confirmationRef.current = undefined; setConfirmation(undefined);
 *       return execute(() => remove(capturedKey));
 *     };
 *     setConfirmation({ confirm, cancel, key: capturedKey });
 *   };
 *   const shownValue = (item: SPFxSiteKeyValueStoreItem): string => typeof item.value === 'string' ? JSON.stringify(item.value) : JSON.stringify(item.value) ?? String(item.value);
 *   const regions: DemoRegions = {
 *     controls: <>
 *       <Field label="Operation"><Select aria-label="Operation" disabled={locked} value={action} onChange={(_, data) => { updateDraft(data.value !== action, () => setAction(data.value as Action)); }}><option value="get">Read</option><option value="list">List</option><option value="save">Save</option><option value="remove">Remove</option></Select></Field>
 *       {action !== 'list' && <Field label="Key"><Input aria-label="Key" disabled={locked} value={key} onChange={(_, data) => { updateDraft(data.value !== key, () => setKey(data.value)); }}/></Field>}
 *       {action === 'save' && <>
 *         <Field label="Value format"><Select aria-label="Value format" disabled={locked} value={format} onChange={(_, data) => { updateDraft(data.value !== format, () => setFormat(data.value as Format)); }}><option value="text">Text</option><option value="json">JSON</option><option value="bigint">Big integer</option><option value="date">Date (ISO)</option></Select></Field>
 *         <Field label="Value"><Input aria-label="Value" disabled={locked} value={value} onChange={(_, data) => { updateDraft(data.value !== value, () => setValue(data.value)); }}/></Field>
 *         <Checkbox label="Replace/clear destination description" disabled={locked} checked={replaceDescription} onChange={(_, data) => { updateDraft((data.checked === true) !== replaceDescription, () => setReplaceDescription(data.checked === true)); }}/>
 *         <Field label="Description"><Input aria-label="Description" disabled={locked || !replaceDescription} value={description} onChange={(_, data) => { updateDraft(data.value !== description, () => setDescription(data.value)); }}/></Field>
 *       </>}
 *       {(action === 'save' || action === 'remove') && <>
 *         <Text as="p">{action} at {session.root}; key {key}; format {format}; value {JSON.stringify(value)}; {replaceDescription ? `description ${JSON.stringify(description)}` : 'preserve destination description'}. Saving may provision or repair the store.</Text>
 *         <Checkbox label="I consent to this exact disposable mutation" disabled={locked || !disposable} checked={consent === target} onChange={(_, data) => { if (!locked && !isLocked() && !confirmationRef.current && valid()) approve(data.checked === true ? target : undefined); }}/>
 *       </>}
 *     </>,
 *     commands: <Button appearance="primary" disabled={!store.isReady || store.isLoading || store.isWriting || locked || (action !== 'list' && !key) || ((action === 'save' || action === 'remove') && (!disposable || approval.current !== target))} onClick={event => run(event.currentTarget)}>Run</Button>,
 *     status: <>{failure ? <Text role="alert">{failure}</Text> : visible?.message ? <Text role="status">{visible.message}</Text> : <Text>Run the selected operation.</Text>}</>,
 *     children: <>
 *       <dl aria-label="Site store state"><dt>Ready</dt><dd>{String(store.isReady)}</dd><dt>Can Write</dt><dd>{String(store.canWrite)}</dd><dt>Loading</dt><dd>{String(Boolean(visible) && store.isLoading)}</dd><dt>Writing</dt><dd>{String(pending)}</dd><dt>Collection root URL</dt><dd>{session.root}</dd><dt>Current web URL</dt><dd>{session.web}</dd></dl>
 *       {!failure && visible?.item && <dl aria-label="Selected site record"><dt>Selected key</dt><dd>{visible.item.key}</dd><dt>Selected value</dt><dd>{shownValue(visible.item)}{visible.item.value === '' ? ' (empty string)' : ''}</dd><dt>Value type</dt><dd>{typeof visible.item.value}</dd><dt>Description</dt><dd>{visible.item.description ?? '(not set)'}</dd></dl>}
 *       {!failure && visible?.items && <><Text>Inventory count: {visible.items.length}</Text><dl aria-label="First ten site records">{visible.items.slice(0, 10).map((item, index) => <React.Fragment key={index}><dt>{item.key}</dt><dd>{shownValue(item)}</dd></React.Fragment>)}</dl></>}
 *       <Dialog open={confirmation !== undefined} onOpenChange={(_, data) => { if (!data.open) confirmation?.cancel(); }}><DialogSurface><DialogBody><DialogTitle>Confirm exact site removal</DialogTitle><DialogContent>Remove {confirmation?.key} at {session.root}. This cannot be undone by Reset.</DialogContent><DialogActions><Button onClick={confirmation?.cancel}>Cancel</Button><Button appearance="primary" onClick={confirmation?.confirm}>Confirm removal</Button></DialogActions></DialogBody></DialogSurface></Dialog>
 *     </>,
 *     diagnostics: <>
 *       <Text as="p">All subsites share the collection-root store. Can Write is an effective indicator, not authorization or a prerequisite for an otherwise consented hook request. SharePoint enforces grants; schema changes may need additional rights. Reads/removal never provision. Normal Save/Remove does not read back automatically. ETags guard the service snapshot, not an earlier Read. Bigint and Date read as strings. Reset hides results; dispatched writes can still complete. This local lock does not serialize other scenarios.</Text>
 *       {!failure && visible && <Accordion collapsible><AccordionItem value="payload"><AccordionHeader>Full site payload</AccordionHeader><AccordionPanel><div role="region" aria-label="Full site payload" tabIndex={0} className={styles.payload}><pre>{JSON.stringify(visible.item ?? visible.items, undefined, 2)}</pre></div></AccordionPanel></AccordionItem></Accordion>}
 *     </>,
 *   };
 *   return present(regions, render);
 * }
 *
 * export function Demo({ render, resetRevision = 0 }: { render?: Presenter; resetRevision?: number } = {}): React.ReactElement {
 *   const page = useSPFxPageContext();
 *   const { client } = useSPFxSPHttpClient();
 *   const info = useSPFxInstanceInfo();
 *   const root = page.site.absoluteUrl.trim().replace(/\/+$/, '');
 *   const web = page.web.absoluteUrl.trim().replace(/\/+$/, '');
 *   const boundary = React.useRef({ revision: resetRevision, root, web, client, page, instance: info.id, generation: {} });
 *   const previous = boundary.current;
 *   if (previous.revision !== resetRevision || previous.root !== root || previous.web !== web || previous.client !== client || previous.page !== page || previous.instance !== info.id) boundary.current = { revision: resetRevision, root, web, client, page, instance: info.id, generation: {} };
 *   const generation = boundary.current.generation;
 *   const [session, setSession] = React.useState<Session>();
 *   const ledger = React.useRef(new Set<Ticket>());
 *   const [, refresh] = React.useState(0);
 *   const mounted = React.useRef(true);
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
 *   const current = (): boolean => mounted.current && boundary.current.generation === generation;
 *   const write = (operation: () => Promise<void>): Promise<void> | undefined => {
 *     if (!current() || ledger.current.size > 0) return undefined;
 *     const ticket: Ticket = {}; ledger.current.add(ticket); refresh(count => count + 1);
 *     const release = (): void => { ledger.current.delete(ticket); if (mounted.current) refresh(count => count + 1); };
 *     try { const promise = operation(); return promise.finally(release); }
 *     catch (error) { release(); throw error; }
 *   };
 *   const active = session?.generation === generation;
 *   const pending = ledger.current.size > 0;
 *   const activate = (): void => {
 *     if (!current() || ledger.current.size > 0 || active || !client || !root) return;
 *     setSession({ generation, root, web });
 *   };
 *   // The outer activation control remains the same DOM node across child teardown.
 *   return <>
 *     <Button disabled={active || pending || !client || !root} onClick={activate}>Activate site store</Button>
 *     {pending && <Text role="status">Previous write still pending</Text>}
 *     {active && session ? <SiteSession session={session} current={current} write={write} pending={pending} isLocked={() => ledger.current.size > 0} render={render}/> : present({ children: <Text>Activate the site store to permit read-only permission initialization. Reset does not cancel server work.</Text> }, render)}
 *   </>;
 * }
 * export const demo = { host: 'webPart', kind: 'write', contentPlacement: 'configuration', observe: 'Activate the site store, run collection-root reads and consented mutations, and compare exact values with the separate read and write errors.' };
 * ```
 */
export function useSPFxSiteKeyValueStore(): SPFxSiteKeyValueStoreResult {
  const { client } = useSPFxSPHttpClient();
  const pageContext = useSPFxPageContext();
  // Read the scalar each render: PageContext may be mutated in place.
  const siteCollectionUrl = (pageContext.site.absoluteUrl || '').trim().replace(/\/+$/, '');
  const service = useMemo(() => client ? createSPFxSiteKeyValueStoreService(client) : undefined, [client]);
  const identity = useMemo(() => ({ service, siteCollectionUrl, reads: 0, writes: 0,
    readId: 0, writeId: 0, permissionId: 0 }), [service, siteCollectionUrl]);
  const identityRef = useRef(identity);
  identityRef.current = identity;
  // Child layout effects can invoke callbacks before this hook's layout effect.
  const mountedRef = useRef(true);
  const initialState = useMemo(() => ({ identity, isLoading: false, isWriting: false,
    error: undefined as Error | undefined, writeError: undefined as Error | undefined, canWrite: false }), [identity]);
  const [state, setState] = useState(initialState);
  const isCurrent = useCallback((): boolean => mountedRef.current && identityRef.current === identity, [identity]);

  // Disable publication as soon as this component unmounts.
  useLayoutEffect(() => {
    mountedRef.current = true;
    return () => { mountedRef.current = false; };
  }, []);

  const refreshPermission = useCallback((): void => {
    if (!service || !siteCollectionUrl || !isCurrent()) return;
    const request = ++identity.permissionId;
    const publish = (canWrite: boolean): void => {
      if (isCurrent() && request === identity.permissionId) {
        setState(previous => isCurrent() ? { ...previous, canWrite } : previous);
      }
    };
    // The service returns false on lookup failures; also guard unexpected rejection.
    service.canCurrentUserWrite(siteCollectionUrl).then(publish, () => publish(false));
  }, [service, siteCollectionUrl, identity, isCurrent]);

  useEffect(() => {
    // A consumer may already have started work in its layout effect.
    setState(previous => previous.identity === identity ? previous : initialState);
    refreshPermission();
  }, [identity, initialState, refreshPermission]);

  const runRead = useCallback(async <T,>(operation: () => Promise<T>, fallback: T): Promise<T> => {
    if (!service || !siteCollectionUrl) throw new Error('SPHttpClient and a nonempty site collection URL are required.');
    const request = ++identity.readId;
    identity.reads++;
    if (isCurrent()) setState(previous => isCurrent() ? { ...(previous.identity === identity ? previous : initialState), isLoading: true, error: undefined } : previous);
    try {
      return await operation();
    } catch (failure) {
      if (isCurrent() && request === identity.readId) {
        const error = failure instanceof Error ? failure : new Error(String(failure));
        setState(previous => isCurrent() ? { ...previous, error } : previous);
      }
      return fallback;
    } finally {
      identity.reads--;
      if (isCurrent()) {
        setState(previous => isCurrent() ? { ...previous, isLoading: identity.reads > 0 } : previous);
        refreshPermission();
      }
    }
  }, [service, siteCollectionUrl, identity, isCurrent, initialState, refreshPermission]);

  const runWrite = useCallback(async (operation: () => Promise<void>): Promise<void> => {
    if (!service || !siteCollectionUrl) throw new Error('SPHttpClient and a nonempty site collection URL are required.');
    const request = ++identity.writeId;
    identity.writes++;
    if (isCurrent()) setState(previous => isCurrent() ? { ...(previous.identity === identity ? previous : initialState), isWriting: true, writeError: undefined } : previous);
    try {
      await operation();
    } catch (failure) {
      if (isCurrent() && request === identity.writeId) {
        const writeError = failure instanceof Error ? failure : new Error(String(failure));
        setState(previous => isCurrent() ? { ...previous, writeError } : previous);
      }
      throw failure;
    } finally {
      identity.writes--;
      if (isCurrent()) {
        setState(previous => isCurrent() ? { ...previous, isWriting: identity.writes > 0 } : previous);
        refreshPermission();
      }
    }
  }, [service, siteCollectionUrl, identity, isCurrent, refreshPermission, initialState]);

  const get = useCallback(<T = unknown>(key: string): Promise<SPFxSiteKeyValueStoreItem<T> | undefined> =>
    runRead(() => service ? service.get<T>(key, siteCollectionUrl) : Promise.reject(new Error('SPHttpClient not available.')), undefined),
  [runRead, service, siteCollectionUrl]);
  const list = useCallback((): Promise<SPFxSiteKeyValueStoreItem<unknown>[]> =>
    runRead(() => service ? service.list(siteCollectionUrl) : Promise.reject(new Error('SPHttpClient not available.')), []),
  [runRead, service, siteCollectionUrl]);
  const save = useCallback(<T = unknown>(key: string, value: T, description?: string): Promise<void> =>
    runWrite(() => service ? service.save<T>(key, value, siteCollectionUrl, description) : Promise.reject(new Error('SPHttpClient not available.'))),
  [runWrite, service, siteCollectionUrl]);
  const remove = useCallback((key: string): Promise<void> =>
    runWrite(() => service ? service.remove(key, siteCollectionUrl) : Promise.reject(new Error('SPHttpClient not available.'))),
  [runWrite, service, siteCollectionUrl]);

  // Hide obsolete state immediately, before effects reset it for the new identity.
  const isLoading = state.identity === identity && state.isLoading;
  const isWriting = state.identity === identity && state.isWriting;
  const error = state.identity === identity ? state.error : undefined;
  const writeError = state.identity === identity ? state.writeError : undefined;
  const canWrite = state.identity === identity && state.canWrite;
  const isReady = service !== undefined && siteCollectionUrl.length > 0;
  return useMemo(() => ({ isLoading, error, isWriting, writeError, canWrite, isReady, get, list, save, remove }),
    [isLoading, error, isWriting, writeError, canWrite, isReady, get, list, save, remove]);
}

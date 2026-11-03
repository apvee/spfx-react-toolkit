/**
 * The legacy tenant list store exposes fallback reads and rejecting mutations.
 *
 * @packageDocumentation
 * @docPage hooks/storage/tenant-key-value-store
 */
// useSPFxTenantKeyValueStore.ts
// Hook to manage tenant-wide key-value pairs using a hidden SharePoint list in the tenant app catalog

import { useState, useCallback, useEffect, useMemo, useRef } from 'react';
import { useAppCatalogUrl } from './useAppCatalogUrl.internal';
import { createSPFxTenantKeyValueStoreService } from '../services/spfx-tenant-key-value-store.service';

// ═══════════════════════════════════════════════════════════════════════════
// TYPES
// ═══════════════════════════════════════════════════════════════════════════

/**
 * A single item from the tenant key-value store
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxTenantKeyValueStoreItem, useSPFxTenantKeyValueStore } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const store = useSPFxTenantKeyValueStore();
 *   const [result, setResult] = React.useState<unknown>();
 *   const [failure, setFailure] = React.useState<string>();
 *   const run = async () => {
 *     setFailure(undefined);
 *     try {
 *       const item: SPFxTenantKeyValueStoreItem<boolean> | undefined = await store.get('featureEnabled');
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
export interface SPFxTenantKeyValueStoreItem<T = unknown> {
    /** Property key (stored in Title column) */
    readonly key: string;
    /** Deserialized property value; the generic does not validate runtime data or preserve all legacy serialization types */
    readonly value: T;
    /** Optional description metadata */
    readonly description: string | undefined;
    /** SharePoint list item ID (internal use) */
    readonly id: number;
}

/**
 * Legacy tenant app-catalog store state and explicit operations with fallback reads and rejecting writes.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxTenantKeyValueStoreResult, useSPFxTenantKeyValueStore } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const store: SPFxTenantKeyValueStoreResult = useSPFxTenantKeyValueStore();
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
export interface SPFxTenantKeyValueStoreResult {
    /**
     * Loading state for read operations (`get`, `list`).
     * True during any active read call.
     */
    readonly isLoading: boolean;

    /**
     * Read error from the latest-started read; cleared when another read starts. Discovery, access, schema, item and paging failures publish this error and return the read fallback.
     */
    readonly error: Error | undefined;

    /**
     * Loading state for write operations (`save`, `remove`).
     * True during any active write call.
     */
    readonly isWriting: boolean;

    /**
     * Write error from the latest-started write; cleared when another write starts.
     */
    readonly writeError: Error | undefined;

    /**
     * Effective Add/Edit/Delete Items indicator for the catalog list. When absent, catalog web Manage Lists is also required. Initially false; checked on discovery and refreshed after operations. Failed checks return false; latest check wins. SharePoint enforces writes and schema repair can require additional permissions.
     */
    readonly canWrite: boolean;

    /**
     * Computed state: true when the SPHttpClient is available and
     * operations can be requested. Does not imply completed discovery, an existing list or write authorization.
     */
    readonly isReady: boolean;

    /**
     * Reads without provisioning or repairing the list. Only a confirmed absent list/key is empty; discovery/access/schema/read failures publish error and return undefined.
     * @typeParam T - Expected deserialized value shape; no runtime validation.
     * @param key - Stored property key.
     * @returns Item or undefined for an absent list/key or failed read; inspect error to distinguish failures.
     * @throws When the SharePoint client is unavailable.
     */
    readonly get: <T = unknown>(key: string) => Promise<SPFxTenantKeyValueStoreItem<T> | undefined>;

    /**
     * Reads every server page in Title order without provisioning. Discovery/access/schema/item/paging failures publish error and return []; no partial results are returned.
     * @returns All items, an empty array for a confirmed absent list, or an empty fallback with error.
     * @throws When the SharePoint client is unavailable.
     */
    readonly list: () => Promise<SPFxTenantKeyValueStoreItem<unknown>[]>;

    /**
     * Save (create or update) a key-value pair.
     * Auto-provisions the hidden list on first write if it doesn't exist.
     *
     * Smart serialization:
     * - Empty strings and strings that parse as JSON are JSON-quoted to preserve their text
     * - bigint becomes quoted decimal text and reads as string without precision loss
     * - Numbers/booleans and legacy null input retain their string representation
     * - Date objects → ISO 8601 string
     * - Objects/arrays → JSON.stringify(value)
     *
     * @param key - Property key
     * @typeParam T - Supplied value type; serialization may change its read-back representation.
     * @param value - Value to store
     * @param description - Omitted or undefined preserves an existing description; empty string clears it.
     *
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { useSPFxTenantKeyValueStore } from '@apvee/spfx-react-toolkit';
     *
     * function Example() {
     *   const store = useSPFxTenantKeyValueStore();
     *   const [result, setResult] = React.useState<unknown>();
     *   const [failure, setFailure] = React.useState<string>();
     *   const run = async () => {
     *     setFailure(undefined);
     *     try {
     *       await store.save<string>('apiEndpoint', 'https://api.example.com', 'Production API');
     *       // Save the complex value only after the first write succeeds.
     *       await store.save('featureFlags', { enableChat: true, maxUsers: 1000 });
     *       setResult('Saved settings');
     *     } catch (failure) {
     *       setFailure(failure instanceof Error ? failure.message : String(failure));
     *     }
     *   };
     *   return <>
     *     <button disabled={!store.isReady || store.isWriting} onClick={() => void run()}>Save settings</button>
     *     {(failure ?? store.writeError?.message) && <span role="alert">{failure ?? store.writeError?.message}</span>}
     *     <pre>{JSON.stringify(result)}</pre>
     *   </>;
     * }
     * ```
     * @returns Resolves after a successful write; updates preserve an omitted description and empty string clears it.
     * @throws Client, catalog discovery, schema, serialization, missing ETag or write/conflict failures; operation failures also publish writeError. No wildcard overwrite is performed.
     */
    readonly save: <T = unknown>(key: string, value: T, description?: string) => Promise<void>;

    /**
     * Deletes without provisioning or schema repair. A confirmed absent list/key is a no-op; changes since the service snapshot reject using ETag preconditions.
     * @param key - Key to remove.
     * @returns Resolves after successful deletion or absent key.
     * @throws Client, discovery, schema, missing ETag or deletion/conflict failures, also captured in writeError.
     */
    readonly remove: (key: string) => Promise<void>;
}

// ═══════════════════════════════════════════════════════════════════════════
// HOOK
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Manages the hidden TenantKeyValueStore list in the tenant app catalog beneath an SPFx provider.
 * Mount discovers the catalog and checks effective permissions without provisioning. Reads/removes
 * never create or repair schema; save ensures validated fields and unique indexed keys. isReady
 * means client availability only. canWrite is an asynchronous effective-permission indicator.
 * New string writes preserve text and bigint reads as exact decimal text; Date remains ISO text.
 * Existing raw data retains its legacy JSON-first interpretation, and T performs no validation.
 * Read failures publish error and return fallbacks; write failures publish writeError and reject.
 * ETag preconditions detect changes during mutations, not edits since a caller's earlier get().
 * Context changes reset state; obsolete/unmounted completions cannot publish state.
 * @returns Client readiness, read/write state, permission indicator and CRUD methods.
 * @throws When called outside an SPFx provider.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxTenantKeyValueStore } from '@apvee/spfx-react-toolkit';
 *
 * function TenantSetting() {
 *   const store = useSPFxTenantKeyValueStore();
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
 * @useWhen legacy tenant-wide key/value configuration needs tracked list CRUD and permission indicators.
 *
 * @avoidWhen new configuration needs rejecting reads and explicit mutation scopes; use useSPFxConfigurationStore.
 *
 *
 * @example Save and read numeric-looking text
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxTenantKeyValueStore } from '@apvee/spfx-react-toolkit';
 *
 * function TenantConfigEditor() {
 *   const { save, get, isReady, isWriting, error, writeError } = useSPFxTenantKeyValueStore();
 *   const [text, setText] = React.useState<string>('');
 *   const run = async (): Promise<void> => {
 *     try {
 *       await save('my-app:version', '123', 'Application version');
 *       const item = await get<unknown>('my-app:version');
 *       if (typeof item?.value === 'string') setText(item.value);
 *     } catch {
 *       // Write failures are rendered through writeError; reads use error + fallback.
 *     }
 *   };
 *   return <>
 *     <button disabled={!isReady || isWriting} onClick={run}>Save and read</button>
 *     <p>{text}</p>
 *     {(error ?? writeError) && <p role="alert">{(error ?? writeError)?.message}</p>}
 *   </>;
 * }
 * ```
 *
 * @example Explore tenant records and exact text persistence
 * ```tsx
 * import * as React from 'react';
 * import { makeStyles, Accordion, AccordionHeader, AccordionItem, AccordionPanel, Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, Select, Text } from '@fluentui/react-components';
 * import { useSPFxInstanceInfo, useSPFxPageContext, useSPFxTenantKeyValueStore, useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
 * import type { SPFxTenantKeyValueStoreItem } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Presenter = (regions: DemoRegions) => React.ReactElement;
 * type Session = { generation: object; root: string; web: string };
 * type Action = 'get' | 'list' | 'save' | 'remove' | 'roundtrip';
 * type Format = 'text' | 'json' | 'bigint' | 'date';
 * const useStyles = makeStyles({ payload: { overflowX: 'auto', maxWidth: '100%', minWidth: 0 } });
 * type Ticket = object;
 *
 * function present(regions: DemoRegions, render?: Presenter): React.ReactElement {
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 *
 * // This bounded child owns the hook session. Reset deliberately unmounts it.
 * function TenantSession({ session, current, write, pending, isLocked, render }: {
 *   session: Session; current: () => boolean; write: (operation: () => Promise<void>) => Promise<void> | undefined;
 *   pending: boolean; isLocked: () => boolean; render?: Presenter;
 * }): React.ReactElement {
 *   const store = useSPFxTenantKeyValueStore();
 *   const styles = useStyles();
 *   const [action, setAction] = React.useState<Action>('get');
 *   const [key, setKey] = React.useState('spfx-toolkit-tenant-demo-disposable');
 *   const [value, setValue] = React.useState('');
 *   const [format, setFormat] = React.useState<Format>('text');
 *   const [replaceDescription, setReplaceDescription] = React.useState(false);
 *   const [description, setDescription] = React.useState('Disposable SPFx React Toolkit tenant demo item');
 *   const [consent, setConsent] = React.useState<object>();
 *   const approval = React.useRef<object>();
 *   approval.current = consent;
 *   const approve = (next: object | undefined): void => { approval.current = next; setConsent(next); };
 *   const [outcome, setOutcome] = React.useState<{ boundary: object; action: Action; message?: string; failure?: string; persisted?: boolean; reading?: boolean; item?: SPFxTenantKeyValueStoreItem; items?: SPFxTenantKeyValueStoreItem[] }>();
 *   const boundary = React.useRef({ fingerprint: '', token: {} });
 *   const fingerprint = JSON.stringify([action, key, value, format, replaceDescription, description]);
 *   if (boundary.current.fingerprint !== fingerprint) boundary.current = { fingerprint, token: {} };
 *   const target = boundary.current.token;
 *   const latest = React.useRef(0);
 *   const sequence = React.useRef<object>();
 *   const [, refreshSession] = React.useState(0);
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
 *   const disposable = key.startsWith('spfx-toolkit-tenant-demo-') && key.length > 'spfx-toolkit-tenant-demo-'.length;
 *   const visible = valid() && outcome?.boundary === target ? outcome : undefined;
 *   const reading = visible?.action === 'get' || visible?.action === 'list' || visible?.reading === true;
 *   // The public read resolves a fallback. Its CURRENT rendered error takes precedence.
 *   const failure = visible ? visible.failure ?? (reading ? store.error?.message : store.writeError?.message) : undefined;
 *   const locked = pending || confirmation !== undefined || sequence.current === target;
 *   const edit = (change: () => void): void => {
 *     if (!valid() || locked || sequence.current === target || isLocked() || confirmationRef.current) return;
 *     // Retire captured consent/commands synchronously, before a batched React commit.
 *     boundary.current = { ...boundary.current, token: {} }; approve(undefined); change(); refreshSession(count => count + 1);
 *   };
 *
 *   const run = async (trigger: HTMLButtonElement): Promise<void> => {
 *     if (!valid() || sequence.current === target || isLocked() || pending || confirmationRef.current || !store.isReady || store.isLoading || store.isWriting) return;
 *     if (action !== 'list' && !key) return;
 *     const mutating = action === 'save' || action === 'remove' || action === 'roundtrip';
 *     if (mutating && (!disposable || approval.current !== target)) return;
 *     const operation = ++latest.current;
 *     const owns = (): boolean => valid() && latest.current === operation;
 *     const execute = async (remove?: () => Promise<void>): Promise<void> => {
 *       if (!owns() || isLocked()) return;
 *       if (mutating) sequence.current = target;
 *       setOutcome({ boundary: target, action });
 *       try {
 *         if (action === 'get') {
 *           const item = await store.get(key);
 *           if (owns()) setOutcome({ boundary: target, action, item, message: item ? `Found ${key}.` : `${key} not found.` });
 *         } else if (action === 'list') {
 *           const items = await store.list();
 *           if (owns()) setOutcome({ boundary: target, action, items, message: `Loaded ${items.length} tenant key-value item(s).` });
 *         } else {
 *           let input: unknown = value;
 *           if (action === 'save') {
 *             if (format === 'json') input = JSON.parse(value);
 *             else if (format === 'bigint') input = BigInt(value);
 *             else if (format === 'date') { const date = new Date(value); if (!Number.isFinite(date.getTime())) throw new Error('Enter a valid Date.'); input = date; }
 *           }
 *           // The parent retains the ORIGINAL public promise through Reset/child unmount.
 *           const promise = write(remove ?? (() => action === 'roundtrip' ? store.save(key, '123', 'Application version') : replaceDescription ? store.save(key, input, description) : store.save(key, input)));
 *           if (!promise) return;
 *           await promise;
 *           if (!owns()) return;
 *           approve(undefined);
 *           if (action === 'roundtrip') {
 *             setOutcome({ boundary: target, action, persisted: true, reading: true });
 *             const item = await store.get<unknown>(key);
 *             if (owns()) setOutcome({ boundary: target, action, persisted: true, reading: true, item, message: item?.value === '123' && typeof item.value === 'string' ? 'Read back exact text 123 (string).' : 'Readback did not contain exact text 123.' });
 *           } else setOutcome({ boundary: target, action, message: `${action === 'save' ? 'Saved' : 'Removed'} ${key}.` });
 *         }
 *       } catch (error) {
 *         if (owns()) setOutcome(previous => ({ ...previous, boundary: target, action, failure: error instanceof Error ? error.message : String(error) }));
 *       } finally { if (sequence.current === target) sequence.current = undefined; if (mounted.current) refreshSession(count => count + 1); }
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
 *   const shownValue = (item: SPFxTenantKeyValueStoreItem): string => typeof item.value === 'string' ? JSON.stringify(item.value) : JSON.stringify(item.value) ?? String(item.value);
 *   const regions: DemoRegions = {
 *     controls: <>
 *       <Field label="Operation"><Select aria-label="Operation" disabled={locked} value={action} onChange={(_, data) => edit(() => setAction(data.value as Action))}><option value="get">Read</option><option value="list">List</option><option value="save">Save</option><option value="remove">Remove</option><option value="roundtrip">Save and read numeric-looking text</option></Select></Field>
 *       {action !== 'list' && <Field label="Key"><Input aria-label="Key" disabled={locked} value={key} onChange={(_, data) => edit(() => setKey(data.value))}/></Field>}
 *       {action === 'save' && <>
 *         <Field label="Value format"><Select aria-label="Value format" disabled={locked} value={format} onChange={(_, data) => edit(() => setFormat(data.value as Format))}><option value="text">Text</option><option value="json">JSON</option><option value="bigint">Big integer</option><option value="date">Date (ISO)</option></Select></Field>
 *         <Field label="Value"><Input aria-label="Value" disabled={locked} value={value} onChange={(_, data) => edit(() => setValue(data.value))}/></Field>
 *         <Checkbox label="Replace/clear destination description" disabled={locked} checked={replaceDescription} onChange={(_, data) => edit(() => setReplaceDescription(data.checked === true))}/>
 *         <Field label="Description"><Input aria-label="Description" disabled={locked || !replaceDescription} value={description} onChange={(_, data) => edit(() => setDescription(data.value))}/></Field>
 *       </>}
 *       {(action === 'save' || action === 'remove' || action === 'roundtrip') && <>
 *         <Text as="p">{action} in the tenant app-catalog store resolved internally for this captured discovery session; key {key}; {action === 'roundtrip' ? 'exact text "123" and description "Application version"' : `format ${format}; value ${JSON.stringify(value)}; ${replaceDescription ? `description ${JSON.stringify(description)}` : 'preserve destination description'}`}. Saving may provision or repair the store.</Text>
 *         <Checkbox label="I consent to this exact disposable mutation" disabled={locked || !disposable} checked={consent === target} onChange={(_, data) => { if (!locked && sequence.current !== target && valid() && !isLocked() && !confirmationRef.current && valid()) approve(data.checked === true ? target : undefined); }}/>
 *       </>}
 *     </>,
 *     commands: <Button appearance="primary" disabled={!store.isReady || store.isLoading || store.isWriting || locked || (action !== 'list' && !key) || ((action === 'save' || action === 'remove' || action === 'roundtrip') && (!disposable || approval.current !== target))} onClick={event => run(event.currentTarget)}>{action === 'roundtrip' ? 'Save and read numeric-looking text' : 'Run'}</Button>,
 *     status: <>{visible?.persisted && <Text role="status">Saved {key}. The subsequent read is a separate step.</Text>}{failure ? <Text role="alert">{failure}</Text> : visible?.message ? <Text role="status">{visible.message}</Text> : <Text>Run the selected operation.</Text>}</>,
 *     children: <>
 *       <dl aria-label="Tenant store state"><dt>Client ready</dt><dd>{String(store.isReady)}</dd><dt>Can Write</dt><dd>{String(store.canWrite)}</dd><dt>Loading</dt><dd>{String(Boolean(visible) && store.isLoading)}</dd><dt>Writing</dt><dd>{String(pending)}</dd><dt>Collection root URL</dt><dd>{session.root}</dd><dt>Discovery web URL</dt><dd>{session.web}</dd></dl>
 *       {!failure && visible?.item && <dl aria-label="Selected tenant record"><dt>Selected key</dt><dd>{visible.item.key}</dd><dt>Selected value</dt><dd>{shownValue(visible.item)}{visible.item.value === '' ? ' (empty string)' : ''}</dd><dt>Value type</dt><dd>{typeof visible.item.value}</dd><dt>Description</dt><dd>{visible.item.description ?? '(not set)'}</dd></dl>}
 *       {!failure && visible?.items && <><Text>Inventory count: {visible.items.length}</Text><dl aria-label="First ten tenant records">{visible.items.slice(0, 10).map((item, index) => <React.Fragment key={index}><dt>{item.key}</dt><dd>{shownValue(item)}</dd></React.Fragment>)}</dl></>}
 *       <Dialog open={confirmation !== undefined} onOpenChange={(_, data) => { if (!data.open) confirmation?.cancel(); }}><DialogSurface><DialogBody><DialogTitle>Confirm exact tenant removal</DialogTitle><DialogContent>Remove {confirmation?.key} in the tenant app-catalog store resolved internally for this captured discovery session. This cannot be undone by Reset.</DialogContent><DialogActions><Button onClick={confirmation?.cancel}>Cancel</Button><Button appearance="primary" onClick={confirmation?.confirm}>Confirm removal</Button></DialogActions></DialogBody></DialogSurface></Dialog>
 *     </>,
 *     diagnostics: <>
 *       <Text as="p">All sites share TenantKeyValueStore in the tenant app catalog. Physical catalog URL is unavailable through this hook API. The discovery-web and collection-root URLs identify the captured client/PageContext/host session; neither is a catalog address. Client ready means client availability, not completed discovery or authorization. Can Write is an effective indicator, not authorization or a prerequisite for an otherwise consented hook request. SharePoint enforces grants; schema changes may need additional rights. Reads/removal never provision. Normal Save/Remove does not read back automatically. ETags guard the service snapshot, not an earlier Read. Bigint and Date read as strings. Reset hides results; dispatched writes can still complete. This local lock does not serialize other scenarios.</Text>
 *       {!failure && visible && <Accordion collapsible><AccordionItem value="payload"><AccordionHeader>Full tenant payload</AccordionHeader><AccordionPanel><div role="region" aria-label="Full tenant payload" tabIndex={0} className={styles.payload}><pre>{JSON.stringify(visible.item ?? visible.items, undefined, 2)}</pre></div></AccordionPanel></AccordionItem></Accordion>}
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
 *   const current = (): boolean => mounted.current && boundary.current.generation === generation && page.web.absoluteUrl.trim().replace(/\/+$/, '') === web && page.site.absoluteUrl.trim().replace(/\/+$/, '') === root;
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
 *     <Button disabled={active || pending || !client || !root} onClick={activate}>Activate tenant store</Button>
 *     {pending && <Text role="status">Previous write still pending</Text>}
 *     {active && session ? <TenantSession session={session} current={current} write={write} pending={pending} isLocked={() => ledger.current.size > 0} render={render}/> : present({ children: <Text>Activate the tenant store to permit read-only catalog discovery and effective permission initialization. Reset does not cancel server work.</Text> }, render)}
 *   </>;
 * }
 * export const demo = { host: 'webPart', kind: 'write', contentPlacement: 'configuration', observe: 'Activate the tenant store, preserve numeric-looking and empty text, and run physical reads or consented mutations with current-only results.' };
 * ```
 */
export function useSPFxTenantKeyValueStore(): SPFxTenantKeyValueStoreResult {
    const { spHttpClient, discoverAppCatalogUrl, checkWritePermission, isMountedRef } = useAppCatalogUrl();
    const tenantKeyValueStoreService = useMemo(
        () => spHttpClient ? createSPFxTenantKeyValueStoreService(spHttpClient) : undefined,
        [spHttpClient]
    );

    // State management
    const [isLoading, setIsLoading] = useState<boolean>(false);
    const [error, setError] = useState<Error | undefined>(undefined);
    const [isWriting, setIsWriting] = useState<boolean>(false);
    const [writeError, setWriteError] = useState<Error | undefined>(undefined);
    const [canWrite, setCanWrite] = useState<boolean>(false);

    const operationRef = useRef({ service: tenantKeyValueStoreService, discover: discoverAppCatalogUrl,
        reads: 0, writes: 0, readId: 0, writeId: 0, permissionId: 0 });
    if (operationRef.current.service !== tenantKeyValueStoreService ||
        operationRef.current.discover !== discoverAppCatalogUrl) {
        operationRef.current = { service: tenantKeyValueStoreService, discover: discoverAppCatalogUrl,
            reads: 0, writes: 0, readId: 0, writeId: 0, permissionId: 0 };
    }
    const currentOperation = operationRef.current;
    useEffect(() => {
        setIsLoading(false);
        setIsWriting(false);
        setError(undefined);
        setWriteError(undefined);
        setCanWrite(false);
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl]);

    // ─────────────────────────────────────────────────────────────────────────
    // Internal helpers
    // ─────────────────────────────────────────────────────────────────────────



    /**
     * Refresh the effective-permission indicator; only the latest check may publish.
     */
    const ensurePermissionChecked = useCallback((catalogUrl: string): void => {
        const identity = currentOperation;
        const request = ++identity.permissionId;
        checkWritePermission(catalogUrl)
            .then(hasPermission => {
                if (isMountedRef.current && operationRef.current === identity && request === identity.permissionId) {
                    setCanWrite(hasPermission);
                }
            })
            .catch(() => {
                if (isMountedRef.current && operationRef.current === identity && request === identity.permissionId) {
                    setCanWrite(false);
                }
            });
    }, [checkWritePermission, isMountedRef, currentOperation]);

    // Discovery and permission lookup are read-only; provisioning belongs to save.
    useEffect(() => {
        if (!tenantKeyValueStoreService) return;

        discoverAppCatalogUrl()
            .then(catalogUrl => {
                if (!isMountedRef.current || operationRef.current !== currentOperation) return;
                ensurePermissionChecked(catalogUrl);
            })
            .catch(() => {
                // Explicit operations retry discovery and publish failures in their own error channel.
            });
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, ensurePermissionChecked, isMountedRef, currentOperation]);

    // ─────────────────────────────────────────────────────────────────────────
    // Public operations
    // ─────────────────────────────────────────────────────────────────────────

    const get = useCallback(async <T = unknown>(key: string): Promise<SPFxTenantKeyValueStoreItem<T> | undefined> => {
        if (!tenantKeyValueStoreService) {
            throw new Error('SPHttpClient not available. Cannot get property.');
        }

        const identity = currentOperation;
        const request = ++identity.readId;
        const isCurrent = (): boolean => isMountedRef.current && operationRef.current === identity &&
            identity.service === tenantKeyValueStoreService;
        identity.reads++;
        if (isCurrent()) { setIsLoading(true); setError(undefined); }

        let catalogUrl: string | undefined;
        try {
            catalogUrl = await discoverAppCatalogUrl();

            return await tenantKeyValueStoreService.get<T>(key, catalogUrl);
        } catch (err) {
            if (isCurrent() && request === identity.readId) {
                const capturedError = err instanceof Error ? err : new Error(String(err));
                setError(capturedError);
                console.error('Failed to get tenant key-value:', capturedError);
            }
            return undefined;
        } finally {
            identity.reads--;
            if (isCurrent()) {
                setIsLoading(identity.reads > 0);
                if (catalogUrl) ensurePermissionChecked(catalogUrl);
            }
        }
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef, currentOperation, ensurePermissionChecked]);

    const list = useCallback(async (): Promise<SPFxTenantKeyValueStoreItem<unknown>[]> => {
        if (!tenantKeyValueStoreService) {
            throw new Error('SPHttpClient not available. Cannot list properties.');
        }

        const identity = currentOperation;
        const request = ++identity.readId;
        const isCurrent = (): boolean => isMountedRef.current && operationRef.current === identity &&
            identity.service === tenantKeyValueStoreService;
        identity.reads++;
        if (isCurrent()) { setIsLoading(true); setError(undefined); }

        let catalogUrl: string | undefined;
        try {
            catalogUrl = await discoverAppCatalogUrl();

            return await tenantKeyValueStoreService.list(catalogUrl);
        } catch (err) {
            if (isCurrent() && request === identity.readId) {
                const capturedError = err instanceof Error ? err : new Error(String(err));
                setError(capturedError);
                console.error('Failed to list tenant key-values:', capturedError);
            }
            return [];
        } finally {
            identity.reads--;
            if (isCurrent()) {
                setIsLoading(identity.reads > 0);
                if (catalogUrl) ensurePermissionChecked(catalogUrl);
            }
        }
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef, currentOperation, ensurePermissionChecked]);

    const save = useCallback(async <T = unknown>(
        key: string,
        value: T,
        description?: string
    ): Promise<void> => {
        if (!tenantKeyValueStoreService) {
            throw new Error('SPHttpClient not available. Cannot save property.');
        }

        const identity = currentOperation;
        const request = ++identity.writeId;
        const isCurrent = (): boolean => isMountedRef.current && operationRef.current === identity &&
            identity.service === tenantKeyValueStoreService;
        identity.writes++;
        if (isCurrent()) { setIsWriting(true); setWriteError(undefined); }

        let catalogUrl: string | undefined;
        try {
            catalogUrl = await discoverAppCatalogUrl();
            await tenantKeyValueStoreService.save<T>(key, value, catalogUrl, description);
        } catch (err) {
            if (isCurrent() && request === identity.writeId) {
                const capturedError = err instanceof Error ? err : new Error(String(err));
                setWriteError(capturedError);
                console.error('Failed to save tenant key-value:', capturedError);
            }
            throw err;
        } finally {
            identity.writes--;
            if (isCurrent()) {
                setIsWriting(identity.writes > 0);
                if (catalogUrl) ensurePermissionChecked(catalogUrl);
            }
        }
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef, currentOperation, ensurePermissionChecked]);

    const remove = useCallback(async (key: string): Promise<void> => {
        if (!tenantKeyValueStoreService) {
            throw new Error('SPHttpClient not available. Cannot remove property.');
        }

        const identity = currentOperation;
        const request = ++identity.writeId;
        const isCurrent = (): boolean => isMountedRef.current && operationRef.current === identity &&
            identity.service === tenantKeyValueStoreService;
        identity.writes++;
        if (isCurrent()) { setIsWriting(true); setWriteError(undefined); }

        let catalogUrl: string | undefined;
        try {
            catalogUrl = await discoverAppCatalogUrl();
            await tenantKeyValueStoreService.remove(key, catalogUrl);
        } catch (err) {
            if (isCurrent() && request === identity.writeId) {
                const capturedError = err instanceof Error ? err : new Error(String(err));
                setWriteError(capturedError);
                console.error('Failed to remove tenant key-value:', capturedError);
            }
            throw err;
        } finally {
            identity.writes--;
            if (isCurrent()) {
                setIsWriting(identity.writes > 0);
                if (catalogUrl) ensurePermissionChecked(catalogUrl);
            }
        }
    }, [tenantKeyValueStoreService, discoverAppCatalogUrl, isMountedRef, currentOperation, ensurePermissionChecked]);

    // Computed: ready when client is available
    const isReady = spHttpClient !== undefined;

    return useMemo(() => ({
        isLoading,
        error,
        isWriting,
        writeError,
        canWrite,
        isReady,
        get,
        list,
        save,
        remove,
    }), [isLoading, error, isWriting, writeError, canWrite, isReady, get, list, save, remove]);
}

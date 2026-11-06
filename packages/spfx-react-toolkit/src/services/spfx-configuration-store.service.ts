/**
 * Compose whole-record site overrides over tenant defaults with explicit physical mutation targets.
 *
 * @packageDocumentation
 * @docPage services/configuration-store
 */
import type { SPHttpClient } from '@microsoft/sp-http';
import { createSPFxAppCatalogService } from './spfx-app-catalog.service';
import { createSPFxSiteKeyValueStoreService } from './spfx-site-key-value-store.service';
import { createSPFxTenantKeyValueStoreService } from './spfx-tenant-key-value-store.service';

/**
 * Physical configuration destination. Site means the collection root, shared by its subsites.
 * @example
 * ```ts
 * import { SPFxConfigurationStoreScope } from '@apvee/spfx-react-toolkit';
 *
 * const scope: SPFxConfigurationStoreScope = 'site';
 * ```
 */
export type SPFxConfigurationStoreScope = 'tenant' | 'site';

/**
 * Allowed physical stores and effective-read policy. Layered reads site first, then tenant on confirmed absence.
 * @example
 * ```ts
 * import { SPFxConfigurationStoreMode } from '@apvee/spfx-react-toolkit';
 *
 * const mode: SPFxConfigurationStoreMode = 'layered';
 * ```
 */
export type SPFxConfigurationStoreMode = SPFxConfigurationStoreScope | 'layered';

/**
 * Required mode and its addressable dependencies; no defaults. URLs are snapshotted and normalized at construction. Validation checks nonempty strings, not HTTP URL validity or whether a supplied site URL is a collection root.
 * @example
 * ```ts
 * import type { WebPartContext } from '@microsoft/sp-webpart-base';
 * import { SPFxConfigurationStoreServiceOptions } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(context: WebPartContext) {
 *   const options: SPFxConfigurationStoreServiceOptions = {
 *     mode: 'layered', discoveryWebUrl: context.pageContext.web.absoluteUrl,
 *     siteCollectionUrl: context.pageContext.site.absoluteUrl,
 *   };
 * }
 * ```
 */
export type SPFxConfigurationStoreServiceOptions =
  | {
    /** Only the tenant store is available. */
    readonly mode: 'tenant';
    /** Current web URL used for lazy tenant catalog discovery; not the catalog URL. */
    readonly discoveryWebUrl: string;
    /** Omit in tenant mode; an explicitly undefined value is also ignored at runtime. */
    readonly siteCollectionUrl?: never;
  }
  | {
    /** Only the site collection store is available. */
    readonly mode: 'site';
    /** Absolute collection-root URL, never a subsite URL. */
    readonly siteCollectionUrl: string;
    /** Omit in site mode; an explicitly undefined value is also ignored at runtime. */
    readonly discoveryWebUrl?: never;
  }
  | {
    /** Site values override complete tenant values by key. */
    readonly mode: 'layered';
    /** Current web URL used for lazy tenant catalog discovery. */
    readonly discoveryWebUrl: string;
    /** Absolute collection-root URL. */
    readonly siteCollectionUrl: string;
  };

/**
 * Hook configuration. The matching SPFx provider supplies the client and mode-specific URLs.
 * @example
 * ```ts
 * import { SPFxConfigurationStoreOptions } from '@apvee/spfx-react-toolkit';
 *
 * const options: SPFxConfigurationStoreOptions = { mode: 'site' };
 * ```
 */
export interface SPFxConfigurationStoreOptions {
  /** Required read policy and allowed physical scopes; no default. */
  readonly mode: SPFxConfigurationStoreMode;
}

/**
 * Explicit mutation destination, independent of a previously read value's source.
 * @example
 * ```ts
 * import { SPFxConfigurationStoreTarget } from '@apvee/spfx-react-toolkit';
 *
 * const target: SPFxConfigurationStoreTarget = { scope: 'site' };
 * ```
 */
export interface SPFxConfigurationStoreTarget {
  /** Required physical scope; must be enabled by the configured mode. */
  readonly scope: SPFxConfigurationStoreScope;
}

/**
 * Save destination and destination-local metadata.
 * @example
 * ```ts
 * import { SPFxConfigurationStoreSaveOptions } from '@apvee/spfx-react-toolkit';
 *
 * const options: SPFxConfigurationStoreSaveOptions = { scope: 'site', description: '' };
 * ```
 */
export interface SPFxConfigurationStoreSaveOptions extends SPFxConfigurationStoreTarget {
  /** Omitted preserves existing destination metadata; empty string clears it. Never inherits metadata from another scope. */
  readonly description?: string;
}

/**
 * One physical record and its provenance. T describes expected data; it does not validate stored content.
 * @typeParam T - Expected deserialized value; defaults to unknown, without runtime validation.
 * @example
 * ```ts
 * import { SPFxConfigurationStoreItem, SPFxConfigurationStoreService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(store: SPFxConfigurationStoreService) {
 *   const item: SPFxConfigurationStoreItem<number> | undefined = await store.get<number>('my-app:pageSize');
 *   console.log(item?.value, item?.source);
 * }
 * ```
 */
export interface SPFxConfigurationStoreItem<T = unknown> {
  /** Stored key with server-provided casing, without normalization. */
  readonly key: string;
  /** Whole deserialized value, retaining the source store's historical interpretation. */
  readonly value: T;
  /** Metadata from the same physical record; absent/empty descriptions read as undefined. */
  readonly description: string | undefined;
  /** Source-list item ID, not a globally unique identity. */
  readonly id: number;
  /** Physical store that supplied this record; does not route later writes. */
  readonly source: SPFxConfigurationStoreScope;
}

/**
 * Configuration operations without React. All failed operations reject; only confirmed absence is empty. False, zero, empty text and legacy null-valued records remain present site overrides. Keys are forwarded without trimming or namespacing; SharePoint determines equality.
 * @example
 * ```ts
 * import { SPFxConfigurationStoreService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(store: SPFxConfigurationStoreService) {
 *   await store.save('my-app:pageSize', 50, { scope: 'site' });
 *   const item = await store.get<number>('my-app:pageSize');
 *   await store.remove('my-app:pageSize', { scope: 'site' });
 * }
 * ```
 */
export interface SPFxConfigurationStoreService {
  /**
   * Reads the mode's effective record without provisioning. Layered reads site first and consults tenant only after confirmed site absence.
   * @typeParam T - Expected deserialized value; defaults to unknown and performs no runtime validation.
   * @param key - String key forwarded unchanged; equality is delegated to SharePoint.
   * @returns Whole winning record with source, or undefined for confirmed absence in all required scopes.
   * @throws Discovery, access, schema or lookup failures; a failed site read never enables tenant fallback.
   */
  get: <T = unknown>(key: string) => Promise<SPFxConfigurationStoreItem<T> | undefined>;
  /**
   * Inventories one physical store, preserving paging order, casing and legacy duplicates; never merges scopes.
   * @param scope - Required enabled physical scope.
   * @returns Complete physical inventory, or [] for confirmed absent list. No partial results on page failures.
   * @throws Disabled/invalid scope, discovery, schema or read failures.
   */
  listScope: (scope: SPFxConfigurationStoreScope) => Promise<SPFxConfigurationStoreItem<unknown>[]>;
  /**
   * Replaces the complete value only at the explicit destination, provisioning that store if necessary. No JSON-property merge or automatic read follows.
   * @typeParam T - Input value shape; defaults to unknown and does not guarantee serializability.
   * @param key - String key forwarded unchanged.
   * @param value - Existing store encoding applies: new text is preserved, bigint reads as decimal string and Date as ISO text. T is not validation.
   * @param options - Required enabled scope and optional destination description.
   * @returns Resolves after successful destination mutation.
   * @throws Invalid target/input, serialization, discovery, setup or mutation failures, including observed-ETag conflicts. No failover or additional retry.
   */
  save: <T = unknown>(key: string, value: T, options: SPFxConfigurationStoreSaveOptions) => Promise<void>;
  /**
   * Removes only the selected physical record without provisioning. Removing a site override reveals tenant data on the next effective read.
   * @param key - String key forwarded unchanged.
   * @param target - Required enabled physical scope, even in single-store modes.
   * @returns Resolves after deletion or confirmed absence; never deletes both scopes.
   * @throws Invalid target, discovery or mutation failures/conflicts.
   */
  remove: (key: string, target: SPFxConfigurationStoreTarget) => Promise<void>;
}

/**
 * Creates an isolated facade over existing list stores. Construction performs no requests. Tenant discovery is lazy and shared while pending within this instance; successful URLs use the existing catalog cache. Site-only operations never need catalog access.
 * Existing APIs/data, nullable-Note legacy differences, serialization and ETag preconditions remain unchanged. No cross-store transaction, value cache or permission indicator is provided.
 * @useWhen You need whole-record site overrides over tenant defaults with an explicit destination for every mutation.
 * @param client - Caller-owned authenticated SharePoint client.
 * @param options - Required mode and relevant URLs; copied at construction, never read again from the caller's object.
 * @returns Effective reads, physical inventories and explicitly scoped mutations.
 * @throws Invalid mode or missing/empty required URLs synchronously. Operation failures reject their Promises. Tenant serialization failures can follow discovery, but never setup/item writes.
 * @example Basic usage
 * ```ts
 * import type { WebPartContext } from '@microsoft/sp-webpart-base';
 * import { createSPFxConfigurationStoreService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(context: WebPartContext) {
 *   const store = createSPFxConfigurationStoreService(context.spHttpClient, {
 *     mode: 'layered',
 *     discoveryWebUrl: context.pageContext.web.absoluteUrl,
 *     siteCollectionUrl: context.pageContext.site.absoluteUrl
 *   });
 *   await store.save('my-app:pageSize', 20, { scope: 'tenant' });
 *   await store.save('my-app:pageSize', 50, { scope: 'site' });
 *   console.log((await store.get<number>('my-app:pageSize'))?.source); // site
 * }
 * ```
 *
 *
 * @example Tenant default site override inventory and removal
 * ```ts
 * import type { SPHttpClient } from '@microsoft/sp-http';
 * import type { PageContext } from '@microsoft/sp-page-context';
 *
 * import { createSPFxConfigurationStoreService } from '@apvee/spfx-react-toolkit';
 * async function demonstrate(context: { spHttpClient: SPHttpClient; pageContext: PageContext }) {
 *   const store = createSPFxConfigurationStoreService(context.spHttpClient, {
 *     mode: 'layered',
 *     discoveryWebUrl: context.pageContext.web.absoluteUrl,
 *     siteCollectionUrl: context.pageContext.site.absoluteUrl,
 *   });
 *   await store.save('my-app:pageSize', 20, { scope: 'tenant' });
 *   await store.save('my-app:pageSize', 50, { scope: 'site', description: 'Local override' });
 *   const item = await store.get<number>('my-app:pageSize'); // value 50, source site
 *   const siteRows = await store.listScope('site'); // physical inventory, not effective merge
 *   await store.remove('my-app:pageSize', { scope: 'site' });
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
 * @scenario Compose tenant defaults and site overrides with the service
 * ```tsx
 * import * as React from 'react';
 * import { Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, Text, makeStyles } from '@fluentui/react-components';
 * import { createSPFxConfigurationStoreService, useSPFxInstanceInfo, useSPFxPageContext, useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
 * import type { SPFxConfigurationStoreItem } from '@apvee/spfx-react-toolkit';
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Ticket = object;
 * type Outcome = { token: object; steps: string[]; message?: string; failure?: string; item?: SPFxConfigurationStoreItem; items?: SPFxConfigurationStoreItem[] };
 * const useStyles = makeStyles({ payload: { overflowX: 'auto', maxWidth: '100%', minWidth: 0 } });
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const page = useSPFxPageContext();
 *   const { client } = useSPFxSPHttpClient();
 *   const instance = useSPFxInstanceInfo();
 *   const styles = useStyles();
 *   const mode = 'layered';
 *   const root = page.site.absoluteUrl.trim().replace(/\/+$/, '');
 *   const web = page.web.absoluteUrl.trim().replace(/\/+$/, '');
 *   // Required mode and both relevant URLs are supplied explicitly; construction sends no requests.
 *   const service = React.useMemo(() => client && root && web ? createSPFxConfigurationStoreService(client, { mode, siteCollectionUrl: root, discoveryWebUrl: web }) : undefined, [client, root, web]);
 *   const [key, setKey] = React.useState('spfx-toolkit-config-demo-service');
 *   const [consent, setConsent] = React.useState<object>();
 *   const approval = React.useRef<object>();
 *   const approve = (next: object | undefined): void => { approval.current = next; setConsent(next); };
 *   const boundary = React.useRef({ resetRevision, client, page, instance: instance.id, root, web, key, token: {} });
 *   const previous = boundary.current;
 *   if (previous.resetRevision !== resetRevision || previous.client !== client || previous.page !== page || previous.instance !== instance.id || previous.root !== root || previous.web !== web || previous.key !== key) boundary.current = { resetRevision, client, page, instance: instance.id, root, web, key, token: {} };
 *   const token = boundary.current.token;
 *   const mounted = React.useRef(true);
 *   const ledger = React.useRef(new Set<Ticket>());
 *   const sequence = React.useRef<object>();
 *   const latest = React.useRef(0);
 *   const [, refresh] = React.useState(0);
 *   const [outcome, setOutcome] = React.useState<Outcome>();
 *   const [confirmation, setConfirmation] = React.useState<{ confirm: () => void; cancel: () => void; key: string }>();
 *   const confirmationRef = React.useRef<{ token: object; owner: object; settle: (value: boolean) => void }>();
 *   const focus = React.useRef<{ trigger: HTMLButtonElement; owns: () => boolean }>();
 *   const current = (): boolean => mounted.current && boundary.current.token === token && page.site.absoluteUrl.trim().replace(/\/+$/, '') === root && page.web.absoluteUrl.trim().replace(/\/+$/, '') === web;
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; confirmationRef.current?.settle(false); confirmationRef.current = undefined; }; }, []);
 *   React.useLayoutEffect(() => { if (confirmationRef.current && confirmationRef.current.owner !== token) { confirmationRef.current.settle(false); confirmationRef.current = undefined; setConfirmation(undefined); } }, [token]);
 *   React.useLayoutEffect(() => { const request = focus.current; if (!confirmation && request && sequence.current === undefined && ledger.current.size === 0 && !request.trigger.disabled) { focus.current = undefined; if (request.owns() && request.trigger.isConnected) request.trigger.focus(); } });
 *   const resetSeen = React.useRef(resetRevision);
 *   React.useLayoutEffect(() => { if (resetSeen.current !== resetRevision) { resetSeen.current = resetRevision; approve(undefined); setOutcome(undefined); setKey('spfx-toolkit-config-demo-service'); } }, [resetRevision]);
 *   const busy = (): boolean => ledger.current.size > 0 || sequence.current !== undefined || confirmationRef.current !== undefined;
 *   const locked = busy();
 *   const disposable = key.startsWith('spfx-toolkit-config-demo-') && key.length > 'spfx-toolkit-config-demo-'.length;
 *   const visible = current() && outcome?.token === token ? outcome : undefined;
 *   const run = async (trigger: HTMLButtonElement, recipe: boolean): Promise<void> => {
 *     if (!current() || busy() || !service || !key || recipe && (!disposable || approval.current !== token)) return;
 *     const operation = ++latest.current, sequenceToken = {};
 *     const owns = (): boolean => current() && latest.current === operation;
 *     const steps: string[] = [];
 *     if (recipe) { sequence.current = sequenceToken; approve(undefined); refresh(count => count + 1); }
 *     setOutcome({ token, steps: [], message: recipe ? 'Running service recipe...' : 'Reading...' });
 *     const mark = (message: string, data: { item?: SPFxConfigurationStoreItem; items?: SPFxConfigurationStoreItem[] } = {}): void => {
 *       if (!owns()) return;
 *       steps.push(message); setOutcome(previousOutcome => ({ ...(previousOutcome?.token === token ? previousOutcome : { token }), ...data, token, steps: [...steps], message }));
 *     };
 *     const write = async (dispatch: () => Promise<void>): Promise<boolean> => {
 *       if (!owns() || ledger.current.size > 0) return false;
 *       const ticket: Ticket = {}; ledger.current.add(ticket); refresh(count => count + 1);
 *       const release = (): void => { ledger.current.delete(ticket); if (mounted.current) refresh(count => count + 1); };
 *       let promise: Promise<void>;
 *       try { promise = dispatch(); } catch (error) { release(); throw error; }
 *       // Keep the ORIGINAL whole public Promise, including discovery/probes/setup/recovery.
 *       await promise.finally(release); return owns();
 *     };
 *     try {
 *       if (!recipe) { const item = await service.get<unknown>(key); if (owns()) mark(item ? `Found ${key}.` : `${key} not found.`, { item }); return; }
 *       if (!await write(() => service.save(key, 20, { scope: 'tenant' }))) return;
 *       mark('Saved tenant default 20.');
 *       if (!owns() || !await write(() => service.save(key, 50, { scope: 'site', description: 'Local override' }))) return;
 *       mark('Saved site override 50.');
 *       if (!owns()) return; const item = await service.get<unknown>(key); if (!owns()) return;
 *       mark('Read the effective site override.', { item });
 *       if (!owns()) return; const items = await service.listScope('site'); if (!owns()) return;
 *       mark(`Loaded ${items.length} site physical item(s).`, { items });
 *       const removeSite = (): Promise<void> => service.remove(key, { scope: 'site' });
 *       const accepted = await new Promise<boolean>(resolve => {
 *         if (!owns()) { resolve(false); return; }
 *         const confirmationToken = {};
 *         confirmationRef.current = { token: confirmationToken, owner: token, settle: resolve };
 *         const close = (accept: boolean): void => {
 *           if (!owns() || confirmationRef.current?.token !== confirmationToken || ledger.current.size > 0) return;
 *           confirmationRef.current = undefined; focus.current = { trigger, owns }; setConfirmation(undefined); resolve(accept);
 *         };
 *         setConfirmation({ key, confirm: () => close(true), cancel: () => close(false) });
 *       });
 *       if (!owns()) return;
 *       if (!accepted) { mark('Site removal cancelled; completed writes remain.'); return; }
 *       if (!await write(removeSite)) return;
 *       mark('Removed only the site override.');
 *       if (!owns()) return; const inherited = await service.get<unknown>(key); if (!owns()) return;
 *       mark('Read the tenant default after removal.', { item: inherited });
 *     } catch (error) { if (owns()) setOutcome(previousOutcome => ({ ...(previousOutcome?.token === token ? previousOutcome : { token, steps: [] }), failure: error instanceof Error ? error.message : String(error) })); }
 *     finally { if (sequence.current === sequenceToken) sequence.current = undefined; if (mounted.current) refresh(count => count + 1); }
 *   };
 *   // Tie Fluent disabledFocusable to the exact disabled flag: return focus after consumed consent without enabling invocation.
 *   const shown = (value: unknown): string => (JSON.stringify(value) ?? String(value)) + (value === '' ? ' (empty string)' : '');
 *   const readDisabled = !service || locked || !key;
 *   const recipeDisabled = !service || locked || !disposable || consent !== token;
 *   const regions: DemoRegions = {
 *     controls: <>
 *       <Field label="Key"><Input aria-label="Key" disabled={locked} value={key} onChange={(_, data) => { if (!current() || busy() || data.value === key) return; boundary.current = { ...boundary.current, key: data.value, token: {} }; approve(undefined); setKey(data.value); }}/></Field>
 *       <Text as="p">Required mode layered, collection root {root}, captured discovery web {web}, provider instance {instance.id}. This public factory composes tenant defaults and site overrides. Consent: exact key {key}, tenant numeric 20 (preserve description), site numeric 50 with Local override description, remove only that site override. Tenant is the captured logical app-catalog session.</Text>
 *       <Checkbox label="I consent to these exact disposable mutations" disabled={locked || !disposable} checked={consent === token} onChange={(_, data) => { if (current() && !busy()) approve(data.checked === true ? token : undefined); }}/>
 *     </>,
 *     commands: <><Button disabled={readDisabled} disabledFocusable={readDisabled} onClick={event => run(event.currentTarget, false)}>Read effective value</Button><Button appearance="primary" disabled={recipeDisabled} disabledFocusable={recipeDisabled} onClick={event => run(event.currentTarget, true)}>Run service recipe</Button></>,
 *     status: <>{ledger.current.size > 0 && <Text role="status">Previous write still pending</Text>}{visible?.failure ? <Text role="alert">{visible.failure}</Text> : <Text role="status">{visible?.message ?? 'Not run'}</Text>}{visible && <ol aria-label="Completed configuration steps">{visible.steps.map((step, index) => <li key={index}>{step}</li>)}</ol>}</>,
 *     children: <><dl><dt>Mode (required)</dt><dd>{mode}</dd><dt>Collection root URL</dt><dd>{root}</dd><dt>Discovery web URL</dt><dd>{web}</dd><dt>Provider instance</dt><dd>{instance.id}</dd>{visible?.item && <><dt>Source</dt><dd>{visible.item.source}</dd><dt>Selected key</dt><dd>{visible.item.key}</dd><dt>Selected value</dt><dd>{shown(visible.item.value)}</dd><dt>Value type</dt><dd>{typeof visible.item.value}</dd><dt>Description</dt><dd>{visible.item.description ?? '(not set)'}</dd><dt>Physical item ID</dt><dd>{visible.item.id}</dd></>}</dl>
 *       {visible?.items && <><Text>Inventory count: {visible.items.length}</Text><dl aria-label="First ten configuration records">{visible.items.slice(0, 10).map((item, index) => <React.Fragment key={index}><dt>{item.key}</dt><dd>{item.source} / {item.id}: {shown(item.value)}</dd></React.Fragment>)}</dl></>}
 *       <Dialog open={confirmation !== undefined} onOpenChange={(_, data) => { if (!data.open) confirmation?.cancel(); }}><DialogSurface><DialogBody><DialogTitle>Confirm site override removal</DialogTitle><DialogContent>Remove exact key {confirmation?.key} only from site collection {root}. Tenant default remains in the captured logical app-catalog session.</DialogContent><DialogActions><Button onClick={confirmation?.cancel}>Cancel</Button><Button appearance="primary" onClick={confirmation?.confirm}>Confirm removal</Button></DialogActions></DialogBody></DialogSurface></Dialog>
 *     </>,
 *     diagnostics: <><Text as="p">Physical catalog URL is unavailable through this facade. Construction is request-free and tenant discovery is lazy. Physical inventory never merges scopes or deduplicates IDs. Whole-record overrides and destination-local metadata apply. Only Save can initialize its chosen store. Completed steps remain facts when a later read fails. No transaction, rollback or cancellation is provided. Reset retains the service and the original pending-write tickets; the lock is local to this mounted Demo. Tenant cleanup is separate, explicitly scoped and consented; this recipe never removes tenant data.</Text><div role="region" aria-label="Configuration service payload" tabIndex={0} className={styles.payload}><pre>{JSON.stringify(visible, undefined, 2)}</pre></div></>,
 *   };
 *   return render ? render(regions) : <section aria-label="Configuration service">{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</section>;
 * }
 * export const demo = { host: 'webPart', kind: 'write', contentPlacement: 'configuration', observe: 'Write a tenant default and site override, inspect their provenance and inventory, then remove only the site override.' };
 * ```
 */
export function createSPFxConfigurationStoreService(
  client: SPHttpClient,
  options: SPFxConfigurationStoreServiceOptions
): SPFxConfigurationStoreService {
  const mode = options?.mode;
  if (mode !== 'tenant' && mode !== 'site' && mode !== 'layered') throw new Error('A valid configuration store mode is required.');
  const normalize = (url: string | undefined): string => {
    if (typeof url !== 'string') throw new Error('A nonempty configuration store URL is required.');
    const normalized = url.trim().replace(/\/+$/, '');
    if (!normalized) throw new Error('A nonempty configuration store URL is required.');
    return normalized;
  };
  const siteUrl = mode !== 'tenant' ? normalize(options.siteCollectionUrl) : undefined;
  const discoveryUrl = mode !== 'site' ? normalize(options.discoveryWebUrl) : undefined;
  const site = siteUrl ? createSPFxSiteKeyValueStoreService(client) : undefined;
  const tenant = discoveryUrl ? createSPFxTenantKeyValueStoreService(client) : undefined;
  const catalog = discoveryUrl ? createSPFxAppCatalogService(client, { web: { absoluteUrl: discoveryUrl } }) : undefined;
  let pendingCatalog: Promise<string> | undefined;
  const catalogUrl = (): Promise<string> => {
    if (!catalog) return Promise.reject(new Error('Tenant scope is not enabled.'));
    if (!pendingCatalog) {
      pendingCatalog = catalog.discoverUrl().finally(() => { pendingCatalog = undefined; });
    }
    return pendingCatalog;
  };
  const scopeOf = (scope: SPFxConfigurationStoreScope): SPFxConfigurationStoreScope => {
    if ((scope !== 'tenant' && scope !== 'site') || (mode !== 'layered' && scope !== mode)) throw new Error('An enabled tenant or site scope is required.');
    return scope;
  };
  const keyOf = (key: string): void => { if (typeof key !== 'string') throw new Error('A string configuration key is required.'); };
  const targetOf = (target: SPFxConfigurationStoreTarget): SPFxConfigurationStoreScope => scopeOf(target?.scope);
  const physicalGet = async <T,>(key: string, scope: SPFxConfigurationStoreScope): Promise<SPFxConfigurationStoreItem<T> | undefined> => {
    const item = scope === 'site' && site && siteUrl ? await site.get<T>(key, siteUrl)
      : tenant ? await tenant.get<T>(key, await catalogUrl()) : undefined;
    return item === undefined ? undefined : { ...item, source: scope };
  };
  return {
    get: async <T,>(key: string): Promise<SPFxConfigurationStoreItem<T> | undefined> => {
      keyOf(key);
      if (mode !== 'layered') return physicalGet<T>(key, mode);
      const local = await physicalGet<T>(key, 'site');
      return local !== undefined ? local : physicalGet<T>(key, 'tenant');
    },
    listScope: async (requestedScope): Promise<SPFxConfigurationStoreItem<unknown>[]> => {
      const scope = scopeOf(requestedScope);
      const items = scope === 'site' && site && siteUrl ? await site.list(siteUrl)
        : tenant ? await tenant.list(await catalogUrl()) : [];
      return items.map(item => ({ ...item, source: scope }));
    },
    save: async <T,>(key: string, value: T, saveOptions: SPFxConfigurationStoreSaveOptions): Promise<void> => {
      keyOf(key);
      const scope = targetOf(saveOptions);
      const description = saveOptions.description;
      if (description !== undefined && typeof description !== 'string') throw new Error('Configuration description must be a string.');
      if (scope === 'site' && site && siteUrl) await site.save(key, value, siteUrl, description);
      else if (tenant) await tenant.save(key, value, await catalogUrl(), description);
    },
    remove: async (key, target): Promise<void> => {
      keyOf(key);
      const scope = targetOf(target);
      if (scope === 'site' && site && siteUrl) await site.remove(key, siteUrl);
      else if (tenant) await tenant.remove(key, await catalogUrl());
    }
  };
}

/**
 * Share collection settings in the root-web list with read-only probes and conditional mutations.
 *
 * Keys are forwarded unchanged; SharePoint field/unique-column comparison
 * determines equality, while returned Title casing is preserved. URLs are
 * trimmed and stripped of trailing slashes, not validated as collection/catalog
 * addresses. Empty URLs reject before requests. Reads validate schema without
 * enforcing or repairing uniqueness: get rejects duplicates, list preserves
 * historical duplicate rows. Only the list existence probe's 404 is absence;
 * field/item/page 404 and access failures reject. Services have no value cache,
 * React provider dependency or disposal method.
 *
 * @packageDocumentation
 * @docPage services/site-key-value-store
 */
import { SPHttpClient } from '@microsoft/sp-http';

import { createSPFxListKeyValueStoreService } from './spfx-list-key-value-store.internal';
import { canWriteKeyValueStore } from './spfx-site-key-value-store.internal';

/**
 * Stored key and deserialized value. T is an expected shape, not runtime validation; historical data keeps its JSON-first interpretation. New text preserves its value; bigint reads as exact decimal text and Date as ISO text. Legacy empty REST Note fields still read as empty strings.
 *
 * @example
 * ```ts
 * import { SPFxSiteKeyValueStoreService, SPFxSiteKeyValueStoreServiceItem } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(siteCollectionUrl: string, store: SPFxSiteKeyValueStoreService) {
 *   const item: SPFxSiteKeyValueStoreServiceItem<boolean> | undefined = await store.get('featureEnabled', siteCollectionUrl);
 *   console.log(item?.value);
 * }
 * ```
 * @template T - Expected application data shape; defaults to unknown and performs no runtime validation.
 */
export interface SPFxSiteKeyValueStoreServiceItem<T = unknown> {
  /**
   * Stored Title/key, preserving its stored casing.
   */
  readonly key: string;
  /**
   * Deserialized value; caller must validate untrusted content.
   */
  readonly value: T;
  /**
   * Optional metadata; empty descriptions are returned as undefined.
   */
  readonly description: string | undefined;
  /**
   * SharePoint list item ID.
   */
  readonly id: number;
}

/**
 * Root-web SiteKeyValueStore operations shared by all subsites. Reads verify schema without provisioning, and only confirmed list absence is treated as empty. Requests reject on failed schema/item/page operations.
 *
 * @example
 * ```ts
 * import type { SPHttpClient } from '@microsoft/sp-http';
 * import { SPFxSiteKeyValueStoreService, createSPFxSiteKeyValueStoreService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(client: SPHttpClient, siteCollectionUrl: string) {
 *   const store: SPFxSiteKeyValueStoreService = createSPFxSiteKeyValueStoreService(client);
 *   await store.save('featureEnabled', true, siteCollectionUrl);
 * }
 * ```
 */
export interface SPFxSiteKeyValueStoreService {
  /**
   * Provisions a hidden generic SiteKeyValueStore (BaseTemplate 100, Hidden/NoCrawl true) or repairs its schema: Text Title plus Note Value/Description and indexed unique Title. Concurrent setup/readiness is shared per normalized URL within this instance; failed setup can retry. Incompatible types/duplicates reject without deleting data. Separate instances recover list/field creation races only through bounded validated rereads; 401/403 never recover.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @returns Resolves when setup succeeds.
   * @throws Provisioning or incompatible-schema failures.
   */
  ensureListReady: (siteCollectionUrl: string) => Promise<void>;
  /**
   * Reads a key without provisioning.
   * @typeParam T - Expected deserialized value; defaults to unknown and performs no runtime validation.
   * @param key - Stored key to match.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @returns Deserialized item, or undefined for an absent key or confirmed absent list.
   * @throws Failed read requests or incompatible schema.
   */
  get: <T = unknown>(key: string, siteCollectionUrl: string) => Promise<SPFxSiteKeyValueStoreServiceItem<T> | undefined>;
  /**
   * Reads every page in Title order, including historical duplicate rows. Both OData continuation properties are accepted; links must remain on the same origin/list items path and must not repeat. Page failures reject without partial data.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @returns Items, or an empty array for a confirmed absent list.
   * @throws Failed reads, schema or continuation validation failures.
   */
  list: (siteCollectionUrl: string) => Promise<SPFxSiteKeyValueStoreServiceItem<unknown>[]>;
  /**
   * Creates or updates a key after ensuring setup. Uses the observed item ETag; conflicts (HTTP 412) and missing ETags reject without a wildcard overwrite or blind retry. Creation collisions allow one validated conditional update and expose its actual error. This protects the service snapshot, not a consumer's earlier get().
   * @typeParam T - Input value shape; defaults to unknown and does not guarantee serializability.
   * @param key - Stored key.
   * @param value - Value to encode: empty/JSON-like text is JSON-quoted, bigint becomes quoted decimal text and Date remains ISO text. Existing data is not rewritten. Circular values, nested bigint and invalid Date throw before provisioning; T does not validate runtime content.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @param description - Omitted preserves an existing description; empty string clears it.
   * @returns Resolves after the write succeeds.
   * @throws Setup, serialization, missing ETag, request or conflict failures.
   */
  save: <T = unknown>(key: string, value: T, siteCollectionUrl: string, description?: string) => Promise<void>;
  /**
   * Deletes using an observed ETag without provisioning; confirmed absent list/key is a no-op. Snapshot/deletion 404 is accepted only after rechecking key absence; concurrent changes reject.
   * @param key - Key to remove.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @returns Resolves after deletion or confirmed absence.
   * @throws Request/schema failures, missing ETag or concurrent changes.
   */
  remove: (key: string, siteCollectionUrl: string) => Promise<void>;
  /**
   * Checks effective Add/Edit/Delete Items on the list, or on the root web plus Manage Lists when the list is absent. This indicator does not authorize writes or schema changes.
   * @param siteCollectionUrl - Absolute site collection root URL (pageContext.site.absoluteUrl), never the current subsite.
   * @returns False for empty URLs or failed lookups; otherwise the effective-permission indicator.
   */
  canCurrentUserWrite: (siteCollectionUrl: string) => Promise<boolean>;
}

/**
 * Creates an isolated, non-React site store. Construction performs no request; SharePoint enforces actual grants.
 * @param spHttpClient - Caller-owned authenticated SharePoint client.
 * @returns Explicit store operations.
 *
 * @example Basic usage
 * ```ts
 * import type { WebPartContext } from '@microsoft/sp-webpart-base';
 * import { createSPFxSiteKeyValueStoreService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(context: WebPartContext) {
 *   const store = createSPFxSiteKeyValueStoreService(context.spHttpClient);
 *   await store.save('featureEnabled', true, context.pageContext.site.absoluteUrl);
 * }
 * ```
 * @useWhen You share settings across a site collection and its subsites through the root-web store.
 *
 *
 * @example Read and conditionally save at collection root
 * ```ts
 * import type { SPHttpClient } from '@microsoft/sp-http';
 * import type { PageContext } from '@microsoft/sp-page-context';
 * import { createSPFxSiteKeyValueStoreService } from '@apvee/spfx-react-toolkit';
 * async function demonstrate(context: { spHttpClient: SPHttpClient; pageContext: PageContext }) {
 *
 *   const store = createSPFxSiteKeyValueStoreService(context.spHttpClient);
 *   const rootUrl = context.pageContext.site.absoluteUrl;
 *   const setting = await store.get<boolean>('featureEnabled', rootUrl);
 *   if (await store.canCurrentUserWrite(rootUrl)) {
 *     await store.save('featureEnabled', true, rootUrl, 'Shared across subsites');
 *   }
 * }
 * ```
 *
 * @example Read a collection root record before a consented save
 * ```tsx
 * import * as React from 'react';
 * import { makeStyles, Button, Checkbox, Field, Input, Text } from '@fluentui/react-components';
 * import { createSPFxSiteKeyValueStoreService, useSPFxInstanceInfo, useSPFxPageContext, useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
 * import type { SPFxSiteKeyValueStoreServiceItem } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * const useStyles = makeStyles({ payload: { overflowX: 'auto', maxWidth: '100%', minWidth: 0 } });
 * type Ticket = { promise?: Promise<void> };
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const page = useSPFxPageContext();
 *   const { client } = useSPFxSPHttpClient();
 *   const info = useSPFxInstanceInfo();
 *   const styles = useStyles();
 *   const service = React.useMemo(() => client ? createSPFxSiteKeyValueStoreService(client) : undefined, [client]);
 *   const root = page.site.absoluteUrl.trim().replace(/\/+$/, '');
 *   const web = page.web.absoluteUrl.trim().replace(/\/+$/, '');
 *   const [key, setKey] = React.useState('spfx-toolkit-site-demo-service');
 *   const [consent, setConsent] = React.useState<object>();
 *   const approval = React.useRef<object>();
 *   approval.current = consent;
 *   const approve = (next: object | undefined): void => { approval.current = next; setConsent(next); };
 *   const [outcome, setOutcome] = React.useState<{ token: object; item?: SPFxSiteKeyValueStoreServiceItem; canWrite?: boolean; message?: string; failure?: string }>();
 *   const boundary = React.useRef({ revision: resetRevision, root, web, client, page, instance: info.id, key, token: {} });
 *   const previous = boundary.current;
 *   if (previous.revision !== resetRevision || previous.root !== root || previous.web !== web || previous.client !== client || previous.page !== page || previous.instance !== info.id || previous.key !== key) boundary.current = { revision: resetRevision, root, web, client, page, instance: info.id, key, token: {} };
 *   const token = boundary.current.token;
 *   const mounted = React.useRef(true);
 *   const ledger = React.useRef(new Set<Ticket>());
 *   const latest = React.useRef(0);
 *   const reading = React.useRef<object>();
 *   const [, refresh] = React.useState(0);
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
 *   const resetSeen = React.useRef(resetRevision);
 *   React.useLayoutEffect(() => {
 *     if (resetSeen.current !== resetRevision) { resetSeen.current = resetRevision; setKey('spfx-toolkit-site-demo-service'); approve(undefined); setOutcome(undefined); }
 *   }, [resetRevision]);
 *   const current = (): boolean => mounted.current && boundary.current.token === token;
 *   const visible = current() && outcome?.token === token ? outcome : undefined;
 *   const pending = ledger.current.size > 0;
 *   const isReading = reading.current === token;
 *   const disposable = key.startsWith('spfx-toolkit-site-demo-') && key.length > 'spfx-toolkit-site-demo-'.length;
 *   const read = async (): Promise<void> => {
 *     if (!current() || !service || !root || ledger.current.size > 0 || reading.current === token) return;
 *     const operation = ++latest.current; reading.current = token; approve(undefined); setOutcome(undefined); refresh(count => count + 1);
 *     const owns = (): boolean => current() && latest.current === operation;
 *     try {
 *       const item = await service.get(key, root);
 *       if (!owns()) return;
 *       const canWrite = await service.canCurrentUserWrite(root);
 *       if (owns()) setOutcome({ token, item, canWrite, message: item ? `Found ${key}.` : `${key} not found.` });
 *     } catch (error) { if (owns()) setOutcome({ token, failure: error instanceof Error ? error.message : String(error) }); }
 *     finally { if (reading.current === token) reading.current = undefined; if (mounted.current) refresh(count => count + 1); }
 *   };
 *   const save = async (): Promise<void> => {
 *     if (!current() || !service || !root || ledger.current.size > 0 || reading.current === token || approval.current !== token || !disposable || visible?.canWrite !== true) return;
 *     const operation = ++latest.current;
 *     const ticket: Ticket = {}; ledger.current.add(ticket); refresh(count => count + 1);
 *     const owns = (): boolean => current() && latest.current === operation;
 *     const release = (): void => { ledger.current.delete(ticket); if (mounted.current) refresh(count => count + 1); };
 *     try {
 *       const promise = service.save(key, true, root, 'Shared across subsites');
 *       ticket.promise = promise;
 *       await promise.finally(release);
 *       if (owns()) { approve(undefined); setOutcome({ token, message: `Saved ${key}.` }); }
 *     } catch (error) { if (owns()) setOutcome({ token, failure: error instanceof Error ? error.message : String(error) }); }
 *     // No read-after-write: the original server operation can finish after Reset.
 *   };
 *   const regions: DemoRegions = {
 *     controls: <>
 *       <Field label="Key"><Input aria-label="Key" disabled={pending} value={key} onChange={(_, data) => { if (ledger.current.size === 0) { approve(undefined); setKey(data.value); } }}/></Field>
 *       <Text as="p">Optional Save: {root}, key {key}, boolean true, description "Shared across subsites". Only disposable keys are accepted; Save may provision/repair the root store.</Text>
 *       <Checkbox label="I consent to this exact disposable mutation" checked={consent === token} disabled={pending || !disposable || visible?.canWrite !== true} onChange={(_, data) => { if (current() && ledger.current.size === 0 && visible?.canWrite === true) approve(data.checked === true ? token : undefined); }}/>
 *     </>,
 *     commands: <><Button disabled={!service || !root || pending || isReading} onClick={read}>Read</Button><Button disabled={!service || pending || isReading || !disposable || approval.current !== token || visible?.canWrite !== true} onClick={save}>Save</Button></>,
 *     status: <>{pending && <Text role="status">Previous write still pending</Text>}{visible?.failure ? <Text role="alert">{visible.failure}</Text> : visible?.message ? <Text role="status">{visible.message}</Text> : <Text>{isReading ? 'Reading...' : 'Read the record and permission indicator before consenting to Save.'}</Text>}</>,
 *     children: <>
 *       <dl aria-label="Site service result"><dt>Collection root URL</dt><dd>{root}</dd><dt>Current web URL</dt><dd>{web}</dd><dt>Can Write</dt><dd>{visible?.canWrite === undefined ? '(not observed)' : String(visible.canWrite)}</dd>
 *       {visible?.item && <><dt>Selected key</dt><dd>{visible.item.key}</dd><dt>Selected value</dt><dd>{JSON.stringify(visible.item.value)}{visible.item.value === '' ? ' (empty string)' : ''}</dd><dt>Value type</dt><dd>{typeof visible.item.value}</dd><dt>Description</dt><dd>{visible.item.description ?? '(not set)'}</dd></>}</dl>
 *     </>,
 *     diagnostics: <><Text as="p">The caller-owned native client constructs a request-free service. Read does not provision; service failures reject. This recipe deliberately requires a current observed true indicator before optional consented Save. The indicator, schema and grants can change; SharePoint authorizes the request. ETags protect the service snapshot, not this earlier Read. Reset restores local UI, not server state; an already dispatched operation and its internal RPCs can finish. The pending lock is local to this mounted Demo, not other scenarios.</Text><div role="region" aria-label="Site service payload" tabIndex={0} className={styles.payload}><pre>{JSON.stringify(visible?.item, undefined, 2)}</pre></div></>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'write', contentPlacement: 'configuration', observe: 'Read the collection-root record and permission indicator before explicitly consenting to the service write.' };
 * ```
 */
export function createSPFxSiteKeyValueStoreService(spHttpClient: SPHttpClient): SPFxSiteKeyValueStoreService {
  const store = createSPFxListKeyValueStoreService(spHttpClient, 'site');
  const canCurrentUserWrite = (siteCollectionUrl: string): Promise<boolean> =>
    canWriteKeyValueStore(spHttpClient, siteCollectionUrl, 'site');
  return { ...store, canCurrentUserWrite };
}

/**
 * Share tenant settings through a caller-selected app-catalog list with observed-ETag mutations.
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
 * @docPage services/tenant-key-value-store
 */
import { SPHttpClient } from '@microsoft/sp-http';

import { createSPFxListKeyValueStoreService } from './spfx-list-key-value-store.internal';

/**
 * Stored key and deserialized value. T is an expected shape, not runtime validation; existing raw values keep legacy JSON-first interpretation. New strings preserve their text; Date reads as ISO text and bigint as exact decimal text. A legacy REST null Note reads as null here, whereas the site store reads it as empty text.
 *
 * @example
 * ```ts
 * import { SPFxTenantKeyValueStoreService, SPFxTenantKeyValueStoreServiceItem } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(catalogUrl: string, store: SPFxTenantKeyValueStoreService) {
 *   const item: SPFxTenantKeyValueStoreServiceItem<boolean> | undefined = await store.get('featureEnabled', catalogUrl);
 *   console.log(item?.value);
 * }
 * ```
 * @template T - Expected application data shape; defaults to unknown and performs no runtime validation.
 */
export interface SPFxTenantKeyValueStoreServiceItem<T = unknown> {
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
 * TenantKeyValueStore operations in an explicitly supplied app catalog. Reads and remove never provision or repair schema; only save/ensureListReady can do so. Requests reject on failed schema/read/write operations. A confirmed missing list is empty.
 *
 * @example
 * ```ts
 * import type { SPHttpClient } from '@microsoft/sp-http';
 * import { SPFxTenantKeyValueStoreService, createSPFxTenantKeyValueStoreService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(client: SPHttpClient, catalogUrl: string) {
 *   const store: SPFxTenantKeyValueStoreService = createSPFxTenantKeyValueStoreService(client);
 *   await store.save('featureEnabled', true, catalogUrl);
 * }
 * ```
 */
export interface SPFxTenantKeyValueStoreService {
  /**
   * Provisions the hidden generic TenantKeyValueStore (BaseTemplate 100, Hidden/NoCrawl true), or repairs Title (Text), Value and Description (Note), requiring indexed unique Title. Normalized catalog URLs share setup within this instance. Failed setup can retry; existing duplicates or incompatible fields reject without deleting data. Separate instances recover list/field creation races only through bounded validated rereads; 401/403 never recover.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Resolves when setup succeeds.
   * @throws Provisioning or incompatible-schema failures.
   */
  ensureListReady: (catalogUrl: string) => Promise<void>;
  /**
   * Reads a unique key after read-only schema validation; duplicate matches reject.
   * @typeParam T - Expected deserialized value; defaults to unknown and performs no runtime validation.
   * @param key - Stored key to match.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Deserialized item, or undefined for a confirmed absent list/key.
   * @throws Failed probes/reads, malformed data or incompatible schema; 401/403/5xx are never absence.
   */
  get: <T = unknown>(key: string, catalogUrl: string) => Promise<SPFxTenantKeyValueStoreServiceItem<T> | undefined>;
  /**
   * Reads every server page in Title order. Continuation links must stay on the same origin/list items path and cannot repeat; partial results are not returned after a failure.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Complete items, including historical duplicates, or an empty array for a confirmed absent list.
   * @throws Failed reads, malformed rows or unsafe/repeated continuation links, without partial results.
   */
  list: (catalogUrl: string) => Promise<SPFxTenantKeyValueStoreServiceItem<unknown>[]>;
  /**
   * Creates or updates a key after ensuring setup. Updates require the ETag of the service's item snapshot; concurrent changes reject (HTTP 412), without a blind retry. Creation collisions allow one validated conditional update. This is not compare-and-swap against a previous get() call.
   * @typeParam T - Input value shape; defaults to unknown and does not guarantee serializability.
   * @param key - Stored key.
   * @param value - Value to serialize. New JSON-like strings are JSON-quoted, bigint becomes quoted decimal text, Date stays ISO text; other inputs retain the legacy format. Circular objects, nested bigint and invalid Date reject before provisioning. T does not validate stored data.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @param description - Omitted preserves an existing description; empty string clears it.
   * @returns Resolves after the write succeeds.
   * @throws Setup, serialization or request failures.
   */
  save: <T = unknown>(key: string, value: T, catalogUrl: string, description?: string) => Promise<void>;
  /**
   * Deletes a key without provisioning, using an observed item ETag. A confirmed absent list/key is a no-op; concurrent changes reject. A vanished snapshot/deletion is accepted only after rechecking key absence, never by deleting a recreated key.
   * @param key - Key to remove.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Resolves after deletion or confirmed absence.
   * @throws Request/schema failures, missing ETag or concurrent changes.
   */
  remove: (key: string, catalogUrl: string) => Promise<void>;
}

/**
 * Creates an isolated, non-React tenant store. Construction performs no request; SharePoint enforces actual grants.
 * @param spHttpClient - Caller-owned authenticated SharePoint client.
 * @returns Explicit store operations.
 *
 * @example Basic usage
 * ```ts
 * import type { WebPartContext } from '@microsoft/sp-webpart-base';
 * import { createSPFxTenantKeyValueStoreService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(context: WebPartContext, catalogUrl: string) {
 *   const store = createSPFxTenantKeyValueStoreService(context.spHttpClient);
 *   await store.save('featureEnabled', true, catalogUrl);
 * }
 * ```
 * @useWhen You share application settings through a TenantKeyValueStore list at an explicit app-catalog URL.
 *
 *
 * @example Preserve numeric-looking text in a tenant list
 * ```ts
 * import type { SPHttpClient } from '@microsoft/sp-http';
 *
 * import { createSPFxTenantKeyValueStoreService } from '@apvee/spfx-react-toolkit';
 * async function demonstrate(context: { spHttpClient: SPHttpClient }, catalogUrl: string) {
 *   const store = createSPFxTenantKeyValueStoreService(context.spHttpClient);
 *   await store.save('my-app:version', '123', catalogUrl);
 *   const version = await store.get<unknown>('my-app:version', catalogUrl);
 *   if (typeof version?.value === 'string') console.log(version.value);
 * }
 * ```
 *
 * @scenario Round trip numeric looking text through a tenant service
 * ```tsx
 * import * as React from 'react';
 * import { makeStyles, Button, Checkbox, Field, Input, Text } from '@fluentui/react-components';
 * import { createSPFxTenantKeyValueStoreService, useSPFxInstanceInfo, useSPFxPageContext, useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
 * import type { SPFxTenantKeyValueStoreServiceItem } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Ticket = object;
 * const useStyles = makeStyles({ payload: { overflowX: 'auto', maxWidth: '100%', minWidth: 0 } });
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const page = useSPFxPageContext();
 *   const { client } = useSPFxSPHttpClient();
 *   const instance = useSPFxInstanceInfo();
 *   const styles = useStyles();
 *   const service = React.useMemo(() => client ? createSPFxTenantKeyValueStoreService(client) : undefined, [client]);
 *   const web = page.web.absoluteUrl.trim().replace(/\/+$/, '');
 *   const root = page.site.absoluteUrl.trim().replace(/\/+$/, '');
 *   const [catalog, setCatalog] = React.useState('');
 *   const [key, setKey] = React.useState('spfx-toolkit-tenant-demo-service');
 *   const [consent, setConsent] = React.useState<object>();
 *   const approval = React.useRef<object>();
 *   const approve = (next: object | undefined): void => { approval.current = next; setConsent(next); };
 *   const boundary = React.useRef({ revision: resetRevision, page, client, web, root, instance: instance.id, catalog, key, token: {} });
 *   const previous = boundary.current;
 *   if (previous.revision !== resetRevision || previous.page !== page || previous.client !== client || previous.web !== web || previous.root !== root || previous.instance !== instance.id || previous.catalog !== catalog || previous.key !== key) boundary.current = { revision: resetRevision, page, client, web, root, instance: instance.id, catalog, key, token: {} };
 *   const token = boundary.current.token;
 *   const mounted = React.useRef(true);
 *   const ledger = React.useRef(new Set<Ticket>());
 *   const sequence = React.useRef<object>();
 *   const [, refresh] = React.useState(0);
 *   const [outcome, setOutcome] = React.useState<{ token: object; persisted?: boolean; item?: SPFxTenantKeyValueStoreServiceItem; message?: string; failure?: string }>();
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
 *   const resetSeen = React.useRef(resetRevision);
 *   React.useLayoutEffect(() => {
 *     if (resetSeen.current !== resetRevision) { resetSeen.current = resetRevision; setCatalog(''); setKey('spfx-toolkit-tenant-demo-service'); approve(undefined); setOutcome(undefined); }
 *   }, [resetRevision]);
 *   const current = (): boolean => mounted.current && boundary.current.token === token && page.web.absoluteUrl.trim().replace(/\/+$/, '') === web && page.site.absoluteUrl.trim().replace(/\/+$/, '') === root;
 *   const visible = current() && outcome?.token === token ? outcome : undefined;
 *   const pending = ledger.current.size > 0;
 *   const locked = pending || sequence.current === token;
 *   const disposable = key.startsWith('spfx-toolkit-tenant-demo-') && key.length > 'spfx-toolkit-tenant-demo-'.length;
 *   const run = async (): Promise<void> => {
 *     if (!current() || !service || !catalog.trim() || !disposable || approval.current !== token || ledger.current.size > 0 || sequence.current === token) return;
 *     let url: string;
 *     try {
 *       const parsed = new URL(catalog.trim());
 *       if (!['http:', 'https:'].includes(parsed.protocol) || parsed.username || parsed.password || parsed.search || parsed.hash) throw new Error('Enter an absolute catalog URL without credentials, query or fragment.');
 *       url = catalog.trim().replace(/\/+$/, '');
 *     } catch (error) { setOutcome({ token, failure: error instanceof Error ? error.message : String(error) }); return; }
 *     sequence.current = token;
 *     const ticket: Ticket = {}; ledger.current.add(ticket); approve(undefined); setOutcome(undefined); refresh(count => count + 1);
 *     const release = (): void => { ledger.current.delete(ticket); if (mounted.current) refresh(count => count + 1); };
 *     let persisted = false;
 *     try {
 *       // Keep the original WHOLE public Save promise, including setup/snapshot/internal RPCs.
 *       let promise: Promise<void>;
 *       try { promise = service.save(key, '123', url); }
 *       catch (error) { release(); throw error; }
 *       await promise.finally(release);
 *       if (!current()) return;
 *       persisted = true; setOutcome({ token, persisted, message: 'Reading back the saved text...' });
 *       const item = await service.get<unknown>(key, url);
 *       if (current()) setOutcome({ token, persisted, item, message: item?.value === '123' && typeof item.value === 'string' ? 'Read back exact text 123 (string).' : 'Readback did not contain exact text 123.' });
 *     } catch (error) { if (current()) setOutcome({ token, persisted, failure: error instanceof Error ? error.message : String(error) }); }
 *     finally { if (sequence.current === token) sequence.current = undefined; if (mounted.current) refresh(count => count + 1); }
 *   };
 *   const regions: DemoRegions = {
 *     controls: <>
 *       <Field label="Catalog URL"><Input aria-label="Catalog URL" disabled={locked} value={catalog} onChange={(_, data) => { if (current() && ledger.current.size === 0 && sequence.current !== token) { boundary.current = { ...boundary.current, catalog: data.value, token: {} }; approve(undefined); setCatalog(data.value); } }}/></Field>
 *       <Field label="Key"><Input aria-label="Key" disabled={locked} value={key} onChange={(_, data) => { if (current() && ledger.current.size === 0 && sequence.current !== token) { boundary.current = { ...boundary.current, key: data.value, token: {} }; approve(undefined); setKey(data.value); } }}/></Field>
 *       <Text as="p">Save exact text "123" at supplied catalog {catalog || '(required)'}, key {key}; preserve destination description by omitting it. Save may provision or repair TenantKeyValueStore.</Text>
 *       <Checkbox label="I consent to this exact disposable mutation" checked={consent === token} disabled={locked || !disposable || !catalog.trim()} onChange={(_, data) => { if (current() && ledger.current.size === 0 && sequence.current !== token) approve(data.checked === true ? token : undefined); }}/>
 *     </>,
 *     commands: <Button appearance="primary" disabled={!service || !catalog.trim() || !disposable || consent !== token || locked} onClick={run}>Save and read numeric-looking text</Button>,
 *     status: <>{pending && <Text role="status">Previous write still pending</Text>}{visible?.persisted && <Text role="status">Saved {key}. The subsequent read is a separate step.</Text>}{visible?.failure ? <Text role="alert">{visible.failure}</Text> : <Text role="status">{visible?.message ?? 'Supply a known catalog URL, then explicitly consent to Save and read.'}</Text>}</>,
 *     children: <dl aria-label="Tenant service result"><dt>Catalog URL</dt><dd>{catalog || '(required)'}</dd><dt>Key</dt><dd>{key}</dd>{visible?.item && <><dt>Selected value</dt><dd>{JSON.stringify(visible.item.value)}</dd><dt>Value type</dt><dd>{typeof visible.item.value}</dd><dt>Description</dt><dd>{visible.item.description ?? '(not set)'}</dd></>}</dl>,
 *     diagnostics: <><Text as="p">This request-free factory uses only the explicitly supplied catalog URL. It does not discover a catalog or expose a permission API. SharePoint authorizes writes. ETags protect the service's snapshot, not an earlier application read. A completed Save followed by failed Read is partial success. Reset clears local inputs/consent/results; it never removes the key, cancels a dispatched write or rolls back server data. Pending writes retain the local lock until their original public promise settles. Other scenarios have independent locks.</Text><div role="region" aria-label="Tenant service payload" tabIndex={0} className={styles.payload}><pre>{JSON.stringify(visible?.item, undefined, 2)}</pre></div></>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'write', contentPlacement: 'configuration', observe: 'Save numeric-looking text through the service at the supplied catalog and read it back as a string.' };
 * ```
 */
export function createSPFxTenantKeyValueStoreService(
  spHttpClient: SPHttpClient
): SPFxTenantKeyValueStoreService {
  return createSPFxListKeyValueStoreService(spHttpClient, 'tenant');
}

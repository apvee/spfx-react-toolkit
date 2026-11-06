/**
 * Read legacy StorageEntity values from a caller-selected tenant app catalog.
 *
 * @packageDocumentation
 * @docPage services/tenant-property
 */
import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';

import {
  deserializeTenantValue,
  escapeODataValue
} from '../helpers/spfx-tenant-value.helpers';

/**
 * Deserialized StorageEntity value and description. T is a compile-time expectation, not runtime validation; legacy serialization may coerce JSON-like strings.
 *
 * @example
 * ```ts
 * import { SPFxTenantProperty, SPFxTenantPropertyService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(catalogUrl: string, service: SPFxTenantPropertyService) {
 *   const property: SPFxTenantProperty<boolean> = await service.get('featureEnabled', catalogUrl);
 *   console.log(property.value);
 * }
 * ```
 * @template T - Expected application data shape; defaults to unknown and performs no runtime validation.
 */
export interface SPFxTenantProperty<T = unknown> {
  /**
   * Deserialized value, or undefined when SharePoint returns no Value or a legacy REST-null Value. A stored string is parsed JSON-first without schema validation or Date reconstruction.
   */
  readonly value: T | undefined;
  /**
   * Stored description passed through unchanged when Value exists; undefined when Value is missing/legacy REST null, even if a description was supplied.
   */
  readonly description: string | undefined;
}

/**
 * Reads StorageEntity properties from an explicitly supplied tenant app catalog. No discovery, cache, provisioning or write method is added.
 *
 * @example
 * ```ts
 * import type { SPHttpClient } from '@microsoft/sp-http';
 * import { SPFxTenantPropertyService, createSPFxTenantPropertyService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(client: SPHttpClient, catalogUrl: string) {
 *   const service: SPFxTenantPropertyService = createSPFxTenantPropertyService(client);
 *   const setting = await service.get<boolean>('featureEnabled', catalogUrl);
 * }
 * ```
 */
export interface SPFxTenantPropertyService {
  /**
   * Reads and deserializes one StorageEntity.
   * @typeParam T - Expected deserialized value; defaults to unknown and performs no runtime validation.
   * @param key - Required property key; OData-escaped and URL-encoded for the path.
   * @param catalogUrl - Absolute tenant app catalog URL.
   * @returns Value and description, possibly undefined.
   * @throws Every unsuccessful HTTP response, including 404/403, and response parsing failures; HTTP failure is never an absent-property success.
   */
  get: <T = unknown>(key: string, catalogUrl: string) => Promise<SPFxTenantProperty<T>>;
}

interface StorageEntityResponse {
  Value?: string;
  Description?: string;
}

function encodeODataPathValue(value: string): string {
  return encodeURIComponent(escapeODataValue(value)).replace(/'/g, '%27');
}

/**
 * Creates explicit StorageEntity reads without discovering the catalog.
 * @param spHttpClient - Caller-owned SharePoint client.
 * @returns A property-reading service; construction performs no request.
 *
 * @example Basic usage
 * ```ts
 * import type { WebPartContext } from '@microsoft/sp-webpart-base';
 * import { createSPFxTenantPropertyService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(context: WebPartContext, catalogUrl: string) {
 *   const service = createSPFxTenantPropertyService(context.spHttpClient);
 *   const setting = await service.get<string>('appVersion', catalogUrl);
 * }
 * ```
 * @useWhen You read existing StorageEntity values from an already known tenant catalog.
 *
 *
 * @example Read and validate an unknown StorageEntity
 * ```ts
 * import type { SPHttpClient } from '@microsoft/sp-http';
 *
 * import { createSPFxTenantPropertyService } from '@apvee/spfx-react-toolkit';
 * async function demonstrate(context: { spHttpClient: SPHttpClient }, catalogUrl: string) {
 *   const service = createSPFxTenantPropertyService(context.spHttpClient);
 *   const property = await service.get<unknown>('my-app:featureEnabled', catalogUrl);
 *   if (typeof property.value === 'boolean') console.log(property.value);
 * }
 * ```
 *
 * @scenario Read and validate a catalog StorageEntity
 * ```tsx
 * import * as React from 'react';
 * import { makeStyles, Button, Field, Input, Text } from '@fluentui/react-components';
 * import { createSPFxTenantPropertyService, useSPFxInstanceInfo, useSPFxPageContext, useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
 * import type { SPFxTenantProperty } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * const useStyles = makeStyles({ payload: { minWidth: 0, maxWidth: '100%', overflowX: 'auto' } });
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const page = useSPFxPageContext();
 *   const { client } = useSPFxSPHttpClient();
 *   const instance = useSPFxInstanceInfo();
 *   const styles = useStyles();
 *   const service = React.useMemo(() => client ? createSPFxTenantPropertyService(client) : undefined, [client]);
 *   const [catalog, setCatalog] = React.useState('');
 *   const [key, setKey] = React.useState('my-app:featureEnabled');
 *   const web = page.web.absoluteUrl.trim().replace(/\/+$/, '');
 *   const root = page.site.absoluteUrl.trim().replace(/\/+$/, '');
 *   const boundary = React.useRef({ page, client, web, root, instance: instance.id, revision: resetRevision, catalog, key, token: {} });
 *   const previous = boundary.current;
 *   if (previous.page !== page || previous.client !== client || previous.web !== web || previous.root !== root || previous.instance !== instance.id || previous.revision !== resetRevision || previous.catalog !== catalog || previous.key !== key) boundary.current = { page, client, web, root, instance: instance.id, revision: resetRevision, catalog, key, token: {} };
 *   const token = boundary.current.token;
 *   const mounted = React.useRef(true);
 *   const latest = React.useRef(0);
 *   const [outcome, setOutcome] = React.useState<{ token: object; operation: number; pending?: boolean; property?: SPFxTenantProperty; failure?: string }>();
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
 *   const resetSeen = React.useRef(resetRevision);
 *   React.useLayoutEffect(() => {
 *     if (resetSeen.current !== resetRevision) { resetSeen.current = resetRevision; setCatalog(''); setKey('my-app:featureEnabled'); setOutcome(undefined); }
 *   }, [resetRevision]);
 *   const current = (): boolean => mounted.current && boundary.current.token === token && page.web.absoluteUrl.trim().replace(/\/+$/, '') === web && page.site.absoluteUrl.trim().replace(/\/+$/, '') === root;
 *   const read = async (): Promise<void> => {
 *     if (!current() || !service || !catalog.trim() || !key) return;
 *     const operation = ++latest.current;
 *     const owns = (): boolean => current() && latest.current === operation;
 *     setOutcome({ token, operation, pending: true });
 *     try {
 *       const url = new URL(catalog.trim());
 *       if (!['https:', 'http:'].includes(url.protocol) || url.username || url.password || url.search || url.hash) throw new Error('Enter an absolute catalog URL without credentials, query or fragment.');
 *       const property = await service.get<unknown>(key, catalog.trim().replace(/\/+$/, ''));
 *       if (owns()) setOutcome({ token, operation, property });
 *     } catch (error) {
 *       if (owns()) setOutcome({ token, operation, failure: error instanceof Error ? error.message : String(error) });
 *     }
 *   };
 *   const visible = current() && outcome?.token === token && outcome.operation === latest.current ? outcome : undefined;
 *   const property = visible?.property;
 *   const message = property ? property.value === undefined ? 'Property not set' : typeof property.value === 'boolean' ? `Validated boolean: ${String(property.value)}` : 'Expected a boolean' : undefined;
 *   const regions: DemoRegions = {
 *     controls: <>
 *       <Field label="Catalog URL"><Input aria-label="Catalog URL" value={catalog} onChange={(_, data) => { if (current()) { boundary.current = { ...boundary.current, catalog: data.value, token: {} }; setCatalog(data.value); } }}/></Field>
 *       <Field label="Property key"><Input aria-label="Property key" value={key} onChange={(_, data) => { if (current()) { boundary.current = { ...boundary.current, key: data.value, token: {} }; setKey(data.value); } }}/></Field>
 *     </>,
 *     commands: <Button appearance="primary" disabled={!service || !catalog.trim() || !key || visible?.pending} onClick={read}>Read</Button>,
 *     status: visible?.failure ? <Text role="alert">{visible.failure}</Text> : <Text role="status">{visible?.pending ? 'Reading StorageEntity...' : message ?? 'Supply the known tenant catalog URL and Read explicitly.'}</Text>,
 *     children: property ? <dl aria-label="StorageEntity result"><dt>Value type</dt><dd>{property.value === undefined ? '(not set)' : typeof property.value}</dd>
 *       <dt>Description</dt><dd>{property.description === undefined ? '(not set)' : typeof property.description === 'string' ? JSON.stringify(property.description) : '(legacy description unavailable)'}</dd></dl> : undefined,
 *     diagnostics: <>
 *       <Text as="p">This factory does not discover the catalog, cache values, provision a list or write data. A current web URL is not a catalog address. HTTP 403/404 and malformed response JSON are failures. A missing or legacy REST-null Value means absence and clears its description. Defined stored text is decoded JSON-first; invalid stored JSON falls back to a string and fails this boolean validator.</Text>
 *       <Text as="p">Description metadata is passed through by the legacy service. Empty descriptions remain visible as quoted empty text; legacy null metadata is shown as unavailable. Reset clears inputs/results and invalidates publication, without cancelling an already-started read.</Text>
 *       {property && <div className={styles.payload} role="region" aria-label="Raw StorageEntity result" tabIndex={0}><pre>{JSON.stringify(property, undefined, 2)}</pre></div>}
 *     </>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Read a StorageEntity from the supplied catalog and distinguish validated boolean data from missing, malformed or failed data.' };
 * ```
 */
export function createSPFxTenantPropertyService(
  spHttpClient: SPHttpClient
): SPFxTenantPropertyService {
  const get = async <T = unknown>(
    key: string,
    catalogUrl: string
  ): Promise<SPFxTenantProperty<T>> => {
    const safeKey = encodeODataPathValue(key);
    const response: SPHttpClientResponse = await spHttpClient.get(
      `${catalogUrl}/_api/web/GetStorageEntity('${safeKey}')`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to read property: ${response.statusText}`);
    }

    const entity: StorageEntityResponse = await response.json();

    if (entity.Value !== undefined && entity.Value !== null) {
      return {
        value: deserializeTenantValue<T>(entity.Value),
        description: entity.Description
      };
    }

    return {
      value: undefined,
      description: undefined
    };
  };

  return {
    get
  };
}

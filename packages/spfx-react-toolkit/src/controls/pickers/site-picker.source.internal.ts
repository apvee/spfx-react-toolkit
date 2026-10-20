import type { SitePickerProps, SitePickerOption } from './picker.types';
import type { SitePickerValue } from './picker-values.types';
import type { PickerSourceContext, PickerSourceRequest, PickerPersonalSuggestions } from './picker-source.types.internal';
import { createMetadataPickerSource } from './picker-metadata-source.internal';
import { createPickerAlias, normalizePickerUrl } from './picker-identity.internal';
import { copyPickerValue, validatePickerOption } from './picker-values-validation.internal';
import { normalizeListId } from '../../helpers/spfx-list-target.helpers.internal';
import { restCollection, restRecord, restString, cachedPickerSpRead, validateRestContinuation } from './picker-rest.internal';

export interface SiteSearchOptions { readonly startRow: number; readonly pageSize: number; readonly trimDuplicates: boolean; }
export type SiteSearchTransport = (query: string, options: SiteSearchOptions) => Promise<{ readonly results: readonly { readonly data: unknown }[]; readonly totalResults: number }>;
const quoteKql = (value: string): string => `"${value.replace(/"/g, '""')}"`;

export function createSitePickerSource(props: SitePickerProps, context: PickerSourceContext, search?: SiteSearchTransport) {
  const mode = props.mode ?? 'site', collection = normalizePickerUrl(context.siteCollectionUrl);
  if (!['site', 'hub', 'associatedSites'].includes(mode)) throw new Error('Unsupported SitePicker mode.');
  if (props.additionalQuery !== undefined && mode !== 'site') throw new Error('additionalQuery is supported only in site mode.');
  const hubId = props.hubSiteId === undefined ? undefined : normalizeListId(props.hubSiteId, '[SitePicker]');
  const cursorScope = (request: PickerSourceRequest): string => JSON.stringify([context.tenantId, context.userId, request.cache.identityKey(context.owner), request.cache.identityKey(context.spClient), request.pageSize]);
  const discoveryIdentity = JSON.stringify(['site', mode, hubId, collection, props.additionalQuery, props.orderBy, props.trimDuplicates]);
  const identity = JSON.stringify([discoveryIdentity, props.includeCurrentSite !== false, props.includeHubSite !== false]);
  const map = (row: Readonly<Record<string, unknown>>): SitePickerOption => {
    const url = normalizePickerUrl(restString(row, 'Path')), title = restString(row, 'Title');
    const siteId = typeof row.SiteId === 'string' && row.SiteId !== '' ? normalizeListId(row.SiteId, '[SitePicker]') : undefined;
    const value: SitePickerValue = { kind: 'site', url, siteId };
    return validatePickerOption({ key: createPickerAlias(value), text: title, value, url, siteId,
      hubSiteId: typeof row.DepartmentId === 'string' && row.DepartmentId !== '' ? normalizeListId(row.DepartmentId, '[SitePicker]') : undefined });
  };
  const accepts = (option: SitePickerOption): boolean => option.value.kind === 'site';
  const siteInfo = async (request: PickerSourceRequest, url: string): Promise<Readonly<Record<string, unknown>>> =>
    restRecord(await cachedPickerSpRead(request, `${url}/_api/site?$select=Id,Url,HubSiteId,IsHubSite`));
  const associatedHubId = (site: Readonly<Record<string, unknown>>): string | undefined => {
    const id = typeof site.HubSiteId === 'string' && site.HubSiteId !== '' && site.HubSiteId !== '00000000-0000-0000-0000-000000000000' ? site.HubSiteId : site.IsHubSite === true ? restString(site, 'Id') : undefined;
    return id === undefined ? undefined : normalizeListId(id, '[SitePicker]');
  };
  async function association(request: PickerSourceRequest): Promise<string | undefined> {
    if (hubId !== undefined) return hubId;
    return associatedHubId(await siteInfo(request, collection));
  }
  async function discover(request: PickerSourceRequest, exactUrl?: string) {
    if (mode === 'hub') {
      const endpoint = `${collection}/_api/HubSites`;
      const result = await request.cache.request(`hubs:${discoveryIdentity}:${request.cache.identityKey(context.owner)}:${request.cache.identityKey(context.spClient)}:${context.tenantId}:${context.userId}`, async () => {
        const rows: Readonly<Record<string, unknown>>[] = [], visited = new Set<string>();
        let uri: string | undefined = `${endpoint}?$top=200`;
        for (let page = 0; page < 5 && uri !== undefined; page++) {
          if (visited.has(uri)) throw new Error('Repeated hub continuation.');
          visited.add(uri);
          const result = restCollection(await cachedPickerSpRead(request, uri)); rows.push(...result.rows);
          uri = result.next === undefined ? undefined : validateRestContinuation(result.next, endpoint);
        }
        return { rows, partial: uri !== undefined };
      }, request.partition);
      const all = result.rows.map(row => map({ Title: restString(row, 'Title'), Path: restString(row, 'SiteUrl'), SiteId: restString(row, 'ID') })).filter(accepts)
        .filter(option => exactUrl === undefined ? option.text.toLowerCase().includes(request.query.toLowerCase()) || option.url.toLowerCase().includes(request.query.toLowerCase()) : option.url === exactUrl);
      all.sort((a, b) => (props.orderBy?.field === 'url' ? a.url.localeCompare(b.url) : a.text.localeCompare(b.text)) * (props.orderBy?.direction === 'desc' ? -1 : 1));
      let offset = 0;
      if (request.continuation !== undefined) {
        const cursor = restRecord(JSON.parse(request.continuation));
        if (cursor.source !== identity || cursor.query !== request.query || cursor.scope !== cursorScope(request) || !Number.isSafeInteger(cursor.offset) || typeof cursor.offset !== 'number' || cursor.offset < 1) throw new Error('Invalid hub continuation.');
        offset = cursor.offset;
      }
      const next = offset + request.pageSize;
      return { options: all.slice(offset, next), continuation: next < all.length ? JSON.stringify({ source: identity, query: request.query, scope: cursorScope(request), offset: next }) : undefined,
        completeness: next < all.length || result.partial ? 'partial' as const : 'complete' as const,
        errors: result.partial ? [{ source: 'sharePoint' as const, retryable: false, error: new Error('Elenco hub parziale: limita l’ambito o usa la modalità site.') }] : undefined };
    }
    if (search === undefined) throw new Error('SharePoint search transport is not ready.');
    const associatedHub = mode === 'associatedSites' ? await association(request) : undefined;
    if (mode === 'associatedSites' && associatedHub === undefined) return { options: [], completeness: 'complete' as const };
    const conditions = ['contentclass:STS_Site'];
    if (associatedHub !== undefined) conditions.push(`DepartmentId:${quoteKql(associatedHub)}`);
    if (props.additionalQuery) conditions.push(`(${props.additionalQuery})`);
    if (exactUrl !== undefined) conditions.push(`Path=${quoteKql(exactUrl)}`);
    else if (request.query) conditions.push(`(Title:${quoteKql(request.query + '*')} OR Path:${quoteKql(request.query + '*')})`);
    const query = conditions.join(' AND ');
    let offset = 0;
    if (request.continuation !== undefined) {
      const cursor = restRecord(JSON.parse(request.continuation));
      if (cursor.source !== identity || cursor.query !== query || cursor.scope !== cursorScope(request) || typeof cursor.offset !== 'number' || !Number.isSafeInteger(cursor.offset) || cursor.offset < 1) throw new Error('Invalid SitePicker continuation.');
      offset = cursor.offset;
    }
    // Title/Path are not sortable in the standard SharePoint schema. Collect a bounded
    // native result set first; sort once before paging, never sort each remote page independently.
    const setKey = JSON.stringify(['site-search-set', discoveryIdentity, query, context.tenantId, context.userId, request.cache.identityKey(context.owner), request.cache.identityKey(context.spClient)]);
    const result = await request.cache.request(setKey, async () => {
      const unique = new Map<string, SitePickerOption>();
      let partial = false;
      for (let page = 0; page < 5; page++) {
        const startRow = page * 200;
        const response = await search(query, { startRow, pageSize: 200, trimDuplicates: props.trimDuplicates ?? true });
        if (!Number.isSafeInteger(response.totalResults) || response.totalResults < 0 || !Array.isArray(response.results)) throw new Error('Unexpected SharePoint search response.');
        response.results.map(result => map(restRecord(result.data))).filter(accepts).filter(option => exactUrl === undefined || option.url === exactUrl).forEach(option => unique.set(option.url, option));
        if (startRow + 200 >= response.totalResults) break;
        if (page === 4) partial = true;
      }
      const options = [...unique.values()].sort((a, b) => ((props.orderBy?.field === 'url' ? a.url.localeCompare(b.url) : a.text.localeCompare(b.text)) || a.url.localeCompare(b.url)) * (props.orderBy?.direction === 'desc' ? -1 : 1));
      return { options, partial };
    }, request.partition);
    const next = offset + request.pageSize;
    return { options: result.options.slice(offset, next), continuation: next < result.options.length ? JSON.stringify({ source: identity, query, scope: cursorScope(request), offset: next }) : undefined,
      completeness: next < result.options.length || result.partial ? 'partial' as const : 'complete' as const,
      errors: result.partial ? [{ source: 'sharePoint' as const, retryable: false, error: new Error('Ricerca site collection parziale (massimo 1000 risultati): restringi i criteri.') }] : undefined };

  }
  const source = createMetadataPickerSource<SitePickerValue, SitePickerOption>({ identity, context, kinds: ['site'], accepts, discover,
    async read(value, request, verify) {
      const url = normalizePickerUrl(value.url);
      if (verify || mode === 'hub') {
        const option = (await discover({ ...request, query: '', continuation: undefined }, url)).options[0];
        return option === undefined || value.siteId !== undefined && option.siteId !== value.siteId ? undefined : option;
      }
      const row = restRecord(await cachedPickerSpRead(request, `${url}/_api/web?$select=Title,Id`));
      const site = restRecord(await cachedPickerSpRead(request, `${url}/_api/site?$select=Id,Url`));
      if (normalizePickerUrl(restString(site, 'Url')) !== url || value.siteId !== undefined && normalizeListId(restString(site, 'Id'), '[SitePicker]') !== value.siteId) return undefined;
      const option = map({ Title: restString(row, 'Title'), Path: url, SiteId: restString(site, 'Id') });
      copyPickerValue(option.value);
      return option;
    },
  });
  const collectionOption = async (request: PickerSourceRequest, url: string): Promise<SitePickerOption> => {
    const [site, web] = await Promise.all([siteInfo(request, url), cachedPickerSpRead(request, `${url}/_api/web?$select=Title,Id`)]);
    if (normalizePickerUrl(restString(site, 'Url')) !== url) throw new Error('SitePicker requires a site collection root.');
    return map({ Title: restString(restRecord(web), 'Title'), Path: url, SiteId: restString(site, 'Id') });
  };
  const personalSuggestions = props.includeCurrentSite === false && props.includeHubSite === false ? undefined : async (request: PickerSourceRequest): Promise<PickerPersonalSuggestions<SitePickerOption>> => {
    const errors: { source: 'currentSite' | 'hubSite'; error: Error }[] = [];
    const eligible = async (option: SitePickerOption): Promise<SitePickerOption | undefined> =>
      mode === 'site' && props.additionalQuery === undefined ? option : source.verifyCandidate(option, request);
    const load = async (kind: 'currentSite' | 'hubSite', work: () => Promise<SitePickerOption | undefined>): Promise<SitePickerOption | undefined> => {
      try { return await work(); }
      catch (error) { errors.push({ source: kind, error: error instanceof Error ? error : new Error(String(error)) }); return undefined; }
    };
    const [current, hub] = await Promise.all([
      props.includeCurrentSite === false ? undefined : load('currentSite', async () => eligible(await collectionOption(request, collection))),
      props.includeHubSite === false ? undefined : load('hubSite', async () => {
        const site = await siteInfo(request, collection);
        // Only a real association qualifies; being a hub is not an association to itself.
        const associated = typeof site.HubSiteId === 'string' && site.HubSiteId !== '' && site.HubSiteId !== '00000000-0000-0000-0000-000000000000' ? normalizeListId(site.HubSiteId, '[SitePicker]') : undefined;
        if (associated === undefined) return undefined;
        const id = normalizeListId(associated, '[SitePicker]');
        const record = restRecord(await cachedPickerSpRead(request, `${collection}/_api/HubSites/GetById('${id}')`));
        const returnedId = normalizeListId(restString(record, typeof record.SiteId === 'string' ? 'SiteId' : 'ID'), '[SitePicker]');
        if (returnedId !== id) throw new Error('Unexpected associated hub identity.');
        const option = await collectionOption(request, normalizePickerUrl(restString(record, 'SiteUrl')));
        if (option.siteId !== id) throw new Error('Associated hub does not match its site collection.');
        return eligible(option);
      }),
    ]);
    return { currentSite: current === undefined ? undefined : [current], hubSite: hub === undefined ? undefined : [hub], errors, completeness: errors.length > 0 ? 'partial' : 'complete' };
  };
  return { ...source, persistentSuggestions: true, personalSuggestions, ready: context.spClient !== undefined && (mode === 'hub' || search !== undefined) };
}

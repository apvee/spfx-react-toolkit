import type { PrincipalPickerProps, PrincipalPickerOption, PrincipalPickerInitialPrincipal, UserPrincipal } from './picker.types';
import type { PrincipalPickerValue } from './picker-values.types';
import type { PickerSource, PickerSourceContext, PickerSourceRequest, PickerSourcePage } from './picker-source.types.internal';
import { createPickerAlias, decodePickerAlias, normalizePickerUrl, pickerIdentityAliases } from './picker-identity.internal';
import { copyPickerValue, validatePickerOption } from './picker-values-validation.internal';
import { graphContinuation, graphRead, graphUri, principalCollectionUrl, principalSpRead } from './principal-rest.internal';
import { graphGroupOption, graphUserOption, sharePointGroupOption, sharePointUserOption } from './principal-models.internal';
import { principalLoginKey } from './principal-login.internal';
import { readInitialPrincipalUser, readPrincipalGraphUser, readPrincipalSharePointUser } from './principal-read-identity.internal';
import { loadPrincipalSuggestions } from './principal-suggestions.internal';
import { principalScopeGroups, principalUserInScope } from './principal-scope.internal';
import { principalTransitiveUsers } from './principal-scope.internal';
import { confirmPrincipalSelection } from './principal-identity.internal';
import { normalizeListId } from '../../helpers/spfx-list-target.helpers.internal';
import { odataString, restCollection, restRecord, restString, restUrl, spRest, validateRestContinuation } from './picker-rest.internal';

const groupSelect = 'id,displayName,groupTypes,securityEnabled,mailEnabled,mail,description';
const errorOf = (error: unknown): Error => error instanceof Error ? error : new Error(String(error));
interface DiscoveryCursor { readonly source: string; readonly query: string; readonly pageSize: number; readonly scope: string; readonly pending: Readonly<Record<string, string>>; }

export function createPrincipalPickerSource(props: PrincipalPickerProps, context: PickerSourceContext): PickerSource<PrincipalPickerValue, PrincipalPickerOption> {
  const types = [...(props.principalTypes ?? ['user'])].sort(), web = normalizePickerUrl(props.webUrl ?? context.webUrl), identityMode = props.userIdentity?.mode ?? 'available';
  if (types.some(type => !['user', 'sharePointGroup', 'microsoft365Group', 'securityGroup', 'distributionList'].includes(type))) throw new Error('Unsupported principal type.');
  if (props.allowUnvalidatedEmail && (!types.includes('user') || identityMode !== 'available')) throw new Error('Unvalidated email requires user type and available identity mode.');
  const identity = JSON.stringify(['principal', types, web, props.userScope, props.useSubstrateSearch, props.loadPersonalSuggestions, props.initialSuggestions?.includeCurrentUser, props.allowUnvalidatedEmail, props.userIdentity]);
  const accepts = (option: PrincipalPickerOption): boolean => {
    try {
      validatePickerOption(option);
      if (option.kind === 'unvalidatedEmail') return props.allowUnvalidatedEmail === true && types.includes('user') && identityMode === 'available';
      if (!types.includes(option.kind)) return false;
      if (option.value.kind === 'user') return option.value.entra === undefined || option.value.entra.tenantId.toLowerCase() === context.tenantId.toLowerCase();
      if (option.value.kind === 'sharePointGroup') return props.webUrl !== undefined && normalizePickerUrl(props.webUrl) !== normalizePickerUrl(context.webUrl) || normalizePickerUrl(option.value.identity.siteCollectionUrl) === normalizePickerUrl(context.siteCollectionUrl);
      return option.value.identity.tenantId.toLowerCase() === context.tenantId.toLowerCase();
    } catch { return false; }
  };
  async function nativeUsers(query: string, request: PickerSourceRequest): Promise<readonly UserPrincipal[]> {
    const scope = props.userScope ?? { kind: 'tenant' as const };
    const target = scope.kind === 'web' ? normalizePickerUrl(scope.webUrl ?? web) : web;
    const endpoint = `${target}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
    const body = { queryParams: { AllowEmailAddresses: false, AllowMultipleEntities: false, AllUrlZones: false, MaximumEntitySuggestions: request.pageSize, PrincipalSource: scope.kind === 'web' ? 1 : 15,
      PrincipalType: 1, QueryString: query, SharePointGroupID: 0, UseSubstrateSearch: props.useSubstrateSearch ?? false } };
    const raw = restRecord(await request.cache.request(`principal-search:${request.cache.identityKey(context.owner)}:${request.cache.identityKey(context.spClient)}:${context.tenantId}:${context.userId}:${endpoint}:${JSON.stringify(body)}`, () => spRest(request, endpoint, body), request.partition));
    // JSON-light returns value; verbose REST uses the named operation after spRest unwraps d.
    const native = raw.value === undefined ? raw.ClientPeoplePickerSearchUser : raw.value;
    const rows: unknown = typeof native === 'string' ? JSON.parse(native) : native;
    if (!Array.isArray(rows)) throw new Error('Invalid native people-picker result.');
    const options: UserPrincipal[] = [];
    for (const input of rows) {
      const row = restRecord(input);
      if (row.IsResolved !== true || row.EntityType !== 'User') continue;
      const data = row.EntityData === undefined ? {} : restRecord(row.EntityData), loginName = restString(row, 'Key'), text = restString(row, 'DisplayText');
      const value: PrincipalPickerValue = { kind: 'user', loginName };
      options.push(validatePickerOption({ key: createPickerAlias(value), text, value, kind: 'user', loginName,
        email: typeof data.Email === 'string' ? data.Email : undefined, userPrincipalName: typeof data.UserPrincipalName === 'string' ? data.UserPrincipalName : undefined }));
    }
    return options;
  }
  async function scopedUsers(query: string, request: PickerSourceRequest, continuation?: string): Promise<PickerSourcePage<PrincipalPickerOption>> {
    const scope = props.userScope;
    if (scope?.kind !== 'groups' && scope?.kind !== 'sharePointGroupName') return { options: await nativeUsers(query, request), completeness: 'complete' };
    const groups = await principalScopeGroups(props, request), options: PrincipalPickerOption[] = [];
    let index = 0, nativeNext: string | undefined;
    if (continuation !== undefined) {
      const cursor = restRecord(JSON.parse(continuation));
      if (typeof cursor.index !== 'number' || !Number.isSafeInteger(cursor.index) || cursor.index < 0 || cursor.index >= groups.length || cursor.next !== undefined && typeof cursor.next !== 'string') throw new Error('Invalid membership continuation.');
      index = cursor.index; nativeNext = cursor.next as string | undefined;
    }
    let pages = 0;
    for (; index < groups.length && pages < 5; index++) {
      const group = groups[index];
      if (group.source === 'entra') {
        const endpoint = `/groups/${group.id}/transitiveMembers/microsoft.graph.user`;
        const data = nativeNext === undefined ? await principalTransitiveUsers(group.id, query, request.pageSize, request) : await graphRead(request, nativeNext, true);
        const rows = restCollection(data).rows;
        options.push(...rows.filter(row => row['@odata.type'] === undefined || row['@odata.type'] === '#microsoft.graph.user').map(row => graphUserOption(row, context)));
        nativeNext = graphContinuation(data, endpoint);
      } else {
        const collection = await principalCollectionUrl(request, group.webUrl), endpoint = `${normalizePickerUrl(group.webUrl)}/_api/web/sitegroups/getById(${group.id})/users`;
        const uri = restUrl(endpoint, { $select: 'Id,Title,LoginName,Email,PrincipalType', $filter: `PrincipalType eq 1 and (startswith(Title,${odataString(query)}) or startswith(Email,${odataString(query)}))`, $top: request.pageSize });
        const data = restCollection(await principalSpRead(request, nativeNext === undefined ? uri : validateRestContinuation(nativeNext, endpoint)));
        options.push(...data.rows.filter(row => row.PrincipalType === 1).map(row => sharePointUserOption(row, collection)));
        nativeNext = data.next;
      }
      pages++;
      if (nativeNext !== undefined) return { options, continuation: JSON.stringify({ index, next: nativeNext }), completeness: 'partial' };
      if (options.length >= request.pageSize) { index++; break; }
    }
    return { options, continuation: index < groups.length ? JSON.stringify({ index }) : undefined, completeness: index < groups.length ? 'partial' : 'complete' };
  }
  async function readValue(value: PrincipalPickerValue, request: PickerSourceRequest): Promise<PrincipalPickerOption> {
    copyPickerValue(value);
    switch (value.kind) {
      case 'user':
        if (value.entra !== undefined) {
          if (value.entra.tenantId.toLowerCase() !== context.tenantId.toLowerCase()) throw new Error('Principal belongs to another tenant.');
          const option = await readPrincipalGraphUser({ id: value.entra.id }, request);
          const registered = value.sharePoint === undefined ? undefined : await readPrincipalSharePointUser({ ...option.value, sharePoint: value.sharePoint }, request, value.sharePoint.siteCollectionUrl);
          return { ...option, value: { ...option.value, sharePoint: registered?.value.sharePoint }, sharePointId: registered?.sharePointId, sharePointWebUrl: registered?.sharePointWebUrl };
        }
        if (value.sharePoint !== undefined) return readPrincipalSharePointUser(value, request, value.sharePoint.siteCollectionUrl);
        return readInitialPrincipalUser({ kind: 'login', loginName: value.loginName }, request, web);
      case 'sharePointGroup': return sharePointGroupOption(await principalSpRead(request, `${normalizePickerUrl(value.identity.siteCollectionUrl)}/_api/web/sitegroups/getById(${value.identity.id})?$select=Id,Title,Description`), normalizePickerUrl(value.identity.siteCollectionUrl));
      case 'microsoft365Group': case 'securityGroup': case 'distributionList': {
        if (value.identity.tenantId.toLowerCase() !== context.tenantId.toLowerCase()) throw new Error('Group belongs to another tenant.');
        const option = graphGroupOption(await graphRead(request, graphUri(`/groups/${normalizeListId(value.identity.id, '[PrincipalPicker]')}`, { $select: groupSelect })), context);
        if (option === undefined) throw new Error('Group classification metadata is unavailable.');
        return option;
      }
      case 'unvalidatedEmail': return validatePickerOption<PrincipalPickerOption>({ key: createPickerAlias(value), text: value.text ?? value.email, value, kind: 'unvalidatedEmail', email: value.email });
    }
  }
  async function initial(reference: PrincipalPickerInitialPrincipal, request: PickerSourceRequest): Promise<PrincipalPickerOption | undefined> {
    if (reference.kind === 'displayOnly') return undefined;
    if (reference.kind !== 'group') return readInitialPrincipalUser(reference, request, web);
    if (reference.group.source === 'entra') return readValue({ kind: 'securityGroup', identity: { tenantId: context.tenantId, id: reference.group.id } }, request);
    const collection = await principalCollectionUrl(request, reference.group.webUrl);
    return readValue({ kind: 'sharePointGroup', identity: { siteCollectionUrl: collection, id: reference.group.id } }, request);
  }
  const source: PickerSource<PrincipalPickerValue, PrincipalPickerOption> = {
    preflight: true, family: 'principal', identity, context, ready: types.length === 0 || context.spClient !== undefined || context.graphClient !== undefined || context.getGraphClient !== undefined, transport: 'sharePoint', accepts,
    async discover(request) {
      const scopeKey = JSON.stringify([context.tenantId, context.userId, request.cache.identityKey(context.owner), request.cache.identityKey(context.spClient), request.cache.identityKey(context.graphClientIdentity ?? context.graphClient)]);
      let pending: Readonly<Record<string, string>> | undefined;
      if (request.continuation !== undefined) {
        const cursor = restRecord(JSON.parse(request.continuation));
        if (cursor.source !== identity || cursor.query !== request.query || cursor.pageSize !== request.pageSize || cursor.scope !== scopeKey) throw new Error('Invalid principal continuation.');
        pending = restRecord(cursor.pending) as Readonly<Record<string, string>>;
      }
      const tasks: { name: string; transport: 'sharePoint' | 'graph'; run: () => Promise<PickerSourcePage<PrincipalPickerOption>> }[] = [];
      if (types.includes('user') && (pending === undefined || pending.users !== undefined)) tasks.push({ name: 'users', transport: props.userScope?.kind === 'groups' ? 'graph' : 'sharePoint', run: () => scopedUsers(request.query, request, pending?.users) });
      if (types.includes('sharePointGroup') && (pending === undefined || pending.sharePointGroups !== undefined)) tasks.push({ name: 'sharePointGroups', transport: 'sharePoint', run: async () => {
        const collection = await principalCollectionUrl(request, web), endpoint = `${web}/_api/web/sitegroups`;
        const uri = restUrl(endpoint, { $select: 'Id,Title,Description', $filter: `startswith(Title,${odataString(request.query)})`, $top: request.pageSize });
        const data = restCollection(await principalSpRead(request, pending?.sharePointGroups === undefined ? uri : validateRestContinuation(pending.sharePointGroups, endpoint)));
        return { options: data.rows.map(row => sharePointGroupOption(row, collection)), continuation: data.next, completeness: data.next === undefined ? 'complete' : 'partial' };
      } });
      if (types.some(type => type !== 'user' && type !== 'sharePointGroup') && (pending === undefined || pending.entraGroups !== undefined)) tasks.push({ name: 'entraGroups', transport: 'graph', run: async () => {
        const uri = graphUri('/groups', { $select: groupSelect, $filter: `startswith(displayName,${odataString(request.query)})`, $top: request.pageSize });
        const data = await graphRead(request, pending?.entraGroups ?? uri);
        return { options: restCollection(data).rows.flatMap(row => { const option = graphGroupOption(row, context); return option === undefined ? [] : [option]; }), continuation: graphContinuation(data, '/groups'), completeness: data['@odata.nextLink'] === undefined ? 'complete' : 'partial' };
      } });
      const results = await Promise.allSettled(tasks.map(task => task.run()));
      const options: PrincipalPickerOption[] = [], next: Record<string, string> = {}, errors: { source: 'sharePoint' | 'graph'; error: Error }[] = [];
      let partial = false;
      results.forEach((result, index) => {
        const task = tasks[index];
        if (result.status === 'rejected') { errors.push({ source: task.transport, error: errorOf(result.reason) }); partial = true; }
        else { options.push(...result.value.options); if (result.value.continuation !== undefined) next[task.name] = result.value.continuation; partial ||= result.value.completeness === 'partial'; }
      });
      if (pending === undefined && props.allowUnvalidatedEmail && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(request.query) && options.every(option => option.kind !== 'user' || option.email?.toLowerCase() !== request.query.toLowerCase())) {
        const value: PrincipalPickerValue = { kind: 'unvalidatedEmail', email: request.query };
        options.push(validatePickerOption({ key: createPickerAlias(value), text: request.query, value, kind: 'unvalidatedEmail', email: request.query }));
      }
      const unique = new Map<string, PrincipalPickerOption>();
      for (const option of options.filter(accepts)) if (!pickerIdentityAliases(option.value, context).some(alias => unique.has(alias))) pickerIdentityAliases(option.value, context).forEach(alias => unique.set(alias, option));
      return { options: [...new Set(unique.values())], continuation: Object.keys(next).length === 0 ? undefined : JSON.stringify({ source: identity, query: request.query, pageSize: request.pageSize, scope: scopeKey, pending: next } satisfies DiscoveryCursor), completeness: partial ? 'partial' : 'complete', errors };
    },
    async resolve(values, request) {
      const resolved: { value: PrincipalPickerValue; option: PrincipalPickerOption }[] = [], unresolved: { value: PrincipalPickerValue; error?: Error }[] = [];
      for (const value of values) {
        try { const option = await readValue(value, request); resolved.push({ value: option.value, option }); }
        catch (error) { unresolved.push({ value, error: errorOf(error) }); }
      }
      return { resolved, unresolved };
    },
    async hydrate(inputs, request) {
      const resolved: { key: string; value: PrincipalPickerValue; option: PrincipalPickerOption }[] = [], unresolved: { key: string; label?: string; error?: Error }[] = [];
      for (const input of inputs) {
        try {
          const value = input.kind === 'value' ? input.value : input.kind === 'key' ? decodePickerAlias(input.key) : undefined;
          const option = input.kind === 'principalReference' ? await initial(input.reference, request) : value === undefined ? undefined :
            ['user', 'sharePointGroup', 'microsoft365Group', 'securityGroup', 'distributionList', 'unvalidatedEmail'].includes(value.kind) ? await readValue(value as PrincipalPickerValue, request) : undefined;
          if (option === undefined) unresolved.push({ key: input.key, label: input.kind === 'principalReference' ? input.reference.text : undefined });
          else resolved.push({ key: input.key, value: option.value, option: { ...option, key: input.key } });
        } catch (error) { unresolved.push({ key: input.key, error: errorOf(error) }); }
      }
      return { resolved, unresolved };
    },
    async verifyCandidate(option, request) {
      if (!accepts(option)) return undefined;
      let verified: PrincipalPickerOption;
      if (option.kind === 'user' && option.value.entra === undefined && option.value.sharePoint === undefined) {
        // Native search accepts the UPN; the membership claim wrapper remains part of the returned identity.
        const users = await nativeUsers(principalLoginKey(option.loginName), { ...request, pageSize: 50 });
        const exact = users.find(user => principalLoginKey(user.loginName) === principalLoginKey(option.loginName));
        if (exact === undefined) return undefined;
        verified = exact;
      } else verified = await readValue(option.value, request);
      if (verified.kind === 'user' && !await principalUserInScope(verified, props, request)) return undefined;
      if (verified.kind === 'sharePointGroup' && verified.value.kind === 'sharePointGroup' && verified.value.identity.siteCollectionUrl !== await principalCollectionUrl(request, web)) return undefined;
      if (!accepts(verified)) return undefined;
      return { ...verified, key: option.key, text: option.text };
    },
    async needsIdentityResolution(option, request) {
      if (option.kind !== 'user' || identityMode === 'available') return false;
      if ((identityMode === 'entra' || identityMode === 'both') && (option.value.entra === undefined || option.value.entra.tenantId.toLowerCase() !== context.tenantId.toLowerCase())) return true;
      if (identityMode === 'sharePoint' || identityMode === 'both') {
        if (option.value.sharePoint === undefined) return true;
        const destination = props.userIdentity !== undefined && 'webUrl' in props.userIdentity ? props.userIdentity.webUrl ?? web : web;
        return normalizePickerUrl(option.value.sharePoint.siteCollectionUrl) !== await principalCollectionUrl(request, normalizePickerUrl(destination));
      }
      return false;
    },
    personalSuggestions: props.loadPersonalSuggestions || props.initialSuggestions?.includeCurrentUser ? request => loadPrincipalSuggestions(props, request) : undefined,
    async confirm(options, request, reason) {
      const verified: PrincipalPickerOption[] = [];
      for (const option of options) {
        if (request.isCurrent?.() === false) throw new Error('Principal confirmation cancelled.');
        const candidate = await source.verifyCandidate(option, request);
        if (candidate === undefined) throw new Error('Principal is outside the configured type or scope.');
        verified.push(candidate);
      }
      return confirmPrincipalSelection(verified, props, request, reason);
    },
  };
  return source;
}

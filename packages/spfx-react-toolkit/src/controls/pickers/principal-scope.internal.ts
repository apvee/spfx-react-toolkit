import type { PrincipalPickerProps, PickerGroupReference, UserPrincipal } from './picker.types';
import type { PickerSourceRequest } from './picker-source.types.internal';
import type { PrincipalLookupBudget } from './principal-rest.internal';
import { graphRead, principalSpRead, principalCollectionUrl, principalGraphClient } from './principal-rest.internal';
import { readPrincipalGraphUser, readPrincipalSharePointUser } from './principal-read-identity.internal';
import { principalLoginKey } from './principal-login.internal';
import { normalizeListId } from '../../helpers/spfx-list-target.helpers.internal';
import { normalizePickerUrl } from './picker-identity.internal';
import { odataString, restCollection, restNumber, restUrl } from './picker-rest.internal';

export async function principalScopeGroups(props: PrincipalPickerProps, request: PickerSourceRequest, budget?: PrincipalLookupBudget): Promise<readonly PickerGroupReference[]> {
  const scope = props.userScope;
  if (scope?.kind === 'groups') {
    return scope.groups.map(group => {
      if (group.source === 'entra') return { source: 'entra' as const, id: normalizeListId(group.id, '[PrincipalPicker]') };
      if (!Number.isSafeInteger(group.id) || group.id <= 0) throw new Error('Invalid SharePoint membership group ID.');
      return { source: 'sharePoint' as const, id: group.id, webUrl: normalizePickerUrl(group.webUrl) };
    });
  }
  if (scope?.kind !== 'sharePointGroupName') return [];
  const web = normalizePickerUrl(scope.webUrl ?? props.webUrl ?? request.context.webUrl);
  const group = await principalSpRead(request, `${web}/_api/web/sitegroups/getByName(${encodeURIComponent(odataString(scope.name))})?$select=Id`, budget);
  const id = restNumber(group, 'Id');
  if (!Number.isSafeInteger(id) || id <= 0) throw new Error('Invalid SharePoint named group ID.');
  return [{ source: 'sharePoint', webUrl: web, id }];
}

/** Candidate-aware read-only membership; current-user permission helpers cannot establish another person's scope. */
export async function principalUserInScope(option: UserPrincipal, props: PrincipalPickerProps, request: PickerSourceRequest, budget?: PrincipalLookupBudget): Promise<boolean> {
  const scope = props.userScope ?? { kind: 'tenant' as const };
  if (option.value.entra !== undefined && option.value.entra.tenantId.toLowerCase() !== request.context.tenantId.toLowerCase()) return false;
  if (scope.kind === 'tenant') return true;
  if (scope.kind === 'web') {
    try {
      await readPrincipalSharePointUser(option.value, request, scope.webUrl ?? props.webUrl ?? request.context.webUrl, budget);
      return true;
    } catch (error) { if (error instanceof Error && 'status' in error && error.status === 404) return false; throw error; }
  }
  const groups = await principalScopeGroups(props, request, budget);
  let failure: unknown;
  const entra = groups.filter((group): group is Extract<PickerGroupReference, { source: 'entra' }> => group.source === 'entra');
  if (entra.length > 0) {
    try {
      const user = await readPrincipalGraphUser(option.value.entra === undefined ? { upn: option.userPrincipalName ?? option.loginName } : { id: option.value.entra.id }, request, budget);
      for (let index = 0; index < entra.length; index += 20) {
        const ids = entra.slice(index, index + 20).map(group => group.id);
        // POST checkMemberGroups is a read-only membership action, never provisioning.
        const uri = `/users/${user.entraId}/checkMemberGroups`;
        const client = await principalGraphClient(request);
        const key = `principal-membership:${request.cache.identityKey(request.context.owner)}:${request.cache.identityKey(client)}:${request.context.tenantId}:${request.context.userId}:${uri}:${JSON.stringify(ids)}`;
        if (!request.cache.has(key) && budget !== undefined && budget.remaining-- <= 0) throw new Error('Personal suggestion scope lookup budget exhausted.');
        const result = await request.cache.request(key, async () => {
          return await client.api(uri).version('v1.0').post({ groupIds: ids }) as unknown;
        }, request.partition);
        const record = result !== null && typeof result === 'object' ? result as Record<string, unknown> : undefined;
        if (record === undefined || !Array.isArray(record.value) || !record.value.every(id => typeof id === 'string' && ids.includes(id))) throw new Error('Invalid candidate group-membership response.');
        if (record.value.length > 0) return true;
      }
    } catch (error) { failure = error; }
  }
  for (const group of groups) {
    if (group.source !== 'sharePoint') continue;
    try {
      const collection = await principalCollectionUrl(request, group.webUrl, budget);
      const registered = await readPrincipalSharePointUser(option.value, request, group.webUrl, budget);
      const uri = restUrl(`${normalizePickerUrl(group.webUrl)}/_api/web/sitegroups/getById(${group.id})/users`, { $select: 'Id,LoginName,PrincipalType', $filter: `Id eq ${registered.sharePointId} and PrincipalType eq 1`, $top: 1 });
      const rows = restCollection(await principalSpRead(request, uri, budget)).rows;
      if (rows.some(row => restNumber(row, 'Id') === registered.sharePointId && principalLoginKey(String(row.LoginName)) === principalLoginKey(option.loginName)) && registered.sharePointWebUrl === collection) return true;
    } catch (error) { failure = error; }
  }
  if (failure !== undefined) throw failure;
  return false;
}

/** Cached transitive user lookup projection; used by source discovery, not signed-in access booleans. */
export async function principalTransitiveUsers(groupId: string, query: string, pageSize: number, request: PickerSourceRequest) {
  const id = normalizeListId(groupId, '[PrincipalPicker]');
  return graphRead(request, `/groups/${id}/transitiveMembers/microsoft.graph.user?$select=id,displayName,userPrincipalName,mail&$filter=${encodeURIComponent(`startswith(displayName,${odataString(query)}) or startswith(userPrincipalName,${odataString(query)})`)}&$count=true&$top=${pageSize}`, true);
}

import type { PrincipalPickerProps, PrincipalPickerOption, UserPrincipal } from './picker.types';
import type { PickerSourcePage, PickerSourceRequest, PickerPersonalSuggestions } from './picker-source.types.internal';
import type { PrincipalLookupBudget } from './principal-rest.internal';
import { graphContinuation, graphRead, graphUri, PrincipalBudgetError, principalCollectionUrl, principalSpRead } from './principal-rest.internal';
import { graphGroupOption, sharePointUserOption } from './principal-models.internal';
import { readPrincipalGraphUser } from './principal-read-identity.internal';
import { principalUserInScope } from './principal-scope.internal';
import { canonicalPickerIdentity } from './picker-identity.internal';
import { restCollection, restRecord } from './picker-rest.internal';

/** Requested delegated endpoints. People is intentionally retained despite its maintenance status. */
export async function loadPrincipalPersonalSuggestions(source: 'people' | 'memberOf', props: PrincipalPickerProps, request: PickerSourceRequest, budget: PrincipalLookupBudget): Promise<PickerSourcePage<PrincipalPickerOption>> {
  const types = props.principalTypes ?? ['user'];
  const users = types.includes('user'), groups = types.some(type => type !== 'user' && type !== 'sharePointGroup');
  if (!props.loadPersonalSuggestions || request.query !== '' || props.showDefaultSuggestions === false || (source === 'people' ? !users : !groups)) return { options: [], completeness: 'complete' };
  const endpoint = source === 'people' ? '/me/people' : '/me/memberOf';
  let uri: string | undefined = graphUri(endpoint, { $top: 50 });
  const unique = new Map<string, PrincipalPickerOption>(), visited = new Set<string>();
  let partial = false;
  for (let page = 0; uri !== undefined && page < 5; page++) {
    if (visited.has(uri)) throw new Error('Repeated personal suggestion continuation.');
    visited.add(uri);
    const data = await graphRead(request, uri), result = restCollection(data);
    for (const row of result.rows) {
      let option: PrincipalPickerOption | undefined;
      try {
        if (source === 'people') {
          if (row.personType === undefined || row.personType === null) continue;
          const type = restRecord(row.personType);
          if (type.class !== 'Person' || type.subclass !== 'OrganizationUser') continue;
          if (typeof row.userPrincipalName !== 'string' || !row.userPrincipalName) { partial = true; continue; }
          // person.id is aggregate-person identity; only /users establishes an Entra user ID.
          const user = await readPrincipalGraphUser({ upn: row.userPrincipalName }, request, budget);
          if (!await principalUserInScope(user, props, request, budget)) continue;
          option = user;
        } else {
          if (row['@odata.type'] !== '#microsoft.graph.group') continue;
          option = graphGroupOption(row, request.context);
          if (option === undefined) { partial = true; continue; }
        }
        if (option.kind === 'unvalidatedEmail' || !types.includes(option.kind)) continue;
        unique.set(canonicalPickerIdentity(option.value, request.context), option);
      } catch (error) {
        if (error instanceof PrincipalBudgetError || error instanceof Error && error.message.includes('budget')) { partial = true; uri = undefined; break; }
        throw error;
      }
    }
    if (uri === undefined) break;
    uri = graphContinuation(data, endpoint);
    if (page === 4 && uri !== undefined) partial = true;
  }
  return { options: [...unique.values()], completeness: partial ? 'partial' : 'complete' };
}

export async function loadPrincipalSuggestions(props: PrincipalPickerProps, request: PickerSourceRequest): Promise<PickerPersonalSuggestions<PrincipalPickerOption>> {
  const budget: PrincipalLookupBudget = { remaining: 20 };
  const data: { currentUser?: readonly UserPrincipal[]; people?: readonly PrincipalPickerOption[]; groups?: readonly PrincipalPickerOption[]; errors: { source: 'people' | 'memberOf' | 'currentUser'; error: Error }[]; completeness: 'complete' | 'partial' } = { errors: [], completeness: 'complete' };
  const current = props.initialSuggestions?.includeCurrentUser === true && (props.principalTypes ?? ['user']).includes('user');
  const sources: readonly ('currentUser' | 'people' | 'memberOf')[] = ['currentUser', 'people', 'memberOf'];
  // Source failures are independent. Ordering is policy, not an assumption about completion timing.
  const results = await Promise.allSettled(sources.map(async source => {
    if (source !== 'currentUser') return loadPrincipalPersonalSuggestions(source, props, request, budget);
    if (!current || props.showDefaultSuggestions === false || request.query !== '') return { options: [] as UserPrincipal[], completeness: 'complete' as const };
    const web = props.webUrl ?? request.context.webUrl, collection = await principalCollectionUrl(request, web, budget);
    const row = await principalSpRead(request, `${web.replace(/\/$/, '')}/_api/web/currentuser?$select=Id,Title,LoginName,Email,PrincipalType`, budget);
    const option = sharePointUserOption(row, collection);
    return { options: await principalUserInScope(option, props, request, budget) ? [option] : [], completeness: 'complete' as const };
  }));
  results.forEach((result, index) => {
    const source = sources[index];
    if (result.status === 'rejected') {
      data.errors.push({ source, error: result.reason instanceof Error ? result.reason : new Error(String(result.reason)) }); data.completeness = 'partial';
    } else {
      if (result.value.completeness === 'partial') data.completeness = 'partial';
      if (source === 'currentUser') data.currentUser = result.value.options.filter((option): option is UserPrincipal => option.kind === 'user');
      else if (source === 'people') data.people = result.value.options;
      else data.groups = result.value.options;
    }
  });
  return data;
}

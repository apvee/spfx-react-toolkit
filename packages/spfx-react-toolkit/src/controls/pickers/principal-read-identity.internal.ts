import type { PrincipalPickerInitialPrincipal, UserPrincipal } from './picker.types';
import type { UserPrincipalValue } from './picker-values.types';
import type { PickerSourceRequest } from './picker-source.types.internal';
import type { PrincipalLookupBudget } from './principal-rest.internal';
import { graphRead, graphUri, principalCollectionUrl, principalSpRead } from './principal-rest.internal';
import { graphUserOption, sharePointUserOption } from './principal-models.internal';
import { principalLoginKey } from './principal-login.internal';
import { odataString, restCollection } from './picker-rest.internal';
import { normalizeListId } from '../../helpers/spfx-list-target.helpers.internal';
import { normalizePickerUrl } from './picker-identity.internal';

/** Read-only lookup of an authoritative directory ID/UPN or unambiguous email reference. */
export async function readPrincipalGraphUser(reference: { readonly id?: string; readonly upn?: string; readonly email?: string }, request: PickerSourceRequest, budget?: PrincipalLookupBudget): Promise<UserPrincipal> {
  const select = 'id,displayName,userPrincipalName,mail';
  if (reference.email !== undefined) {
    const response = await graphRead(request, graphUri('/users', { $select: select, $filter: `mail eq ${odataString(reference.email)} or userPrincipalName eq ${odataString(reference.email)}`, $top: 2 }), false, budget);
    const options = restCollection(response).rows.map(row => graphUserOption(row, request.context)).filter(option => option.email?.toLowerCase() === reference.email?.toLowerCase() || option.userPrincipalName?.toLowerCase() === reference.email?.toLowerCase());
    const unique = new Map(options.map(option => [option.entraId, option]));
    if (unique.size !== 1 || response['@odata.nextLink'] !== undefined) throw new Error('Email does not identify exactly one directory user.');
    return [...unique.values()][0];
  }
  const id = reference.id === undefined ? undefined : normalizeListId(reference.id, '[PrincipalPicker]');
  const upn = reference.upn === undefined ? undefined : principalLoginKey(reference.upn);
  if (id === undefined && (!upn || !upn.includes('@'))) throw new Error('A verified Entra ID or UPN is required for directory resolution.');
  const response = await graphRead(request, graphUri(`/users/${encodeURIComponent(id ?? upn ?? '')}`, { $select: select }), false, budget);
  const option = graphUserOption(response, request.context);
  if (id !== undefined ? option.entraId !== id : principalLoginKey(option.loginName) !== upn) throw new Error('Directory response does not match the requested user identity.');
  return option;
}

/** Reads an existing registration only. No discovery/hydration code can register a user. */
export async function readPrincipalSharePointUser(value: UserPrincipalValue, request: PickerSourceRequest, webUrl: string, budget?: PrincipalLookupBudget): Promise<UserPrincipal> {
  const web = normalizePickerUrl(webUrl), collection = await principalCollectionUrl(request, web, budget);
  const useId = value.sharePoint !== undefined && normalizePickerUrl(value.sharePoint.siteCollectionUrl) === collection;
  const uri = useId ? `${web}/_api/web/siteusers/getById(${value.sharePoint?.id})?$select=Id,LoginName,Title,Email,PrincipalType` :
    `${web}/_api/web/siteusers/getByLoginName(@login)?@login=${encodeURIComponent(odataString(value.loginName))}&$select=Id,LoginName,Title,Email,PrincipalType`;
  const row = await principalSpRead(request, uri, budget), option = sharePointUserOption(row, collection);
  if (principalLoginKey(option.loginName) !== principalLoginKey(value.loginName)) throw new Error('SharePoint registration does not match the requested login.');
  return option;
}
export async function readInitialPrincipalUser(reference: Extract<PrincipalPickerInitialPrincipal, { kind: 'login' | 'email' | 'entraUser' }>, request: PickerSourceRequest, webUrl: string): Promise<UserPrincipal> {
  if (reference.kind === 'email') return readPrincipalGraphUser({ email: reference.email }, request);
  if (reference.kind === 'entraUser') return readPrincipalGraphUser({ id: reference.id }, request);
  try { return await readPrincipalSharePointUser({ kind: 'user', loginName: reference.loginName }, request, webUrl); }
  catch { return readPrincipalGraphUser({ upn: reference.loginName }, request); }
}

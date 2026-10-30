import type { PrincipalPickerOption, UserPrincipal } from './picker.types';
import type { PrincipalPickerValue, UserPrincipalValue } from './picker-values.types';
import type { PickerSourceContext } from './picker-source.types.internal';
import { createPickerAlias } from './picker-identity.internal';
import { copyPickerValue, validatePickerOption } from './picker-values-validation.internal';
import { normalizeListId } from '../../helpers/spfx-list-target.helpers.internal';
import { restBoolean, restNumber, restString } from './picker-rest.internal';
import { optionalGraphText } from './principal-rest.internal';

/** Identity comes from an authenticated /users response, never from a People person.id or a guessed email. */
export function graphUserOption(row: Readonly<Record<string, unknown>>, context: PickerSourceContext): UserPrincipal {
  const id = normalizeListId(restString(row, 'id'), '[PrincipalPicker]'), upn = restString(row, 'userPrincipalName'), text = restString(row, 'displayName');
  const value: UserPrincipalValue = { kind: 'user', loginName: upn, entra: { tenantId: normalizeListId(context.tenantId, '[PrincipalPicker]'), id } };
  return validatePickerOption({ key: createPickerAlias(value), text, value, kind: 'user', loginName: upn, userPrincipalName: upn, email: optionalGraphText(row.mail), entraId: id, entraTenantId: value.entra?.tenantId });
}
export function sharePointUserOption(row: Readonly<Record<string, unknown>>, siteCollectionUrl: string): UserPrincipal {
  if (row.PrincipalType !== 1) throw new Error('SharePoint principal is not a user.');
  const loginName = restString(row, 'LoginName'), id = restNumber(row, 'Id'), text = restString(row, 'Title');
  const value: UserPrincipalValue = { kind: 'user', loginName, sharePoint: { siteCollectionUrl, id } };
  return validatePickerOption({ key: createPickerAlias(value), text, value, kind: 'user', loginName, email: optionalGraphText(row.Email), sharePointId: id, sharePointWebUrl: siteCollectionUrl });
}
export function graphGroupOption(row: Readonly<Record<string, unknown>>, context: PickerSourceContext): PrincipalPickerOption | undefined {
  if (!Array.isArray(row.groupTypes) || !row.groupTypes.every(type => typeof type === 'string') || typeof row.securityEnabled !== 'boolean' || typeof row.mailEnabled !== 'boolean' || typeof row.displayName !== 'string' || !row.displayName.trim()) return undefined;
  const kind = row.groupTypes.includes('Unified') ? 'microsoft365Group' : row.securityEnabled ? 'securityGroup' : row.mailEnabled ? 'distributionList' : undefined;
  if (kind === undefined) return undefined;
  const id = normalizeListId(restString(row, 'id'), '[PrincipalPicker]'), tenantId = normalizeListId(context.tenantId, '[PrincipalPicker]');
  const value: PrincipalPickerValue = { kind, identity: { tenantId, id } };
  copyPickerValue(value);
  return validatePickerOption<PrincipalPickerOption>({ key: createPickerAlias(value), text: row.displayName, value, kind, source: 'entra', id, tenantId, mail: optionalGraphText(row.mail), description: optionalGraphText(row.description), securityEnabled: restBoolean(row, 'securityEnabled'), mailEnabled: restBoolean(row, 'mailEnabled') });
}
export function sharePointGroupOption(row: Readonly<Record<string, unknown>>, collection: string): PrincipalPickerOption {
  const id = restNumber(row, 'Id'), text = restString(row, 'Title');
  const value: PrincipalPickerValue = { kind: 'sharePointGroup', identity: { siteCollectionUrl: collection, id } };
  return validatePickerOption({ key: createPickerAlias(value), text, value, kind: 'sharePointGroup', source: 'sharePoint', id, webUrl: collection, description: optionalGraphText(row.Description) });
}

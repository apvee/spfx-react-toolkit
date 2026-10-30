import type { PrincipalPickerProps, PrincipalPickerOption, UserPrincipal } from './picker.types';
import type { UserPrincipalValue } from './picker-values.types';
import type { PickerSourceRequest } from './picker-source.types.internal';
import { copyPickerValue, validatePickerOption } from './picker-values-validation.internal';
import { readPrincipalGraphUser, readPrincipalSharePointUser } from './principal-read-identity.internal';
import { principalCollectionUrl } from './principal-rest.internal';
import { sharePointUserOption } from './principal-models.internal';
import { principalLoginKey } from './principal-login.internal';
import { normalizePickerUrl } from './picker-identity.internal';
import { principalUserInScope } from './principal-scope.internal';
import { restRecord, spRest } from './picker-rest.internal';

const assertCurrent = (request: PickerSourceRequest): void => { if (request.isCurrent?.() === false) throw new Error('Principal confirmation cancelled.'); };
/** Only explicit confirmation/resolution may call ensureUser. Reads, seeds, discovery and config changes cannot. */
export async function confirmPrincipalSelection(options: readonly PrincipalPickerOption[], props: PrincipalPickerProps, request: PickerSourceRequest, reason: 'selection' | 'explicitResolution'): Promise<readonly PrincipalPickerOption[]> {
  if (!['selection', 'explicitResolution'].includes(reason)) throw new Error('Identity writes require an explicit selection or resolution action.');
  const policy = props.userIdentity ?? { mode: 'available' as const };
  if (!['available', 'entra', 'sharePoint', 'both'].includes(policy.mode)) throw new Error('Unsupported user identity mode.');
  const ensure = 'ensureUser' in policy && policy.ensureUser === true;
  if (ensure && policy.mode !== 'sharePoint' && policy.mode !== 'both') throw new Error('ensureUser requires SharePoint identity resolution.');
  assertCurrent(request);
  const result: PrincipalPickerOption[] = [];
  for (const input of options) {
    validatePickerOption(input);
    assertCurrent(request);
    if (input.kind !== 'user') {
      if (input.kind === 'unvalidatedEmail' && (!props.allowUnvalidatedEmail || policy.mode !== 'available')) throw new Error('Unvalidated email is disabled for this identity mode.');
      result.push({ ...input, value: copyPickerValue(input.value) } as PrincipalPickerOption);
      continue;
    }
    let value: UserPrincipalValue = { ...input.value }, enriched: UserPrincipal = input;
    if (!await principalUserInScope(input, props, request)) throw new Error('User is outside the configured scope.');
    assertCurrent(request);
    if (policy.mode === 'entra' || policy.mode === 'both') {
      if (value.entra !== undefined && value.entra.tenantId.toLowerCase() !== request.context.tenantId.toLowerCase()) throw new Error('User ID belongs to another Entra tenant.');
      const directory = await readPrincipalGraphUser(value.entra === undefined ? { upn: input.userPrincipalName ?? value.loginName } : { id: value.entra.id }, request);
      assertCurrent(request);
      if (value.entra === undefined && principalLoginKey(directory.loginName) !== principalLoginKey(value.loginName)) throw new Error('Directory user does not match selected login.');
      value = { ...value, entra: directory.value.entra };
      enriched = { ...enriched, email: directory.email ?? enriched.email, userPrincipalName: directory.userPrincipalName };
    }
    if (policy.mode === 'sharePoint' || policy.mode === 'both') {
      const web = normalizePickerUrl(policy.webUrl ?? props.webUrl ?? request.context.webUrl);
      const collection = await principalCollectionUrl(request, web);
      assertCurrent(request);
      let registration: UserPrincipal;
      try { registration = await readPrincipalSharePointUser(value, request, web); }
      catch (error) {
        if (!ensure || !(error instanceof Error) || !('status' in error) || error.status !== 404) throw error;
        assertCurrent(request);
        // Native ensureUser is a server write. An already-sent request cannot be rolled back on cancellation.
        const row = restRecord(await spRest(request, `${web}/_api/web/ensureuser`, { logonName: value.loginName }));
        assertCurrent(request);
        registration = sharePointUserOption(row, collection);
        if (principalLoginKey(registration.loginName) !== principalLoginKey(value.loginName)) throw new Error('ensureUser returned another principal.');
        request.cache.invalidate(`principal-sp-get:${request.cache.identityKey(request.context.owner)}:${request.cache.identityKey(request.context.spClient)}:${request.context.tenantId}:${request.context.userId}:${web}/_api/web/siteusers`);
      }
      assertCurrent(request);
      value = { ...value, sharePoint: registration.value.sharePoint };
      enriched = { ...enriched, email: registration.email ?? enriched.email };
    }
    const confirmed: UserPrincipal = { ...enriched, key: input.key, value, loginName: value.loginName,
      entraId: value.entra?.id, entraTenantId: value.entra?.tenantId, sharePointId: value.sharePoint?.id, sharePointWebUrl: value.sharePoint?.siteCollectionUrl };
    result.push(validatePickerOption(confirmed));
  }
  assertCurrent(request);
  return result;
}

import type { PickerValue } from './picker-values.types';
import type { PickerOption } from './picker.types';
import { copyPickerValue, validatePickerOption } from './picker-values-validation.internal';
import { principalLoginKey } from './principal-login.internal';

export interface PickerIdentityContext { readonly tenantId?: string; }
const lower = (value: string): string => value.toLowerCase();
export const normalizePickerUrl = (url: string): string => new URL(url).href.replace(/\/$/, '').toLowerCase();
const tuple = (...values: readonly (string | number | undefined)[]): string => JSON.stringify(values);

/** Scoped canonical entity key; names, snapshots and custom item keys never determine identity. */
export function canonicalPickerIdentity(input: PickerValue, context: PickerIdentityContext = {}): string {
  const value = copyPickerValue(input);
  switch (value.kind) {
    case 'site': return tuple('site', normalizePickerUrl(value.url));
    case 'list': return tuple('list', normalizePickerUrl(value.webUrl), lower(value.listId));
    case 'field': return tuple('field', normalizePickerUrl(value.webUrl), value.listId === undefined ? undefined : lower(value.listId), lower(value.fieldId));
    case 'contentType': return tuple('contentType', value.scope, normalizePickerUrl(value.webUrl), value.listId === undefined ? undefined : lower(value.listId), lower(value.contentTypeId));
    case 'view': return tuple('view', normalizePickerUrl(value.webUrl), lower(value.listId), lower(value.viewId));
    case 'item': case 'file': case 'folder': return tuple('item', normalizePickerUrl(value.webUrl), lower(value.listId), value.itemId);
    case 'user':
      if (value.entra !== undefined) return tuple('entraUser', lower(value.entra.tenantId), lower(value.entra.id));
      if (value.sharePoint !== undefined) return tuple('spUser', normalizePickerUrl(value.sharePoint.siteCollectionUrl), value.sharePoint.id);
      if (context.tenantId === undefined) throw new Error('User login identity requires a verified tenant context.');
      return tuple('login', lower(context.tenantId), principalLoginKey(value.loginName));
    case 'sharePointGroup': return tuple('spGroup', normalizePickerUrl(value.identity.siteCollectionUrl), value.identity.id);
    case 'microsoft365Group': case 'securityGroup': case 'distributionList': return tuple('entraGroup', lower(value.identity.tenantId), lower(value.identity.id));
    case 'unvalidatedEmail': return tuple('unvalidatedEmail', context.tenantId === undefined ? undefined : lower(context.tenantId), lower(value.email));
  }
}

/** Verified aliases for one model, used only after source/scope validation. */
export function pickerIdentityAliases(value: PickerValue, context: PickerIdentityContext = {}): readonly string[] {
  const aliases = [canonicalPickerIdentity(value, context)];
  if (value.kind === 'user') {
    if (context.tenantId !== undefined && (value.entra === undefined || lower(value.entra.tenantId) === lower(context.tenantId))) aliases.push(tuple('login', lower(context.tenantId), principalLoginKey(value.loginName)));
    if (value.sharePoint !== undefined) aliases.push(tuple('spUser', normalizePickerUrl(value.sharePoint.siteCollectionUrl), value.sharePoint.id));
  }
  return Array.from(new Set(aliases));
}

/** Versioned opaque renderer alias containing a validated scoped reference, never an SDK object. */
export function createPickerAlias(value: PickerValue): string {
  return 'picker:v1:' + encodeURIComponent(JSON.stringify([value.kind, copyPickerValue(value)], (key, member: unknown) =>
    ['text', 'customKey', 'extension', 'serverRelativeUrl'].includes(key) || value.kind === 'site' && key === 'siteId' ? undefined : member));
}

/** Unknown consumer aliases remain unresolved; only our own validated versioned tuple is decoded. */
export function decodePickerAlias(key: string): PickerValue | undefined {
  if (!key.startsWith('picker:v1:')) return undefined;
  try {
    const input: unknown = JSON.parse(decodeURIComponent(key.slice('picker:v1:'.length)));
    if (!Array.isArray(input) || input.length !== 2) return undefined;
    const value = copyPickerValue(input[1]);
    return input[0] === value.kind ? value : undefined;
  } catch { return undefined; }
}

export interface PickerAliasState<T extends PickerOption> {
  readonly options: ReadonlyMap<string, T>;
  readonly errors: readonly Error[];
}

export function arePickerValuesCompatible(previous: PickerValue, next: PickerValue, context: PickerIdentityContext): boolean {
  if (previous.kind === 'user' && next.kind === 'user') {
    if (previous.entra !== undefined && next.entra !== undefined && tuple(lower(previous.entra.tenantId), lower(previous.entra.id)) !== tuple(lower(next.entra.tenantId), lower(next.entra.id))) return false;
    if (previous.sharePoint !== undefined && next.sharePoint !== undefined) {
      const sameCollection = normalizePickerUrl(previous.sharePoint.siteCollectionUrl) === normalizePickerUrl(next.sharePoint.siteCollectionUrl);
      if (sameCollection && previous.sharePoint.id !== next.sharePoint.id) return false;
      if (!sameCollection && (previous.entra === undefined || next.entra === undefined)) return false;
    }
  }
  const aliases = new Set(pickerIdentityAliases(previous, context));
  return pickerIdentityAliases(next, context).some(alias => aliases.has(alias));
}

/** Retains known models when candidate sets shrink; caller prunes unselected model history at its lifetime boundary. */
export function reconcilePickerAliases<T extends PickerOption>(previous: PickerAliasState<T> | undefined, incoming: readonly T[], context: PickerIdentityContext = {}): PickerAliasState<T> {
  const options = new Map(previous?.options);
  const errors: Error[] = [];
  for (const option of incoming) {
    try {
      validatePickerOption(option);
      const old = options.get(option.key);
      if (old !== undefined && !arePickerValuesCompatible(old.value, option.value, context)) throw new Error(`Picker alias ${option.key} cannot be rebound to another identity.`);
      if (old?.value.kind === 'user' && option.value.kind === 'user') {
        // A reduced model does not erase IDs already verified for a retained identity.
        const value = { ...option.value, entra: option.value.entra ?? old.value.entra, sharePoint: option.value.sharePoint ?? old.value.sharePoint };
        const enriched = { ...old, ...option, value };
        options.set(option.key, enriched);
      } else options.set(option.key, option);
    } catch (error) { errors.push(error instanceof Error ? error : new Error('Invalid picker option.')); }
  }
  return { options, errors };
}

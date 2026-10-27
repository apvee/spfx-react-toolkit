import type { PickerValue, SitePickerValue, ListPickerValue, FieldPickerValue, ViewPickerValue, ContentTypePickerValue, ListItemPickerValue, PrincipalPickerValue } from './picker-values.types';
import type { PickerOption } from './picker.types';

type Check = (input: unknown) => unknown;
interface Member { readonly check: Check; readonly optional?: boolean; }
type Schema = Readonly<Record<string, Check | Member>>;
const hasOwn = (object: object, key: PropertyKey): boolean => Object.prototype.hasOwnProperty.call(object, key);
const invalid = (): never => { throw new TypeError('Invalid picker reference: expected a plain scoped JSON value.'); };
const string: Check = input => typeof input === 'string' && input.trim().length > 0 ? input : invalid();
const label: Check = input => typeof input === 'string' ? input : invalid();
const integer: Check = input => typeof input === 'number' && Number.isSafeInteger(input) && input > 0 ? input : invalid();
const guid: Check = input => typeof input === 'string' && /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(input) ? input : invalid();
/** SharePoint lineage consists of nonzero bytes or 00 followed by a complete unpunctuated GUID.
 * The special System root is 0x; descendants require complete lineage segments.
 */
export function validateContentTypeId(input: unknown): string {
  if (typeof input === 'string' && /^0x$/i.test(input)) return input;
  if (typeof input !== 'string' || !/^0x(?:[0-9a-f]{2})+$/i.test(input) || input.length > 1026) return invalid();
  for (let offset = 2; offset < input.length;) {
    const length = input.slice(offset, offset + 2) === '00' ? 34 : 2;
    if (offset + length > input.length) return invalid();
    offset += length;
  }
  return input;
}
const url: Check = input => {
  if (typeof input !== 'string' || input.trim() !== input || !/^https?:\/\//i.test(input) || input.includes('\\')) return invalid();
  try {
    const parsed = new URL(input);
    if (!['https:', 'http:'].includes(parsed.protocol) || parsed.username || parsed.password || parsed.search || parsed.hash) return invalid();
    if (parsed.href.replace(/\/$/, '') !== input.replace(/\/$/, '')) return invalid();
    return input;
  } catch { return invalid(); }
};
const serverPath: Check = input => typeof input === 'string' && /^\/(?!\/)[^\\?#]*$/.test(input) && input.trim() === input && !input.split('/').some(segment => segment === '..' || segment === '.') ? input : invalid();
const optional = (check: Check): Member => ({ check, optional: true });

function record(input: unknown, schema: Schema): Record<string, unknown> {
  if (typeof input !== 'object' || input === null || Object.getPrototypeOf(input) !== Object.prototype) return invalid();
  const descriptors = Object.getOwnPropertyDescriptors(input);
  for (const key of Reflect.ownKeys(descriptors)) {
    if (typeof key !== 'string' || !hasOwn(schema, key)) return invalid();
    const descriptor = descriptors[key];
    if (!hasOwn(descriptor, 'value') || !descriptor.enumerable) return invalid();
  }
  const output: Record<string, unknown> = {};
  for (const key of Object.keys(schema)) {
    const member = schema[key];
    const check = typeof member === 'function' ? member : member.check;
    const descriptor = descriptors[key];
    if (descriptor === undefined || descriptor.value === undefined) {
      if (typeof member !== 'function' && member.optional) continue;
      return invalid();
    }
    output[key] = check(descriptor.value);
  }
  return output;
}
const entra: Check = input => record(input, { tenantId: guid, id: guid });
const sharePoint: Check = input => record(input, { siteCollectionUrl: url, id: integer });
const custom: Check = input => record(input, { field: string, value: value => typeof value === 'string' || typeof value === 'number' && Number.isFinite(value) ? value : invalid() });
const common = { kind: string, text: optional(label) };
const item = { webUrl: url, listId: guid, itemId: integer, customKey: optional(custom) };
const schemas: Readonly<Record<PickerValue['kind'], Schema>> = {
  site: { ...common, url, siteId: optional(guid) },
  list: { ...common, webUrl: url, listId: guid },
  field: { ...common, webUrl: url, listId: optional(guid), fieldId: guid, internalName: string },
  view: { ...common, webUrl: url, listId: guid, viewId: guid },
  contentType: { ...common, scope: value => value === 'site' || value === 'list' ? value : invalid(), webUrl: url, listId: optional(guid), contentTypeId: validateContentTypeId },
  item: { ...common, ...item },
  file: { ...common, ...item, uniqueId: guid, serverRelativeUrl: optional(serverPath), extension: optional(string) },
  folder: { ...common, ...item, uniqueId: guid, serverRelativeUrl: optional(serverPath) },
  user: { ...common, loginName: string, entra: optional(entra), sharePoint: optional(sharePoint) },
  sharePointGroup: { ...common, identity: sharePoint },
  microsoft365Group: { ...common, identity: entra },
  securityGroup: { ...common, identity: entra },
  distributionList: { ...common, identity: entra },
  unvalidatedEmail: { ...common, email: value => typeof value === 'string' && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value) ? value : invalid() },
};

/** Validate and copy a reference, stripping omitted optional members; rejects extra fields and accessors.
 * @param value A scoped JSON reference; throws TypeError for malformed input.
 * @returns A new reference with independently copied nested identities.
 * @example const site = copyPickerValue({ kind: 'site', url: 'https://tenant.sharepoint.com/sites/hr' });
 */
export function copyPickerValue(value: PickerValue): PickerValue {
  const descriptor = typeof value === 'object' && value !== null ? Object.getOwnPropertyDescriptor(value, 'kind') : undefined;
  const kind: unknown = descriptor?.value;
  if (typeof kind !== 'string' || !hasOwn(schemas, kind)) return invalid();
  // Every returned member has passed its domain-specific schema; external fields are never spread.
  const output = record(value, schemas[kind as PickerValue['kind']]);
  if (kind === 'contentType' && (output.scope === 'list' ? output.listId === undefined : output.listId !== undefined)) return invalid();
  return output as unknown as PickerValue;
}

function values(input: unknown, allowed: readonly PickerValue['kind'][]): readonly PickerValue[] {
  if (!Array.isArray(input) || Object.getPrototypeOf(input) !== Array.prototype) return invalid();
  const descriptors = Object.getOwnPropertyDescriptors(input);
  if (Reflect.ownKeys(descriptors).length !== input.length + 1) return invalid();
  const result: PickerValue[] = [];
  for (let index = 0; index < input.length; index++) {
    const descriptor = descriptors[String(index)];
    if (descriptor === undefined || !hasOwn(descriptor, 'value') || !descriptor.enumerable) return invalid();
    const value = copyPickerValue(descriptor.value);
    if (!allowed.includes(value.kind)) return invalid();
    result.push(value);
  }
  return result;
}

/** Validate external site references; throws TypeError for malformed or different-domain values.
 * @param input Unknown JSON array, including an empty selection.
 * @returns Fresh site references. Valid unresolved references remain valid.
 * @example const sites = validateSitePickerValues([{kind:'site',url:'https://tenant.sharepoint.com/sites/hr'}]);
 */
export function validateSitePickerValues(input: unknown): readonly SitePickerValue[] { return values(input, ['site']) as SitePickerValue[]; }
/** Validate external list references; throws TypeError for malformed input.
 * @param input Unknown JSON array. @returns Fresh scoped list references.
 * @example const lists = validateListPickerValues([]);
 */
export function validateListPickerValues(input: unknown): readonly ListPickerValue[] { return values(input, ['list']) as ListPickerValue[]; }
/** Validate external field references, including site columns without listId.
 * @param input Unknown JSON array; malformed input throws TypeError. @returns Fresh field references.
 * @example const fields = validateFieldPickerValues([]);
 */
export function validateFieldPickerValues(input: unknown): readonly FieldPickerValue[] { return values(input, ['field']) as FieldPickerValue[]; }
/** Validate external view references; each view requires web/list scope.
 * @param input Unknown JSON array; malformed input throws TypeError. @returns Fresh view references.
 * @example const views = validateViewPickerValues([]);
 */
export function validateViewPickerValues(input: unknown): readonly ViewPickerValue[] { return values(input, ['view']) as ViewPickerValue[]; }
/** Validate external collection/list content type references; throws TypeError for malformed IDs,
 * extra metadata or inconsistent scope/listId. Does not certify server existence or permissions.
 * @param input Unknown JSON array; empty arrays represent an empty selection in either mode.
 * @returns Fresh scoped references retaining ID casing and optional text snapshots.
 * @example const types = validateContentTypePickerValues([{kind:'contentType',scope:'site',webUrl:'https://tenant.sharepoint.com/sites/hr',contentTypeId:'0x0101'}]);
 */
export function validateContentTypePickerValues(input: unknown): readonly ContentTypePickerValue[] { return values(input, ['contentType']) as ContentTypePickerValue[]; }
/** Validate item/file/folder references with authoritative positive item IDs.
 * @param input Unknown JSON array; malformed input throws TypeError. @returns Fresh item references.
 * @example const items = validateListItemPickerValues([]);
 */
export function validateListItemPickerValues(input: unknown): readonly ListItemPickerValue[] { return values(input, ['item', 'file', 'folder']) as ListItemPickerValue[]; }
/** Validate principal references; ID requirements at confirmation are separate from JSON validity.
 * @param input Unknown JSON array; malformed input throws TypeError. @returns Fresh principal references.
 * @example const principals = validatePrincipalPickerValues([{kind:'user',loginName:'user@tenant.com'}]);
 */
export function validatePrincipalPickerValues(input: unknown): readonly PrincipalPickerValue[] { return values(input, ['user', 'sharePointGroup', 'microsoft365Group', 'securityGroup', 'distributionList', 'unvalidatedEmail']) as PrincipalPickerValue[]; }

/** Validate consumer models before adopting an alias; metadata cannot invent or contradict scoped identity. */
export function validatePickerOption<T extends PickerOption>(input: T): T {
  if (typeof input !== 'object' || input === null || Object.getPrototypeOf(input) !== Object.prototype) return invalid();
  const descriptors = Object.getOwnPropertyDescriptors(input);
  for (const descriptor of Object.values(descriptors)) if (!hasOwn(descriptor, 'value')) return invalid();
  string(input.key);
  label(input.text);
  const value = copyPickerValue(input.value);
  const metadata: Readonly<Record<string, unknown>> = input as unknown as Record<string, unknown>;
  const same = (key: string, expected: unknown): void => {
    if (metadata[key] !== expected) return invalid();
  };
  const optionalSame = (key: string, expected: unknown): void => {
    if (metadata[key] !== undefined) same(key, expected);
  };
  switch (value.kind) {
    case 'site': same('url', value.url); optionalSame('siteId', value.siteId); break;
    case 'list': same('webUrl', value.webUrl); same('id', value.listId); break;
    case 'field': same('webUrl', value.webUrl); optionalSame('listId', value.listId); same('id', value.fieldId); same('internalName', value.internalName); break;
    case 'contentType':
      same('kind', value.kind); same('webUrl', value.webUrl); same('id', value.contentTypeId);
      string(metadata.name); label(metadata.group);
      if (metadata.description !== undefined) label(metadata.description);
      for (const flag of ['hidden', 'readOnly', 'sealed']) if (typeof metadata[flag] !== 'boolean') return invalid();
      break;
    case 'view': same('webUrl', value.webUrl); same('listId', value.listId); same('id', value.viewId); break;
    case 'item': case 'file': case 'folder':
      same('kind', value.kind); same('webUrl', value.webUrl); same('listId', value.listId); same('itemId', value.itemId);
      if (metadata.keyField !== undefined && metadata.keyField !== (value.customKey?.field ?? 'Id')) return invalid();
      same('keyValue', value.customKey?.value ?? value.itemId); break;
    case 'user':
      same('kind', value.kind); same('loginName', value.loginName);
      optionalSame('entraId', value.entra?.id); optionalSame('entraTenantId', value.entra?.tenantId);
      optionalSame('sharePointId', value.sharePoint?.id); optionalSame('sharePointWebUrl', value.sharePoint?.siteCollectionUrl); break;
    case 'sharePointGroup': same('kind', value.kind); same('source', 'sharePoint'); same('id', value.identity.id); same('webUrl', value.identity.siteCollectionUrl); break;
    case 'microsoft365Group': case 'securityGroup': case 'distributionList':
      same('kind', value.kind); same('source', 'entra'); same('id', value.identity.id); optionalSame('tenantId', value.identity.tenantId); break;
    case 'unvalidatedEmail': same('kind', value.kind); same('email', value.email); break;
  }
  if (value.kind === 'file') optionalSame('fileExtension', value.extension);
  else if (['item', 'folder'].includes(value.kind) && metadata.fileExtension !== undefined) return invalid();
  return input;
}

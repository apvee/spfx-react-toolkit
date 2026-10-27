import type { FileFolderPickerValue } from './file-folder-picker.types';

type Check = (value: unknown) => unknown;
interface Member { readonly check: Check; readonly optional?: boolean; }
type Schema = Readonly<Record<string, Check | Member>>;
const own = (object: object, key: PropertyKey): boolean => Object.prototype.hasOwnProperty.call(object, key);
const invalid = (): never => { throw new TypeError('Invalid file/folder picker reference: expected a plain scoped JSON value.'); };
const hasControl = (input: string): boolean => Array.from(input).some(character => character.charCodeAt(0) < 32 || character.charCodeAt(0) === 127);
const text: Check = input => typeof input === 'string' ? input : invalid();
const opaque: Check = input => typeof input === 'string' && input.trim().length > 0 && !hasControl(input) ? input : invalid();
const guid: Check = input => typeof input === 'string' && /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(input) ? input : invalid();
const integer: Check = input => typeof input === 'number' && Number.isSafeInteger(input) && input > 0 ? input : invalid();
const optional = (check: Check): Member => ({ check, optional: true });

function httpsUrl(input: unknown, navigation = false): URL {
  if (typeof input !== 'string' || input.trim() !== input || !/^https:\/\//i.test(input) || (hasControl(input) || /[\\ ]/.test(input))) return invalid();
  let parsed: URL;
  try { parsed = new URL(input); } catch { return invalid(); }
  if (parsed.protocol !== 'https:' || parsed.username || parsed.password || parsed.port || (!navigation && (parsed.search || parsed.hash || /[?#]/.test(input)))) return invalid();
  // URL normalizes traversal away; inspect the original decoded segments before accepting it.
  const path = input.slice(input.indexOf('://') + 3).replace(/^[^/]+/, '').split(/[?#]/)[0];
  for (const segment of path.split('/')) {
    let decoded: string;
    try { decoded = decodeURIComponent(segment); } catch { return invalid(); }
    if (decoded === '.' || decoded === '..' || (hasControl(decoded) || /[\\/]/.test(decoded))) return invalid();
  }
  return parsed;
}

/** Normalize a syntactically valid SPO v2 endpoint. Contextual tenant/provenance checks remain required. */
export function normalizeFileFolderPickerEndpoint(input: unknown): string {
  const parsed = httpsUrl(input);
  // URL.port drops an explicit default :443; reject port syntax in the original authority too.
  const authority = typeof input === 'string' ? input.slice(input.indexOf('://') + 3).split('/')[0] : '';
  if (authority.includes(':')) return invalid();
  if (!/^[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.sharepoint\.com$/i.test(parsed.hostname)
    || !/^(?:\/[^/]+)*\/_api\/v2\.0\/?$/.test(parsed.pathname)) return invalid();
  return `${parsed.origin}${parsed.pathname.replace(/\/$/, '')}`;
}

function record(input: unknown, schema: Schema): Record<string, unknown> {
  // null is rejected here as a legacy JSON input; new absence is undefined.
  if (typeof input !== 'object' || input === null || Object.getPrototypeOf(input) !== Object.prototype) return invalid();
  const descriptors = Object.getOwnPropertyDescriptors(input);
  for (const key of Reflect.ownKeys(descriptors)) {
    if (typeof key !== 'string' || !own(schema, key)) return invalid();
    const descriptor = descriptors[key];
    if (!own(descriptor, 'value') || !descriptor.enumerable) return invalid();
  }
  const result: Record<string, unknown> = {};
  for (const key of Object.keys(schema)) {
    const member = schema[key];
    const descriptor = descriptors[key];
    if (descriptor === undefined || descriptor.value === undefined) {
      if (typeof member !== 'function' && member.optional) continue;
      return invalid();
    }
    result[key] = (typeof member === 'function' ? member : member.check)(descriptor.value);
  }
  return result;
}
const sharePoint: Check = input => record(input, {
  siteCollectionId: guid, webId: guid, listId: guid, listItemId: optional(integer), uniqueId: optional(guid),
});
const schema: Schema = {
  kind: input => input === 'file' || input === 'folder' ? input : invalid(),
  tenantId: guid,
  endpoint: normalizeFileFolderPickerEndpoint,
  driveId: opaque,
  itemId: opaque,
  name: optional(text),
  webUrl: optional(input => { httpsUrl(input, true); return input; }),
  sharePoint: optional(sharePoint),
};

/** Validate an untrusted JSON reference and copy all known nested members without invoking accessors. */
export function copyFileFolderPickerValue(input: unknown): FileFolderPickerValue {
  // Each member has passed the complete schema; unknown fields/prototypes were rejected before copying.
  return record(input, schema) as unknown as FileFolderPickerValue;
}

/** Validate the complete ordered array atomically and return independent references; never coerce legacy values. */
export function validateFileFolderPickerValues(input: unknown): readonly FileFolderPickerValue[] {
  if (!Array.isArray(input) || Object.getPrototypeOf(input) !== Array.prototype) return invalid();
  const descriptors = Object.getOwnPropertyDescriptors(input);
  if (Reflect.ownKeys(descriptors).length !== input.length + 1) return invalid();
  const result: FileFolderPickerValue[] = [];
  for (let index = 0; index < input.length; index++) {
    const descriptor = descriptors[String(index)];
    if (!descriptor || !own(descriptor, 'value') || !descriptor.enumerable) return invalid();
    result.push(copyFileFolderPickerValue(descriptor.value));
  }
  return result;
}

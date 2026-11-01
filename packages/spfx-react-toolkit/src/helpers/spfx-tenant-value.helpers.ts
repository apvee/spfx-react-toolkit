/**
 * Historical tenant-value serialization and OData string-literal escaping.
 *
 * @packageDocumentation
 * @docPage helpers/tenant-value
 */
/**
 * Doubles single quotes for the content of an OData string literal. This does not add surrounding quotes, escape a full query or encode a URL.
 *
 * @param value - Raw string-literal content.
 * @returns Content with each single quote doubled.
 *
 * @example
 * ```ts
 * import { escapeODataValue } from '@apvee/spfx-react-toolkit';
 *
 * const filter = `key eq '${escapeODataValue("owner's-settings")}'`;
 * ```
 * @useWhen You embed user text inside an OData single-quoted string literal.
 */
export function escapeODataValue(value: string): string {
  return value.replace(/'/g, "''");
}

/**
 * Converts a tenant-store value into a string: primitives become String(value), Date uses ISO text, and other values use JSON.stringify with a String fallback when JSON returns undefined. Legacy null input becomes the literal null string. Circular objects, nested BigInts and invalid Dates can throw.
 *
 * @param value - Value to serialize without a schema.
 * @returns Stored text; serialization is not lossless for every input type.
 * @throws Native serialization errors for unsupported JSON values or invalid Dates.
 *
 * @example Basic usage
 * ```ts
 * import { serializeTenantValue } from '@apvee/spfx-react-toolkit';
 *
 * const stored = serializeTenantValue({ enabled: true });
 * ```
 * @useWhen You must preserve the historical standalone tenant-value string format.
 * @avoidWhen You write new list-store values and need exact empty/JSON-like strings or bigint text; use the tenant/site store service encoding instead.
 *
 *
 * @example Compare legacy serialization with validated unknown data
 * ```ts
 * import { serializeTenantValue, deserializeTenantValue, escapeODataValue } from '@apvee/spfx-react-toolkit';
 * const filter = `Title eq '${escapeODataValue("owner's-settings")}'`;
 * // Title eq 'owner''s-settings'
 * const raw = serializeTenantValue({ enabled: true });
 * const value = deserializeTenantValue<unknown>(raw);
 * if (typeof value === 'object' && value !== undefined && value !== null) console.log(value);
 * const legacy = deserializeTenantValue<unknown>('false'); // boolean false
 * ```
 */
export function serializeTenantValue(value: unknown): string {
  if (value === null) {
    return String(value);
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  const type = typeof value;
  if (type === 'string' || type === 'number' || type === 'boolean' || type === 'bigint') {
    return String(value);
  }

  const serializedValue = JSON.stringify(value);

  return serializedValue ?? String(value);
}

/**
 * Parses stored text as JSON, falling back to the unchanged string when parsing fails. JSON-like strings become primitives, bigint decimal strings can lose precision as numbers, and Dates remain text. T is only a compile-time assertion: no schema validation, Date reconstruction or runtime type checking occurs.
 *
 * @typeParam T - Expected result type; callers must validate untrusted data.
 * @param rawValue - Stored string to parse.
 * @returns Parsed JSON value, or original string when parsing fails.
 *
 * @example
 * ```tsx
 * import { deserializeTenantValue } from '@apvee/spfx-react-toolkit';
 *
 * const value = deserializeTenantValue<unknown>('false'); // boolean false
 * ```
 * @useWhen You read historical tenant text and can validate its JSON-first interpretation in the caller.
 * @avoidWhen You need a schema-validated or lossless type round trip; validate the unknown result in the caller.
 */
export function deserializeTenantValue<T>(rawValue: string): T {
  try {
    return JSON.parse(rawValue) as T;
  } catch {
    return rawValue as unknown as T;
  }
}

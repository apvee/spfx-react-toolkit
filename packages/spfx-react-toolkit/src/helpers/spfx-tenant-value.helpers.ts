/**
 * Doubles single quotes for the content of an OData string literal. This does not add surrounding quotes, escape a full query or encode a URL.
 *
 * @param value - Raw string-literal content.
 * @returns Content with each single quote doubled.
 *
 * @example
 * ```ts
 * const filter = `key eq '${escapeODataValue("owner's-settings")}'`;
 * ```
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
 * @example
 * ```ts
 * const stored = serializeTenantValue({ enabled: true });
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
 * Parses stored text as JSON, falling back to the unchanged string when parsing fails. T is only a compile-time assertion: no schema validation, Date reconstruction or runtime type checking occurs.
 *
 * @typeParam T - Expected result type; callers must validate untrusted data.
 * @param rawValue - Stored string to parse.
 * @returns Parsed JSON value, or original string when parsing fails.
 *
 * @example
 * ```tsx
 * const value = deserializeTenantValue<unknown>('false'); // boolean false
 * ```
 */
export function deserializeTenantValue<T>(rawValue: string): T {
  try {
    return JSON.parse(rawValue) as T;
  } catch {
    return rawValue as unknown as T;
  }
}

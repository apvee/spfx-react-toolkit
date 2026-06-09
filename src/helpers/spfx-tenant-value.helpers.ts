/**
 * Escapes a string value for use in OData filter expressions.
 *
 * @param value - Raw string value
 * @returns Escaped string safe for OData filters
 */
export function escapeODataValue(value: string): string {
  return value.replace(/'/g, "''");
}

/**
 * Serializes a tenant-scoped value for SharePoint string storage.
 *
 * @param value - Value to serialize
 * @returns Serialized string representation
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
 * Deserializes a tenant-scoped SharePoint string value.
 *
 * @param rawValue - Stored string value
 * @returns Parsed value, or the raw string when JSON parsing fails
 */
export function deserializeTenantValue<T>(rawValue: string): T {
  try {
    return JSON.parse(rawValue) as T;
  } catch {
    return rawValue as unknown as T;
  }
}

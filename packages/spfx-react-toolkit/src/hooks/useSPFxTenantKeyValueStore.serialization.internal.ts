export function escapeODataValue(value: string): string {
    return value.replace(/'/g, "''");
}

export function serializeValue(value: unknown): string {
    if (value === null) return String(value);
    if (value instanceof Date) return value.toISOString();

    const type = typeof value;
    if (type === 'string' || type === 'number' || type === 'boolean' || type === 'bigint') {
        return String(value);
    }

    return JSON.stringify(value);
}

export function deserializeValue<T>(rawValue: string): T {
    try {
        return JSON.parse(rawValue) as T;
    } catch {
        return rawValue as unknown as T;
    }
}

# Tenant Value Helpers

[Helpers overview](../INDEX.md#helpers-and-styles) · [Non-React composition](../../guides/helpers-and-services.md)


Import:

```ts
import {
  deserializeTenantValue,
  escapeODataValue,
  serializeTenantValue,
} from '@apvee/spfx-react-toolkit';
```

## `escapeODataValue`

```ts
function escapeODataValue(value: string): string;
```

Escapes single quotes for OData string filters by replacing `'` with `''`.

## `serializeTenantValue`

```ts
function serializeTenantValue(value: unknown): string;
```

Serializes tenant-scoped values for SharePoint string storage:

| Input | Output |
|-------|--------|
| `null` | `'null'` |
| `Date` | ISO string |
| `string`, `number`, `boolean`, `bigint` | `String(value)` |
| object/array | `JSON.stringify(value)` when available |
| non-JSON fallback | `String(value)` |

## `deserializeTenantValue`

```ts
function deserializeTenantValue<T>(rawValue: string): T;
```

Parses JSON when possible. If parsing fails, returns the raw string cast to `T`. This legacy format is not a complete type round trip: numeric/boolean/null-like strings become JSON primitives, bigints may lose precision when parsed as numbers, and dates remain strings. `T` does not perform runtime validation; validate deserialized values in the caller. No new envelope or stored-data migration is introduced.

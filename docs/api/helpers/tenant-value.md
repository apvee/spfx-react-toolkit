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

> These public helpers preserve the historical format. New tenant and site list-store writes additionally quote empty/JSON-like strings and bigint decimal text; see [list-store saved-data compatibility](../hooks/storage/tenant-key-value-store.md#serialization-and-saved-data-compatibility). Historical stored values and StorageEntity reads retain their existing interpretation, including each list store's documented legacy REST-null mapping.

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

## Serialization errors and usage

The helpers require no provider/client and perform no persistence. `serializeTenantValue` does not suppress native serialization errors: circular structures, nested bigint and invalid Date values can throw. When JSON serialization yields undefined (for example top-level undefined, functions or symbols), the helper falls back to `String(value)`; this does not make such values losslessly serializable. `escapeODataValue` escapes string-literal content only; it adds no quotes and performs no URL encoding.

```ts
const filter = `Title eq '${escapeODataValue("owner's-settings")}'`;
// Title eq 'owner''s-settings'
const raw = serializeTenantValue({ enabled: true });
const value = deserializeTenantValue<unknown>(raw);
if (typeof value === 'object' && value !== undefined) console.log(value);
const legacy = deserializeTenantValue<unknown>('false'); // boolean false
```

Use `unknown` and validate application-specific shapes before relying on stored data. Use the [tenant](../services/tenant-key-value-store.md) or [site](../services/site-key-value-store.md) service for new list-store writes; their safe text encoding intentionally differs from these historical standalone helpers.

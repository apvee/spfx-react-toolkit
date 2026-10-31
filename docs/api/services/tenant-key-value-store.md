# Tenant Key-Value Store Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

Import this factory and its public types from `@apvee/spfx-react-toolkit/services` or the package root. Callers supply the initialized client/context; these services require no React provider.

## `createSPFxTenantKeyValueStoreService`

```ts
function createSPFxTenantKeyValueStoreService(
  spHttpClient: SPHttpClient
): SPFxTenantKeyValueStoreService;
```

Required public types:

- `SPFxTenantKeyValueStoreServiceItem`
- `SPFxTenantKeyValueStoreService`

Returned methods:

| Method | Description |
|--------|-------------|
| `ensureListReady(catalogUrl)` | Explicitly provisions/repairs the hidden list, validates field types and indexed unique Title keys |
| `get<T>(key, catalogUrl)` | Reads one unique key without provisioning; confirmed list/key absence returns undefined |
| `list(catalogUrl)` | Reads all pages in Title order, validating continuation links and rejecting partial failures |
| `save<T>(key, value, catalogUrl, description?)` | Ensures strict setup, then creates or conditionally updates using the observed ETag |
| `remove(key, catalogUrl)` | Deletes with an observed ETag without provisioning; confirmed absence is a no-op |

`SPFxTenantKeyValueStoreServiceItem<T>` contains `key`, `value`, `description`, and `id`.

Provisioning is shared per normalized catalog URL within each service instance. Schema/uniqueness failures reject and are retryable; data is never removed to repair existing duplicates. Separate instances recover list/field/item creation races through bounded validated rereads. Reads and deletion never repair schema; only confirmed list/key absence is empty. Failed probes, incompatible schema, duplicate keys and malformed/failed pages reject.

Tenant writes preserve JSON-like strings using JSON quoting; bigint reads as exact decimal text, and Date remains ISO text. Existing raw values retain JSON-first interpretation and need no migration. Both list stores use this encoding for new writes; public legacy helpers remain unchanged and previously stored values retain their interpretation. See [tenant serialization and compatibility](../hooks/storage/tenant-key-value-store.md#serialization-and-saved-data-compatibility) for the complete format and limitations. T does not validate runtime data.

Updates/deletion require an observed ETag (HTTP 412 conflicts reject without blind retry). This protects the service's snapshot-to-mutation interval, not a consumer's earlier get-to-save interval. Omitting description preserves it; an empty string clears it. Actual catalog/list permissions authorize requests. Construction performs no HTTP and the non-React service does not discover the catalog:

```ts
const store = createSPFxTenantKeyValueStoreService(context.spHttpClient);
await store.save('my-app:version', '123', catalogUrl);
const version = await store.get<unknown>('my-app:version', catalogUrl);
if (typeof version?.value === 'string') console.log(version.value);
```

## Item shape and compatibility

`SPFxTenantKeyValueStoreServiceItem<T = unknown>` has readonly `key: string`, `value: T`, `description: string | undefined` (empty metadata becomes undefined), and source-list `id: number`. `SPFxTenantKeyValueStoreService` exposes the five Promise-based operations above; get defaults to `T = unknown`, list returns an array of unknown-valued items, and setup/save/remove resolve void. There is no tenant-service permission-indicator method; catalog administrator status is not a substitute for the list's effective permissions.

Use the actual tenant catalog URL for every operation. Empty URLs reject before HTTP. The hidden generic `TenantKeyValueStore` uses `BaseTemplate: 100`, `Hidden: true`, `NoCrawl: true`, Text Title keys and multiline Note Value/Description fields. Reads validate types without enforcing/repairing uniqueness; get rejects duplicate matches, and list preserves server order and historical duplicate rows. Setup/write requires indexed unique Title. Key equality follows SharePoint field/unique-column comparison, including case-insensitive matching, while returned Title casing is preserved. Failures are never transformed into confirmed absence.

Only a list probe's 404 means an absent store; field/item/continuation failures, including their 404s, reject. Pagination accepts validated same-origin/same-list item links and rejects malformed/repeated/foreign continuations without partial results. New writes preserve empty and JSON-like strings using JSON quoting; simple text retains its raw text format. Bigint becomes exact decimal string, Date stays ISO text, and objects use JSON. Circular objects, nested bigint or invalid Dates reject before setup. Existing raw values keep JSON-first interpretation: prior numeric/boolean-like text may still become primitives, and lost historical precision cannot be recovered. Tenant's legacy REST-null Value and the stored text `'null'` both decode to legacy `null`; this differs from the site store's empty-string handling of a REST-null Note.

Omitting a description preserves the destination's existing metadata, while empty text clears it. Updates/removal obtain observed ETags, using a fresh item/header fallback if needed; missing/unsafe ETags or HTTP 412 reject without wildcard overwrite or blind retry. Creation collision recovery can attempt one conditional update only after confirming the competing unique key; its actual update failure is propagated. Removal accepts a vanished snapshot/deletion only after confirming key absence, never deleting a recreated key. These safeguards protect the service snapshot-to-mutation interval, not a consumer's earlier get-to-save interval.

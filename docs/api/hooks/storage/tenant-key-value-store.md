# useSPFxTenantKeyValueStore

[Storage overview](../../INDEX.md#storage-hooks) · [Usage guide](../../../guides/storage.md)

```ts
import { useSPFxTenantKeyValueStore } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).

Reads and writes shared tenant configuration in the hidden `TenantKeyValueStore` list in the tenant app catalog. Requires a matching host-specific SPFx provider, its authenticated `SPHttpClient`, page context, and a provisioned tenant app catalog. All applications share the same key namespace; use application-specific key prefixes.

## Signature and result

```typescript
function useSPFxTenantKeyValueStore(): SPFxTenantKeyValueStoreResult

interface SPFxTenantKeyValueStoreResult {
  readonly isLoading: boolean;
  readonly error: Error | undefined;
  readonly isWriting: boolean;
  readonly writeError: Error | undefined;
  readonly canWrite: boolean;
  readonly isReady: boolean;
  readonly get: <T = unknown>(key: string) => Promise<SPFxTenantKeyValueStoreItem<T> | undefined>;
  readonly list: () => Promise<SPFxTenantKeyValueStoreItem<unknown>[]>;
  readonly save: <T = unknown>(key: string, value: T, description?: string) => Promise<void>;
  readonly remove: (key: string) => Promise<void>;
}

interface SPFxTenantKeyValueStoreItem<T = unknown> {
  readonly key: string;
  readonly value: T;
  readonly description: string | undefined;
  readonly id: number;
}
```

## Initialization, permissions and errors

Mount discovers the catalog and checks permissions using read-only requests. `get`, `list` and `remove` never create the list or repair its schema. Only `save` (or the non-React service's explicit `ensureListReady`) provisions/repairs it.

`isReady` retains its client-availability meaning: it does **not** mean discovery has finished, the list exists or the user can write. `canWrite` starts false and checks effective Add/Edit/Delete Items permissions on the list. If the list is confirmed absent, it checks those permissions plus Manage Lists on the catalog web. It refreshes after operations; failed checks return false and older checks cannot overwrite newer results. Schema repair can require additional permissions even when item-write permissions are present. SharePoint authorizes every request; this flag is only an indicator, not a security boundary or a guarantee of success.

Only a confirmed missing list/key is considered absent. Access denials, server errors, incompatible fields, duplicate lookup results, malformed data and paging failures publish `error`; `get` returns `undefined` and `list` returns `[]` for compatibility. Inspect `error` to distinguish failure from absence. Missing client availability rejects the operation. Writes/removal publish `writeError` and reject; callers must handle rejections. Background discovery failures are retried and surfaced by explicit operations.

Read/write loading flags independently count pending operations. Starting an operation clears its channel error, and only the latest-started operation may publish an error. Obsolete contexts and unmounted completions cannot update hook state. Remote requests already dispatched are not cancelled.

## Storage, paging and concurrency

The list keeps its existing schema: `Title` (Text/key), `Value` (Note/serialized text), `Description` (Note/optional metadata). Setup verifies field types and requires indexed, unique Title keys on both existing and newly created lists. Failed uniqueness/schema checks reject and are retryable; existing data is never deleted to repair schema. Resolve pre-existing duplicate keys administratively before retrying a write. Concurrent setup is shared within each service instance per normalized catalog URL; separate instances recover creation races only after bounded schema rereads.

`list()` follows all supported SharePoint continuation pages in server Title order. Links must stay on the same origin and list-items path, without credentials/fragments or repeated URLs. Failed later pages reject in the service and produce the hook's empty fallback plus `error`; partial data is never presented as a complete list. SharePoint list thresholds and the user's actual access still apply.

`save()` updates existing keys and preserves an omitted description; `description: ''` clears it. Updates and deletion use the observed item's ETag, fetching a single-item snapshot if collection metadata omits it. A missing ETag fails safely; HTTP 412 conflicts reject without wildcard overwrite or blind retry. This detects changes between the service's snapshot and mutation, **not** edits since a consumer's earlier `get()` call. A concurrent creation can recover through one validated conditional update. Removal of a concurrently deleted key is accepted only after confirming absence.

## Serialization and saved-data compatibility

No migration, new column or tagged envelope is required. Existing raw values keep their historical JSON-first interpretation, including REST Note `null` reading as legacy `null`. Previously ambiguous data cannot be reconstructed automatically: a historical raw `123` still reads as number `123`, even if the original caller supplied a string. Existing public legacy serialization helpers and StorageEntity reads remain unchanged. Both list stores use the safe encoding below for new writes, while preserving their respective historical read behavior.

New tenant and site `save()` values use:

| Input | Storage/readback |
|-------|------------------|
| Empty text, or text that parses as JSON, such as `"123"`, `"true"`, `"null"` or JSON-looking text | JSON-quoted text; reads back as the original string |
| Other text | Raw text; reads back unchanged |
| `number`, `boolean`, legacy `null` input | Legacy string representation, decoded with `JSON.parse` |
| `bigint` | JSON-quoted decimal text; reads as an exact **string**, without numeric precision loss |
| `Date` | ISO text; reads as a **string**, not a Date instance |
| Objects/arrays | `JSON.stringify`, then `JSON.parse` on read |

Older JSON-first readers can read new strings and bigint text correctly. Generics describe the expected shape and do not validate runtime data. This remains a text/JSON store, not a lossless arbitrary-JavaScript serializer: nested bigint, circular objects and invalid Dates throw before provisioning. Non-JSON values such as undefined/functions, non-finite numbers and custom prototypes are not guaranteed to round-trip. Validate application configuration after reading.

## Example: save and read text safely

```tsx
import * as React from 'react';
import { useSPFxTenantKeyValueStore } from '@apvee/spfx-react-toolkit';

function TenantConfigEditor() {
  const { save, get, isReady, isWriting, error, writeError } = useSPFxTenantKeyValueStore();
  const [text, setText] = React.useState<string>('');
  const run = async (): Promise<void> => {
    try {
      await save('my-app:version', '123', 'Application version');
      const item = await get<unknown>('my-app:version');
      if (typeof item?.value === 'string') setText(item.value);
    } catch {
      // Write failures are rendered through writeError; reads use error + fallback.
    }
  };
  return <>
    <button disabled={!isReady || isWriting} onClick={run}>Save and read</button>
    <p>{text}</p>
    {(error ?? writeError) && <p role="alert">{(error ?? writeError)?.message}</p>}
  </>;
}
```

For a bigint, `await save('my-app:identifier', BigInt('9007199254740993'))` is read using `get<string>()`. Reconstruct a BigInt or Date explicitly only after validating the returned text. Existing legacy numeric values may already have lost precision and are not silently rewritten.

## Demo and verification

The **Tenant → Tenant storage** sample runs CRUD only on explicit Run. Select Text/JSON/Big integer/Date to exercise serialization and inspect the read value's type. Empty description clears metadata. The list result reports the total fetched count and previews the first ten entries. Use disposable keys and remove them afterward. Local automated tests execute the real hook and panel with simulated service boundaries; authenticated tenant checks are tracked separately in [SharePoint validation](../../../development/sharepoint-validation/INDEX.md).

## Source

[View source](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxTenantKeyValueStore.ts)

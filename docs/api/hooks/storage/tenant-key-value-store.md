# useSPFxTenantKeyValueStore

[Storage overview](../../INDEX.md#storage-hooks) · [Usage guide](../../../guides/storage.md)

```ts
import { useSPFxTenantKeyValueStore } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).


Tenant-level key-value store backed by a hidden SharePoint list in the tenant app catalog. Provides CRUD operations (get, list, save, remove) using the existing string/JSON serialization format.

This hook is an alternative to tenant properties (StorageEntity) for scenarios requiring read/write access via REST, since Microsoft has blocked the SetStorageEntity and RemoveStorageEntity REST endpoints.

## Signature

```typescript
function useSPFxTenantKeyValueStore(): SPFxTenantKeyValueStoreResult
```

## Returns

```typescript
interface SPFxTenantKeyValueStoreResult {
  /** Loading state for read operations */
  readonly isLoading: boolean;
  /** Last error from read operations */
  readonly error: Error | undefined;
  /** Loading state for write operations */
  readonly isWriting: boolean;
  /** Last error from write operations */
  readonly writeError: Error | undefined;
  /** Whether the current user can write (Site Collection Admin on app catalog) */
  readonly canWrite: boolean;
  /** True when SPHttpClient is available */
  readonly isReady: boolean;

  /** Get a single item by key */
  readonly get: <T = unknown>(key: string) => Promise<SPFxTenantKeyValueStoreItem<T> | undefined>;
  /** List all items (sorted by key) */
  readonly list: () => Promise<SPFxTenantKeyValueStoreItem<unknown>[]>;
  /** Create or update a key-value pair (auto-provisions list on first write) */
  readonly save: <T = unknown>(key: string, value: T, description?: string) => Promise<void>;
  /** Remove an item by key (no-op if not found) */
  readonly remove: (key: string) => Promise<void>;
}

interface SPFxTenantKeyValueStoreItem<T = unknown> {
  /** Property key */
  readonly key: string;
  /** Deserialized property value */
  readonly value: T;
  /** Optional description metadata */
  readonly description: string | undefined;
  /** SharePoint list item ID */
  readonly id: number;
}
```

## Error and loading behavior

Read operations remain pending until their service promises settle. `get()` returns `undefined` and `list()` returns `[]` on read failures while publishing `error`; these fallbacks alone cannot distinguish absence from failure. Writes publish `writeError` and reject their returned promises on failure. Handle write rejections in the caller.

Read and write channels have independent pending counts: their loading flags remain true while any current-identity operation in that channel is pending. Starting an operation clears its channel error; only the latest-started operation may publish that error. Old catalog/service and unmounted completions do not update the current hook. The server still controls authorization and remote ordering.

## Storage Details

The store uses a hidden list named `TenantKeyValueStore` in the tenant app catalog:
- **Title** column: key (indexed, unique)
- **Value** column: multiline text (stores serialized data)
- **Description** column: multiline text (optional metadata)

The list is auto-provisioned on the first `save()` call. Read operations (`get`, `list`) return `undefined`/`[]` gracefully if the list doesn't exist yet.

## Serialization

| Input type | Stored as |
|------------|-----------|
| `string` | Raw string |
| `number`, `boolean`, `bigint` | `String(value)` |
| `null` | `"null"` |
| `Date` | ISO 8601 string |
| Objects/arrays | `JSON.stringify()` |

Deserialization attempts `JSON.parse()` first; falls back to raw string. This legacy format does not preserve every input type: strings such as `"123"`, `"true"` or `"null"` become JSON primitives; a bigint may become a lossy number; dates remain ISO strings. The generic `T` is a compile-time annotation, not runtime validation. Existing stored values retain this format; validate read values in application code.

## Requirements

- Tenant app catalog must be provisioned
- **Read**: Requires the current user to have access to the catalog/list; actual tenant permissions apply
- **Write/Remove**: Site Collection Administrator role on the tenant app catalog site

## Example: Basic CRUD

```tsx
import { useSPFxTenantKeyValueStore } from '@apvee/spfx-react-toolkit';

function TenantConfigEditor() {
  const store = useSPFxTenantKeyValueStore();
  const [value, setValue] = React.useState('');

  const handleLoad = async () => {
    const item = await store.get<string>('apiEndpoint');
    if (item) setValue(item.value);
  };

  const handleSave = async () => {
    await store.save<string>('apiEndpoint', value, 'Production API endpoint');
  };

  const handleDelete = async () => {
    await store.remove('apiEndpoint');
    setValue('');
  };

  return (
    <Stack tokens={{ childrenGap: 10 }}>
      <TextField value={value} onChange={(_, v) => setValue(v ?? '')} />
      <Stack horizontal tokens={{ childrenGap: 5 }}>
        <PrimaryButton onClick={handleLoad}>Load</PrimaryButton>
        <PrimaryButton onClick={handleSave} disabled={store.isWriting}>Save</PrimaryButton>
        <DefaultButton onClick={handleDelete}>Delete</DefaultButton>
      </Stack>
      {store.error && <MessageBar messageBarType={MessageBarType.error}>{store.error.message}</MessageBar>}
    </Stack>
  );
}
```

## Example: Heterogeneous Types

```tsx
const store = useSPFxTenantKeyValueStore();

// String
await store.save<string>('appVersion', '2.1.0');

// Number
await store.save<number>('maxUploadSize', 10485760);

// Complex object
interface FeatureFlags { enableChat: boolean; maxUsers: number; }
await store.save<FeatureFlags>('featureFlags', { enableChat: true, maxUsers: 500 });

// Read back with type safety
const flags = await store.get<FeatureFlags>('featureFlags');
if (flags?.value.enableChat) { /* ... */ }
```

## Example: Property Dashboard

```tsx
function AllPropertiesView() {
  const store = useSPFxTenantKeyValueStore();
  const [items, setItems] = React.useState<SPFxTenantKeyValueStoreItem<unknown>[]>([]);

  React.useEffect(() => {
    store.list().then(setItems);
  }, []);

  if (store.isLoading) return <Spinner />;

  return (
    <DetailsList
      items={items.map(i => ({
        key: i.key,
        value: typeof i.value === 'object' ? JSON.stringify(i.value) : String(i.value),
        description: i.description ?? '',
      }))}
      columns={[
        { key: 'key', name: 'Key', fieldName: 'key', minWidth: 150 },
        { key: 'value', name: 'Value', fieldName: 'value', minWidth: 200 },
        { key: 'desc', name: 'Description', fieldName: 'description', minWidth: 200 },
      ]}
    />
  );
}
```

## Source

[View source](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxTenantKeyValueStore.ts)

---

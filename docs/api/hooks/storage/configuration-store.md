# useSPFxConfigurationStore

[Storage overview](../../INDEX.md#storage-hooks) · [Usage guide](../../../guides/storage.md)

```ts
import { useSPFxConfigurationStore } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).

Resolves application configuration using the existing tenant and site collection lists. Requires a matching host-specific SPFx provider and its authenticated client/page context. It is additive: the existing tenant/site hooks, public signatures, lists, values and legacy decoding remain unchanged.

```ts
useSPFxConfigurationStore(options: SPFxConfigurationStoreOptions): SPFxConfigurationStoreResult;
// options: { mode: 'tenant' | 'site' | 'layered' }; mode is required.
```

| Mode | Effective `get(key)` | Allowed physical scopes |
| --- | --- | --- |
| `tenant` | Tenant app catalog only | `tenant` |
| `site` | Current site collection root only | `site` |
| `layered` | Site first; tenant only after confirmed site absence | `tenant`, `site` |

`site` always means `pageContext.site.absoluteUrl`, shared by any subsites. Tenant catalog discovery uses a snapshot of `pageContext.web.absoluteUrl`. Construction, mount and equivalent renders perform no reads, permission checks or provisioning. Tenant discovery is lazy and shared while pending within the service instance; failures can retry through an explicit operation. A layered site hit and site-targeted operations make no catalog requests.

## Operations and state

```ts
get<T = unknown>(key: string): Promise<SPFxConfigurationStoreItem<T> | undefined>;
listScope(scope: 'tenant' | 'site'): Promise<SPFxConfigurationStoreItem<unknown>[]>;
save<T = unknown>(key: string, value: T, options: {
  scope: 'tenant' | 'site'; description?: string;
}): Promise<void>;
remove(key: string, target: { scope: 'tenant' | 'site' }): Promise<void>;
```

Every mutation requires a physical scope, even in a single-store mode. Invalid/disabled scopes reject before I/O; no scope is inferred from permissions, availability or the source of an earlier read. There is no default write target or cross-scope failover.

Items contain `key`, `value`, `description`, `id` and `source: 'tenant' | 'site'`. Metadata and server key casing come from the same winning record. IDs are source-list-local; `source` is provenance, not a routing instruction. `get` replaces the entire value by key, including objects/arrays; there is no JSON-property merge. A present site record wins even for `false`, `0`, `''` or legacy stored `null`. Only the absence of the **record** enables inheritance.

`listScope` is a complete **physical inventory** of one store, preserving server order/casing and any legacy duplicate rows. It does not merge or deduplicate keys. A page failure rejects without partial output. There is no effective `list()`: SharePoint determines key equality, which a JavaScript Map/lowercase merge cannot faithfully reproduce for arbitrary saved keys. The existing unique-lookup error remains observable through `get`.

State consists of `isReady`, `isLoading`, `error`, `isWriting`, `writeError`. `isReady` means the client and mode-specific local URLs are available; it does not establish successful discovery, schema or access. Read/write loading flags count all current-identity pending operations in their channel. Errors are cleared on starting another operation in the same channel; only the latest-started operation can publish its channel error. A mode/client/relevant URL change immediately hides obsolete state, and old/unmounted completions cannot publish current state. Retained callbacks keep their original targets.

**All failed operations reject their promises.** Unlike the historical store hooks, failed reads also reject while publishing `error`; they never return absence-like fallbacks. Handle both read and write rejections. Non-Error rejections are normalized to Error for state while their original rejection is preserved. There is no cached `data`, `canWrite`, automatic load or refetch. SharePoint authorizes every request.

## Example: tenant defaults and a site override

```tsx
import * as React from 'react';
import { useSPFxConfigurationStore } from '@apvee/spfx-react-toolkit/hooks';
import type { SPFxConfigurationStoreItem } from '@apvee/spfx-react-toolkit/hooks';

function ConfigurationExample() {
  const store = useSPFxConfigurationStore({ mode: 'layered' });
  const [item, setItem] = React.useState<SPFxConfigurationStoreItem<number>>();
  const run = async (): Promise<void> => {
    try {
      await store.save('my-app:pageSize', 20, { scope: 'tenant' });
      await store.save('my-app:pageSize', 50, { scope: 'site' });
      setItem(await store.get<number>('my-app:pageSize')); // 50, source site
      await store.remove('my-app:pageSize', { scope: 'site' });
      setItem(await store.get<number>('my-app:pageSize')); // 20, source tenant
    } catch {
      setItem(undefined); // Render the current hook error below.
    }
  };
  return <>
    <button disabled={!store.isReady || store.isLoading || store.isWriting} onClick={run}>Run configuration example</button>
    <p>{item ? `${item.source}: ${item.value}` : 'No result'}</p>
    {(store.error ?? store.writeError) && <p role="alert">{(store.error ?? store.writeError)?.message}</p>}
  </>;
}
```

This example intentionally writes configuration. Use your own key namespace and grants. Removing a site override reveals tenant data on the next read; removing tenant leaves an existing site override intact. Neither operation deletes both stores or writes a tombstone. A successful mutation remains successful if a subsequent explicit read fails.

Only `save` can initialize/repair its selected list. Reads, inventories and removal never provision. Omitted description preserves existing **destination** metadata; `''` clears it. Creating a site override never copies the tenant description. Existing [safe serialization and saved-data compatibility](tenant-key-value-store.md#serialization-and-saved-data-compatibility), nullable-Note differences and ETag preconditions apply unchanged. T is not runtime validation. A tenant serialization failure may occur after catalog discovery but before any setup/item writes. No cross-store transaction or caller get-to-save compare-and-swap is implied.

The **Configuration** test web part tab exposes manual modes, effective Read, explicitly scoped Save/Remove and Physical inventory. Advanced diagnostics provide Text/JSON/Big integer/Date and description inputs; an empty description explicitly clears metadata. Results identify their current parameters, value type and source. Inventory reports the complete fetched count and previews ten rows. Use unique `spfx-toolkit-config-demo-<run>-` keys and remove them from every written scope. Local tests exercise the real hook/service facade with simulated boundaries; [authenticated checks](../../../development/sharepoint-validation/checklists/configuration-store.md#configuration-store) are separate.

See the [non-React factory](../../services/configuration-store.md). Source: [hook](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxConfigurationStore.ts), [canonical types and service](../../../../packages/spfx-react-toolkit/src/services/spfx-configuration-store.service.ts).

## Public configuration types

The hook re-exports the canonical service types; import them from `/hooks`, `/services` or the package root. Their common method contracts and service-only address union are documented by the [configuration service](../../services/configuration-store.md).

| Type | Contract |
| --- | --- |
| `SPFxConfigurationStoreScope` | `'tenant'` or `'site'`, always a physical destination |
| `SPFxConfigurationStoreMode` | A physical scope or `'layered'`; no default |
| `SPFxConfigurationStoreOptions` | Required `mode`; the hook obtains local URLs/client from its provider |
| `SPFxConfigurationStoreTarget` | Required `scope`, enabled by the current mode |
| `SPFxConfigurationStoreSaveOptions` | Required destination scope plus optional destination-local `description`; omission preserves it and `''` clears it |
| `SPFxConfigurationStoreItem<T>` | Required `key`, `value`, `description`, `id`, `source`; `description` can be undefined and source identifies the winning physical record |
| `SPFxConfigurationStoreService` | `get`, `listScope`, `save` and `remove`; all failed operations reject |
| `SPFxConfigurationStoreServiceOptions` | Factory-only discriminated URL dependencies: tenant requires `discoveryWebUrl`, site requires `siteCollectionUrl`, layered requires both |
| `SPFxConfigurationStoreResult` | The four service methods plus independent read/write state and `isReady`; no cached value or permission flag |

An invalid `mode` throws during rendering. Unavailable client/required URL operations reject in the relevant error channel. Equivalent inline option objects preserve callback identity; mode and only relevant URL/client changes create another identity. Options, target and item types describe expected values without validating stored application data.

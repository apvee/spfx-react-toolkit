# useSPFxSiteKeyValueStore

[Storage overview](../../INDEX.md#storage-hooks) · [Usage guide](../../../guides/storage.md)

```ts
import { useSPFxSiteKeyValueStore } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).

Shares a hidden `SiteKeyValueStore` in `pageContext.site.absoluteUrl`, the root web of the current site collection. Root pages and every subsite use the same data; another site collection has its own store. The current web URL and tenant app catalog are not used as this store's target. New writes use the same safe encoding and ETag preconditions as the tenant store; historical site data keeps its previous interpretation without migration.

## Signature and public types

```typescript
function useSPFxSiteKeyValueStore(): SPFxSiteKeyValueStoreResult

interface SPFxSiteKeyValueStoreItem<T = unknown> {
  readonly key: string;
  readonly value: T;
  readonly description: string | undefined;
  readonly id: number;
}

interface SPFxSiteKeyValueStoreResult {
  readonly isLoading: boolean;
  readonly error: Error | undefined;
  readonly isWriting: boolean;
  readonly writeError: Error | undefined;
  readonly canWrite: boolean;
  readonly isReady: boolean;
  readonly get: <T = unknown>(key: string) => Promise<SPFxSiteKeyValueStoreItem<T> | undefined>;
  readonly list: () => Promise<SPFxSiteKeyValueStoreItem<unknown>[]>;
  readonly save: <T = unknown>(key: string, value: T, description?: string) => Promise<void>;
  readonly remove: (key: string) => Promise<void>;
}

```

## Readiness, errors and lifecycle

Mount checks permissions without provisioning or fetching values. `isReady` means the SPHttpClient and a nonempty normalized collection URL are available; it does not mean the list exists or the user has access. Operations invoked before readiness reject.

`get` and `list` never provision. Confirmed absent list/key returns `undefined`; an absent list returns `[]`. Other read failures set `error` and resolve with those same fallbacks. A fallback alone cannot establish absence. Display the current render's `error` before displaying a missing result; inspecting the error captured by a callback immediately after `await` can read stale React state. The standalone [service](../../services/site-key-value-store.md#createspfxsitekeyvaluestoreservice) rejects read failures instead.

`save` provisions or repairs if needed. `remove` never provisions and confirmed absence is a no-op. Failed writes set `writeError` and reject: catch the returned promise. Only `save` and the standalone service's explicit `ensureListReady` perform setup.

`isLoading` and `isWriting` independently track all pending operations for the current identity. Starting a read/write clears that channel's error; only its latest-started operation can publish a failure. Client or collection changes reset visible state and permission results; subsite navigation within the collection retains storage identity. Old callbacks keep their original collection/client, and obsolete completions or unmount do not publish state. Dispatched requests still settle and remote writes are not cancelled or serialized.

## Storage, grants and serialization

The root list uses generic list template 100, `Hidden: true`, `NoCrawl: true`, indexed unique `Title`, and multiline text `Value` and `Description`. Reads validate field types without repair. Writes reject incompatible schema or failed uniqueness enforcement without deleting existing data. `list` follows all continuation pages in server Title order and rejects incomplete, foreign or repeated continuations; a later-page failure becomes the hook's error/fallback, not partial success.

Read access depends on actual root-web/list permissions. `canWrite` checks effective Add, Edit and Delete Items on an existing list, including unique list permissions. An absent list also requires Manage Lists on the root web. The hook refreshes this advisory indicator on mount and after each current-identity read/write settles, including failures. Failed checks yield false; only the latest permission check can update state, and obsolete/unmounted operations never start a refresh. SharePoint enforces authorization, schema-management rights and item restrictions even when `canWrite` is true; subsite membership alone is insufficient.

Keys follow SharePoint's case-insensitive server collation and unique-column comparison; returned keys retain stored Title casing. Updates and deletion require the observed item ETag. If collection metadata omits it, the service retrieves a fresh item snapshot and its header ETag, preserving metadata from that same snapshot. Missing/invalid ETags and HTTP 412 conflicts reject without a wildcard overwrite or blind retry. This protects the service's snapshot-to-mutation interval, not changes since a consumer's previous get(). Creation recovery allows one validated conditional update and exposes that update's actual failure. Snapshot/deletion 404 succeeds only after verifying absence; a recreated key is not deleted. Omitted update descriptions preserve existing metadata; `''` clears it.

New site writes use the [safe list-store encoding](tenant-key-value-store.md#serialization-and-saved-data-compatibility): empty and JSON-like strings are JSON-quoted so their text survives readback; bigint becomes exact decimal **string**, Date remains ISO **string**, and JSON-compatible objects/arrays retain JSON encoding. Historical JSON-first readers can read these new values. Circular objects, nested bigint and invalid Dates throw before any setup request. Other non-JSON values/prototypes have no complete round-trip guarantee; `T` is not runtime validation.

No existing rows, fields or public signatures are migrated. Historical raw `123`/`true` still decode as number/boolean, and any previously lost bigint precision cannot be reconstructed. The site-specific blank Note contract is preserved: REST `Value: null` reads as `''`, while the stored text `'null'` decodes to legacy `null`. New empty text is stored as `'""'`, avoiding SharePoint's empty-Note normalization. The tenant's historical REST-null behavior remains unchanged.

## Observable sample

The **Site** panel offers Text, JSON, Big integer and Date inputs for explicit Save. Read the key afterward to inspect the value and its runtime type; empty description explicitly clears metadata. Try Text `123`/empty text and Big integer `9007199254740993` to observe string-preserving/exact readback. Date reads as ISO string. List reports the complete count and previews the first ten entries. Conflicts/errors never appear as successful saves or missing records. Tests execute the actual panel/hook locally with simulated service boundaries; [authenticated root/subsite acceptance](../../../development/sharepoint-validation/INDEX.md) remains separate.

## Example: Manual read and save

```tsx
import * as React from 'react';
import { useSPFxSiteKeyValueStore, SPFxSiteKeyValueStoreItem } from '@apvee/spfx-react-toolkit';

function SiteSettings() {
  const store = useSPFxSiteKeyValueStore();
  const [item, setItem] = React.useState<SPFxSiteKeyValueStoreItem<unknown>>();
  const [hasRead, setHasRead] = React.useState(false);
  const [message, setMessage] = React.useState('');

  const load = async () => {
    setHasRead(false);
    setMessage('');
    try {
      setItem(await store.get('toolkit-site-demo-setting'));
      setHasRead(true);
    } catch (failure) {
      setMessage(String(failure)); // For example, an operation before readiness.
    }
  };

  const save = async () => {
    setMessage('');
    try {
      await store.save('toolkit-site-demo-setting', { enabled: true }, 'Collection setting');
      setMessage('Saved');
    } catch (failure) {
      setMessage(String(failure));
    }
  };

  return (
    <div>
      <button onClick={load} disabled={!store.isReady || store.isLoading}>Read</button>
      <button onClick={save} disabled={!store.isReady || !store.canWrite || store.isWriting}>Save</button>
      {store.error ? <p role="alert">{store.error.message}</p> :
        hasRead && !store.isLoading && <p>{item ? JSON.stringify(item.value) : 'Key not found'}</p>}
      {store.writeError && <p role="alert">{store.writeError.message}</p>}
      {message && <p>{message}</p>}
    </div>
  );
}
```

Use a disposable key prefix for testing and catch `remove` rejections just as for `save`. The [Site sample and real-host checklist](../../../development/sharepoint-validation/checklists/site-key-value-store.md#site-collection-key-value-store) cover CRUD, sharing and permissions.

## Source

[View hook source](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxSiteKeyValueStore.ts)

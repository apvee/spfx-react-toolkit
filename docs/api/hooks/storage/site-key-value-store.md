# useSPFxSiteKeyValueStore

[Storage overview](../../INDEX.md#storage-hooks) · [Usage guide](../../../guides/storage.md)

```ts
import { useSPFxSiteKeyValueStore } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).


Shares a hidden `SiteKeyValueStore` in `pageContext.site.absoluteUrl`, the root web of the current site collection. Root pages and every subsite use the same data; another site collection has its own store. The current web URL and tenant app catalog are not used as this store's target. Existing tenant store APIs and behavior remain unchanged.

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

Read access depends on actual root-web/list permissions. `canWrite` checks effective Add, Edit and Delete Items on an existing list, including unique list permissions. An absent list also requires Manage Lists on the root web. The hook refreshes this advisory indicator after successful writes. SharePoint enforces authorization, schema-management rights and item restrictions even when `canWrite` is true; subsite membership alone is insufficient.

Keys follow SharePoint's case-insensitive server collation and unique-column comparison; returned keys retain stored Title casing. Updates use wildcard ETags with last-writer-wins behavior, without optimistic locking. Omitted update descriptions preserve existing metadata; `''` clears it.

The [legacy serialization format](tenant-key-value-store.md#serialization) applies: JSON-like strings can become primitives, dates remain strings, and bigint can lose precision. `T` provides no runtime validation or complete round-trip guarantee. For this site store, a blank SharePoint Note field returned as `Value: null` reads as `''`; the stored text `'null'` deserializes to `null`. Tenant mapping remains unchanged.

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

---

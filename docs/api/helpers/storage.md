# Storage Helpers

[Helpers overview](../INDEX.md#helpers-and-styles) · [Non-React composition](../../guides/helpers-and-services.md)

Import:

```ts
import { createScopedSPFxStorageKey } from '@apvee/spfx-react-toolkit';
```

## `createScopedSPFxStorageKey`

```ts
function createScopedSPFxStorageKey(instanceId: string, key: string): string;
```

Builds keys in the exact format `spfx:<instanceId>:<key>`.

Both string inputs are required. They are concatenated unchanged, including whitespace, colons and empty strings; the helper does not validate uniqueness, encode values, read/write storage or inspect SPFx context. Keep the instance ID and unscoped key stable to retrieve the same entry. The prefix scopes names only; it is not access control or encryption.

```ts
const scoped = createScopedSPFxStorageKey('webpart-42', 'filters');
// spfx:webpart-42:filters
```

The [local storage hook](../hooks/storage/local-storage.md) and [session storage hook](../hooks/storage/session-storage.md) own JSON storage, events, initialization and write failures.

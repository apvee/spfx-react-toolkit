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

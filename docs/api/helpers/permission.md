# Permission Helpers

[Helpers overview](../INDEX.md#helpers-and-styles) · [Non-React composition](../../guides/helpers-and-services.md)


Import:

```ts
import { hasSPFxPermission } from '@apvee/spfx-react-toolkit';
```

## `hasSPFxPermission`

```ts
function hasSPFxPermission(
  permissionSet: SPPermission | undefined,
  permission: SPPermission
): boolean;
```

Returns `false` when `permissionSet` is undefined. Otherwise delegates to `SPPermission.hasPermission(permission)`.

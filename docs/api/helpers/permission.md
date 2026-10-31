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

Pass a host `SPPermission` set and the requested SharePoint permission bit. This synchronous capability check performs no I/O, requires no provider and does not acquire Graph/API scopes. An available permission flag is an indication from the supplied snapshot; SharePoint still authorizes actual requests. The helper does not catch failures from a supplied object's `hasPermission` implementation.

```ts
import { SPPermission } from '@microsoft/sp-page-context';
import { hasSPFxPermission } from '@apvee/spfx-react-toolkit/helpers';

const canEdit = hasSPFxPermission(pageContext.web.permissions, SPPermission.editListItems);
const unavailable = hasSPFxPermission(undefined, SPPermission.editListItems); // false
```

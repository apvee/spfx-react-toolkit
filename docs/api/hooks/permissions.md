# Permission Hooks

Public hook signatures and result shapes match the library source. External types (React, SPFx and PnPjs) come from their respective packages. All hooks require a matching SPFx provider.

## useSPFxPermissions

```typescript
export function useSPFxPermissions(): SPFxPermissionsInfo;
export interface SPFxPermissionsInfo {
    readonly sitePermissions: SPPermission | undefined;
    readonly webPermissions: SPPermission | undefined;
    readonly listPermissions: SPPermission | undefined;
    readonly hasWebPermission: (permission: SPPermission) => boolean;
    readonly hasSitePermission: (permission: SPPermission) => boolean;
    readonly hasListPermission: (permission: SPPermission) => boolean;
}
```

Provides current-site/web/list permission sets and specific helper checks. Missing permission sets return false through the helpers. Use `SPPermission` constants from @microsoft/sp-page-context; server authorization still controls each request.

```tsx
import * as React from 'react';
import { useSPFxPermissions } from '@apvee/spfx-react-toolkit';

import { SPPermission } from '@microsoft/sp-page-context';

function PermissionStatus() {
  const { hasWebPermission } = useSPFxPermissions();
  return <p>{hasWebPermission(SPPermission.addListItems) ? 'Can add items' : 'Read only'}</p>;
}
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxPermissions.ts)

## useSPFxCrossSitePermissions

```typescript
export function useSPFxCrossSitePermissions(siteUrl?: string, options?: SPFxCrossSitePermissionsOptions): SPFxCrossSitePermissionsInfo;
export interface SPFxCrossSitePermissionsOptions {
    webUrl?: string;
    listId?: string;
}
export interface SPFxCrossSitePermissionsInfo {
    sitePermissions?: SPPermission;
    webPermissions?: SPPermission;
    listPermissions?: SPPermission;
    hasWebPermission: (permission: SPPermission) => boolean;
    hasSitePermission: (permission: SPPermission) => boolean;
    hasListPermission: (permission: SPPermission) => boolean;
    isLoading: boolean;
    error?: Error;
}
```

Changing site/web/list target clears prior permission state and starts a new request. Empty/undefined siteUrl returns idle state without fetching. Obsolete target and unmounted completions cannot publish results. Options may specify a web URL or list ID; the returned helpers check each respective permission set. Server authorization still applies.

```tsx
import * as React from 'react';
import { useSPFxCrossSitePermissions } from '@apvee/spfx-react-toolkit';

import { SPPermission } from '@microsoft/sp-page-context';

function TargetPermissions({ siteUrl }: { siteUrl?: string }) {
  const { isLoading, error, hasWebPermission } = useSPFxCrossSitePermissions(siteUrl);
  if (isLoading) return <p>Loading permissions</p>;
  if (error) return <p>{error.message}</p>;
  return <p>{hasWebPermission(SPPermission.addListItems) ? 'Can add items' : 'Cannot add items'}</p>;
}
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxCrossSitePermissions.ts)

## See Also

- [API index](../../INDEX.md)
- [Services](../services/INDEX.md)
- [SharePoint validation](../../SHAREPOINT-VALIDATION.md)

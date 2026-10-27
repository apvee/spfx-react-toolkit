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

The legacy string overload remains supported. Undefined, empty or whitespace-only `siteUrl` returns idle state without fetching; `hasSitePermission`, `hasWebPermission` and `hasListPermission` then return false. With a URL it retrieves site-collection root-web permissions, selected-web permissions and optionally list permissions. `options.webUrl` defaults to `siteUrl`; `listId` is a hyphenated GUID with optional paired braces. The result names and optional permission sets are unchanged.

The site set uses `/_api/site/rootweb/effectivebasepermissions`. This endpoint follows the documented [SP.Site.rootWeb resource](https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-visio/jj245543(v=office.15)) to [SP.Web.EffectiveBasePermissions](https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-visio/jj245395(v=office.15)); the former `/_api/site/effectivebasepermissions` request addressed a property absent from SP.Site. Site permissions therefore describe the collection's root web, including when the URL points to a subweb.

Provider, native client, current user and target changes hide previous sets during render, before effects start their replacement request. Obsolete and unmounted completions cannot publish results. HTTP failures and malformed permission masks expose `error` and no grant. Server authorization still applies.

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

### Scoped resource overload

Pass an options object to select one resource:

```typescript
export function useSPFxCrossSitePermissions(options: SPFxScopedPermissionsOptions): SPFxScopedPermissionsInfo;
export interface SPFxScopedPermissionsInfo {
    readonly permissions: SPPermission | undefined;
    readonly hasPermission: (permission: SPPermission) => boolean;
    readonly isLoading: boolean;
    readonly error: Error | undefined;
}
```

`SPFxScopedPermissionsOptions` is a discriminated union. These are its valid shapes (each includes optional `webUrl`):

| Scope | Required selector | Additional input |
| --- | --- | --- |
| `web` (default) | None; list selectors are forbidden | None |
| `list` | Exactly one of `listId`, `listTitle`, `listPath` | None |
| `item` | Exactly one of `listId`, `listTitle`, `listPath` | Required positive integer `itemId` |
| `folder` | Exactly one of `listId`, `listTitle`, `listPath` | Required `folderPath` |

`webUrl` defaults to the current provider web and must be absolute HTTP(S) without credentials, query or fragment. `listId` is a hyphenated GUID, optionally in paired braces. `listTitle` matches the exact nonblank title, including apostrophes/spaces. `listPath` is a **decoded web-relative list root** such as `Shared Documents` or `Lists/Tasks`, without a leading slash. `folderPath` is **decoded server-relative**, such as `/sites/team/Shared Documents/Reports`. Supply literal spaces, percent signs and hash characters; do not URL-encode input. Paths reject schemes, backslashes, queries and dot segments. `itemId` is valid only for item scope, and `folderPath` only for folder scope. Runtime validation also rejects invalid JavaScript callers before any HTTP dispatch.

```tsx
import * as React from 'react';
import { SPPermission } from '@microsoft/sp-page-context';
import { useSPFxCrossSitePermissions } from '@apvee/spfx-react-toolkit';
import type { SPFxScopedPermissionsOptions } from '@apvee/spfx-react-toolkit';

function ItemPermission({ itemId }: { itemId: number }) {
  const target: SPFxScopedPermissionsOptions = { scope: 'item', listTitle: 'Tasks', itemId };
  const info = useSPFxCrossSitePermissions(target);
  if (info.error) return <p>{info.error.message}</p>;
  if (info.isLoading) return <p>Loading item permissions</p>;
  return <p>{info.hasPermission(SPPermission.editListItems) ? 'Can edit' : 'Read only'}</p>;
}

function FolderPermission() {
  const info = useSPFxCrossSitePermissions({
    scope: 'folder',
    listPath: 'Shared Documents',
    folderPath: '/sites/team/Shared Documents/Reports',
  });
  return <p>{info.error?.message ?? String(info.hasPermission(SPPermission.viewListItems))}</p>;
}
```

For current-web scope, an available mask is read through `useSPFxPermissions` without REST. If it is unavailable, or the target web differs, the hook uses the native SPHttpClient. List and item scope request their own effective mask, without requesting unrelated site/web masks. Folder scope reads selected-list identity and the folder's list item with its permission mask and parent-list identity; a mismatching or missing parent list is an error. It does not infer folder permissions from the list or a path prefix. A library's root folder may lack a list item; use list scope for that resource.

`permissions` is undefined during loading, validation/HTTP/payload errors and target replacement. `hasPermission` returns false when the mask is unavailable and otherwise requires every bit in the supplied native `SPPermission`. `isLoading` is false for native-mask success and validation errors. `error` holds only the current failure. A newly allocated object with equivalent scalar values does not repeat requests. Concurrent identical requests may share an inflight promise **only within one provider and matching client/user identity**; completed decisions are not cached and there is no polling or refresh timer. Changing provider, client, user or target immediately discards old asynchronous decisions. In-place host changes still require a React render; the hook does not subscribe independently to every PageContext field.

Requires a matching SPFx provider and its native PageContext/SPHttpClient; it imports no PnP feature registrations. Local behavior/type tests use native boundary doubles and do not establish authenticated SharePoint results. The Context demo includes a staged scoped target and explicit Run/repeat, alongside the unchanged legacy staged URL view. The [SharePointPermissionTrim](../controls/sharepoint-permission-trim.md) sample composes this same hook; [authenticated scoped checks](../../SHAREPOINT-VALIDATION.md) remain separate. A permission check is a UI decision; SharePoint remains responsible for authorization.

## See Also

- [API index](../../INDEX.md)
- [Services](../services/INDEX.md)
- [SharePoint validation](../../SHAREPOINT-VALIDATION.md)

Qualified overload result members: `SPFxCrossSitePermissionsOptions.webUrl`, `SPFxCrossSitePermissionsOptions.listId` select the optional legacy web/list. `SPFxCrossSitePermissionsInfo.sitePermissions`, `SPFxCrossSitePermissionsInfo.webPermissions`, `SPFxCrossSitePermissionsInfo.listPermissions` are absent while unavailable; `SPFxCrossSitePermissionsInfo.hasSitePermission`, `SPFxCrossSitePermissionsInfo.hasWebPermission`, `SPFxCrossSitePermissionsInfo.hasListPermission` check those masks and return false when absent. `SPFxCrossSitePermissionsInfo.isLoading` reports pending reads and `SPFxCrossSitePermissionsInfo.error` reports the current failure.

`SPFxScopedPermissionsInfo.permissions` is the selected mask or undefined; `SPFxScopedPermissionsInfo.hasPermission` checks every requested bit or returns false while absent; `SPFxScopedPermissionsInfo.isLoading` reports pending reads and `SPFxScopedPermissionsInfo.error` holds the current error or undefined. `SPFxScopedPermissionsOptions.scope`, `SPFxScopedPermissionsOptions.webUrl`, `SPFxScopedPermissionsOptions.listId`, `SPFxScopedPermissionsOptions.listTitle`, `SPFxScopedPermissionsOptions.listPath`, `SPFxScopedPermissionsOptions.itemId` and `SPFxScopedPermissionsOptions.folderPath` follow the discriminated scope/selector table above.

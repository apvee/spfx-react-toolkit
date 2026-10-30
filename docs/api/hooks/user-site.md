# User and Site Hooks

[React and SPFx hooks](../INDEX.md#react-and-spfx-hooks)

Public hook signatures and result shapes match the library source. External types (React, SPFx and PnPjs) come from their respective packages. All hooks require a matching SPFx provider.

## useSPFxUserInfo

```typescript
export function useSPFxUserInfo(): SPFxUserInfo;
export interface SPFxUserInfo {
    readonly loginName: string;
    readonly displayName: string;
    readonly email?: string;
    readonly isExternal: boolean;
}
```

Maps the current PageContext user. Optional email depends on host metadata; isExternal falls back to false when the host omits the guest flag.

`SPFxUserInfo.loginName` and `SPFxUserInfo.displayName` are native user strings, `SPFxUserInfo.email` is optional host metadata, and `SPFxUserInfo.isExternal` reflects `isExternalGuestUser` (false if omitted). The hook maps metadata without a profile request.

```tsx
const user = useSPFxUserInfo();
return <span>{user.displayName} ({user.loginName}) {user.email ?? 'Email unavailable'}</span>;
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxUserInfo.ts)

## useSPFxUserPhoto

```typescript
export function useSPFxUserPhoto(options?: SPFxUserPhotoOptions): SPFxUserPhotoResult;
export type SPFxUserPhotoSize = '48x48' | '64x64' | '96x96' | '120x120' | '240x240' | '360x360' | '432x432' | '504x504' | '648x648';
export interface SPFxUserPhotoOptions {
    userId?: string;
    email?: string;
    size?: SPFxUserPhotoSize;
    autoFetch?: boolean;
}
export interface SPFxUserPhotoResult {
    readonly photoUrl: string | undefined;
    readonly photoBlob: Blob | undefined;
    readonly isLoading: boolean;
    readonly error: Error | undefined;
    readonly reload: () => Promise<void>;
    readonly isReady: boolean;
}
```

Loads a Graph photo for the specified user ID/email or current user, with default size 240x240 and autoFetch true. Changing Graph service, user ID/email or size clears prior photo state and invalidates previous request ownership. Only the current request publishes photo/error/loading results. Obsolete blob URLs are revoked; active URLs are cleaned up on replacement or unmount. Requests are not cancelled. Handle unavailable photos and authentication/permission errors in the UI.

```tsx
import * as React from 'react';
import { useSPFxUserPhoto } from '@apvee/spfx-react-toolkit';

function Avatar() {
  const { photoUrl, isLoading, error } = useSPFxUserPhoto();
  if (isLoading) return <p>Loading photo</p>;
  if (error || !photoUrl) return <p>No photo available</p>;
  return <img src={photoUrl} alt="Current user" />;
}
```

`SPFxUserPhotoOptions.userId` takes precedence over `email`; if neither is truthy the service requests `/me`. `size` defaults to `'240x240'` and permits exactly the exported `SPFxUserPhotoSize` values above; `autoFetch` defaults to true. The Graph client and approved permissions appropriate to the selected user are required. See [Graph client setup](http-clients/graph-client.md).

`SPFxUserPhotoResult.photoBlob` is the downloaded Blob and `SPFxUserPhotoResult.photoUrl` is its hook-owned object URL. Do not retain that URL after identity replacement/unmount. `SPFxUserPhotoResult.reload()` loads explicitly and captures failures in `error` rather than rethrowing them. `SPFxUserPhotoResult.isLoading` tracks the latest request and `SPFxUserPhotoResult.isReady` means a URL exists with no pending load or error. A 404 is a displayed “photo not found” error; 403 and 401 become permission and authentication messages. There is no separate Graph initialization error field in this hook.

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxUserPhoto.ts)

## useSPFxSiteInfo

```typescript
export function useSPFxSiteInfo(): SPFxSiteInfo;
export interface SPFxGroupInfo {
    readonly id: string;
    readonly isPublic: boolean;
}
export interface SPFxSiteInfo {
    readonly webId: string;
    readonly webUrl: string;
    readonly webServerRelativeUrl: string;
    readonly title: string;
    readonly languageId: number;
    readonly logoUrl?: string;
    readonly siteId: string;
    readonly siteUrl: string;
    readonly siteServerRelativeUrl: string;
    readonly siteClassification?: string;
    readonly siteGroup?: SPFxGroupInfo;
}
```

Maps site collection and web metadata from PageContext. Classification, logo and group metadata are optional.

| Result member | Meaning/default |
| --- | --- |
| `SPFxSiteInfo.webId`, `SPFxSiteInfo.siteId` | Native web/site GUIDs as strings |
| `SPFxSiteInfo.webUrl`, `SPFxSiteInfo.siteUrl` | Current web and collection absolute URLs |
| `SPFxSiteInfo.webServerRelativeUrl`, `SPFxSiteInfo.siteServerRelativeUrl` | Corresponding server-relative paths |
| `SPFxSiteInfo.title` | Current web title |
| `SPFxSiteInfo.languageId` | Native web LCID; 1033 if missing |
| `SPFxSiteInfo.logoUrl` | Optional host web logo URL |
| `SPFxSiteInfo.siteClassification` | Optional legacy classification |
| `SPFxSiteInfo.siteGroup` | Associated group or undefined; `SPFxGroupInfo.id` is its GUID string and `SPFxGroupInfo.isPublic` defaults to false if the host omits the flag |

```tsx
const site = useSPFxSiteInfo();
return <a href={site.webUrl}>{site.title} ({site.languageId})</a>;
```

The hook maps PageContext without REST. It memoizes by native site/web/legacy object identity; an in-place field mutation may require those references to change before the derived snapshot updates.

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxSiteInfo.ts)

## useSPFxHubSiteInfo

```typescript
export function useSPFxHubSiteInfo(): SPFxHubSiteInfo;
export interface SPFxHubSiteInfo {
    readonly isHubSite: boolean;
    readonly hubSiteId: string | undefined;
    readonly hubSiteUrl: string | undefined;
    readonly isLoading: boolean;
    readonly error: Error | undefined;
}
```

`isHubSite` means legacy metadata associates the site with a nonempty, nonzero hub ID; it does not mean the current site is the hub itself. The hook uses current-site metadata when the current site is identified as the hub, otherwise reads `/_api/web/hubsitedata(false)`. Loading starts in an effect. Requests are not cancelled or guarded against stale completion, and previously resolved URL/error state can remain when association changes. Check `isHubSite`, `isLoading`, `error` and the optional URL together before rendering a link.

`SPFxHubSiteInfo.hubSiteId` is undefined for absent, blank or all-zero IDs; `SPFxHubSiteInfo.hubSiteUrl` is optional resolved hub data. `SPFxHubSiteInfo.isLoading` tracks its effect-driven lookup and `SPFxHubSiteInfo.error` captures lookup/JSON failures. There are no options, manual reload or refresh timer. The hook parses the native hub response without automatically checking HTTP `response.ok`.

```tsx
const hub = useSPFxHubSiteInfo();
return hub.isHubSite && !hub.isLoading && !hub.error && hub.hubSiteUrl
  ? <a href={hub.hubSiteUrl}>Hub</a> : <span>Hub unavailable</span>;
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxHubSiteInfo.ts)

## useSPFxListInfo

```typescript
export function useSPFxListInfo(): SPFxListInfo | undefined;
export interface SPFxListInfo {
    readonly id: string;
    readonly title: string;
    readonly serverRelativeUrl: string;
    readonly baseTemplate?: number;
    readonly isDocumentLibrary?: boolean;
}
```

Returns undefined when list metadata or its ID is absent. `SPFxListInfo.id` is the native GUID string; `SPFxListInfo.title` defaults to `'Unknown List'`, `SPFxListInfo.serverRelativeUrl` to `''`, and `SPFxListInfo.baseTemplate` remains undefined if missing. Although its declaration permits omission, `SPFxListInfo.isDocumentLibrary` is always populated by the current mapper: true for template 101, false for other or missing templates. The hook memoizes by native list object identity and does not fetch list details.

```tsx
const list = useSPFxListInfo();
return <p>{list ? `${list.title}: ${list.isDocumentLibrary ? 'document library' : 'list'}` : 'No list context'}</p>;
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxListInfo.ts)

## See Also

- [API index](../INDEX.md)
- [Services](../INDEX.md#non-react-services)
- [SharePoint validation](../../development/sharepoint-validation/INDEX.md)

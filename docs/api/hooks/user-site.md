# User and Site Hooks

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

Loads hub information when available. Inspect isLoading/error and optional hubSiteId/hubSiteUrl rather than assuming every site is connected to a hub.

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

Returns undefined when the page has no list context. baseTemplate and isDocumentLibrary are optional because host metadata can omit the template.

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxListInfo.ts)

## See Also

- [API index](../../INDEX.md)
- [Services](../services/INDEX.md)
- [SharePoint validation](../../SHAREPOINT-VALIDATION.md)

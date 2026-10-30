# User Photo Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

Import this factory and its public types from `@apvee/spfx-react-toolkit/services` or the package root. Callers supply the initialized client/context; these services require no React provider.

## `createSPFxUserPhotoService`

```ts
function createSPFxUserPhotoService(
  graphClient: MSGraphClientV3
): SPFxUserPhotoService;
```

Required public types:

- `SPFxUserPhotoServiceSize`
- `SPFxUserPhotoServiceOptions`
- `SPFxUserPhotoService`

Returned method:

| Method | Description |
|--------|-------------|
| `getPhotoBlob(options?)` | Gets a profile photo blob from Microsoft Graph |

`SPFxUserPhotoServiceOptions` accepts `userId`, `email`, and `size`. If no user identifier is supplied, the endpoint targets `/me`. The default size is `240x240`.

## Usage and public types

```ts
const photos = createSPFxUserPhotoService(graphClient);
const ownPhoto = await photos.getPhotoBlob();
const selected = await photos.getPhotoBlob({ email: 'user@contoso.com', size: '96x96' });
```

`SPFxUserPhotoServiceSize` is `'48x48' | '64x64' | '96x96' | '120x120' | '240x240' | '360x360' | '432x432' | '504x504' | '648x648'`. `SPFxUserPhotoServiceOptions` has readonly optional `userId: string`, `email: string`, and `size: SPFxUserPhotoServiceSize`. A nonempty userId wins over email; empty identifiers fall through to the next choice, then `/me`. Identifiers are interpolated unchanged by the [path helper](../helpers/graph-path.md#builduserphotoendpoint); use valid path-safe Graph identifiers.

`SPFxUserPhotoService.getPhotoBlob(options?)` returns `Promise<Blob>` through the supplied Graph client. All Graph failures, including missing photos and insufficient grants, reject. Size availability depends on the stored photo. Construction performs no requests; the caller owns authentication, permissions and any object URL created from the result. This service does not cache photos, create/revoke object URLs, or convert failures into empty state; the [user photo hook](../hooks/user-site.md#usespfxuserphoto) adds React lifecycle management.

# Graph Path Helpers

[Helpers overview](../INDEX.md#helpers-and-styles) · [Non-React composition](../../guides/helpers-and-services.md)

Import:

```ts
import {
  buildOneDriveAppDataPath,
  buildUserPhotoEndpoint,
} from '@apvee/spfx-react-toolkit';
```

## `buildOneDriveAppDataPath`

```ts
function buildOneDriveAppDataPath(fileName: string, folder?: string): string;
```

Builds Microsoft Graph app-root content paths:

| Input | Output |
|-------|--------|
| `('settings.json')` | `/me/drive/special/approot:/settings.json:/content` |
| `('settings.json', 'dashboard')` | `/me/drive/special/approot:/dashboard/settings.json:/content` |

Folder names are sanitized by replacing characters outside `a-z`, `A-Z`, `0-9`, `-`, and `_` with `-`. Omitted or empty folder uses the app root. File names are interpolated unchanged, without path encoding or validation; provide a path-safe name. The helper does not create folders or perform I/O.

## `buildUserPhotoEndpoint`

```ts
function buildUserPhotoEndpoint(options?: {
  readonly userId?: string;
  readonly email?: string;
  readonly size?: '48x48' | '64x64' | '96x96' | '120x120' | '240x240' | '360x360' | '432x432' | '504x504' | '648x648';
}): string;
```

Builds Microsoft Graph profile photo endpoints:

| Options | Base path |
|---------|-----------|
| `{ userId }` | `/users/{userId}` |
| `{ email }` | `/users/{email}` |
| omitted | `/me` |

The default size is `240x240`; the full endpoint ends in `/photos/<size>/$value`. Nonempty `userId` takes precedence over `email`; identifiers are interpolated unchanged without encoding or validation.

## Usage and boundaries

Both helpers are synchronous string builders with no provider, client or request dependency. They do not discover permissions, create a OneDrive folder, verify the selected user or test whether a photo exists. Graph rejects unsupported identifiers, paths or permissions when a caller uses the result.

```ts
const rootFile = buildOneDriveAppDataPath('settings.json');
const scopedFile = buildOneDriveAppDataPath('settings.json', 'Sales Reports');
// /me/drive/special/approot:/Sales-Reports/settings.json:/content
const ownPhoto = buildUserPhotoEndpoint();
// /me/photos/240x240/$value
const userPhoto = buildUserPhotoEndpoint({ userId: 'user-object-id', email: 'ignored@contoso.com', size: '96x96' });
// /users/user-object-id/photos/96x96/$value
```

The photo options shape is part of the function signature; its source alias is not a separately exported public type. The service exports its own equivalent [photo option and size types](../services/user-photo.md). See the [OneDrive service](../services/onedrive-app-data.md) for browser download behavior; its read operation derives a metadata path instead of issuing an authenticated `/content` download.

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

Folder names are sanitized by replacing characters outside `a-z`, `A-Z`, `0-9`, `-`, and `_` with `-`.

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

The default size is `240x240`.

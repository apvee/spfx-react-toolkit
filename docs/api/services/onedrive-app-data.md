# OneDrive App Data Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

```ts
import { createSPFxOneDriveAppDataService } from '@apvee/spfx-react-toolkit/services';
```


## `createSPFxOneDriveAppDataService`

```ts
function createSPFxOneDriveAppDataService(
  graphClient: MSGraphClientV3
): SPFxOneDriveAppDataService;
```

Required public types:

- `SPFxOneDriveAppDataReadResult`
- `SPFxOneDriveAppDataService`

Returned methods:

| Method | Description |
|--------|-------------|
| `read<T>(fileName, folder?)` | Reads JSON content from the user's OneDrive app root |
| `write<T>(fileName, content, folder?)` | Writes JSON content to the user's OneDrive app root |

`read` returns `{ data: undefined, isNotFound: true }` for recognized not-found Graph errors. It throws JSON parse errors for invalid stored content and rethrows non-not-found Graph errors.

# User Photo Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

```ts
import { createSPFxUserPhotoService } from '@apvee/spfx-react-toolkit/services';
```


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

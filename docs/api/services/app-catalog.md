# App Catalog Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

```ts
import { createSPFxAppCatalogService } from '@apvee/spfx-react-toolkit/services';
```


## `createSPFxAppCatalogService`

```ts
function createSPFxAppCatalogService(
  spHttpClient: SPHttpClient | undefined,
  pageContext: SPFxAppCatalogPageContextLike | undefined
): SPFxAppCatalogService;
```

Required public types:

- `SPFxAppCatalogPageContextLike`
- `SPFxAppCatalogService`

Returned methods:

| Method | Description |
|--------|-------------|
| `discoverUrl()` | Reads tenant settings and caches `CorporateCatalogUrl` |
| `canCurrentUserWrite(catalogUrl)` | Checks whether current user is site admin for the catalog |

`discoverUrl` throws when the client/page context is missing, the tenant settings request fails, or the tenant app catalog is not provisioned. `canCurrentUserWrite` returns `false` when the client is missing or the permission check fails.

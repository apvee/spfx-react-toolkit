# App Catalog Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

Import this factory and its public types from `@apvee/spfx-react-toolkit/services` or the package root. Callers supply the initialized client/context; these services require no React provider.

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

## Usage and public contracts

```ts
const catalog = createSPFxAppCatalogService(context.spHttpClient, context.pageContext);
const url = await catalog.discoverUrl();
const isCatalogAdmin = await catalog.canCurrentUserWrite(url);
```

`SPFxAppCatalogPageContextLike` requires `{ web: { absoluteUrl: string } }` for the web used to query `/_api/SP_TenantSettings_Current`. `SPFxAppCatalogService` returns `discoverUrl(): Promise<string>` and `canCurrentUserWrite(catalogUrl: string): Promise<boolean>`. Construction performs no request; missing client or page data is tolerated until discovery, which rejects. Failed requests, response parsing and an empty `CorporateCatalogUrl` reject and are not cached. Successful discovery caches only the URL per service instance; concurrent first calls are not deduplicated by this service.

The write method checks only `/_api/web/currentuser?$select=IsSiteAdmin` at the supplied catalog. It does not check fine-grained list grants, item restrictions or schema permissions. False includes missing client, non-success HTTP and thrown lookup/parsing errors. SharePoint still enforces actual permissions; use the relevant store's semantics for storage operations. The service does not provision a catalog or mutate tenant settings.

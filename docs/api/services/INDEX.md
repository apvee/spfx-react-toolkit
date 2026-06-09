# Services API Reference

> Non-React services for composing SPFx, PnPjs, SharePoint REST, and Microsoft Graph operations outside hooks.

Services do not read React context and do not manage React state. Callers provide initialized SPFx dependencies such as `SPFI`, `SPHttpClient`, `MSGraphClientV3`, `PageContext`, or SPFx context.

## Services

| Service factory | Purpose |
|-----------------|---------|
| `createSPFxPnPContextService` | Resolve SharePoint site URLs and create configured PnPjs `SPFI` instances. |
| `createSPFxPnPService` | Run generic `invoke` and `batch` operations against an `SPFI` instance. |
| `createSPFxPnPListService` | Query SharePoint lists and perform CRUD/batch operations. |
| `createSPFxPnPSearchService` | Run SharePoint Search and suggestions through PnPjs. |
| `createSPFxAppCatalogService` | Discover the tenant app catalog URL and check write permission. |
| `createSPFxTenantPropertyService` | Read SharePoint StorageEntity tenant properties. |
| `createSPFxTenantKeyValueStoreService` | Provision and use the hidden tenant key-value store list. |
| `createSPFxOneDriveAppDataService` | Read and write JSON files in the Graph OneDrive app folder. |
| `createSPFxUserPhotoService` | Retrieve user photo blobs from Microsoft Graph. |

## Usage Pattern

Hooks remain the recommended React API. Use services when composing operations outside React hooks or when building custom abstractions.

```ts
const list = createSPFxPnPListService<Task>(sp, 'Tasks', 50);
const result = await list.query(q => q.select('Id', 'Title'));

const store = createSPFxTenantKeyValueStoreService(spHttpClient);
await store.save('featureFlags', { enabled: true }, appCatalogUrl);
```

## Boundaries

- Services may perform I/O.
- Services do not call hooks.
- Services do not depend on Jotai or React lifecycle.
- Services do not replace providers; callers are responsible for obtaining SPFx dependencies.

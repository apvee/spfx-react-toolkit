# Tenant Property Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

```ts
import { createSPFxTenantPropertyService } from '@apvee/spfx-react-toolkit/services';
```


## `createSPFxTenantPropertyService`

```ts
function createSPFxTenantPropertyService(
  spHttpClient: SPHttpClient
): SPFxTenantPropertyService;
```

Required public types:

- `SPFxTenantProperty`
- `SPFxTenantPropertyService`

Returned method:

| Method | Description |
|--------|-------------|
| `get<T>(key, catalogUrl)` | Reads a SharePoint StorageEntity from the tenant app catalog |

The service deserializes the stored value with `deserializeTenantValue`. It throws when the SharePoint REST request fails.

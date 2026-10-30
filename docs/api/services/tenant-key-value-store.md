# Tenant Key-Value Store Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

```ts
import { createSPFxTenantKeyValueStoreService } from '@apvee/spfx-react-toolkit/services';
```


## `createSPFxTenantKeyValueStoreService`

```ts
function createSPFxTenantKeyValueStoreService(
  spHttpClient: SPHttpClient
): SPFxTenantKeyValueStoreService;
```

Required public types:

- `SPFxTenantKeyValueStoreServiceItem`
- `SPFxTenantKeyValueStoreService`

Returned methods:

| Method | Description |
|--------|-------------|
| `ensureListReady(catalogUrl)` | Ensures the hidden `TenantKeyValueStore` list and required fields exist |
| `get<T>(key, catalogUrl)` | Reads one item by key |
| `list(catalogUrl)` | Lists stored key-value items ordered by key |
| `save<T>(key, value, catalogUrl, description?)` | Creates or updates a key |
| `remove(key, catalogUrl)` | Deletes a key when it exists |

`SPFxTenantKeyValueStoreServiceItem<T>` contains `key`, `value`, `description`, and `id`.

Provisioning is guarded per catalog URL so concurrent calls share the same in-flight setup promise. Values are serialized with `serializeTenantValue` and read with `deserializeTenantValue`. The [legacy tenant serialization limits](../helpers/tenant-value.md#deserializetenantvalue) apply; the generic return type does not validate stored values. Actual catalog/list permissions determine whether requests succeed.

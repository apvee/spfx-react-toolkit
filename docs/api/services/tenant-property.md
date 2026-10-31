# Tenant Property Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

Import this factory and its public types from `@apvee/spfx-react-toolkit/services` or the package root. Callers supply the initialized client/context; these services require no React provider.

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

## Usage and public contracts

```ts
const service = createSPFxTenantPropertyService(context.spHttpClient);
const property = await service.get<unknown>('my-app:featureEnabled', catalogUrl);
if (typeof property.value === 'boolean') console.log(property.value);
```

`SPFxTenantPropertyService.get<T = unknown>(key: string, catalogUrl: string)` returns `Promise<SPFxTenantProperty<T>>`. `SPFxTenantProperty<T>` has readonly `value: T | undefined` and `description: string | undefined`. A missing/legacy REST-null `Value` returns both fields undefined; a supplied value is parsed JSON-first with [legacy tenant-value helpers](../helpers/tenant-value.md), and its description is passed through. The generic type performs no runtime validation or Date reconstruction.

Both strings are required; the key is OData-escaped and URL-encoded for `GetStorageEntity`. The caller supplies the actual tenant catalog URL, not the current site URL. Construction performs no request; there is no discovery, cache, provisioning or write method. Every unsuccessful HTTP response, including 404/403, rejects rather than silently returning an absent property; response parsing failures reject too.

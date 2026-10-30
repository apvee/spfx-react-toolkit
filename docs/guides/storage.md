# Choose a storage scope

[Documentation home](../INDEX.md) · [Storage hook index](../api/INDEX.md#storage-hooks)

Choose storage by lifetime, sharing and permissions before selecting an API:

| Need | React API | Non-React option |
| --- | --- | --- |
| Browser persistence scoped to one component instance | [useSPFxLocalStorage](../api/hooks/storage/local-storage.md) | Host browser storage with caller-owned namespacing |
| Browser-session persistence scoped to one instance | [useSPFxSessionStorage](../api/hooks/storage/session-storage.md) | Host session storage with caller-owned namespacing |
| JSON files in the current user's OneDrive app root | [useSPFxOneDriveAppData](../api/hooks/storage/onedrive-app-data.md) | [OneDrive service](../api/services/onedrive-app-data.md) |
| Read tenant storage entities | [useSPFxTenantProperty](../api/hooks/storage/tenant-property.md) | [Tenant property service](../api/services/tenant-property.md) |
| Read/write settings shared through the tenant app catalog | [useSPFxTenantKeyValueStore](../api/hooks/storage/tenant-key-value-store.md) | [Tenant store service](../api/services/tenant-key-value-store.md) |
| Tenant defaults with optional site overrides and explicit write scope | [useSPFxConfigurationStore](../api/hooks/storage/configuration-store.md) | [Configuration service](../api/services/configuration-store.md) |
| Read/write settings shared by subsites in one site collection | [useSPFxSiteKeyValueStore](../api/hooks/storage/site-key-value-store.md) | [Site store service](../api/services/site-key-value-store.md) |

## Browser and remote storage

Browser storage is instance-namespaced and best effort. Its defaults, removal and storage-event behavior are documented separately for local and session storage. It is not a shared cross-component synchronization bus.

Remote hooks need a matching provider and the relevant authenticated client. OneDrive operations require the documented Graph scopes. Tenant data uses the tenant app catalog; site data uses `pageContext.site.absoluteUrl`, the collection root, rather than the current subsite. Actual SharePoint/Graph grants authorize requests; permission indicators do not grant write access.

## Plan reads, writes and failures

Site-store mount, reads and removal do not provision a list; save or explicit service setup can provision it. Hook read failures return their documented fallback and publish error state; write rejections must be caught. Standalone services reject failures. Follow each reference for lifecycle ownership, pagination, schema enforcement and conflict behavior rather than assuming that every storage API has the same error contract.

Generic TypeScript parameters do not validate remote data. New list-store writes preserve text values and full bigint precision using explicit stored metadata; dates and historical unmarked values retain documented read limitations. Read the [serialization contract](../api/helpers/tenant-value.md) before choosing a format. Tenant/site writes use observed ETags and expose stale-update conflicts; retry after reading current data rather than assuming wildcard overwrites. A plain tenant/site store has one scope. The configuration API selects a site override when present and otherwise a tenant default; it does not deep-merge stored values.

Use the [Graph and storage checklist](../development/sharepoint-validation/checklists/graph-storage-permissions.md) and [site collection checklist](../development/sharepoint-validation/checklists/site-key-value-store.md) to record real-host behavior separately from local checks.

## Layered configuration

Choose the configuration hook/service when the same key can have tenant defaults and collection-specific overrides. Reads discover stores lazily without creating missing lists. Returned metadata identifies the winning scope; layered mode follows the documented fallback, while tenant/site mode restricts the target. Save and remove require an explicit scope, so a caller cannot accidentally write to whichever scope happened to satisfy a read. Removing an override reveals a tenant default on a later read; it does not delete that default.

Use the [configuration reference](../api/hooks/storage/configuration-store.md) for actual options, discovery, pending/error state, ownership and nullish-value behavior. Validate real grants and cross-instance visibility with the [configuration host checklist](../development/sharepoint-validation/checklists/configuration-store.md).

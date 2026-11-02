# Choose a storage scope

Choose storage by lifetime, sharing and permissions, then define read, write and failure ownership.

[Documentation home](../INDEX.md) · [Storage hook index](../../api/INDEX.md#references)

Choose storage by lifetime, sharing and permissions before selecting an API:

| Need | React API | Non-React option |
| --- | --- | --- |
| Browser persistence scoped to one component instance | [useSPFxLocalStorage](../../api/hooks/storage/local-storage.md) | Host browser storage with caller-owned namespacing |
| Browser-session persistence scoped to one instance | [useSPFxSessionStorage](../../api/hooks/storage/session-storage.md) | Host session storage with caller-owned namespacing |
| JSON files in the current user's OneDrive app root | [useSPFxOneDriveAppData](../../api/hooks/storage/onedrive-app-data.md) | [OneDrive service](../../api/services/onedrive-app-data.md) |
| Read tenant storage entities | [useSPFxTenantProperty](../../api/hooks/storage/tenant-property.md) | [Tenant property service](../../api/services/tenant-property.md) |
| Read/write settings shared through the tenant app catalog | [useSPFxTenantKeyValueStore](../../api/hooks/storage/tenant-key-value-store.md) | [Tenant store service](../../api/services/tenant-key-value-store.md) |
| Tenant defaults with optional site overrides and explicit write scope | [useSPFxConfigurationStore](../../api/hooks/storage/configuration-store.md) | [Configuration service](../../api/services/configuration-store.md) |
| Read/write settings shared by subsites in one site collection | [useSPFxSiteKeyValueStore](../../api/hooks/storage/site-key-value-store.md) | [Site store service](../../api/services/site-key-value-store.md) |

## Browser and remote storage

Browser storage is instance-namespaced and best effort. Its defaults, removal and storage-event behavior are documented separately for local and session storage. It is not a shared cross-component synchronization bus.

Remote hooks need a matching provider and the relevant authenticated client. OneDrive operations require the documented Graph scopes. Tenant data uses the tenant app catalog; site data uses `pageContext.site.absoluteUrl`, the collection root, rather than the current subsite. Actual SharePoint/Graph grants authorize requests; permission indicators do not grant write access.

## Plan reads, writes and failures

Site-store mount, reads and removal do not provision a list; save or explicit service setup can provision it. Hook read failures return their documented fallback and publish error state; write rejections must be caught. Standalone services reject failures. Follow each reference for lifecycle ownership, pagination, schema enforcement and conflict behavior rather than assuming that every storage API has the same error contract.

Generic TypeScript parameters do not validate remote data. New list-store writes preserve text values and full bigint precision using explicit stored metadata; dates and historical unmarked values retain documented read limitations. Read the [serialization contract](../../api/helpers/tenant-value.md) before choosing a format. Tenant/site writes use observed ETags and expose stale-update conflicts; retry after reading current data rather than assuming wildcard overwrites. A plain tenant/site store has one scope. The configuration API selects a site override when present and otherwise a tenant default; it does not deep-merge stored values.

Use the [Graph and storage checklist](../development/sharepoint-validation/INDEX.md#permissions-denial-and-absence) and [site collection checklist](../development/sharepoint-validation/INDEX.md#storage-destinations-and-persistence) to record real-host behavior separately from local checks.

## Layered configuration

Choose the configuration hook/service when the same key can have tenant defaults and collection-specific overrides. Reads discover stores lazily without creating missing lists. Returned metadata identifies the winning scope; layered mode follows the documented fallback, while tenant/site mode restricts the target. Save and remove require an explicit scope, so a caller cannot accidentally write to whichever scope happened to satisfy a read. Removing an override reveals a tenant default on a later read; it does not delete that default.

Use the [configuration reference](../../api/hooks/storage/configuration-store.md) for actual options, discovery, pending/error state, ownership and nullish-value behavior. Validate real grants and cross-instance visibility with the [configuration host checklist](../development/sharepoint-validation/INDEX.md#storage-destinations-and-persistence).

<!-- toolkit:include symbol="createSPFxTenantPropertyService" section="example" title="Read and validate an unknown StorageEntity" -->

<!-- toolkit:include symbol="createSPFxTenantKeyValueStoreService" section="example" title="Preserve numeric-looking text in a tenant list" -->

<!-- toolkit:include symbol="createSPFxSiteKeyValueStoreService" section="example" title="Read and conditionally save at collection root" -->

<!-- toolkit:include symbol="createSPFxConfigurationStoreService" section="example" title="Tenant default site override inventory and removal" -->

<!-- toolkit:include symbol="createSPFxOneDriveAppDataService" section="example" title="Read and replace app-folder JSON" -->

<!-- toolkit:include symbol="serializeTenantValue" section="example" title="Compare legacy serialization with validated unknown data" -->

<!-- toolkit:include symbol="useSPFxLocalStorage" section="example" title="Edit local preferences" -->

<!-- toolkit:include symbol="useSPFxLocalStorage" section="example" title="Keep and clear the last ten searches" -->

<!-- toolkit:include symbol="useSPFxSessionStorage" section="example" title="Save a session draft before clearing it" -->

<!-- toolkit:include symbol="useSPFxSessionStorage" section="example" title="Resume and complete a session wizard" -->

<!-- toolkit:include symbol="useSPFxConfigurationStore" section="example" title="Read a site override then reveal its tenant default" -->

<!-- toolkit:include symbol="useSPFxTenantKeyValueStore" section="example" title="Save and read numeric-looking text" -->

<!-- toolkit:include symbol="useSPFxSiteKeyValueStore" section="example" title="Read and save with separate hook errors" -->

<!-- toolkit:include symbol="useSPFxOneDriveAppData" section="example" title="Create preferences only for a missing app file" -->

<!-- toolkit:include symbol="useSPFxOneDriveAppData" section="example" title="Load an app-folder cache on demand" -->

<!-- toolkit:include symbol="useSPFxOneDriveAppData" section="example" title="Add toggle and remove persisted todos" -->

<!-- toolkit:include symbol="useSPFxOneDriveAppData" section="example" title="Edit app-folder configuration" -->

## Explicit Graph workflows and local reset

The two complete host scenarios below keep Graph initialization behind Start and business requests behind Run. They require the matching SPFx provider and approved endpoint grants. The photo requests the default 240x240 bytes and displays a 96px image or named fallback. OneDrive uses the fixed demo app-folder file, with no automatic fetch or missing-file provisioning; only its explicit Write replaces content. Its generic JSON type does not validate arbitrary stored shapes.

Reset preserves the operational child, original hook callbacks and genuine pending locks. It restores local drafts and hides stale publication; it does not clear hook state, cancel requests, revoke a still-owned photo URL or undo accepted OneDrive content. A current photo request may finish with a hidden owned URL after Reset, released by normal hook cleanup. Ordinary feature navigation can leave pending work, with unmount suppressing late UI publication. Real writes/cleanup require an authorized destination; metadata/download failures are not confirmed absence.

<!-- toolkit:include symbol="useSPFxUserPhoto" section="example" title="Read the current user photo with managed URL cleanup" -->

<!-- toolkit:include symbol="useSPFxOneDriveAppData" section="example" title="Read and explicitly replace OneDrive app data" -->

## Stable browser names

Keep the instance ID and unscoped key stable to retrieve the same browser entry. {@link createScopedSPFxStorageKey} creates a name such as `spfx:webpart-42:filters`; it does not encrypt data or synchronize components. The browser hooks own initialization, JSON conversion, storage events and best-effort writes. Local and session persistence have different lifetimes even when their APIs look similar.

A session draft should clear only after the caller's save succeeds. Wizard step completion, navigation eligibility and accumulated data are application state; the complete recipe above persists those choices and exposes the completion action. A local search history can deduplicate the current term, retain the last ten searches and clear deliberately, without implying a backend search or cross-instance bus.

## Historical helper encodings

The standalone {@link serializeTenantValue} and {@link deserializeTenantValue} helpers preserve their historical JSON-first interpretation. New tenant/site list writers deliberately quote JSON-like text so `'123'` stays text. Do not pre-serialize with a legacy helper before passing a value to a new list writer. Existing raw records need no automatic migration, but previously lost numeric precision cannot be reconstructed. Date reads remain ISO text; bigint reads as exact decimal text. Use `unknown` and validate the application shape rather than treating a generic type as validation.

{@link escapeODataValue} escapes a string literal's quote content; it supplies neither surrounding quotes nor URL encoding. The combined source example keeps escaping, serialization, parsing and validation separate.

## StorageEntities and list stores

Choose {@link createSPFxTenantPropertyService} to read existing StorageEntities, supplying the actual catalog URL. It does not discover that URL or add writes. Choose the tenant list store for application-managed shared settings. A current web URL is not an interchangeable catalog address, and a failed request is not evidence that a property is absent.

## Historical data and tenant-list conflicts

Both list stores preserve historical values while using explicit encoding for new writes. The tenant store preserves a legacy REST-null Note as legacy `null`; the site store maps a nullable empty Note to an empty string. Stored text `'null'` still reads as legacy `null`. This difference is an interoperability constraint, not a new absence default.

Provisioning/readiness is per service and normalized URL, while values and permission decisions remain operation-owned. SharePoint key equality follows its field/unique-column comparison, including case-insensitive matching while preserving returned casing. Failed uniqueness/schema validation must not delete data to make the store usable. See [Microsoft's unique-column comparison rules](https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ee536168%28v%3Doffice.14%29).

Observed ETags protect the service's snapshot-to-mutation interval. They do not protect a consumer's earlier `get`-to-`save` interval. Design editing conflicts accordingly: inspect current data before retrying, and do not interpret bounded creation-collision recovery as a transaction or blind overwrite. Missing/unsafe ETags and HTTP 412 remain failures. Catalog site-administrator status is not a tenant-store permission indicator or a substitute for actual list grants.

## Collection root and conditional mutations

Use `pageContext.site.absoluteUrl` for collection-wide settings shared across subsites. A current subsite has a different lifetime/scope from the collection root; user-specific data belongs in browser or OneDrive storage instead. Read and permission probes do not provision a list. Save or explicit setup may provision, requiring the actual schema-management rights in addition to item writes.

The complete service recipe reads a boolean, checks `canCurrentUserWrite`, then awaits save at the same root. Treat the indicator as UI guidance: actual SharePoint grants and item restrictions still decide the mutation. The React hook captures read failures in state and returns its documented fallback; the standalone service rejects. Distinguish an absent key from that fallback/error before telling a reader a setting does not exist.

## Explicit collection-root exploration

The Site hook scenario requires explicit activation before its permission initialization. Read/List never provision; Save and confirmed Remove require consent to the displayed disposable key and exact value/description policy. Text, JSON, bigint and Date inputs expose their different read-back types; omitting a description preserves existing metadata while explicit empty text clears it. A failed fallback-returning hook read remains an error, not absence or an empty inventory. Can Write is advisory for the hook.

The separate service scenario explicitly reads the root record and observes effective permissions before its optional consented boolean save. It deliberately requires a current true indicator, which does not guarantee later grants/schema success. Neither normal mutation automatically reads back. Observed ETags guard the service snapshot rather than an earlier UI Read.

Reset preserves the outer Demo and actual pending-write ledger, hides results/consent, and deactivates the hook session. Reactivation is explicit and waits for the original public write promise to settle. Internal service requests may continue after Reset; neither Reset nor navigation cancels or rolls back server work. The lock belongs to that still-mounted Demo and does not serialize another scenario. Real tenant writes and cleanup require separate authorization.

<!-- toolkit:include symbol="useSPFxSiteKeyValueStore" section="example" title="Explore collection root records with guarded mutations" -->

<!-- toolkit:include symbol="createSPFxSiteKeyValueStoreService" section="example" title="Read a collection root record before a consented save" -->

## Layered whole-record configuration

A site record replaces the **whole** tenant record for its key. False, zero, empty text and legacy `null` remain present overrides; fallback occurs only for confirmed site absence. There is no JSON-property merge and a failed site read must not silently turn into a tenant default. Physical inventory identifies scope-local IDs and rows, rather than an effective merged list.

The complete service recipe writes a tenant default, writes a described site override, reads the winning record, lists physical site rows, then removes only that override. A later read can reveal the tenant default. Omitted descriptions preserve destination metadata; explicit empty text clears it. Select the mutation scope explicitly even in a single-store mode.

The hook takes mode-only configuration from its provider; the standalone service needs actual discovery-web/collection URLs. Discovery is lazy, and a layered site hit need not access the tenant catalog. This isolation is useful where site readers lack catalog rights. It does not add a background read, cross-store transaction, global cache or authorization layer. See {@link createSPFxConfigurationStoreService} and {@link useSPFxConfigurationStore} for their distinct operational contracts.

## Temporary browser download URLs

Browser reads follow [Microsoft's download guidance](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0#downloading-files-in-javascript-apps): obtain complete Graph metadata, then immediately fetch the preauthenticated HTTPS download URL without Graph authorization headers or cookies. An authenticated `/content` redirect can fail browser CORS preflight. Do not project away the download annotation, retain its temporary URL, log it or forward credentials to the download host.

Only a recognized metadata not-found result is absence. An expired/failed download, including a download 404, remains failure. Treating it as absence could make a create-if-missing policy overwrite an existing file. The complete service example reads/isNotFound and replaces JSON; the hook examples separately exercise existing configuration, namespaced lazy reads, missing-file preferences and todo edits. Those examples require approved app-folder grants, JSON-compatible application data and real initialized clients; they do not create folders or validate arbitrary stored schemas.

## Picker references and storage

A saved picker reference is identity data, not a durable download token or a guaranteed current display model. Persist the complete accepted JSON array, including unresolved references, and resolve display separately. Do not replace it with the visible/resolved subset or save personal access tokens/history through a picker. See [reference selection](pickers.md#selecting-sharepoint-values) and the canonical picker contracts for source identity and stale-reference behavior.

<!-- toolkit:include symbol="useSPFxTenantProperty" section="example" title="Compose feature flags endpoint configuration and branding" -->

## Run read-only tenant property examples

Select the tenant-property hook example and explicitly activate its manual session before Run. Activation makes no HTTP request; Run loads both configured keys concurrently through independent catalog caches. Reset or a changed host/web identity requires explicit activation again and hides obsolete results. Already-started logical reads may still finish; current read errors take precedence over retained values.

<!-- toolkit:include symbol="useSPFxTenantProperty" section="example" title="Read tenant version and counter with explicit activation" -->

For an already known tenant catalog, use the direct factory example. Supply the actual catalog URL; the factory adds no discovery or writes. It reads unknown data, validates a boolean and distinguishes absence, a wrong decoded type and a failed HTTP/response parse.

<!-- toolkit:include symbol="createSPFxTenantPropertyService" section="example" title="Read and validate a catalog StorageEntity" -->

## Explore tenant list records safely

Activate the tenant-store hook session explicitly before its initial catalog discovery and effective permission checks. Client readiness does not mean discovery has completed or that a write is authorized. The logical scope is the tenant app-catalog store resolved internally for the captured discovery-web, collection, client, PageContext and host session. The hook does not expose its physical catalog URL; the discovery-web address is not a catalog address. A changed session requires explicit reactivation.

Use disposable keys with a nonempty `spfx-toolkit-tenant-demo-` suffix and consent to the displayed exact operation, value and description policy. Ordinary Save and Remove have no automatic readback. The separate numeric-text recipe saves `'123'` with `'Application version'`, then reads explicitly; a failed read after a completed save is partial success. The direct service recipe uses only an explicitly supplied known catalog URL and has no discovery or permission API.

Reset restores local UI and consent, retaining the original pending public-write lock even when its hook child unmounts. Already-dispatched setup, mutation and recovery requests may finish; Reset does not cancel, roll back or remove a stored record. No stale sequence may begin a later public read. Locks belong to each still-mounted Demo and do not serialize other scenarios.

<!-- toolkit:include symbol="useSPFxTenantKeyValueStore" section="example" title="Explore tenant records and exact text persistence" -->

<!-- toolkit:include symbol="createSPFxTenantKeyValueStoreService" section="example" title="Round trip numeric looking text through a tenant service" -->

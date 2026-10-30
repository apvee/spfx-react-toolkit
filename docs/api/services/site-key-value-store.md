# Site Key-Value Store Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

```ts
import { createSPFxSiteKeyValueStoreService } from '@apvee/spfx-react-toolkit/services';
```


## `createSPFxSiteKeyValueStoreService`

```ts
function createSPFxSiteKeyValueStoreService(
  spHttpClient: SPHttpClient
): SPFxSiteKeyValueStoreService;
```

Public item and service contracts:

```ts
interface SPFxSiteKeyValueStoreServiceItem<T = unknown> {
  readonly key: string;
  readonly value: T;
  readonly description: string | undefined;
  readonly id: number;
}

interface SPFxSiteKeyValueStoreService {
  ensureListReady: (siteCollectionUrl: string) => Promise<void>;
  get: <T = unknown>(key: string, siteCollectionUrl: string) => Promise<SPFxSiteKeyValueStoreServiceItem<T> | undefined>;
  list: (siteCollectionUrl: string) => Promise<SPFxSiteKeyValueStoreServiceItem<unknown>[]>;
  save: <T = unknown>(key: string, value: T, siteCollectionUrl: string, description?: string) => Promise<void>;
  remove: (key: string, siteCollectionUrl: string) => Promise<void>;
  canCurrentUserWrite: (siteCollectionUrl: string) => Promise<boolean>;
}
```

| Method | Description |
|--------|-------------|
| `ensureListReady(siteCollectionUrl)` | Provisions or repairs the hidden root-web list and required schema |
| `get<T>(key, siteCollectionUrl)` | Reads one exact key without provisioning; returns `undefined` for confirmed absent list/key |
| `list(siteCollectionUrl)` | Reads every page in server Title order; returns `[]` for a confirmed absent list |
| `save<T>(key, value, siteCollectionUrl, description?)` | Provisions if necessary, then creates or updates the key |
| `remove(key, siteCollectionUrl)` | Deletes an existing key without provisioning; confirmed absence is a no-op |
| `canCurrentUserWrite(siteCollectionUrl)` | Checks effective permissions; returns `false` when lookup fails |

Use the collection root URL from `pageContext.site.absoluteUrl`. Empty collection URLs reject before any HTTP request; their permission check returns `false`. All subsites share one `SiteKeyValueStore` there. The service does not discover an app catalog or use the current subsite URL. Standalone usage needs no React provider or hooks:

```ts
import { createSPFxSiteKeyValueStoreService } from '@apvee/spfx-react-toolkit';

const store = createSPFxSiteKeyValueStoreService(context.spHttpClient);
const rootUrl = context.pageContext.site.absoluteUrl;
const setting = await store.get<boolean>('featureEnabled', rootUrl);
if (await store.canCurrentUserWrite(rootUrl)) {
  await store.save('featureEnabled', true, rootUrl, 'Shared across subsites');
}
```

Only a list-level 404 confirms an absent store. Failed field, item, and later-page requests reject, including their 404 responses. Reads verify required field types without making schema changes. Pagination accepts both OData continuation properties and same-list relative/absolute links; invalid, foreign, and repeated links reject instead of returning partial results.

Provisioning uses a generic list (`BaseTemplate: 100`, `Hidden: true`, `NoCrawl: true`), indexed unique `Title`, and multiline text (`FieldTypeKind: 3`) `Value`/`Description` fields. Incompatible fields or a failed uniqueness constraint reject without deleting data. Setup promises and readiness are isolated per service and URL after trimming whitespace/trailing slashes; failed setup can retry. Values and permissions are not cached.

Existing-list write capability requires effective Add, Edit, and Delete Items grants on that list, including its unique permissions. An absent list additionally requires Manage Lists on the root web. This indicator does not authorize operations: SharePoint still enforces actual grants, schema-management rights, and item restrictions.

Key equality follows SharePoint's field comparison and unique-column semantics, including case-insensitive matching; reads preserve stored Title casing. See [Microsoft's unique-column comparison rules](https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ee536168%28v%3Doffice.14%29). Updates use wildcard ETags and last-writer-wins semantics. Omitting the update description preserves it; an explicit empty string clears it. Creation recovery is bounded and requires a verified competing resource; 401/403 never recover. An item collision can attempt one MERGE only after confirming a unique schema and exact key. A delete 404 succeeds only after verifying absence.

The [legacy tenant serialization limits](../helpers/tenant-value.md#deserializetenantvalue) also apply here. Generic types do not validate runtime values. JSON-like strings can deserialize as numbers/booleans, dates remain strings, and bigint does not gain a lossless round trip. SharePoint's nullable empty Note values read as empty strings; a saved `null` remains the serialized string `'null'` and reads as `null`.

The React [site store hook](../hooks/storage/site-key-value-store.md#usespfxsitekeyvaluestore) captures read failures in state and returns fallbacks; the standalone service rejects failed reads and writes. Only `save` and `ensureListReady` perform setup; `get`, `list`, `remove` and permission checks do not provision.

[View service source](../../../packages/spfx-react-toolkit/src/services/spfx-site-key-value-store.service.ts)

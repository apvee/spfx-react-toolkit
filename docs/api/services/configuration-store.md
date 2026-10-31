# Configuration Store Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

Import this factory and its public types from `@apvee/spfx-react-toolkit/services` or the package root. Callers supply the initialized client/context; these services require no React provider.

`createSPFxConfigurationStoreService(client, options)` composes the existing tenant/site list services without React or a new persistence format. Import from `@apvee/spfx-react-toolkit/services` or root. Construction makes no request. Required options are discriminated by mode:

```ts
function createSPFxConfigurationStoreService(
  client: SPHttpClient,
  options: SPFxConfigurationStoreServiceOptions
): SPFxConfigurationStoreService;
```

```ts
type SPFxConfigurationStoreServiceOptions =
  | { readonly mode: 'tenant'; readonly discoveryWebUrl: string; readonly siteCollectionUrl?: never }
  | { readonly mode: 'site'; readonly siteCollectionUrl: string; readonly discoveryWebUrl?: never }
  | { readonly mode: 'layered'; readonly discoveryWebUrl: string; readonly siteCollectionUrl: string };
```

`discoveryWebUrl` is the current web from which to discover the tenant catalog, not the catalog URL itself. `siteCollectionUrl` must be the collection root. Required URLs are copied, trimmed and stripped of trailing slashes; empty/invalid mode dependencies throw synchronously. There are no per-operation URL overrides. Only enabled scopes may be read/listed/mutated.

Canonical public types are exported by this service module and re-exported by the hook module:

| Type | Contract |
| --- | --- |
| `SPFxConfigurationStoreScope` | Physical `tenant` or `site` destination. |
| `SPFxConfigurationStoreMode` | Required `tenant`, `site` or `layered` read policy and scope restriction. |
| `SPFxConfigurationStoreOptions` | Hook input with required `mode`; the provider supplies URLs/client. |
| `SPFxConfigurationStoreTarget` | Required `scope` for one removal; no default or inferred destination. |
| `SPFxConfigurationStoreSaveOptions` | Inherits required target `scope`; optional `description` preserves destination metadata when omitted and clears it when empty. |
| `SPFxConfigurationStoreItem<T = unknown>` | Readonly `key`, `value: T`, `description: string \| undefined`, source-list `id` and physical `source`. T performs no runtime validation. |
| `SPFxConfigurationStoreService` | The four Promise-based operations documented below; failures reject. |

```ts
const store = createSPFxConfigurationStoreService(context.spHttpClient, {
  mode: 'layered',
  discoveryWebUrl: context.pageContext.web.absoluteUrl,
  siteCollectionUrl: context.pageContext.site.absoluteUrl,
});
await store.save('my-app:pageSize', 20, { scope: 'tenant' });
await store.save('my-app:pageSize', 50, { scope: 'site', description: 'Local override' });
const item = await store.get<number>('my-app:pageSize'); // value 50, source site
const siteRows = await store.listScope('site'); // physical inventory, not effective merge
await store.remove('my-app:pageSize', { scope: 'site' });
```

| Method | Observable behavior |
| --- | --- |
| `get<T>(key)` | Whole winning record plus `source`, or undefined for confirmed absence. Layered reads site first, then tenant only on absence. Site errors never enable fallback. |
| `listScope(scope)` | Complete physical inventory with provenance. Preserves stored casing, order and legacy duplicates; page failures reject without partial data. |
| `save<T>(key,value,{scope,description?})` | Explicit single destination; whole-value replacement. Only this operation may provision the chosen list. |
| `remove(key,{scope})` | Explicit single destination, no provisioning. Site removal reveals tenant data on the next layered get. |

All operation failures reject unchanged. Invalid scope/key/description arguments fail before I/O. Stored keys are forwarded without normalization, and the underlying server determines equality. No effective merged list, JSON-property merge, permission indicator, automatic read, global cache or multi-store transaction is provided. Item fields `key`, `value`, `description`, `id` and `source` describe one physical record; IDs are not globally unique. False, zero, empty text and legacy null values remain present overrides.

Tenant discovery is lazy, uses a frozen scalar discovery-web URL and shares its in-flight Promise within the instance. Successful URLs use the existing catalog cache; a failed discovery can retry. Site operations, including a layered site hit, require no tenant discovery/access. Existing serialization, destination-local omitted/empty descriptions, schema enforcement and observed ETag conflicts are inherited unchanged. Serialization can fail after tenant discovery but before setup/item writes. See the [hook contract](../hooks/storage/configuration-store.md#usespfxconfigurationstore) for all types, examples, errors and compatibility limits.

[View source](../../../packages/spfx-react-toolkit/src/services/spfx-configuration-store.service.ts)

## Complete operation signatures

```ts
interface SPFxConfigurationStoreService {
  get<T = unknown>(key: string): Promise<SPFxConfigurationStoreItem<T> | undefined>;
  listScope(scope: SPFxConfigurationStoreScope): Promise<SPFxConfigurationStoreItem<unknown>[]>;
  save<T = unknown>(key: string, value: T, options: SPFxConfigurationStoreSaveOptions): Promise<void>;
  remove(key: string, target: SPFxConfigurationStoreTarget): Promise<void>;
}
```

`SPFxConfigurationStoreServiceOptions` is the required discriminated URL/options union above. `SPFxConfigurationStoreOptions` is `{ readonly mode: SPFxConfigurationStoreMode }` for the corresponding React hook only; it is not sufficient as standalone service input. `SPFxConfigurationStoreTarget` is `{ readonly scope: SPFxConfigurationStoreScope }`; save extends it with optional readonly `description?: string`. There is no default mode or mutation scope, even for a single-store instance.

The client is caller-owned and required; the factory does not acquire it. URL checks ensure required values are strings and nonempty after trimming/trailing-slash removal, not that they are valid HTTP URLs or actual collection roots. Supply correct absolute SharePoint URLs. Key validation requires a string but permits empty strings/whitespace; it does not trim or namespace keys. Disabled/invalid scopes and nonstring descriptions reject asynchronously before discovery/storage I/O. No requested scope is inferred from a prior record's source.

`get` never provisions either store; a false/zero/empty string or legacy null-valued record still blocks fallback. `listScope` intentionally returns physical rows, including legacy duplicates; it is not a merged/effective configuration list. Save/remove target one physical destination and inherit the [tenant](tenant-key-value-store.md) or [site](site-key-value-store.md) service's serialization/schema/observed-ETag behavior. The service adds no compare-and-swap against a previous caller read, transaction, cross-scope rollback, retry, permission probe, background fetch or automatic reread after a mutation.

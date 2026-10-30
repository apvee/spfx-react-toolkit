# Compose helpers and services outside React

[Documentation home](../INDEX.md) · [Helpers](../api/INDEX.md#helpers-and-styles) · [Services](../api/INDEX.md#non-react-services)

Helpers receive explicit inputs, do not read React context or manage provider state, and do not perform I/O. Token-evaluation helpers also read the clock; shared theme objects must be treated as immutable. The styles facade additionally exposes the React `useSx` hook, which composes descriptors through the host renderer.

Services are non-React factories. Callers provide initialized dependencies such as `SPFI`, `SPHttpClient`, `MSGraphClientV3`, page context or SPFx context data. Use services in command handlers, utility modules, tests or other code outside a React component; hooks remain the usual React-facing API.

| API type | Use when | Dependency |
|----------|----------|------------|
| Hooks | A React component is wrapped by an SPFx provider | React provider context |
| Helpers | You need pure mapping, formatting, or conversion logic | Direct input values |
| Services | You need reusable non-React I/O operations | SPFx clients, PnPjs, or Graph clients |

## Helper example

```ts
import {
  createScopedSPFxStorageKey,
  getSPFxContainerSize,
  getSPFxUserInfo,
} from '@apvee/spfx-react-toolkit';

const user = getSPFxUserInfo(pageContext);
const scopedKey = createScopedSPFxStorageKey(instanceId, 'filters');
const size = getSPFxContainerSize(720);
```

```ts
import {
  buildOneDriveAppDataPath,
  buildUserPhotoEndpoint,
} from '@apvee/spfx-react-toolkit';

const appDataPath = buildOneDriveAppDataPath('settings.json', 'dashboard');
const photoPath = buildUserPhotoEndpoint({ email: 'user@contoso.com', size: '240x240' });
```



## Dependency matrix

| Factory | Required dependency |
|---------|---------------------|
| `createSPFxPnPContextService` | `ISPFXContext`, page context-like object |
| `createSPFxPnPService` | `SPFI` |
| `createSPFxPnPListService` | `SPFI`, title string or explicit list selector |
| `createSPFxPnPSearchService` | `SPFI` |
| `createSPFxAppCatalogService` | `SPHttpClient`, page context-like object |
| `createSPFxTenantPropertyService` | `SPHttpClient` |
| `createSPFxTenantKeyValueStoreService` | `SPHttpClient` |
| `createSPFxSiteKeyValueStoreService` | `SPHttpClient` |
| `createSPFxOneDriveAppDataService` | `MSGraphClientV3` |
| `createSPFxUserPhotoService` | `MSGraphClientV3` |
| `createSPFxApiPermissionPrecheckService` | `SPFxAadTokenProviderLike` |

## Service examples

```ts
import {
  createSPFxPnPContextService,
  createSPFxPnPListService,
} from '@apvee/spfx-react-toolkit';

const contextService = createSPFxPnPContextService(spfxContext, pageContext);
const sp = contextService.createSPFI('/sites/projects', {
  cache: { enabled: true, storage: 'session' },
});

const tasks = createSPFxPnPListService<{ Id: number; Title: string }>(sp, 'Tasks', 50);
const result = await tasks.query(items => items.select('Id', 'Title'));
```

```ts
import {
  createSPFxAppCatalogService,
  createSPFxTenantKeyValueStoreService,
} from '@apvee/spfx-react-toolkit';

const appCatalog = createSPFxAppCatalogService(spHttpClient, pageContext);
const catalogUrl = await appCatalog.discoverUrl();

const store = createSPFxTenantKeyValueStoreService(spHttpClient);
await store.save('featureFlags', { enabled: true }, catalogUrl);
```

```ts
import {
  createSPFxOneDriveAppDataService,
  createSPFxUserPhotoService,
} from '@apvee/spfx-react-toolkit';

const appData = createSPFxOneDriveAppDataService(graphClient);
await appData.write('settings.json', { theme: 'dark' }, 'dashboard');

const photo = createSPFxUserPhotoService(graphClient);
const blob = await photo.getPhotoBlob({ email: 'user@contoso.com', size: '240x240' });
```

## Package import migration

Root and clean domain imports use the canonical implementations; historical `/lib/...` imports remain supported. Unrelated root imports no longer install incidental PnP features. Direct consumers must import every feature used by their operations, including callbacks passed to `invoke`/`batch`; toolkit-owned registrations do not cover arbitrary caller features. See [package imports](package-imports.md) for aliases, resolver and peer requirements, and the local-versus-tenant verification boundary.

## PnP feature registrations

Each PnP service module registers the features its operations need. A standalone list or search factory works with a supplied `SPFI` without first importing or calling `createSPFxPnPContextService`. Supply authentication and request behaviors on that client; registrations do not configure credentials or grant permissions.

| Factory module | Guaranteed PnP registrations |
| --- | --- |
| `createSPFxPnPContextService` | Webs and batching; `createSPFI` also applies the SPFx request behaviors |
| `createSPFxPnPListService` | Webs, lists, items and batching |
| `createSPFxPnPSearchService` | Search and suggestions |
| `createSPFxPnPService` | Batching |

Direct PnP operations, including callbacks passed to generic `invoke` or `batch`, require explicit imports for the features used by the consumer. Do not depend on incidental registrations from unrelated toolkit imports.

```typescript
import '@pnp/sp/webs'; // The callback below uses sp.web directly.
import { createSPFxPnPService } from '@apvee/spfx-react-toolkit/lib/services/spfx-pnp.service';

const service = createSPFxPnPService(sp); // sp is an authenticated, configured SPFI.
const web = await service.batch(batched => batched.web());
```

These factories retain their existing selector, query, cache, error and batch contracts. See the [PnP hooks](../api/hooks/pnpjs.md#pnp-feature-registrations) for the same registration policy when using hooks.

## Service boundaries

- Services may perform I/O.
- Services do not call hooks or depend on provider runtime state or React lifecycle.
- Services do not replace providers; callers obtain the necessary SPFx dependencies.

Style descriptors are immutable data, while `useSx` composes them using the existing host Griffel renderer. See the [style reference](../api/helpers/styles.md) for its separate cache and runtime contract.

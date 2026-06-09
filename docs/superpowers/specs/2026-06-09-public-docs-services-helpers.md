# Public Documentation Coverage For Helpers And Services

Date: 2026-06-09

## Goal

Bring the public documentation in line with the current public API before the next release, with focused changes that preserve the existing documentation structure and fully document the new public `helpers` and `services` exports.

## Context

The toolkit now exports these public modules from `src/index.ts`:

```ts
export * from './core';
export * from './hooks';
export * from './services';
export * from './helpers';
```

The existing public documentation already has module entry points:

- `docs/INDEX.md`
- `docs/INTRODUCTION.md`
- `docs/api/core/INDEX.md`
- `docs/api/hooks/INDEX.md`
- `docs/api/helpers/INDEX.md`
- `docs/api/services/INDEX.md`

The helper and service pages currently exist, but they are overview-level only. They do not yet document every public function, interface, input, return shape, boundary, and usage pattern exposed by the new public barrels.

## Current Findings

The current public docs need targeted cleanup:

- `docs/api/hooks/INDEX.md` references APIs that are not exported by this package:
  - `useSPFxPropertyPane`
  - `useSPFxPnPSP`
  - `useSPFxPnPGraph`
- `docs/api/helpers/INDEX.md` lists helper groups but does not document each helper signature and behavior.
- `docs/api/services/INDEX.md` lists service factories but does not document each service interface, method, expected dependency, error behavior, and usage.
- `docs/INDEX.md` links to helpers/services, but the module descriptions should explicitly treat them as public API.
- `docs/INTRODUCTION.md` should mention that hooks are the React API, while helpers and services are public non-React composition APIs.
- Public docs should not reference removed implementation concepts:
  - Jotai
  - atoms
  - the old generic `SPFxProvider` name where a host-specific provider is intended

## Scope

Modify only public documentation and documentation verification support.

In scope:

- Enrich `docs/api/helpers/INDEX.md`.
- Enrich `docs/api/services/INDEX.md`.
- Fix stale hook references in `docs/api/hooks/INDEX.md`.
- Update `docs/INDEX.md` and `docs/INTRODUCTION.md` to position helpers/services as public API.
- Add or update a documentation verification script that catches missing helper/service docs and stale API references.
- Keep existing docs folder layout.

Out of scope:

- No runtime code changes.
- No API signature changes.
- No new services or helpers.
- No generated documentation tool migration.
- No broad rewrite of all hook docs unless required to fix stale references.

## Documentation Structure

Keep the current structure:

```text
docs/
  INDEX.md
  INTRODUCTION.md
  api/
    core/
      INDEX.md
      providers.md
      types.md
    hooks/
      INDEX.md
      context.md
      environment.md
      http-clients.md
      performance.md
      permissions.md
      pnpjs.md
      properties.md
      storage.md
      theming.md
      user-site.md
    helpers/
      INDEX.md
    services/
      INDEX.md
```

Do not create helper/service sub-pages in this pass. The targeted change is to make the existing helper and service module pages complete enough for public API consumption.

## Helpers Documentation Requirements

`docs/api/helpers/INDEX.md` must document every export from `src/helpers/index.ts`.

### Page Requirements

The page must include:

- A short module summary.
- A "When To Use Helpers" section.
- A "Helpers vs Hooks vs Services" boundary table.
- One section per helper group.
- One subsection per exported helper function.
- Copy-pasteable TypeScript examples.
- Required imports from `@apvee/spfx-react-toolkit`.

### Required Helper Coverage

Document these exports:

| Helper | Source | Required docs |
|--------|--------|---------------|
| `getSPFxUserInfo` | `src/helpers/spfx-page-context.helpers.ts` | Signature, input `PageContext`, returned user fields, example |
| `getSPFxSiteInfo` | `src/helpers/spfx-page-context.helpers.ts` | Signature, returned site/web fields, optional `siteGroup`, example |
| `getSPFxListInfo` | `src/helpers/spfx-page-context.helpers.ts` | Signature, undefined when no list context, document-library detection, example |
| `getSPFxLocaleInfo` | `src/helpers/spfx-page-context.helpers.ts` | Signature, culture/timezone fields, example |
| `getSPFxEnvironmentInfo` | `src/helpers/spfx-page-context.helpers.ts` | Signature, environment type values, Teams/Office/Outlook flags, example |
| `getSPFxPageTypeInfo` | `src/helpers/spfx-page-context.helpers.ts` | Signature, page type values, boolean helpers, example |
| `getSPFxCorrelationInfo` | `src/helpers/spfx-page-context.helpers.ts` | Signature, correlation/tenant IDs, example |
| `hasSPFxPermission` | `src/helpers/spfx-permissions.helpers.ts` | Signature, `SPPermission` usage, undefined permission set behavior |
| `getSPFxContainerSize` | `src/helpers/spfx-container.helpers.ts` | Breakpoint categories and exact thresholds |
| `createScopedSPFxStorageKey` | `src/helpers/spfx-storage.helpers.ts` | Exact key format `spfx:<instanceId>:<key>` |
| `escapeODataValue` | `src/helpers/spfx-tenant-value.helpers.ts` | Single quote escaping behavior |
| `serializeTenantValue` | `src/helpers/spfx-tenant-value.helpers.ts` | Date, primitive, object, and fallback behavior |
| `deserializeTenantValue` | `src/helpers/spfx-tenant-value.helpers.ts` | JSON parse behavior and raw-string fallback |
| `buildOneDriveAppDataPath` | `src/helpers/spfx-graph-path.helpers.ts` | Graph app-root path format, folder sanitization |
| `buildUserPhotoEndpoint` | `src/helpers/spfx-graph-path.helpers.ts` | `/me` vs `/users/{id/email}`, supported sizes |
| `createFluent9ThemeFromSPFxTheme` | `src/helpers/spfx-theme.helpers.ts` | Default `webLightTheme` fallback and conversion |
| `getTeamsFluentTheme` | `src/helpers/spfx-theme.helpers.ts` | Theme name mapping |

### Helper Examples

The page must include examples for:

```ts
import {
  createScopedSPFxStorageKey,
  getSPFxUserInfo,
  getSPFxContainerSize,
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

## Services Documentation Requirements

`docs/api/services/INDEX.md` must document every export from `src/services/index.ts`.

### Page Requirements

The page must include:

- A short module summary.
- A "When To Use Services" section.
- A "Services vs Hooks" boundary table.
- A dependency matrix listing what each service factory needs.
- One section per service factory.
- One subsection per returned service method.
- Error behavior notes where the service throws or returns a fallback.
- Copy-pasteable TypeScript examples.
- Required imports from `@apvee/spfx-react-toolkit`.

### Required Service Coverage

Document these exports:

| Service factory | Source | Required docs |
|-----------------|--------|---------------|
| `createSPFxPnPContextService` | `src/services/spfx-pnp-context.service.ts` | `spfxContext`, page context dependency, `resolveSiteUrl`, `createSPFI`, cache/header config |
| `createSPFxPnPService` | `src/services/spfx-pnp.service.ts` | `invoke`, `batch`, `SPFI` ownership |
| `createSPFxPnPListService` | `src/services/spfx-pnp-list.service.ts` | `query`, `loadMore`, CRUD, batch methods, pagination result, partial failure behavior |
| `createSPFxPnPSearchService` | `src/services/spfx-pnp-search.service.ts` | `search`, `suggest`, default options, refiners, pagination |
| `createSPFxAppCatalogService` | `src/services/spfx-app-catalog.service.ts` | `discoverUrl`, `canCurrentUserWrite`, app catalog not provisioned error |
| `createSPFxTenantPropertyService` | `src/services/spfx-tenant-property.service.ts` | `get`, catalog URL dependency, deserialization |
| `createSPFxTenantKeyValueStoreService` | `src/services/spfx-tenant-key-value-store.service.ts` | `ensureListReady`, `get`, `list`, `save`, `remove`, hidden list behavior |
| `createSPFxOneDriveAppDataService` | `src/services/spfx-onedrive-app-data.service.ts` | `read`, `write`, `isNotFound`, JSON parse errors |
| `createSPFxUserPhotoService` | `src/services/spfx-user-photo.service.ts` | `getPhotoBlob`, options, Graph endpoint behavior |

### Services Dependency Matrix

The docs must include this matrix:

| Factory | Required dependency |
|---------|---------------------|
| `createSPFxPnPContextService` | `ISPFXContext`, page context-like object |
| `createSPFxPnPService` | `SPFI` |
| `createSPFxPnPListService` | `SPFI`, list title |
| `createSPFxPnPSearchService` | `SPFI` |
| `createSPFxAppCatalogService` | `SPHttpClient`, page context-like object |
| `createSPFxTenantPropertyService` | `SPHttpClient` |
| `createSPFxTenantKeyValueStoreService` | `SPHttpClient` |
| `createSPFxOneDriveAppDataService` | `MSGraphClientV3` |
| `createSPFxUserPhotoService` | `MSGraphClientV3` |

### Service Examples

The page must include examples for:

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

## Public Navigation Requirements

Update these files:

- `docs/INDEX.md`
- `docs/INTRODUCTION.md`
- `docs/api/hooks/INDEX.md`
- `docs/api/helpers/INDEX.md`
- `docs/api/services/INDEX.md`

Required navigation behavior:

- `docs/INDEX.md` must list Helpers and Services as first-class API modules in the table of contents.
- `docs/INDEX.md` Quick Links must point to `./api/helpers/INDEX.md` and `./api/services/INDEX.md`.
- `docs/INTRODUCTION.md` must mention helpers/services in the API overview.
- `docs/api/hooks/INDEX.md` must remove stale references to unavailable APIs.
- `docs/api/hooks/INDEX.md` must link to the actual PnP hooks:
  - `useSPFxPnP`
  - `useSPFxPnPContext`
  - `useSPFxPnPList`
  - `useSPFxPnPSearch`
- `docs/api/hooks/INDEX.md` must use `useSPFxIsEdit` instead of `useSPFxPropertyPane`.

## Public Wording Rules

The updated public docs must not include:

- `Jotai`
- `atom`
- `atoms`
- `SPFxProvider` as a generic provider name when a host-specific provider should be used
- references to non-exported hooks

Allowed exceptions:

- Words inside unrelated third-party package names in `package-lock.json` are irrelevant because this spec targets docs.
- `docs/superpowers/**` may contain historical planning/spec references.

## Verification Requirements

Add a focused documentation verification script:

```text
scripts/verify-public-docs.cjs
```

The script must:

1. Read `src/helpers/index.ts` and resolve exported helper modules.
2. Read each helper source file and collect exported function names.
3. Assert each helper function name appears in `docs/api/helpers/INDEX.md`.
4. Read `src/services/index.ts` and resolve exported service modules.
5. Read each service source file and collect exported factory function names and exported interfaces.
6. Assert each service factory name appears in `docs/api/services/INDEX.md`.
7. Assert `docs/api/hooks/INDEX.md` does not contain:
   - `useSPFxPropertyPane`
   - `useSPFxPnPSP`
   - `useSPFxPnPGraph`
8. Assert docs outside `docs/superpowers/**` do not contain:
   - `Jotai`
   - `atom`
   - `atoms`
9. Assert public docs contain links to:
   - `./api/helpers/INDEX.md`
   - `./api/services/INDEX.md`

Add this npm script:

```json
"verify:public-docs": "node scripts/verify-public-docs.cjs"
```

Do not rename existing verification scripts.

## Manual Review Checklist

After implementation, manually inspect:

- `docs/INDEX.md`
- `docs/INTRODUCTION.md`
- `docs/api/hooks/INDEX.md`
- `docs/api/helpers/INDEX.md`
- `docs/api/services/INDEX.md`

Review for:

- consistent import path `@apvee/spfx-react-toolkit`;
- no deprecated provider wording;
- no stale hook names;
- no implementation-only runtime store references in public docs;
- examples that compile conceptually with the public API;
- clear distinction between hooks, helpers, and services.

## Acceptance Criteria

This work is complete when:

- Every helper export is documented in `docs/api/helpers/INDEX.md`.
- Every service factory export is documented in `docs/api/services/INDEX.md`.
- Public navigation includes helpers and services as first-class modules.
- Public hook docs no longer mention non-exported hooks.
- `scripts/verify-public-docs.cjs` passes.
- Existing checks still pass:
  - `npx tsc --noEmit --pretty false`
  - `npx eslint src --ext .ts,.tsx --max-warnings=0`
  - `git diff --check`
- No runtime source changes are required.

## Risks

- Large documentation pages can become hard to scan. Keep sections compact and use tables for signatures, dependencies, and return fields.
- Some service behavior performs I/O and has nuanced error handling. Document actual behavior from source, not intended behavior.
- Over-documenting internals can expose unstable contracts. Only document exports reachable from the root public barrel.

## Recommended Implementation Order

1. Add failing `scripts/verify-public-docs.cjs`.
2. Run it and confirm it fails on current docs.
3. Fix `docs/api/hooks/INDEX.md` stale API references.
4. Expand `docs/api/helpers/INDEX.md`.
5. Expand `docs/api/services/INDEX.md`.
6. Update `docs/INDEX.md` and `docs/INTRODUCTION.md`.
7. Run `node scripts/verify-public-docs.cjs`.
8. Run the existing verification commands.
9. Perform manual review against the checklist.

# Services and Helpers Stabilization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Introduce stable non-React services and pure helpers for the agreed reusable SPFx/PnP/Graph operations, then remap the corresponding hooks to consume those services internally without changing public hook signatures.

**Architecture:** Add a new `src/services/` public layer for I/O-oriented SPFx, PnP, SharePoint REST, and Graph operations. Add a new `src/helpers/` public layer for pure mapping, path-building, serialization, permission, theme, and sizing utilities. Keep existing hooks as React facades that own lifecycle, local state, memoization, Jotai usage for now, and backward-compatible public APIs.

**Tech Stack:** SPFx 1.21.1, React 17, TypeScript 5.3, PnPjs v4, SPHttpClient, MSGraphClientV3, Jotai temporarily retained.

---

## Non-Goals

- Do not remove Jotai in this plan. Jotai removal is a separate follow-up refactor.
- Do not rename or remove existing public hooks, provider components, or public hook return types.
- Do not change runtime behavior intentionally.
- Do not introduce new runtime dependencies.
- Do not move React/SPFx dependency categories in `package.json`.
- Do not create services for hooks we agreed not to expose: provider internals, atom-backed state, theme subscription lifecycle, resize observer lifecycle, or Jotai internals.

## Public API Target

New public imports should work from the package root:

```ts
import {
  createSPFxPnPContextService,
  createSPFxPnPService,
  createSPFxPnPListService,
  createSPFxPnPSearchService,
  createSPFxAppCatalogService,
  createSPFxTenantPropertyService,
  createSPFxTenantKeyValueStoreService,
  createSPFxOneDriveAppDataService,
  createSPFxUserPhotoService,
  getSPFxUserInfo,
  getSPFxSiteInfo,
  getSPFxListInfo,
  getSPFxLocaleInfo,
  getSPFxEnvironmentInfo,
  getSPFxPageTypeInfo,
  getSPFxCorrelationInfo,
  hasSPFxPermission,
  getSPFxContainerSize,
  createScopedSPFxStorageKey,
  serializeTenantValue,
  deserializeTenantValue,
  escapeODataValue,
  buildOneDriveAppDataPath,
  buildUserPhotoEndpoint,
  createFluent9ThemeFromSPFxTheme,
  getTeamsFluentTheme,
} from '@apvee/spfx-react-toolkit';
```

Existing public hook imports must remain unchanged:

```ts
import {
  useSPFxPnPContext,
  useSPFxPnP,
  useSPFxPnPList,
  useSPFxPnPSearch,
  useSPFxTenantProperty,
  useSPFxTenantKeyValueStore,
  useSPFxOneDriveAppData,
  useSPFxUserPhoto,
} from '@apvee/spfx-react-toolkit';
```

## File Structure

Create these public service files:

- `src/services/index.ts`
  - Barrel exports for all service factories and service types.
- `src/services/spfx-pnp-context.service.ts`
  - Resolves site URLs and creates configured `SPFI` instances from SPFx context.
- `src/services/spfx-pnp.service.ts`
  - Provides generic `invoke` and `batch` wrappers around `SPFI`.
- `src/services/spfx-pnp-list.service.ts`
  - Provides list query, pagination, CRUD, and batch operations.
- `src/services/spfx-pnp-search.service.ts`
  - Provides SharePoint search execution, suggestions, refiners, and parsing.
- `src/services/spfx-app-catalog.service.ts`
  - Discovers tenant app catalog URL and checks current-user write permission.
- `src/services/spfx-tenant-property.service.ts`
  - Reads tenant storage entities through SharePoint REST.
- `src/services/spfx-tenant-key-value-store.service.ts`
  - Provides tenant key-value store provisioning and CRUD over hidden list.
- `src/services/spfx-onedrive-app-data.service.ts`
  - Provides Graph OneDrive app folder JSON read/write.
- `src/services/spfx-user-photo.service.ts`
  - Provides Graph user photo blob retrieval.

Create these public helper files:

- `src/helpers/index.ts`
  - Barrel exports for all helper functions and helper types.
- `src/helpers/spfx-page-context.helpers.ts`
  - Pure mappers from `PageContext`: user, site, list, locale, environment, page type, correlation.
- `src/helpers/spfx-permissions.helpers.ts`
  - Pure permission checks.
- `src/helpers/spfx-container.helpers.ts`
  - Pure container size breakpoint helper.
- `src/helpers/spfx-storage.helpers.ts`
  - Pure scoped storage key helper.
- `src/helpers/spfx-tenant-value.helpers.ts`
  - Pure tenant value serialization/deserialization and OData escaping.
- `src/helpers/spfx-graph-path.helpers.ts`
  - Pure OneDrive and user photo Graph endpoint builders.
- `src/helpers/spfx-theme.helpers.ts`
  - Pure SPFx/Teams to Fluent UI 9 theme mapping.

Modify these existing public barrels:

- `src/index.ts`
  - Export `./services` and `./helpers`.
- `package.json`
  - Add `lib/services/**/*` and `lib/helpers/**/*` to `files`.

Modify these hooks to consume services/helpers internally:

- `src/hooks/useSPFxPnPContext.ts`
- `src/hooks/useSPFxPnP.ts`
- `src/hooks/useSPFxPnPList.ts`
- `src/hooks/useSPFxPnPSearch.ts`
- `src/hooks/useAppCatalogUrl.internal.ts`
- `src/hooks/useSPFxTenantProperty.ts`
- `src/hooks/useSPFxTenantKeyValueStore.ts`
- `src/hooks/useSPFxOneDriveAppData.ts`
- `src/hooks/useSPFxUserPhoto.ts`
- Pure-mapping hooks only if low-risk and directly covered by helpers:
  - `src/hooks/useSPFxUserInfo.ts`
  - `src/hooks/useSPFxSiteInfo.ts`
  - `src/hooks/useSPFxListInfo.ts`
  - `src/hooks/useSPFxLocaleInfo.ts`
  - `src/hooks/useSPFxEnvironmentInfo.ts`
  - `src/hooks/useSPFxPageType.ts`
  - `src/hooks/useSPFxCorrelationInfo.ts`
  - `src/hooks/useSPFxPermissions.ts`
  - `src/hooks/useSPFxContainerSize.ts`

## Validation Commands

Use these commands after each meaningful slice:

```bash
npx tsc --noEmit --pretty false
npx eslint src/services src/helpers src/hooks --ext .ts,.tsx --max-warnings=0
npm run build
npm pack --dry-run
```

Expected result:

- TypeScript exits `0`.
- ESLint exits `0`.
- Build exits `0`.
- `npm pack --dry-run` includes `lib/services/**/*` and `lib/helpers/**/*`.

---

### Task 1: Establish Baseline and Public API Snapshot

**Files:**
- No source edits.
- Generated output outside repo: `/tmp/spfx-react-toolkit-api-services-before`

- [ ] **Step 1: Record current worktree state**

Run:

```bash
git branch --show-current
git status --short
```

Expected:

- Current branch is known.
- Dirty files are recorded before refactor starts.
- Do not revert unrelated dirty files.

- [ ] **Step 2: Generate declaration baseline**

Run:

```bash
npx tsc --emitDeclarationOnly --declaration --declarationMap false --outDir /tmp/spfx-react-toolkit-api-services-before --pretty false
```

Expected: exits `0`.

- [ ] **Step 3: Save focused declaration list**

Run:

```bash
find /tmp/spfx-react-toolkit-api-services-before -type f \
  \( -path '*/hooks/*.d.ts' -o -path '*/core/*.d.ts' -o -name 'index.d.ts' \) \
  -print
```

Expected:

- Existing hook declarations are present.
- There are no `services` or `helpers` declarations yet.

---

### Task 2: Add Public Service and Helper Barrels

**Files:**
- Create: `src/services/index.ts`
- Create: `src/helpers/index.ts`
- Modify: `src/index.ts`
- Modify: `package.json`

- [ ] **Step 1: Create empty service barrel**

Create `src/services/index.ts`:

```ts
/**
 * SPFx React Toolkit services
 *
 * Non-React service factories for composing SPFx, SharePoint REST, PnPjs,
 * and Microsoft Graph operations outside React hooks.
 *
 * @module services
 */
```

- [ ] **Step 2: Create empty helper barrel**

Create `src/helpers/index.ts`:

```ts
/**
 * SPFx React Toolkit helpers
 *
 * Pure helper functions for mapping SPFx context data, building Graph paths,
 * serializing values, checking permissions, and deriving UI metadata.
 *
 * @module helpers
 */
```

- [ ] **Step 3: Export services and helpers from root**

In `src/index.ts`, keep existing exports and add:

```ts
export * from './services';
export * from './helpers';
```

- [ ] **Step 4: Include generated service/helper artifacts in package files**

In `package.json`, update `files` to include:

```json
"lib/services/**/*",
"lib/helpers/**/*"
```

The resulting `files` array should include:

```json
[
  "lib/index.*",
  "lib/core/**/*",
  "lib/hooks/**/*",
  "lib/utils/**/*",
  "lib/services/**/*",
  "lib/helpers/**/*",
  "README.md",
  "LICENSE"
]
```

- [ ] **Step 5: Validate baseline compile**

Run:

```bash
npx tsc --noEmit --pretty false
```

Expected: exits `0`.

---

### Task 3: Add Pure Page Context, Permission, Container, Storage, Graph Path, Tenant Value, and Theme Helpers

**Files:**
- Create: `src/helpers/spfx-page-context.helpers.ts`
- Create: `src/helpers/spfx-permissions.helpers.ts`
- Create: `src/helpers/spfx-container.helpers.ts`
- Create: `src/helpers/spfx-storage.helpers.ts`
- Create: `src/helpers/spfx-tenant-value.helpers.ts`
- Create: `src/helpers/spfx-graph-path.helpers.ts`
- Create: `src/helpers/spfx-theme.helpers.ts`
- Modify: `src/helpers/index.ts`
- Modify low-risk pure hooks listed below.

- [ ] **Step 1: Create page context helper contracts**

Create `src/helpers/spfx-page-context.helpers.ts` with exported types matching existing hook return interfaces where practical:

```ts
import type { PageContext } from '@microsoft/sp-page-context';

export type SPFxEnvironmentType =
  | 'Local'
  | 'SharePoint'
  | 'SharePointOnPrem'
  | 'Teams'
  | 'Office'
  | 'Outlook';

export type SPFxPageType =
  | 'homePage'
  | 'sitePage'
  | 'newsPage'
  | 'listView'
  | 'listForm'
  | 'documentLibrary'
  | 'searchPage'
  | 'unknown';

export interface SPFxUserInfoData {
  readonly loginName: string;
  readonly displayName: string;
  readonly email: string;
  readonly isExternal: boolean;
}

export interface SPFxSiteInfoData {
  readonly webId: string;
  readonly webUrl: string;
  readonly webServerRelativeUrl: string;
  readonly title: string;
  readonly languageId: number;
  readonly logoUrl: string | undefined;
  readonly siteId: string;
  readonly siteUrl: string;
  readonly siteServerRelativeUrl: string;
  readonly siteClassification: string | undefined;
  readonly siteGroup: { id: string; isPublic: boolean } | undefined;
}

export interface SPFxListInfoData {
  readonly id: string;
  readonly title: string;
  readonly serverRelativeUrl: string;
  readonly itemId: number | undefined;
}

export interface SPFxTimeZoneData {
  readonly description: string;
  readonly id: number;
  readonly offset: number;
}

export interface SPFxLocaleInfoData {
  readonly currentCultureName: string;
  readonly currentUICultureName: string;
  readonly isRightToLeft: boolean;
  readonly timeZone: SPFxTimeZoneData | undefined;
}

export interface SPFxEnvironmentInfoData {
  readonly type: SPFxEnvironmentType;
  readonly isLocal: boolean;
  readonly isWorkbench: boolean;
  readonly isSharePoint: boolean;
  readonly isSharePointOnPrem: boolean;
  readonly isTeams: boolean;
  readonly isOffice: boolean;
  readonly isOutlook: boolean;
}

export interface SPFxPageTypeInfoData {
  readonly type: SPFxPageType;
  readonly isHomePage: boolean;
  readonly isSitePage: boolean;
  readonly isNewsPage: boolean;
  readonly isListView: boolean;
  readonly isListForm: boolean;
  readonly isDocumentLibrary: boolean;
  readonly isSearchPage: boolean;
}

export interface SPFxCorrelationInfoData {
  readonly correlationId: string | undefined;
  readonly requestId: string | undefined;
  readonly traceId: string | undefined;
}
```

- [ ] **Step 2: Implement page context helper functions by moving logic from existing hooks**

Implement these functions in `src/helpers/spfx-page-context.helpers.ts`:

```ts
export function getSPFxUserInfo(pageContext: PageContext): SPFxUserInfoData;
export function getSPFxSiteInfo(pageContext: PageContext): SPFxSiteInfoData;
export function getSPFxListInfo(pageContext: PageContext): SPFxListInfoData | undefined;
export function getSPFxLocaleInfo(pageContext: PageContext): SPFxLocaleInfoData;
export function getSPFxEnvironmentInfo(pageContext: PageContext): SPFxEnvironmentInfoData;
export function getSPFxPageTypeInfo(pageContext: PageContext): SPFxPageTypeInfoData;
export function getSPFxCorrelationInfo(pageContext: PageContext): SPFxCorrelationInfoData;
```

Move existing pure logic from:

- `src/hooks/useSPFxUserInfo.ts`
- `src/hooks/useSPFxSiteInfo.ts`
- `src/hooks/useSPFxListInfo.ts`
- `src/hooks/useSPFxLocaleInfo.ts`
- `src/hooks/useSPFxEnvironmentInfo.ts`
- `src/hooks/useSPFxPageType.ts`
- `src/hooks/useSPFxCorrelationInfo.ts`

Preserve existing fallback values, legacy context access, and string comparisons exactly.

- [ ] **Step 3: Create permission helper**

Create `src/helpers/spfx-permissions.helpers.ts`:

```ts
import type { SPPermission } from '@microsoft/sp-page-context';

export function hasSPFxPermission(
  permissionSet: SPPermission | undefined,
  permission: SPPermission
): boolean {
  if (!permissionSet) {
    return false;
  }

  return permissionSet.hasPermission(permission);
}
```

- [ ] **Step 4: Create container helper**

Create `src/helpers/spfx-container.helpers.ts` using the exact breakpoint logic currently inside `src/hooks/useSPFxContainerSize.ts`:

```ts
export type SPFxContainerSize =
  | 'small'
  | 'medium'
  | 'large'
  | 'xLarge'
  | 'xxLarge'
  | 'xxxLarge';

export function getSPFxContainerSize(width: number): SPFxContainerSize {
  if (width < 480) return 'small';
  if (width < 640) return 'medium';
  if (width < 1024) return 'large';
  if (width < 1366) return 'xLarge';
  if (width < 1920) return 'xxLarge';
  return 'xxxLarge';
}
```

If the current hook uses different thresholds, copy the current hook thresholds instead of these values.

- [ ] **Step 5: Create storage key helper**

Create `src/helpers/spfx-storage.helpers.ts`:

```ts
export function createScopedSPFxStorageKey(instanceId: string, key: string): string {
  return 'spfx:' + instanceId + ':' + key;
}
```

- [ ] **Step 6: Create tenant value helper**

Create `src/helpers/spfx-tenant-value.helpers.ts`:

```ts
export function escapeODataValue(value: string): string {
  return value.replace(/'/g, "''");
}

export function serializeTenantValue(value: unknown): string {
  if (value === null) return String(value);
  if (value instanceof Date) return value.toISOString();

  const type = typeof value;
  if (type === 'string' || type === 'number' || type === 'boolean' || type === 'bigint') {
    return String(value);
  }

  return JSON.stringify(value);
}

export function deserializeTenantValue<T>(rawValue: string): T {
  try {
    return JSON.parse(rawValue) as T;
  } catch {
    return rawValue as unknown as T;
  }
}
```

- [ ] **Step 7: Create Graph path helper**

Create `src/helpers/spfx-graph-path.helpers.ts`:

```ts
export interface SPFxUserPhotoEndpointOptions {
  readonly userId?: string;
  readonly email?: string;
  readonly size?: string;
}

export function buildOneDriveAppDataPath(fileName: string, folderName?: string): string {
  const basePath = '/me/drive/special/approot:';

  if (folderName) {
    const safeFolderName = folderName.replace(/[^a-zA-Z0-9-_]/g, '-');
    return `${basePath}/${safeFolderName}/${fileName}:/content`;
  }

  return `${basePath}/${fileName}:/content`;
}

export function buildUserPhotoEndpoint(options?: SPFxUserPhotoEndpointOptions): string {
  const userId = options?.userId;
  const email = options?.email;
  const size = options?.size ?? '240x240';

  let basePath: string;
  if (userId) {
    basePath = `/users/${userId}`;
  } else if (email) {
    basePath = `/users/${email}`;
  } else {
    basePath = '/me';
  }

  return `${basePath}/photos/${size}/$value`;
}
```

- [ ] **Step 8: Create theme helper**

Create `src/helpers/spfx-theme.helpers.ts`:

```ts
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { Theme } from '@fluentui/react-theme';
import {
  teamsDarkTheme,
  teamsHighContrastTheme,
  teamsLightTheme,
} from '@fluentui/react-theme';
import { createV9Theme } from '@fluentui/react-migration-v8-v9';

export type TeamsThemeName = 'default' | 'dark' | 'highContrast';

export function getTeamsFluentTheme(teamsThemeName: TeamsThemeName): Theme {
  switch (teamsThemeName) {
    case 'dark':
      return teamsDarkTheme;
    case 'highContrast':
      return teamsHighContrastTheme;
    case 'default':
    default:
      return teamsLightTheme;
  }
}

export function createFluent9ThemeFromSPFxTheme(spfxTheme: IReadonlyTheme): Theme {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return createV9Theme(spfxTheme as any);
}
```

- [ ] **Step 9: Export helpers from barrel**

Update `src/helpers/index.ts`:

```ts
/**
 * SPFx React Toolkit helpers
 *
 * Pure helper functions for mapping SPFx context data, building Graph paths,
 * serializing values, checking permissions, and deriving UI metadata.
 *
 * @module helpers
 */
export * from './spfx-page-context.helpers';
export * from './spfx-permissions.helpers';
export * from './spfx-container.helpers';
export * from './spfx-storage.helpers';
export * from './spfx-tenant-value.helpers';
export * from './spfx-graph-path.helpers';
export * from './spfx-theme.helpers';
```

- [ ] **Step 10: Remap pure hooks to helpers**

For each pure hook, replace duplicated mapping logic with the corresponding helper:

- `useSPFxUserInfo` returns `useMemo(() => getSPFxUserInfo(pageContext), [pageContext])`.
- `useSPFxSiteInfo` returns `useMemo(() => getSPFxSiteInfo(pageContext), [pageContext])`.
- `useSPFxListInfo` returns `useMemo(() => getSPFxListInfo(pageContext), [pageContext])`.
- `useSPFxLocaleInfo` returns `useMemo(() => getSPFxLocaleInfo(pageContext), [pageContext])`.
- `useSPFxEnvironmentInfo` returns `useMemo(() => getSPFxEnvironmentInfo(pageContext), [pageContext])`.
- `useSPFxPageType` returns `useMemo(() => getSPFxPageTypeInfo(pageContext), [pageContext])`.
- `useSPFxCorrelationInfo` returns `useMemo(() => getSPFxCorrelationInfo(pageContext), [pageContext])`.
- `useSPFxPermissions` uses `hasSPFxPermission`.
- `useSPFxContainerSize` uses `getSPFxContainerSize`.
- `useSPFxStorage` uses `createScopedSPFxStorageKey`.
- `useSPFxFluent9ThemeInfo` uses `getTeamsFluentTheme` and `createFluent9ThemeFromSPFxTheme`.

Keep each hook return type name and exported function name unchanged.

- [ ] **Step 11: Validate helper slice**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/helpers src/hooks --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 4: Add PnP Context and Generic PnP Services

**Files:**
- Create: `src/services/spfx-pnp-context.service.ts`
- Create: `src/services/spfx-pnp.service.ts`
- Modify: `src/services/index.ts`
- Modify: `src/hooks/useSPFxPnPContext.ts`
- Modify: `src/hooks/useSPFxPnP.ts`

- [ ] **Step 1: Create PnP context service**

Create `src/services/spfx-pnp-context.service.ts`:

```ts
import { spfi, SPFI } from '@pnp/sp';
import { SPFx } from '@pnp/sp';
import { Caching } from '@pnp/queryable';
import { InjectHeaders } from '@pnp/queryable';
import type { PageContext } from '@microsoft/sp-page-context';
import type { SPFxContextType } from '../core/types';

import '@pnp/sp/webs';
import '@pnp/sp/batching';

export interface PnPContextConfig {
  readonly cache?: {
    readonly enabled: boolean;
    readonly storage?: 'session' | 'local';
    readonly timeout?: number;
    readonly keyFactory?: (url: string) => string;
  };
  readonly batch?: {
    readonly enabled: boolean;
    readonly maxRequests?: number;
  };
  readonly headers?: Record<string, string>;
}

export interface SPFxPnPContextService {
  readonly resolveSiteUrl: (siteUrl?: string) => string;
  readonly createSPFI: (siteUrl?: string, config?: PnPContextConfig) => SPFI;
  readonly getConfigKey: (config?: PnPContextConfig) => string;
}

function createDefaultPnPCacheKey(url: string): string {
  let hash = 0;
  for (let i = 0; i < url.length; i++) {
    const char = url.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }
  return `pnp-cache-${Math.abs(hash)}`;
}

export function createSPFxPnPContextService(
  spfxContext: SPFxContextType,
  pageContext: PageContext
): SPFxPnPContextService {
  const currentWebUrl = pageContext.web.absoluteUrl;

  const resolveSiteUrl = (siteUrl?: string): string => {
    if (!siteUrl) {
      return currentWebUrl;
    }

    const trimmed = siteUrl.charAt(siteUrl.length - 1) === '/'
      ? siteUrl.slice(0, -1)
      : siteUrl;

    if (trimmed.charAt(0) === '/') {
      const origin = new URL(currentWebUrl).origin;
      return `${origin}${trimmed}`;
    }

    return trimmed;
  };

  const getConfigKey = (config?: PnPContextConfig): string => {
    return JSON.stringify(config || {});
  };

  const createSPFI = (siteUrl?: string, config?: PnPContextConfig): SPFI => {
    const effectiveSiteUrl = resolveSiteUrl(siteUrl);
    let instance = spfi(effectiveSiteUrl).using(SPFx(spfxContext));

    if (config?.cache?.enabled) {
      instance = instance.using(Caching({
        store: config.cache.storage || 'session',
        keyFactory: config.cache.keyFactory || createDefaultPnPCacheKey,
        timeout: config.cache.timeout || 300000,
      }));
    }

    if (config?.headers) {
      instance = instance.using(InjectHeaders(config.headers));
    }

    return instance;
  };

  return {
    resolveSiteUrl,
    createSPFI,
    getConfigKey,
  };
}
```

- [ ] **Step 2: Remap `useSPFxPnPContext` to the service**

In `src/hooks/useSPFxPnPContext.ts`:

- Import `createSPFxPnPContextService`.
- Re-export or reuse `PnPContextConfig` from the service so existing type import paths continue working.
- Replace local URL resolution and SPFI creation logic with the service.

The hook must keep this signature:

```ts
export function useSPFxPnPContext(
  siteUrl?: string,
  config?: PnPContextConfig
): PnPContextInfo
```

The hook still owns:

- `useSPFxContext()`
- `useSPFxPageContext()`
- `useState<Error | undefined>`
- `useMemo`
- returning `{ sp, isInitialized, error, siteUrl }`

- [ ] **Step 3: Create generic PnP service**

Create `src/services/spfx-pnp.service.ts`:

```ts
import type { SPFI } from '@pnp/sp';

export interface SPFxPnPService {
  readonly sp: SPFI;
  readonly invoke: <T>(fn: (sp: SPFI) => Promise<T>) => Promise<T>;
  readonly batch: <T>(fn: (batchedSP: SPFI) => Promise<T>) => Promise<T>;
}

export function createSPFxPnPService(sp: SPFI): SPFxPnPService {
  return {
    sp,
    invoke: async <T>(fn: (sp: SPFI) => Promise<T>): Promise<T> => {
      return fn(sp);
    },
    batch: async <T>(fn: (batchedSP: SPFI) => Promise<T>): Promise<T> => {
      const [batchedSP, execute] = sp.batched();
      const resultPromise = fn(batchedSP);
      await execute();
      return resultPromise;
    },
  };
}
```

- [ ] **Step 4: Remap `useSPFxPnP` to generic service**

In `src/hooks/useSPFxPnP.ts`:

- Import `createSPFxPnPService`.
- Build `const service = useMemo(() => sp ? createSPFxPnPService(sp) : undefined, [sp]);`.
- Keep local `isLoading`, `invokeError`, `clearError`.
- `invoke` should call `service.invoke(fn)`.
- `batch` should call `service.batch(fn)`.
- Preserve existing error messages for uninitialized `sp`.

- [ ] **Step 5: Export services**

Update `src/services/index.ts`:

```ts
export * from './spfx-pnp-context.service';
export * from './spfx-pnp.service';
```

- [ ] **Step 6: Validate PnP context slice**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/services/spfx-pnp-context.service.ts src/services/spfx-pnp.service.ts src/hooks/useSPFxPnPContext.ts src/hooks/useSPFxPnP.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 5: Add PnP List Service and Remap List Hook

**Files:**
- Create: `src/services/spfx-pnp-list.service.ts`
- Modify: `src/services/index.ts`
- Modify: `src/hooks/useSPFxPnPList.ts`
- Keep or move from current internals:
  - `src/hooks/useSPFxPnPList.query.internal.ts`
  - `src/hooks/useSPFxPnPList.batch.internal.ts`

- [ ] **Step 1: Create service contract**

Create `src/services/spfx-pnp-list.service.ts`:

```ts
import type { SPFI } from '@pnp/sp';
import type { IItems } from '@pnp/sp/items';
import {
  collectCreatedIds,
  collectRejectedReasons,
  createBatchError,
} from '../hooks/useSPFxPnPList.batch.internal';
import {
  createMonitoredListQuery,
  hasMoreListItems,
  resolveEffectiveListQuery,
} from '../hooks/useSPFxPnPList.query.internal';

export interface SPFxPnPListQueryResult<T> {
  readonly items: T[];
  readonly effectivePageSize: number | undefined;
  readonly hasMore: boolean;
}

export interface SPFxPnPListLoadMoreOptions {
  readonly queryBuilder: (items: IItems) => IItems;
  readonly currentSkip: number;
  readonly pageSize: number;
}

export interface SPFxPnPListService<T = unknown> {
  readonly query: (
    queryBuilder?: (items: IItems) => IItems,
    options?: { pageSize?: number }
  ) => Promise<SPFxPnPListQueryResult<T>>;
  readonly loadMore: (options: SPFxPnPListLoadMoreOptions) => Promise<SPFxPnPListQueryResult<T>>;
  readonly getById: (id: number) => Promise<T>;
  readonly create: (item: Partial<T>) => Promise<number>;
  readonly update: (id: number, item: Partial<T>) => Promise<void>;
  readonly remove: (id: number) => Promise<void>;
  readonly createBatch: (items: Partial<T>[]) => Promise<number[]>;
  readonly updateBatch: (updates: Array<{ id: number; item: Partial<T> }>) => Promise<void>;
  readonly removeBatch: (ids: number[]) => Promise<void>;
}
```

- [ ] **Step 2: Implement service factory**

In `src/services/spfx-pnp-list.service.ts`, implement:

```ts
export function createSPFxPnPListService<T = unknown>(
  sp: SPFI,
  listTitle: string,
  defaultPageSize?: number
): SPFxPnPListService<T> {
  const getItems = (): IItems => sp.web.lists.getByTitle(listTitle).items;

  return {
    query: async (
      queryBuilder?: (items: IItems) => IItems,
      options?: { pageSize?: number }
    ): Promise<SPFxPnPListQueryResult<T>> => {
      const pageSize = options?.pageSize ?? defaultPageSize;
      const tracker = { top: undefined as number | undefined };
      const monitored = createMonitoredListQuery(getItems(), tracker);
      const userQuery = queryBuilder ? queryBuilder(monitored) : monitored;

      if (tracker.top !== undefined && pageSize !== undefined) {
        console.warn(
          `[useSPFxPnPList] Both .top(${tracker.top}) and pageSize(${pageSize}) specified. ` +
          `Using .top(${tracker.top}).`
        );
      }

      const effective = resolveEffectiveListQuery(userQuery, tracker, pageSize);
      const result = await effective.query() as T[];

      return {
        items: result,
        effectivePageSize: effective.pageSize,
        hasMore: hasMoreListItems(result.length, effective.pageSize),
      };
    },

    loadMore: async (options: SPFxPnPListLoadMoreOptions): Promise<SPFxPnPListQueryResult<T>> => {
      const tracker = { top: undefined as number | undefined };
      const monitored = createMonitoredListQuery(getItems(), tracker);
      const userQuery = options.queryBuilder(monitored);
      const result = await userQuery.skip(options.currentSkip).top(options.pageSize)() as T[];

      return {
        items: result,
        effectivePageSize: options.pageSize,
        hasMore: hasMoreListItems(result.length, options.pageSize),
      };
    },

    getById: async (id: number): Promise<T> => {
      return getItems().getById(id)() as Promise<T>;
    },

    create: async (item: Partial<T>): Promise<number> => {
      const result = await getItems().add(item as Record<string, unknown>);
      return result.data.Id;
    },

    update: async (id: number, item: Partial<T>): Promise<void> => {
      await getItems().getById(id).update(item as Record<string, unknown>);
    },

    remove: async (id: number): Promise<void> => {
      await getItems().getById(id).delete();
    },

    createBatch: async (itemsToCreate: Partial<T>[]): Promise<number[]> => {
      const [batchedSP, execute] = sp.batched();
      const list = batchedSP.web.lists.getByTitle(listTitle);
      const operations = itemsToCreate.map(itemToCreate =>
        list.items.add(itemToCreate as Record<string, unknown>)
      );

      await execute();
      const settled = await Promise.allSettled(operations);
      const { ids, errors } = collectCreatedIds(settled);

      if (errors.length > 0) {
        console.error('Batch create summary:', errors);
        throw createBatchError('create', errors.length, itemsToCreate.length);
      }

      return ids;
    },

    updateBatch: async (updates: Array<{ id: number; item: Partial<T> }>): Promise<void> => {
      const [batchedSP, execute] = sp.batched();
      const list = batchedSP.web.lists.getByTitle(listTitle);
      const operations = updates.map(updateItem =>
        list.items.getById(updateItem.id).update(updateItem.item as Record<string, unknown>)
      );

      await execute();
      const settled = await Promise.allSettled(operations);
      const errors = collectRejectedReasons(settled);

      if (errors.length > 0) {
        console.error('Batch update summary:', errors);
        throw createBatchError('update', errors.length, updates.length);
      }
    },

    removeBatch: async (ids: number[]): Promise<void> => {
      const [batchedSP, execute] = sp.batched();
      const list = batchedSP.web.lists.getByTitle(listTitle);
      const operations = ids.map(id => list.items.getById(id).delete());

      await execute();
      const settled = await Promise.allSettled(operations);
      const errors = collectRejectedReasons(settled);

      if (errors.length > 0) {
        console.error('Batch delete summary:', errors);
        throw createBatchError('delete', errors.length, ids.length);
      }
    },
  };
}
```

- [ ] **Step 3: Remap `useSPFxPnPList` to service**

In `src/hooks/useSPFxPnPList.ts`:

- Import `createSPFxPnPListService`.
- Create service with:

```ts
const service = useMemo(() => {
  return sp && context?.isInitialized
    ? createSPFxPnPListService<T>(sp, listTitle, defaultPageSize)
    : undefined;
}, [sp, context?.isInitialized, listTitle, defaultPageSize]);
```

- Replace direct PnP operations with service calls.
- Keep hook-owned state:
  - `items`
  - `loading`
  - `loadingMore`
  - `error`
  - `hasMore`
  - `lastQueryBuilder`
  - `lastEffectivePageSize`
  - `currentSkip`
  - debounced refetch timeout
  - mounted cleanup

Preserve exact public behavior:

- `getById` still returns `undefined` on not found/error and sets `error`.
- `create/update/remove/createBatch/updateBatch/removeBatch` still trigger debounced refetch.
- `query` still returns `T[]`.
- `loadMore` still returns `T[]`.

- [ ] **Step 4: Export list service**

Update `src/services/index.ts`:

```ts
export * from './spfx-pnp-list.service';
```

- [ ] **Step 5: Validate list slice**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/services/spfx-pnp-list.service.ts src/hooks/useSPFxPnPList.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 6: Add PnP Search Service and Remap Search Hook

**Files:**
- Create: `src/services/spfx-pnp-search.service.ts`
- Modify: `src/services/index.ts`
- Modify: `src/hooks/useSPFxPnPSearch.ts`
- Keep or move from current internals:
  - `src/hooks/useSPFxPnPSearch.query.internal.ts`
  - `src/hooks/useSPFxPnPSearch.results.internal.ts`

- [ ] **Step 1: Create search service contract**

Create `src/services/spfx-pnp-search.service.ts`:

```ts
import type { SPFI } from '@pnp/sp';
import type { ISearchBuilder, ISuggestResult } from '@pnp/sp/search';
import '@pnp/sp/search';
import {
  buildRefinementFilters,
  parseSearchRefiners,
  parseSearchResults,
} from '../hooks/useSPFxPnPSearch.results.internal';
import { buildSearchQuery } from '../hooks/useSPFxPnPSearch.query.internal';
import type {
  SearchRefiner,
  SearchResult,
  UseSPFxPnPSearchOptions,
} from '../hooks/useSPFxPnPSearch';

export type SearchQueryBuilderFn = (builder: ISearchBuilder) => ISearchBuilder;

export interface SPFxPnPSearchExecuteOptions {
  readonly pageSize?: number;
  readonly startRow?: number;
  readonly refiners?: Map<string, string[]>;
}

export interface SPFxPnPSearchExecuteResult<T = Record<string, string>> {
  readonly results: SearchResult<T>[];
  readonly totalRows: number;
  readonly refiners: SearchRefiner[];
}

export interface SPFxPnPSearchService<T = Record<string, string>> {
  readonly search: (
    query: string | SearchQueryBuilderFn,
    options?: SPFxPnPSearchExecuteOptions
  ) => Promise<SPFxPnPSearchExecuteResult<T>>;
  readonly suggest: (queryText: string) => Promise<string[]>;
}
```

- [ ] **Step 2: Implement search service factory**

In `src/services/spfx-pnp-search.service.ts`, implement:

```ts
export function createSPFxPnPSearchService<T = Record<string, string>>(
  sp: SPFI,
  defaultOptions?: UseSPFxPnPSearchOptions
): SPFxPnPSearchService<T> {
  const defaultPageSize = defaultOptions?.pageSize ?? 50;

  return {
    search: async (
      query: string | SearchQueryBuilderFn,
      options?: SPFxPnPSearchExecuteOptions
    ): Promise<SPFxPnPSearchExecuteResult<T>> => {
      const pageSize = options?.pageSize ?? defaultPageSize;
      const startRow = options?.startRow ?? 0;

      let builder = buildSearchQuery(query, defaultOptions);
      builder = builder.rowLimit(pageSize);

      if (startRow > 0) {
        builder = builder.startRow(startRow);
      }

      const refinementFilters = buildRefinementFilters(options?.refiners ?? new Map());
      if (refinementFilters.length > 0) {
        builder = builder.refinementFilters(...refinementFilters);
      }

      const searchResults = await sp.search(builder);
      const rawResults = searchResults.PrimarySearchResults ?? [];
      const totalRows = searchResults.TotalRows ?? 0;
      const refinerResults = searchResults.RawSearchResults?.PrimaryQueryResult?.RefinementResults?.Refiners ?? [];

      return {
        results: parseSearchResults<T>(rawResults),
        totalRows,
        refiners: parseSearchRefiners(refinerResults),
      };
    },

    suggest: async (queryText: string): Promise<string[]> => {
      const result: ISuggestResult = await sp.searchSuggest(queryText);
      return result.Queries ?? [];
    },
  };
}
```

- [ ] **Step 3: Remap `useSPFxPnPSearch` to service**

In `src/hooks/useSPFxPnPSearch.ts`:

- Import `createSPFxPnPSearchService`.
- Create service with:

```ts
const service = useMemo(() => {
  return sp && context?.isInitialized
    ? createSPFxPnPSearchService<T>(sp, options)
    : undefined;
}, [sp, context?.isInitialized, options]);
```

- Replace direct `sp.search` and `sp.searchSuggest` calls with `service.search` and `service.suggest`.
- Keep hook-owned state:
  - `results`
  - `totalResults`
  - `refiners`
  - `loading`
  - `loadingMore`
  - `error`
  - `hasMore`
  - `lastQueryBuilder`
  - `lastQueryText`
  - `lastPageSize`
  - `currentStartRow`
  - `appliedRefiners`
  - mounted ref

Preserve exact public behavior:

- `search` returns `SearchResult<T>[]`.
- `suggest` returns `string[]`.
- `loadMore`, `refetch`, and `applyRefiner` keep the same error messages and state behavior.

- [ ] **Step 4: Export search service**

Update `src/services/index.ts`:

```ts
export * from './spfx-pnp-search.service';
```

- [ ] **Step 5: Validate search slice**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/services/spfx-pnp-search.service.ts src/hooks/useSPFxPnPSearch.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 7: Add App Catalog, Tenant Property, and Tenant Key-Value Store Services

**Files:**
- Create: `src/services/spfx-app-catalog.service.ts`
- Create: `src/services/spfx-tenant-property.service.ts`
- Create: `src/services/spfx-tenant-key-value-store.service.ts`
- Modify: `src/services/index.ts`
- Modify: `src/hooks/useAppCatalogUrl.internal.ts`
- Modify: `src/hooks/useSPFxTenantProperty.ts`
- Modify: `src/hooks/useSPFxTenantKeyValueStore.ts`
- Modify or remove internal helper imports where replaced:
  - `src/hooks/useSPFxTenantKeyValueStore.serialization.internal.ts`
  - `src/hooks/useSPFxTenantKeyValueStore.sharepoint.internal.ts`

- [ ] **Step 1: Create app catalog service**

Create `src/services/spfx-app-catalog.service.ts`:

```ts
import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';
import type { PageContext } from '@microsoft/sp-page-context';

interface TenantAppCatalogResponse {
  CorporateCatalogUrl: string;
}

export interface SPFxAppCatalogService {
  readonly discoverUrl: () => Promise<string>;
  readonly canCurrentUserWrite: (catalogUrl: string) => Promise<boolean>;
}

export function createSPFxAppCatalogService(
  spHttpClient: SPHttpClient,
  pageContext: PageContext
): SPFxAppCatalogService {
  let cachedUrl: string | undefined;

  return {
    discoverUrl: async (): Promise<string> => {
      if (cachedUrl) {
        return cachedUrl;
      }

      const response: SPHttpClientResponse = await spHttpClient.get(
        `${pageContext.web.absoluteUrl}/_api/SP_TenantSettings_Current`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to discover app catalog: ${response.statusText}`);
      }

      const data: TenantAppCatalogResponse = await response.json();
      if (!data.CorporateCatalogUrl) {
        throw new Error('Tenant app catalog is not provisioned. Please provision the app catalog first.');
      }

      cachedUrl = data.CorporateCatalogUrl;
      return cachedUrl;
    },

    canCurrentUserWrite: async (catalogUrl: string): Promise<boolean> => {
      try {
        const response: SPHttpClientResponse = await spHttpClient.get(
          `${catalogUrl}/_api/web/currentuser?$select=IsSiteAdmin`,
          SPHttpClient.configurations.v1
        );

        if (!response.ok) return false;

        const user = await response.json();
        return user.IsSiteAdmin === true;
      } catch {
        return false;
      }
    },
  };
}
```

- [ ] **Step 2: Remap `useAppCatalogUrl.internal` to app catalog service**

In `src/hooks/useAppCatalogUrl.internal.ts`:

- Import `createSPFxAppCatalogService`.
- Keep hook state: `appCatalogUrl`, `appCatalogUrlRef`, `isMountedRef`.
- Create service with `useMemo` when `spHttpClient` and `pageContext` exist.
- `discoverAppCatalogUrl` should call `service.discoverUrl()`, update refs/state, and preserve the existing wrapped error message:

```ts
throw new Error(`App catalog discovery failed: ${err instanceof Error ? err.message : String(err)}`);
```

- `checkWritePermission` should call `service.canCurrentUserWrite(catalogUrl)` and return `false` if service is missing.

- [ ] **Step 3: Create tenant property service**

Create `src/services/spfx-tenant-property.service.ts`:

```ts
import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';
import { deserializeTenantValue } from '../helpers/spfx-tenant-value.helpers';

interface StorageEntity {
  Value?: string;
  Description?: string;
}

export interface SPFxTenantProperty<T = unknown> {
  readonly data: T | undefined;
  readonly description: string | undefined;
}

export interface SPFxTenantPropertyService {
  readonly get: <T = unknown>(key: string, catalogUrl: string) => Promise<SPFxTenantProperty<T>>;
}

export function createSPFxTenantPropertyService(
  spHttpClient: SPHttpClient
): SPFxTenantPropertyService {
  return {
    get: async <T = unknown>(key: string, catalogUrl: string): Promise<SPFxTenantProperty<T>> => {
      const response: SPHttpClientResponse = await spHttpClient.get(
        `${catalogUrl}/_api/web/GetStorageEntity('${encodeURIComponent(key)}')`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to read property: ${response.statusText}`);
      }

      const entity: StorageEntity = await response.json();

      if (entity.Value) {
        return {
          data: deserializeTenantValue<T>(entity.Value),
          description: entity.Description,
        };
      }

      return {
        data: undefined,
        description: undefined,
      };
    },
  };
}
```

- [ ] **Step 4: Remap `useSPFxTenantProperty` to tenant property service**

In `src/hooks/useSPFxTenantProperty.ts`:

- Import `createSPFxTenantPropertyService`.
- Remove local `deserializeValue`.
- Create service from `spHttpClient`.
- `load` calls:

```ts
const catalogUrl = await discoverAppCatalogUrl();
const property = await service.get<T>(key, catalogUrl);
```

- Preserve `autoFetch`, loading state, error handling, and `isReady`.

- [ ] **Step 5: Create tenant key-value store service**

Create `src/services/spfx-tenant-key-value-store.service.ts` by moving the existing SharePoint REST list functions from `src/hooks/useSPFxTenantKeyValueStore.sharepoint.internal.ts` and value serialization from `src/helpers/spfx-tenant-value.helpers.ts`.

The public contract should be:

```ts
import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';
import {
  deserializeTenantValue,
  escapeODataValue,
  serializeTenantValue,
} from '../helpers/spfx-tenant-value.helpers';

export interface SPFxTenantKeyValueStoreItem<T = unknown> {
  readonly key: string;
  readonly value: T;
  readonly description: string | undefined;
  readonly id: number;
}

export interface SPFxTenantKeyValueStoreService {
  readonly ensureListReady: (catalogUrl: string) => Promise<void>;
  readonly get: <T = unknown>(key: string, catalogUrl: string) => Promise<SPFxTenantKeyValueStoreItem<T> | undefined>;
  readonly list: (catalogUrl: string) => Promise<SPFxTenantKeyValueStoreItem<unknown>[]>;
  readonly save: <T = unknown>(key: string, value: T, catalogUrl: string, description?: string) => Promise<void>;
  readonly remove: (key: string, catalogUrl: string) => Promise<void>;
}
```

The service should own:

- `TENANT_KEY_VALUE_STORE_LIST_TITLE`
- `getListApiUrl`
- `createField`
- `ensureListReady`
- `findByKey`
- `mapItem`
- `get`
- `list`
- `save`
- `remove`

The hook should own:

- loading state
- write state
- read error state
- write error state
- `canWrite`
- eager initialization effect
- provisioning mutex refs, unless moved into service instance safely

- [ ] **Step 6: Remap `useSPFxTenantKeyValueStore` to service**

In `src/hooks/useSPFxTenantKeyValueStore.ts`:

- Import `createSPFxTenantKeyValueStoreService`.
- Create service from `spHttpClient`.
- Keep the current provisioning mutex behavior if service does not internalize it.
- Replace direct REST calls with:

```ts
await service.ensureListReady(catalogUrl);
const item = await service.get<T>(key, catalogUrl);
const items = await service.list(catalogUrl);
await service.save(key, value, catalogUrl, description);
await service.remove(key, catalogUrl);
```

Preserve existing hook behavior:

- `get` returns `undefined` when provisioning fails for read-only users.
- `list` returns `[]` when provisioning fails for read-only users.
- `save/remove` throw on write failures.
- `canWrite` still comes from app catalog permission check.

- [ ] **Step 7: Export app catalog and tenant services**

Update `src/services/index.ts`:

```ts
export * from './spfx-app-catalog.service';
export * from './spfx-tenant-property.service';
export * from './spfx-tenant-key-value-store.service';
```

- [ ] **Step 8: Validate tenant slice**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/services/spfx-app-catalog.service.ts src/services/spfx-tenant-property.service.ts src/services/spfx-tenant-key-value-store.service.ts src/hooks/useAppCatalogUrl.internal.ts src/hooks/useSPFxTenantProperty.ts src/hooks/useSPFxTenantKeyValueStore.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 8: Add Graph OneDrive App Data and User Photo Services

**Files:**
- Create: `src/services/spfx-onedrive-app-data.service.ts`
- Create: `src/services/spfx-user-photo.service.ts`
- Modify: `src/services/index.ts`
- Modify: `src/hooks/useSPFxOneDriveAppData.ts`
- Modify: `src/hooks/useSPFxUserPhoto.ts`

- [ ] **Step 1: Create OneDrive app data service**

Create `src/services/spfx-onedrive-app-data.service.ts`:

```ts
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { buildOneDriveAppDataPath } from '../helpers/spfx-graph-path.helpers';

export interface SPFxOneDriveAppDataReadResult<T> {
  readonly data: T | undefined;
  readonly isNotFound: boolean;
}

export interface SPFxOneDriveAppDataService {
  readonly read: <T = unknown>(fileName: string, folder?: string) => Promise<SPFxOneDriveAppDataReadResult<T>>;
  readonly write: <T = unknown>(fileName: string, content: T, folder?: string) => Promise<void>;
}

export function isSPFxGraphNotFoundError(err: unknown): boolean {
  const anyErr = err as {
    statusCode?: number;
    status?: number;
    code?: string;
    message?: string;
    body?: { error?: { code?: string; message?: string } };
  };

  if (anyErr?.statusCode === 404 || anyErr?.status === 404) return true;

  const code = anyErr?.code ?? anyErr?.body?.error?.code;
  if (code && /itemnotfound/i.test(code)) return true;

  const message = anyErr?.message ?? anyErr?.body?.error?.message;
  if (message && /(\b404\b|not found|itemnotfound)/i.test(message)) return true;

  return false;
}

export function createSPFxOneDriveAppDataService(
  graphClient: MSGraphClientV3
): SPFxOneDriveAppDataService {
  return {
    read: async <T = unknown>(fileName: string, folder?: string): Promise<SPFxOneDriveAppDataReadResult<T>> => {
      try {
        const apiPath = buildOneDriveAppDataPath(fileName, folder);
        const fileContent = await graphClient.api(apiPath).get();

        if (typeof fileContent === 'string') {
          return {
            data: JSON.parse(fileContent) as T,
            isNotFound: false,
          };
        }

        return {
          data: fileContent as T,
          isNotFound: false,
        };
      } catch (err) {
        if (isSPFxGraphNotFoundError(err)) {
          return {
            data: undefined,
            isNotFound: true,
          };
        }

        throw err;
      }
    },

    write: async <T = unknown>(fileName: string, content: T, folder?: string): Promise<void> => {
      const apiPath = buildOneDriveAppDataPath(fileName, folder);
      await graphClient
        .api(apiPath)
        .header('Content-Type', 'application/json')
        .put(JSON.stringify(content));
    },
  };
}
```

- [ ] **Step 2: Remap `useSPFxOneDriveAppData` to service**

In `src/hooks/useSPFxOneDriveAppData.ts`:

- Import `createSPFxOneDriveAppDataService`.
- Remove local `buildApiPath` and `isNotFoundError`.
- Create service from `client`.
- `write` calls `service.write(fileName, content, folder)`.
- `load` calls `service.read<T>(fileName, folder)`.
- Keep hook-owned:
  - `data`
  - `isLoading`
  - `error`
  - `isWriting`
  - `writeError`
  - `isNotFound`
  - `autoFetch`
  - `createIfMissing` effect
  - mounted ref

Preserve existing user-facing error semantics, including parse JSON errors.

- [ ] **Step 3: Create user photo service**

Create `src/services/spfx-user-photo.service.ts`:

```ts
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { buildUserPhotoEndpoint } from '../helpers/spfx-graph-path.helpers';

export interface SPFxUserPhotoServiceOptions {
  readonly userId?: string;
  readonly email?: string;
  readonly size?: string;
}

export interface SPFxUserPhotoService {
  readonly getPhotoBlob: (options?: SPFxUserPhotoServiceOptions) => Promise<Blob>;
}

export function createSPFxUserPhotoService(
  graphClient: MSGraphClientV3
): SPFxUserPhotoService {
  return {
    getPhotoBlob: async (options?: SPFxUserPhotoServiceOptions): Promise<Blob> => {
      return graphClient.api(buildUserPhotoEndpoint(options)).get() as Promise<Blob>;
    },
  };
}
```

- [ ] **Step 4: Remap `useSPFxUserPhoto` to service**

In `src/hooks/useSPFxUserPhoto.ts`:

- Import `createSPFxUserPhotoService`.
- Remove local `buildPhotoEndpoint`.
- Create service from `graphClient`.
- `load` calls:

```ts
const blob = await service.getPhotoBlob({ userId, email, size });
```

- Keep hook-owned:
  - `photoUrl`
  - `photoBlob`
  - `isLoading`
  - `error`
  - blob URL creation
  - `URL.revokeObjectURL`
  - mounted ref
  - `autoFetch`

Preserve enhanced error messages for 404, 403, and 401.

- [ ] **Step 5: Export Graph services**

Update `src/services/index.ts`:

```ts
export * from './spfx-onedrive-app-data.service';
export * from './spfx-user-photo.service';
```

- [ ] **Step 6: Validate Graph slice**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/services/spfx-onedrive-app-data.service.ts src/services/spfx-user-photo.service.ts src/hooks/useSPFxOneDriveAppData.ts src/hooks/useSPFxUserPhoto.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 9: Review Internal Helper Placement and Remove Duplication

**Files:**
- Review:
  - `src/hooks/useSPFxPnPList.query.internal.ts`
  - `src/hooks/useSPFxPnPList.batch.internal.ts`
  - `src/hooks/useSPFxPnPSearch.query.internal.ts`
  - `src/hooks/useSPFxPnPSearch.results.internal.ts`
  - `src/hooks/useSPFxTenantKeyValueStore.serialization.internal.ts`
  - `src/hooks/useSPFxTenantKeyValueStore.sharepoint.internal.ts`
- Optional move/create:
  - `src/services/spfx-pnp-list-query.helpers.ts`
  - `src/services/spfx-pnp-list-batch.helpers.ts`
  - `src/services/spfx-pnp-search-query.helpers.ts`
  - `src/services/spfx-pnp-search-results.helpers.ts`

- [ ] **Step 1: Search for duplicated logic**

Run:

```bash
rg -n "serializeValue|deserializeValue|escapeODataValue|buildApiPath|buildPhotoEndpoint|createMonitoredListQuery|parseSearchResults|parseSearchRefiners|getContainerSize" src
```

Expected:

- No duplicate implementations remain outside services/helpers/internal modules.
- Hooks import services/helpers instead of owning copied logic.

- [ ] **Step 2: Decide whether current hook internals should move**

If a helper is needed by both public service and hook, move it out of `src/hooks/*.internal.ts` into `src/helpers` or `src/services`.

Recommended cleanup:

- Move PnP list query/batch internals from `src/hooks/` to `src/services/` because `SPFxPnPListService` is the primary owner.
- Move PnP search query/results internals from `src/hooks/` to `src/services/` because `SPFxPnPSearchService` is the primary owner.
- Replace `src/hooks/useSPFxTenantKeyValueStore.serialization.internal.ts` with `src/helpers/spfx-tenant-value.helpers.ts`.
- Replace `src/hooks/useSPFxTenantKeyValueStore.sharepoint.internal.ts` with `src/services/spfx-tenant-key-value-store.service.ts`.

- [ ] **Step 3: Remove obsolete internal files only after imports are gone**

Before deleting any internal file, run:

```bash
rg -n "useSPFxTenantKeyValueStore\\.serialization\\.internal|useSPFxTenantKeyValueStore\\.sharepoint\\.internal|useSPFxPnPList\\.query\\.internal|useSPFxPnPList\\.batch\\.internal|useSPFxPnPSearch\\.query\\.internal|useSPFxPnPSearch\\.results\\.internal" src
```

Expected:

- Only files intentionally retained still appear.
- Delete only obsolete files that have no imports.

- [ ] **Step 4: Validate no duplicate logic remains**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/services src/helpers src/hooks --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 10: Documentation and Package Verification

**Files:**
- Modify: `docs/api/hooks/*.md` only where examples mention internal implementation details.
- Create: `docs/api/services/INDEX.md`
- Create: `docs/api/helpers/INDEX.md`
- Modify: `docs/INDEX.md`
- Modify: `README.md` only if public API overview needs updating.

- [ ] **Step 1: Create services API docs**

Create `docs/api/services/INDEX.md` with sections for:

- PnP context service
- PnP service
- PnP list service
- PnP search service
- App catalog service
- Tenant property service
- Tenant key-value store service
- OneDrive app data service
- User photo service

Each section should state:

- The service is non-React.
- The caller must provide already-initialized SPFx/PnP/Graph dependencies.
- Hooks remain the recommended React API.

- [ ] **Step 2: Create helpers API docs**

Create `docs/api/helpers/INDEX.md` with sections for:

- Page context mapping helpers
- Permission helpers
- Container helpers
- Storage helpers
- Tenant value helpers
- Graph path helpers
- Theme helpers

Each section should state:

- Helpers are pure functions.
- Helpers do not read React context or provider state.

- [ ] **Step 3: Update docs index**

In `docs/INDEX.md`, add links to:

```md
- [Services](./api/services/INDEX.md)
- [Helpers](./api/helpers/INDEX.md)
```

- [ ] **Step 4: Verify package dry-run includes new artifacts**

Run:

```bash
npm run build
npm pack --dry-run
```

Expected:

- Build exits `0`.
- Dry-run output includes generated files under:
  - `lib/services/`
  - `lib/helpers/`

---

### Task 11: Final Declaration Diff and Regression Check

**Files:**
- Generated output outside repo:
  - `/tmp/spfx-react-toolkit-api-services-after`

- [ ] **Step 1: Generate post-refactor declaration output**

Run:

```bash
npx tsc --emitDeclarationOnly --declaration --declarationMap false --outDir /tmp/spfx-react-toolkit-api-services-after --pretty false
```

Expected: exits `0`.

- [ ] **Step 2: Diff existing hook declarations**

Run:

```bash
diff -ru /tmp/spfx-react-toolkit-api-services-before/hooks /tmp/spfx-react-toolkit-api-services-after/hooks
```

Expected:

- Existing public hook declarations should be unchanged except for import path differences caused by shared exported types.
- No hook function names, parameter lists, or return type names should disappear.

- [ ] **Step 3: Verify new declarations exist**

Run:

```bash
find /tmp/spfx-react-toolkit-api-services-after/services /tmp/spfx-react-toolkit-api-services-after/helpers -type f -name '*.d.ts' -print
```

Expected:

- Declarations exist for all new service and helper modules.

- [ ] **Step 4: Run full validation**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src --ext .ts,.tsx --max-warnings=0
npm run build
npm pack --dry-run
```

Expected: all commands exit `0`.

- [ ] **Step 5: Review final worktree**

Run:

```bash
git status --short
git diff --stat
```

Expected:

- Changes are limited to services, helpers, remapped hooks, docs, barrels, and package files.
- No unrelated formatting churn.

---

## Execution Order Recommendation

Implement in this order:

1. Task 1 baseline.
2. Task 2 barrels/package surface.
3. Task 3 helpers and pure hook remaps.
4. Task 4 PnP context/generic services.
5. Task 5 PnP list service and hook remap.
6. Task 6 PnP search service and hook remap.
7. Task 7 tenant/app catalog services and hook remap.
8. Task 8 Graph services and hook remap.
9. Task 9 cleanup duplicate internals.
10. Task 10 docs/package verification.
11. Task 11 final declaration and regression check.

## Self-Review

- Spec coverage: The plan covers all services previously selected for public exposure and all helpers selected for public exposure. It explicitly excludes all hooks/features we decided not to service-ize.
- Public compatibility: Existing hooks remain public facades with the same names and signatures.
- Duplication control: Each hook touched by service/helper extraction must consume the new service/helper internally.
- Jotai: intentionally unchanged in this plan and isolated for a future refactor.
- Risk: This is broader than a patch-level cleanup because it adds new public API and package artifacts; target a minor release.

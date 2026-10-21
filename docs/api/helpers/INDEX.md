# Helpers API Reference

> Pure functions for mapping SPFx context data, checking permissions, building keys and Graph paths, serializing tenant values, and converting themes.

Helpers are public non-React utilities exported from `@apvee/spfx-react-toolkit`. They do not read React context, do not perform I/O, and do not manage provider state.

## Styles and useSx

See [Styles and useSx](./styles.md) for every qualified layout, dimension, spacing, color, preset, typography, border, overflow and scrollbar member, plus states and container/viewport scopes. Descriptors are immutable pure data; the React hook composes them through the host Griffel renderer.

## When To Use Helpers

Use helpers when you already have the required SPFx input and want the same transformation logic used by the hooks.

| API type | Use when | Dependency |
|----------|----------|------------|
| Hooks | A React component is wrapped by an SPFx provider | React provider context |
| Helpers | You need pure mapping, formatting, or conversion logic | Direct input values |
| Services | You need reusable non-React I/O operations | SPFx clients, PnPjs, or Graph clients |

## Quick Examples

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

## Page Context Helpers

Import:

```ts
import {
  getSPFxCorrelationInfo,
  getSPFxEnvironmentInfo,
  getSPFxListInfo,
  getSPFxLocaleInfo,
  getSPFxPageTypeInfo,
  getSPFxSiteInfo,
  getSPFxUserInfo,
} from '@apvee/spfx-react-toolkit';
```

### `getSPFxUserInfo`

```ts
function getSPFxUserInfo(pageContext: PageContext): {
  readonly loginName: string;
  readonly displayName: string;
  readonly email?: string;
  readonly isExternal: boolean;
}
```

Maps `pageContext.user` to the toolkit user shape. `isExternal` falls back to `false` when SPFx does not expose the guest flag.

### `getSPFxSiteInfo`

```ts
function getSPFxSiteInfo(pageContext: PageContext): {
  readonly webId: string;
  readonly webUrl: string;
  readonly webServerRelativeUrl: string;
  readonly title: string;
  readonly languageId: number;
  readonly logoUrl?: string;
  readonly siteId: string;
  readonly siteUrl: string;
  readonly siteServerRelativeUrl: string;
  readonly siteClassification?: string;
  readonly siteGroup?: { readonly id: string; readonly isPublic: boolean };
}
```

Maps SPFx site and web metadata. `siteClassification`, `logoUrl`, and `siteGroup` are optional because SPFx does not always provide them.

### `getSPFxListInfo`

```ts
function getSPFxListInfo(pageContext: PageContext): {
  readonly id: string;
  readonly title: string;
  readonly serverRelativeUrl: string;
  readonly baseTemplate?: number;
  readonly isDocumentLibrary?: boolean;
} | undefined;
```

Returns `undefined` when the current page has no list context. `isDocumentLibrary` is `true` when `baseTemplate === 101`.

### `getSPFxLocaleInfo`

```ts
function getSPFxLocaleInfo(pageContext: PageContext): {
  readonly locale: string;
  readonly uiLocale: string;
  readonly timeZone: {
    readonly id: number;
    readonly offset: number;
    readonly description: string;
    readonly daylightOffset: number;
    readonly standardOffset: number;
  } | undefined;
  readonly ianaTimeZone?: string;
  readonly isRtl: boolean;
}
```

Maps culture, UI culture, timezone, and right-to-left metadata from `PageContext`. `getSPFxLocaleInfo` returns optional readonly `ianaTimeZone` in addition to the unchanged raw `timeZone` object, preserved by reference. No parameters were added.

`ianaTimeZone` is derived solely from `timeZone.id` through a private static lookup. It represents the standard numeric group using CLDR territory `001`; it does not identify a physical city, infer from culture/offset/description, or perform network calls. Group 4 is `Europe/Berlin`, 13 is `America/Los_Angeles`, and 17 is `Pacific/Auckland`. Missing, noninteger, nonfinite, unknown and excluded IDs return undefined, which means unavailable mapping rather than UTC.

The audited 111-candidate inventory includes 104 verified groups; IDs 7, 16, 30, 32, 41, 81 and 92 are excluded. It is not a promise of complete current SharePoint Online coverage. SharePoint Server permits customized `TIMEZONE.XML`; standard IDs cannot prove equivalence to those rules. IANA aliases from the pinned CLDR release are retained, and consumers require suitable Intl time-zone data.

The [versioned audit fixture](../../../tests/fixtures/spfx-time-zone-mapping.json) records every candidate's status/reason, source revisions, snapshot SHA-256 values and license notices. Sources are [PnP Core revision b6ff7b3](https://raw.githubusercontent.com/pnp/pnpcore/b6ff7b30dc8a1dfa8202bca50b75a08e3d3d6c43/src/sdk/PnP.Core/Model/SharePoint/Core/Internal/TimeZone.cs), [CLDR release 48 revision acd6d88](https://raw.githubusercontent.com/unicode-org/cldr/acd6d88ae493633240e19a87a721076a8a75c310/common/supplemental/windowsZones.xml), and [Microsoft numeric groups](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-wssfo2/4b41f00e-6782-40f0-9db2-172502446d4c) published 2024-01-17 and snapshotted 2026-10-08. PnP MIT and Unicode-3.0 notices accompany the fixture and derived source.

```tsx
const info = getSPFxLocaleInfo(pageContext);
// Explicit policy: no site-zone date when the standard mapping is unavailable.
const fullDate = info.ianaTimeZone
  ? new Intl.DateTimeFormat(info.uiLocale, { timeZone: info.ianaTimeZone, dateStyle: 'full' }).format(modified)
  : 'Site time zone unavailable';
```

Alternatively, passing undefined to [RelativeTime](../controls/relative-time.md) uses the runtime zone; label that fallback explicitly. The [locale hook](../hooks/environment.md#usespfxlocaleinfo) exposes the same result under an SPFx provider.

### `getSPFxEnvironmentInfo`

```ts
function getSPFxEnvironmentInfo(pageContext: PageContext): {
  readonly type: 'Local' | 'SharePoint' | 'SharePointOnPrem' | 'Teams' | 'Office' | 'Outlook';
  readonly isLocal: boolean;
  readonly isWorkbench: boolean;
  readonly isSharePoint: boolean;
  readonly isSharePointOnPrem: boolean;
  readonly isTeams: boolean;
  readonly isOffice: boolean;
  readonly isOutlook: boolean;
}
```

Detects host environment from the page URL, legacy page context, and available SPFx SDK objects.

### `getSPFxPageTypeInfo`

```ts
function getSPFxPageTypeInfo(pageContext: PageContext): {
  readonly pageType: 'sitePage' | 'webPartPage' | 'listPage' | 'listFormPage' | 'profilePage' | 'searchPage' | 'unknown';
  readonly isModernPage: boolean;
  readonly isSitePage: boolean;
  readonly isListPage: boolean;
  readonly isListFormPage: boolean;
  readonly isWebPartPage: boolean;
}
```

Detects modern and legacy SharePoint page types. List form pages are detected when list context and form metadata are present.

### `getSPFxCorrelationInfo`

```ts
function getSPFxCorrelationInfo(pageContext: PageContext): {
  readonly correlationId: string | undefined;
  readonly tenantId: string | undefined;
}
```

Extracts request correlation and tenant identifiers when SPFx exposes them.

## Permission Helpers

Import:

```ts
import { hasSPFxPermission } from '@apvee/spfx-react-toolkit';
```

### `hasSPFxPermission`

```ts
function hasSPFxPermission(
  permissionSet: SPPermission | undefined,
  permission: SPPermission
): boolean;
```

Returns `false` when `permissionSet` is undefined. Otherwise delegates to `SPPermission.hasPermission(permission)`.

## API Permission Precheck Helpers

Import:

```ts
import {
  buildSPFxApiPermissionAdminMessage,
  classifySPFxApiPermissionError,
  createSPFxApiPermissionRequirementId,
  decodeSPFxJwtPayload,
  evaluateSPFxApiPermissionRequirement,
  extractSPFxDelegatedScopes,
  normalizeSPFxApiPermissionRequirements,
  summarizeSPFxApiPermissionResults,
  validateSPFxApiPermissionRequirement,
} from '@apvee/spfx-react-toolkit';
```

These helpers power the API permission precheck hooks and services without reading React context or acquiring tokens themselves. They normalize Graph and custom API configuration, decode delegated `scp` scopes from JWT payloads, classify token acquisition errors, and summarize results for UI buckets and remediation.

| Function | Purpose |
|----------|---------|
| `normalizeSPFxApiPermissionRequirements` | Expands Graph shorthand, custom API shorthand, and explicit requirements into deduplicated delegated-scope requirements. |
| `buildSPFxApiPermissionAdminMessage` | Builds administrator-facing remediation text for a package resource and scope. |
| `createSPFxApiPermissionRequirementId` | Creates stable requirement identifiers such as `graph:Sites.Read.All` or custom API IDs. |
| `validateSPFxApiPermissionRequirement` | Returns an invalid-result object when a requirement is incomplete or uses an unsupported permission kind. |
| `decodeSPFxJwtPayload` | Decodes the payload section of a JWT without validating or exposing the raw token. |
| `extractSPFxDelegatedScopes` | Reads delegated scopes from the JWT `scp` claim. |
| `evaluateSPFxApiPermissionRequirement` | Evaluates one requirement against a decoded token payload and returns a check result. |
| `classifySPFxApiPermissionError` | Maps token acquisition failures to precheck statuses such as consent, interaction, timeout, and transient failure. |
| `summarizeSPFxApiPermissionResults` | Groups check results into `available`, `missing`, `warnings`, and `unknown` summaries with an overall configuration state. |

`normalizeSPFxApiPermissionRequirements` also carries `packageResource` and `scope` into each requirement so remediation can point back to the same resource/scope values used in `webApiPermissionRequests`.

## Container Helpers

Import:

```ts
import { getSPFxContainerSize } from '@apvee/spfx-react-toolkit';
```

### `getSPFxContainerSize`

```ts
function getSPFxContainerSize(
  width: number
): 'small' | 'medium' | 'large' | 'xLarge' | 'xxLarge' | 'xxxLarge';
```

| Width | Result |
|-------|--------|
| `< 480` | `small` |
| `480-639` | `medium` |
| `640-1023` | `large` |
| `1024-1365` | `xLarge` |
| `1366-1919` | `xxLarge` |
| `>= 1920` | `xxxLarge` |

## Storage Helpers

Import:

```ts
import { createScopedSPFxStorageKey } from '@apvee/spfx-react-toolkit';
```

### `createScopedSPFxStorageKey`

```ts
function createScopedSPFxStorageKey(instanceId: string, key: string): string;
```

Builds keys in the exact format `spfx:<instanceId>:<key>`.

## Tenant Value Helpers

Import:

```ts
import {
  deserializeTenantValue,
  escapeODataValue,
  serializeTenantValue,
} from '@apvee/spfx-react-toolkit';
```

### `escapeODataValue`

```ts
function escapeODataValue(value: string): string;
```

Escapes single quotes for OData string filters by replacing `'` with `''`.

> These public helpers preserve the historical format. New tenant and site list-store writes additionally quote empty/JSON-like strings and bigint decimal text; see [list-store saved-data compatibility](../hooks/storage.md#serialization-and-saved-data-compatibility). Historical stored values and StorageEntity reads retain their existing interpretation, including each list store's documented legacy REST-null mapping.

### `serializeTenantValue`

```ts
function serializeTenantValue(value: unknown): string;
```

Serializes tenant-scoped values for SharePoint string storage:

| Input | Output |
|-------|--------|
| `null` | `'null'` |
| `Date` | ISO string |
| `string`, `number`, `boolean`, `bigint` | `String(value)` |
| object/array | `JSON.stringify(value)` when available |
| non-JSON fallback | `String(value)` |

### `deserializeTenantValue`

```ts
function deserializeTenantValue<T>(rawValue: string): T;
```

Parses JSON when possible. If parsing fails, returns the raw string cast to `T`. This legacy format is not a complete type round trip: numeric/boolean/null-like strings become JSON primitives, bigints may lose precision when parsed as numbers, and dates remain strings. `T` does not perform runtime validation; validate deserialized values in the caller. No new envelope or stored-data migration is introduced.

## Graph Path Helpers

Import:

```ts
import {
  buildOneDriveAppDataPath,
  buildUserPhotoEndpoint,
} from '@apvee/spfx-react-toolkit';
```

### `buildOneDriveAppDataPath`

```ts
function buildOneDriveAppDataPath(fileName: string, folder?: string): string;
```

Builds Microsoft Graph app-root content paths:

| Input | Output |
|-------|--------|
| `('settings.json')` | `/me/drive/special/approot:/settings.json:/content` |
| `('settings.json', 'dashboard')` | `/me/drive/special/approot:/dashboard/settings.json:/content` |

Folder names are sanitized by replacing characters outside `a-z`, `A-Z`, `0-9`, `-`, and `_` with `-`.

### `buildUserPhotoEndpoint`

```ts
function buildUserPhotoEndpoint(options?: {
  readonly userId?: string;
  readonly email?: string;
  readonly size?: '48x48' | '64x64' | '96x96' | '120x120' | '240x240' | '360x360' | '432x432' | '504x504' | '648x648';
}): string;
```

Builds Microsoft Graph profile photo endpoints:

| Options | Base path |
|---------|-----------|
| `{ userId }` | `/users/{userId}` |
| `{ email }` | `/users/{email}` |
| omitted | `/me` |

The default size is `240x240`.

## Theme Helpers

Import:

```ts
import {
  createFluent9ThemeFromSPFxTheme,
  getTeamsFluentTheme,
} from '@apvee/spfx-react-toolkit';
```

### `createFluent9ThemeFromSPFxTheme`

```ts
function createFluent9ThemeFromSPFxTheme(spfxTheme: IReadonlyTheme | undefined): Theme;
```

Converts an SPFx Fluent UI 8 theme to a Fluent UI 9 theme. When `spfxTheme` is undefined, returns `webLightTheme`.

The existing Fluent migration shim supplies the conversion. Two accessible neutral stroke state tokens use the host palette so interactions remain observable:

| Fluent UI 9 token | SPFx palette source |
| --- | --- |
| `colorNeutralStrokeAccessible` | Existing shim mapping (`neutralSecondary`) |
| `colorNeutralStrokeAccessibleHover` | `neutralPrimary` |
| `colorNeutralStrokeAccessiblePressed` | `neutralDark` |

Every other converted token keeps the shim's mapping. Missing, empty or whitespace-only hover/pressed palette values retain the corresponding converted fallback. The helper does not mutate the input theme. Custom/inverted palettes use their supplied colors; palettes with equal colors can still have equal state tokens and require their own contrast/feedback checks. Teams themes returned by `getTeamsFluentTheme` remain the original shared theme objects. The existing Fluent theme/migration peers supply the conversion; no extra dependency is required.

```tsx
const theme = createFluent9ThemeFromSPFxTheme(spfxTheme);
return <FluentProvider theme={theme}>{children}</FluentProvider>;
```

For example, the default Fluent UI 8 host palette provides `#605e5c` for the base accessible stroke, `#323130` for hover and `#201f1e` for pressed. [`scrollbar.fluent`](./styles.md#overflow-and-the-single-scrollbar-recipe) consumes these scoped tokens for its thumb feedback. `useSPFxFluent9ThemeInfo` uses this conversion in SharePoint and keeps its existing Teams/fallback selection order.

### `getTeamsFluentTheme`

```ts
function getTeamsFluentTheme(teamsThemeName: string | undefined): Theme;
```

| Teams theme name | Fluent UI 9 theme |
|------------------|-------------------|
| `dark` | `teamsDarkTheme` |
| `contrast`, `highcontrast` | `teamsHighContrastTheme` |
| `default`, undefined, other values | `teamsLightTheme` |

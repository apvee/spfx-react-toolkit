# Helpers API Reference

> Pure functions for mapping SPFx context data, checking permissions, building keys and Graph paths, serializing tenant values, and converting themes.

Helpers are public non-React utilities exported from `@apvee/spfx-react-toolkit`. They do not read React context, do not perform I/O, and do not manage provider state.

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
  readonly isRtl: boolean;
}
```

Maps culture, UI culture, timezone, and right-to-left metadata from `PageContext`.

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

Parses JSON when possible. If parsing fails, returns the raw string cast to `T`.

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

### `getTeamsFluentTheme`

```ts
function getTeamsFluentTheme(teamsThemeName: string | undefined): Theme;
```

| Teams theme name | Fluent UI 9 theme |
|------------------|-------------------|
| `dark` | `teamsDarkTheme` |
| `contrast`, `highcontrast` | `teamsHighContrastTheme` |
| `default`, undefined, other values | `teamsLightTheme` |

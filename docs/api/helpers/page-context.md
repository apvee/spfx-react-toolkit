# Page Context Helpers

[Helpers overview](../INDEX.md#helpers-and-styles) · [Non-React composition](../../guides/helpers-and-services.md)

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

## `getSPFxUserInfo`

```ts
function getSPFxUserInfo(pageContext: PageContext): {
  readonly loginName: string;
  readonly displayName: string;
  readonly email?: string;
  readonly isExternal: boolean;
}
```

Maps `pageContext.user` to the toolkit user shape. `isExternal` falls back to `false` when SPFx does not expose the guest flag.

## `getSPFxSiteInfo`

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

Maps SPFx site and web metadata, serializing GUIDs as strings. Missing web language defaults to `1033`; a supplied group with no visibility flag defaults to `isPublic: false`. `siteClassification`, `logoUrl`, and `siteGroup` are optional because SPFx does not always provide them.

## `getSPFxListInfo`

```ts
function getSPFxListInfo(pageContext: PageContext): {
  readonly id: string;
  readonly title: string;
  readonly serverRelativeUrl: string;
  readonly baseTemplate?: number;
  readonly isDocumentLibrary?: boolean;
} | undefined;
```

Returns `undefined` when there is no list or no list ID. Missing title defaults to `Unknown List` and missing server-relative URL to an empty string. `isDocumentLibrary` is `true` when `baseTemplate === 101`, and false for any other or missing template.

## `getSPFxLocaleInfo`

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

Maps formatting culture, UI culture, raw time-zone settings and right-to-left metadata from the required `PageContext`. The raw `timeZone` object is preserved by reference; `ianaTimeZone` is an optional representative for Intl formatting.

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

## `getSPFxEnvironmentInfo`

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

Maps only the supplied `PageContext`: URL substrings identify local/workbench, `legacyPageContext.isOnPremises` identifies on-premises, and SDK-shaped fields on that object set Teams/Office/Outlook flags. Category precedence is Local, Teams, Outlook, Office, SharePointOnPrem, SharePoint. Flags can overlap. This synchronous heuristic does not acquire SDK context or read the provider’s native context; use [`useSPFxEnvironmentInfo`](../hooks/environment.md#usespfxenvironmentinfo) for the asynchronous host-aware React result.

## `getSPFxPageTypeInfo`

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

Checks modern `page.type` for site/web-part pages, then legacy `pageType` for site/web-part/list/profile/search categories, then legacy `listId` as a list fallback. A present legacy `formType` (including zero) makes a list form page. `isModernPage` means the chosen type is `sitePage`; it is not an independent host capability check. Unrecognized metadata returns `unknown` with false category flags.

## `getSPFxCorrelationInfo`

```ts
function getSPFxCorrelationInfo(pageContext: PageContext): {
  readonly correlationId: string | undefined;
  readonly tenantId: string | undefined;
}
```

Extracts request correlation and tenant identifiers when SPFx exposes them.

## Inputs, return shapes and usage

Every function requires a caller-supplied SPFx `PageContext`; these synchronous helpers neither read React context nor acquire asynchronous Teams SDK context. They return fresh objects with readonly TypeScript fields and do not mutate the input; the return objects are not frozen at runtime. The locale `timeZone` field deliberately preserves the host object's reference. Missing optional metadata has the defaults described above; the helpers do not validate a malformed PageContext or catch errors from its required fields/GUID conversions. The descriptive source aliases for nested group/time-zone and category shapes are not separately exported names; the inline return contracts above are public.

```ts
const user = getSPFxUserInfo(pageContext);
const site = getSPFxSiteInfo(pageContext);
const list = getSPFxListInfo(pageContext);
const locale = getSPFxLocaleInfo(pageContext);
const environment = getSPFxEnvironmentInfo(pageContext);
const page = getSPFxPageTypeInfo(pageContext);
const correlation = getSPFxCorrelationInfo(pageContext);
console.log(user.displayName, site.siteUrl, list?.title, locale.uiLocale);
console.log(environment.type, page.pageType, correlation.correlationId);
```

`correlationId` comes from `pageContext.site.correlationId`; `tenantId` comes from `pageContext.aadInfo.tenantId`. Missing identifiers return undefined. UI resource culture, regional formatting and IANA time-zone selection have separate responsibilities; see [localization](../../guides/localization.md).

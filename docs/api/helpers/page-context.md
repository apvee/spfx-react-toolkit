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

Maps SPFx site and web metadata. `siteClassification`, `logoUrl`, and `siteGroup` are optional because SPFx does not always provide them.

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

Returns `undefined` when the current page has no list context. `isDocumentLibrary` is `true` when `baseTemplate === 101`.

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
  readonly isRtl: boolean;
}
```

Maps culture, UI culture, timezone, and right-to-left metadata from `PageContext`.

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

Detects host environment from the page URL, legacy page context, and available SPFx SDK objects.

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

Detects modern and legacy SharePoint page types. List form pages are detected when list context and form metadata are present.

## `getSPFxCorrelationInfo`

```ts
function getSPFxCorrelationInfo(pageContext: PageContext): {
  readonly correlationId: string | undefined;
  readonly tenantId: string | undefined;
}
```

Extracts request correlation and tenant identifiers when SPFx exposes them.

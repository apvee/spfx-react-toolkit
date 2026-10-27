import type { PageContext } from '@microsoft/sp-page-context';
import { resolveSPFxIanaTimeZone } from './spfx-time-zone.internal';

/**
 * Microsoft 365 group metadata returned only when the site exposes an associated group.
 *
 * @example
 * ```ts
 * const group = getSPFxSiteInfo(pageContext).siteGroup;
 * if (group) console.log(group.id, group.isPublic);
 * ```
 */
type SPFxGroupInfoResult = {
  /**
   * Associated Microsoft 365 group identifier, serialized as a string.
   */
  readonly id: string;
  /**
   * Group visibility flag; missing host metadata defaults to false.
   */
  readonly isPublic: boolean;
};

/**
 * Raw SharePoint regional time-zone metadata preserved from the host. Numeric offsets describe SharePoint settings; prefer the separately resolved IANA zone for Intl formatting.
 *
 * @example
 * ```ts
 * const zone = getSPFxLocaleInfo(pageContext).timeZone;
 * if (zone) console.log(zone.id, zone.description);
 * ```
 */
type SPFxTimeZoneResult = {
  /**
   * SharePoint numeric regional time-zone group ID, used for the optional IANA lookup.
   */
  readonly id: number;
  /**
   * Raw regional offset supplied by SharePoint; preserved without normalization.
   */
  readonly offset: number;
  /**
   * Host-provided description of the regional time-zone setting.
   */
  readonly description: string;
  /**
   * Raw daylight offset supplied by SharePoint; preserved without normalization.
   */
  readonly daylightOffset: number;
  /**
   * Raw standard offset supplied by SharePoint; preserved without normalization.
   */
  readonly standardOffset: number;
};

/**
 * Best-effort host category derived by `getSPFxEnvironmentInfo` from supplied PageContext URLs, legacy flags and SDK-shaped fields. It is not asynchronous Teams host detection.
 *
 * @example
 * ```ts
 * const host = getSPFxEnvironmentInfo(pageContext).type;
 * ```
 */
type SPFxEnvironmentTypeResult =
  | 'Local'
  | 'SharePoint'
  | 'SharePointOnPrem'
  | 'Teams'
  | 'Office'
  | 'Outlook';

/**
 * Best-effort page classification from modern and legacy page metadata. Unknown/missing metadata produces `unknown`.
 *
 * @example
 * ```ts
 * const type = getSPFxPageTypeInfo(pageContext).pageType;
 * ```
 */
type SPFxPageTypeResult =
  | 'sitePage'
  | 'webPartPage'
  | 'listPage'
  | 'listFormPage'
  | 'profilePage'
  | 'searchPage'
  | 'unknown';

/**
 * Maps SPFx PageContext user data to the toolkit user info shape.
 *
 * @param pageContext - SPFx page context
 * @returns User information
 *
 * @example
 * ```ts
 * const info = getSPFxUserInfo(pageContext);
 * ```
 */
export function getSPFxUserInfo(pageContext: PageContext): {
  /**
   * Current user login name as supplied by PageContext.
   */
  readonly loginName: string;
  /**
   * Current user display name as supplied by PageContext.
   */
  readonly displayName: string;
  /**
   * User email when the host supplies it; otherwise undefined.
   */
  readonly email?: string;
  /**
   * Guest-user flag; defaults to false when the host omits it.
   */
  readonly isExternal: boolean;
} {
  const user = pageContext.user;

  return {
    loginName: user.loginName,
    displayName: user.displayName,
    email: user.email,
    isExternal: user.isExternalGuestUser ?? false,
  };
}

/**
 * Maps SPFx PageContext site and web data to the toolkit site info shape.
 *
 * @param pageContext - SPFx page context
 * @returns Site collection and web information
 *
 * @example
 * ```ts
 * const info = getSPFxSiteInfo(pageContext);
 * ```
 */
export function getSPFxSiteInfo(pageContext: PageContext): {
  /**
   * Current web GUID, serialized as a string.
   */
  readonly webId: string;
  /**
   * Absolute URL of the current web.
   */
  readonly webUrl: string;
  /**
   * Server-relative path of the current web.
   */
  readonly webServerRelativeUrl: string;
  /**
   * Current web title.
   */
  readonly title: string;
  /**
   * Web language LCID; defaults to 1033 when missing.
   */
  readonly languageId: number;
  /**
   * Optional host web logo URL; undefined when unavailable.
   */
  readonly logoUrl?: string;
  /**
   * Site collection GUID, serialized as a string.
   */
  readonly siteId: string;
  /**
   * Absolute URL of the site collection.
   */
  readonly siteUrl: string;
  /**
   * Server-relative path of the site collection.
   */
  readonly siteServerRelativeUrl: string;
  /**
   * Legacy classification label when supplied; undefined otherwise.
   */
  readonly siteClassification?: string;
  /**
   * Associated group metadata, or undefined when the site has no group.
   */
  readonly siteGroup?: SPFxGroupInfoResult;
} {
  const siteObj = pageContext.site;
  const webObj = pageContext.web;
  const legacy = (pageContext as unknown as {
    legacyPageContext?: {
      /**
       * Legacy classification label when supplied; undefined otherwise.
       */
      siteClassification?: string;
    };
  }).legacyPageContext;

  return {
    webId: webObj.id.toString(),
    webUrl: webObj.absoluteUrl,
    webServerRelativeUrl: webObj.serverRelativeUrl,
    title: webObj.title,
    languageId: webObj.language ?? 1033,
    logoUrl: (webObj as unknown as { /**
     * Optional host web logo URL; undefined when unavailable.
     */
    logoUrl?: string }).logoUrl,
    siteId: siteObj.id.toString(),
    siteUrl: siteObj.absoluteUrl,
    siteServerRelativeUrl: siteObj.serverRelativeUrl,
    siteClassification: legacy?.siteClassification,
    siteGroup: siteObj.group ? {
      id: siteObj.group.id.toString(),
      isPublic: siteObj.group.isPublic ?? false,
    } : undefined,
  };
}

/**
 * Maps SPFx PageContext list data to the toolkit list info shape.
 *
 * @param pageContext - SPFx page context
 * @returns List information, or undefined when there is no list context
 *
 * @example
 * ```ts
 * const info = getSPFxListInfo(pageContext);
 * ```
 */
export function getSPFxListInfo(pageContext: PageContext): {
  /**
   * Current list GUID, serialized as a string.
   */
  readonly id: string;
  /**
   * Current list title; defaults to Unknown List when the host omits it.
   */
  readonly title: string;
  /**
   * List server-relative URL; defaults to an empty string when missing.
   */
  readonly serverRelativeUrl: string;
  /**
   * Host list template number, or undefined when unavailable.
   */
  readonly baseTemplate?: number;
  /**
   * True only when baseTemplate equals 101; false for other or missing templates.
   */
  readonly isDocumentLibrary?: boolean;
} | undefined {
  const list = (pageContext as unknown as {
    list?: {
      /**
       * Current list GUID, serialized as a string.
       */
      id?: { toString: () => string };
      /**
       * Current list title; defaults to Unknown List when the host omits it.
       */
      title?: string;
      /**
       * List server-relative URL; defaults to an empty string when missing.
       */
      serverRelativeUrl?: string;
      /**
       * Host list template number, or undefined when unavailable.
       */
      baseTemplate?: number;
    };
  }).list;

  if (!list || !list.id) {
    return undefined;
  }

  const baseTemplate = list.baseTemplate;
  const isDocumentLibrary = baseTemplate === 101;

  return {
    id: list.id.toString(),
    title: list.title ?? 'Unknown List',
    serverRelativeUrl: list.serverRelativeUrl ?? '',
    baseTemplate,
    isDocumentLibrary,
  };
}

/**
 * Maps SPFx PageContext culture and regional data to the toolkit locale info shape.
 * The original timeZone object is preserved. The optional IANA value is the
 * audited CLDR territory001 representative for timeZone.id, not a physical city,
 * user profile zone or guarantee of customized SharePoint Server rules.
 * Unknown or unverified IDs return undefined without a browser/offset fallback.
 *
 * @param pageContext - SPFx page context
 * @returns Locale, regional metadata and an optional representative IANA time zone
 * @example
 * ```ts
 * const { uiLocale, ianaTimeZone } = getSPFxLocaleInfo(pageContext);
 * if (ianaTimeZone !== undefined) {
 *   const formatter = new Intl.DateTimeFormat(uiLocale, { timeZone: ianaTimeZone });
 *   const siteDate = formatter.format(new Date());
 * }
 * ```
 */
export function getSPFxLocaleInfo(pageContext: PageContext): {
  /**
   * Current formatting culture name from PageContext.
   */
  readonly locale: string;
  /**
   * Current UI culture name from PageContext.
   */
  readonly uiLocale: string;
  /**
   * Raw host time-zone metadata preserved by reference, or undefined when unavailable.
   */
  readonly timeZone: SPFxTimeZoneResult | undefined;
  /** Audited standard-group IANA representative; undefined for unresolved IDs. */
  readonly ianaTimeZone?: string;
  /**
   * Right-to-left flag from the current host culture.
   */
  readonly isRtl: boolean;
} {
  const cultureInfo = pageContext.cultureInfo;
  const timeZone = (pageContext.web as { timeZoneInfo?: SPFxTimeZoneResult }).timeZoneInfo;

  return {
    locale: cultureInfo.currentCultureName,
    uiLocale: cultureInfo.currentUICultureName,
    timeZone,
    ianaTimeZone: resolveSPFxIanaTimeZone(timeZone?.id),
    isRtl: cultureInfo.isRightToLeft,
  };
}

/**
 * Maps SPFx PageContext host data to the toolkit environment info shape.
 *
 * @param pageContext - SPFx page context
 * @returns Environment information
 *
 * @example
 * ```ts
 * const info = getSPFxEnvironmentInfo(pageContext);
 * ```
 */
export function getSPFxEnvironmentInfo(pageContext: PageContext): {
  /**
   * Primary category with local, Teams, Outlook, Office, on-premises, then SharePoint precedence.
   */
  readonly type: SPFxEnvironmentTypeResult;
  /**
   * Whether the web URL contains localhost or 127.0.0.1; a heuristic, not origin validation.
   */
  readonly isLocal: boolean;
  /**
   * True for local URLs or URLs containing workbench.aspx.
   */
  readonly isWorkbench: boolean;
  /**
   * True when the primary category is SharePoint.
   */
  readonly isSharePoint: boolean;
  /**
   * True when the primary category is SharePointOnPrem.
   */
  readonly isSharePointOnPrem: boolean;
  /**
   * Whether the supplied PageContext exposes a microsoftTeams SDK-shaped field.
   */
  readonly isTeams: boolean;
  /**
   * Whether the supplied PageContext exposes an office SDK-shaped field.
   */
  readonly isOffice: boolean;
  /**
   * Whether the supplied PageContext exposes an outlook SDK-shaped field.
   */
  readonly isOutlook: boolean;
} {
  const legacy = (pageContext as unknown as {
    legacyPageContext?: {
      isSPO?: boolean;
      isOnPremises?: boolean;
      webAbsoluteUrl?: string;
    };
  }).legacyPageContext;

  const sdks = (pageContext as unknown as {
    sdks?: {
      microsoftTeams?: unknown;
      office?: unknown;
      outlook?: unknown;
    };
  }).sdks;

  const isTeams = sdks?.microsoftTeams !== undefined;
  const isOffice = sdks?.office !== undefined;
  const isOutlook = sdks?.outlook !== undefined;
  const webUrl = pageContext.web.absoluteUrl.toLowerCase();
  const isLocal = webUrl.indexOf('localhost') !== -1 ||
                  webUrl.indexOf('127.0.0.1') !== -1;
  const isWorkbench = isLocal ||
                      webUrl.indexOf('workbench.aspx') !== -1 ||
                      webUrl.indexOf('_layouts/15/workbench.aspx') !== -1;
  const isOnPrem = legacy?.isOnPremises ?? false;

  let type: SPFxEnvironmentTypeResult;
  if (isLocal) {
    type = 'Local';
  } else if (isTeams) {
    type = 'Teams';
  } else if (isOutlook) {
    type = 'Outlook';
  } else if (isOffice) {
    type = 'Office';
  } else if (isOnPrem) {
    type = 'SharePointOnPrem';
  } else {
    type = 'SharePoint';
  }

  return {
    type,
    isLocal,
    isWorkbench,
    isSharePoint: type === 'SharePoint',
    isSharePointOnPrem: type === 'SharePointOnPrem',
    isTeams,
    isOffice,
    isOutlook,
  };
}

/**
 * Detects the current SharePoint page type from modern and legacy page context data.
 *
 * @param pageContext - SPFx page context
 * @returns Page type information
 *
 * @example
 * ```ts
 * const info = getSPFxPageTypeInfo(pageContext);
 * ```
 */
export function getSPFxPageTypeInfo(pageContext: PageContext): {
  /**
   * Detected category; unknown when neither modern nor legacy metadata identifies the page.
   */
  readonly pageType: SPFxPageTypeResult;
  /**
   * True for the detected sitePage category; not an independent modern-page capability test.
   */
  readonly isModernPage: boolean;
  /**
   * True when pageType is sitePage.
   */
  readonly isSitePage: boolean;
  /**
   * True when pageType is listPage.
   */
  readonly isListPage: boolean;
  /**
   * True when pageType is listFormPage.
   */
  readonly isListFormPage: boolean;
  /**
   * True when pageType is webPartPage.
   */
  readonly isWebPartPage: boolean;
} {
  const modernPage = (pageContext as unknown as {
    page?: {
      type?: string;
    };
  }).page;

  const legacy = (pageContext as unknown as {
    legacyPageContext?: {
      /**
       * Detected category; unknown when neither modern nor legacy metadata identifies the page.
       */
      pageType?: string;
      listId?: string;
      formType?: string | number;
    };
  }).legacyPageContext;

  let pageType: SPFxPageTypeResult = 'unknown';
  const modernPageType = modernPage?.type?.toLowerCase();

  if (modernPageType) {
    if (modernPageType.indexOf('sitepage') !== -1) {
      pageType = 'sitePage';
    } else if (modernPageType.indexOf('webpartpage') !== -1) {
      pageType = 'webPartPage';
    }
  }

  if (pageType === 'unknown') {
    const legacyPageType = legacy?.pageType?.toLowerCase();
    if (legacyPageType) {
      if (legacyPageType.indexOf('sitepage') !== -1) {
        pageType = 'sitePage';
      } else if (legacyPageType.indexOf('webpartpage') !== -1) {
        pageType = 'webPartPage';
      } else if (legacyPageType.indexOf('list') !== -1) {
        if (legacy?.formType !== undefined && legacy.formType !== null) {
          pageType = 'listFormPage';
        } else {
          pageType = 'listPage';
        }
      } else if (legacyPageType.indexOf('profile') !== -1) {
        pageType = 'profilePage';
      } else if (legacyPageType.indexOf('search') !== -1) {
        pageType = 'searchPage';
      }
    }
  }

  if (pageType === 'unknown') {
    if (legacy?.listId) {
      if (legacy.formType !== undefined && legacy.formType !== null) {
        pageType = 'listFormPage';
      } else {
        pageType = 'listPage';
      }
    }
  }

  const isSitePage = pageType === 'sitePage';
  const isWebPartPage = pageType === 'webPartPage';
  const isListPage = pageType === 'listPage';
  const isListFormPage = pageType === 'listFormPage';
  const isModernPage = isSitePage;

  return {
    pageType,
    isModernPage,
    isSitePage,
    isListPage,
    isListFormPage,
    isWebPartPage,
  };
}

/**
 * Extracts correlation and tenant IDs from SPFx PageContext.
 *
 * @param pageContext - SPFx page context
 * @returns Correlation and tenant identifiers
 *
 * @example
 * ```ts
 * const info = getSPFxCorrelationInfo(pageContext);
 * ```
 */
export function getSPFxCorrelationInfo(pageContext: PageContext): {
  /**
   * Site correlation GUID serialized as a string, or undefined when unavailable.
   */
  readonly correlationId: string | undefined;
  /**
   * AAD tenant GUID serialized as a string, or undefined when unavailable.
   */
  readonly tenantId: string | undefined;
} {
  const correlationId = pageContext.site?.correlationId?.toString();
  const aadInfo = pageContext.aadInfo as unknown as
    { /**
     * AAD tenant GUID serialized as a string, or undefined when unavailable.
     */
    tenantId?: { toString(): string } } | undefined;
  const tenantId = aadInfo?.tenantId?.toString();

  return {
    correlationId,
    tenantId,
  };
}

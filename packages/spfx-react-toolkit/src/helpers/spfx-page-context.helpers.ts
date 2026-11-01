/**
 * Snapshot mappings for supplied SPFx user, site, list, regional and environment metadata.
 *
 * Each helper requires a real PageContext and returns a fresh, unfrozen object.
 * Readonly is a TypeScript constraint; locale timeZone deliberately retains the
 * host reference. No React provider, asynchronous host detection or input
 * validation is added. Errors from malformed required fields/GUID conversion
 * propagate. Optional output properties may be present with undefined values.
 *
 * @packageDocumentation
 * @docPage helpers/page-context
 */
import type { PageContext } from '@microsoft/sp-page-context';
import { resolveSPFxIanaTimeZone } from './spfx-time-zone.internal';

/**
 * Microsoft 365 group metadata returned only when the site exposes an associated group.
 *
 * @example
 * ```ts
 * import type { PageContext } from '@microsoft/sp-page-context';
 * import { getSPFxSiteInfo } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(pageContext: PageContext) {
 *   const group = getSPFxSiteInfo(pageContext).siteGroup;
 *   if (group) console.log(group.id, group.isPublic);
 * }
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
 * import type { PageContext } from '@microsoft/sp-page-context';
 * import { getSPFxLocaleInfo } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(pageContext: PageContext) {
 *   const zone = getSPFxLocaleInfo(pageContext).timeZone;
 *   if (zone) console.log(zone.id, zone.description);
 * }
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
 * Best-effort host category derived by `getSPFxEnvironmentInfo` from supplied PageContext URLs, legacy flags and SDK-shaped fields. It is not asynchronous Teams host detection. Use actual host SDK initialization for authoritative host detection.
 *
 * @example
 * ```ts
 * import type { PageContext } from '@microsoft/sp-page-context';
 * import { getSPFxEnvironmentInfo } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(pageContext: PageContext) {
 *   const host = getSPFxEnvironmentInfo(pageContext).type;
 * }
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
 * import type { PageContext } from '@microsoft/sp-page-context';
 * import { getSPFxPageTypeInfo } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(pageContext: PageContext) {
 *   const type = getSPFxPageTypeInfo(pageContext).pageType;
 * }
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
 * import type { PageContext } from '@microsoft/sp-page-context';
 * import { getSPFxUserInfo } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(pageContext: PageContext) {
 *   const info = getSPFxUserInfo(pageContext);
 * }
 * ```
 * @useWhen You map a supplied PageContext to login, display-name and guest metadata outside React.
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
 * @example Basic usage
 * ```ts
 * import type { PageContext } from '@microsoft/sp-page-context';
 * import { getSPFxSiteInfo } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(pageContext: PageContext) {
 *   const info = getSPFxSiteInfo(pageContext);
 * }
 * ```
 * @useWhen You need both current-web and collection-root metadata from a supplied PageContext.
 *
 *
 * @example Read combined host metadata
 * ```ts
 * import { getSPFxUserInfo, getSPFxSiteInfo, getSPFxListInfo, getSPFxLocaleInfo, getSPFxEnvironmentInfo, getSPFxPageTypeInfo, getSPFxCorrelationInfo } from '@apvee/spfx-react-toolkit';
 * import type { PageContext } from '@microsoft/sp-page-context';
 * function inspectHost(pageContext: PageContext) {
 * const user = getSPFxUserInfo(pageContext);
 * const site = getSPFxSiteInfo(pageContext);
 * const list = getSPFxListInfo(pageContext);
 * const locale = getSPFxLocaleInfo(pageContext);
 * const environment = getSPFxEnvironmentInfo(pageContext);
 * const page = getSPFxPageTypeInfo(pageContext);
 * const correlation = getSPFxCorrelationInfo(pageContext);
 * console.log(user.displayName, site.siteUrl, list?.title, locale.uiLocale);
 * console.log(environment.type, page.pageType, correlation.correlationId);
 * }
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
 * import type { PageContext } from '@microsoft/sp-page-context';
 * import { getSPFxListInfo } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(pageContext: PageContext) {
 *   const info = getSPFxListInfo(pageContext);
 * }
 * ```
 * @useWhen You inspect optional list context and can handle pages that do not expose a list ID.
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
   * Always returned when list context exists despite the optional declaration: true only when baseTemplate equals 101; false for other or missing templates. No undefined flag is produced by this mapper.
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
 * Missing, noninteger and nonfinite IDs are unresolved, not UTC. Mapping uses
 * only the numeric ID, never culture, description or offset. Consumers need
 * suitable Intl zone data; pinned CLDR aliases are retained.
 *
 * @param pageContext - SPFx page context
 * @returns Locale, regional metadata and an optional representative IANA time zone
 * @example Basic usage
 * ```ts
 * import type { PageContext } from '@microsoft/sp-page-context';
 * import { getSPFxLocaleInfo } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(pageContext: PageContext) {
 *   const { uiLocale, ianaTimeZone } = getSPFxLocaleInfo(pageContext);
 *   if (ianaTimeZone !== undefined) {
 *     const formatter = new Intl.DateTimeFormat(uiLocale, { timeZone: ianaTimeZone });
 *     const siteDate = formatter.format(new Date());
 *   }
 * }
 * ```
 * @useWhen You format site-local dates using the supplied host culture and an audited regional-zone representative.
 *
 *
 * @example Choose an explicit missing site-zone policy
 * ```ts
 * import { getSPFxLocaleInfo } from '@apvee/spfx-react-toolkit';
 * import type { PageContext } from '@microsoft/sp-page-context';
 * function formatModified(pageContext: PageContext, modified: Date) {
 * const info = getSPFxLocaleInfo(pageContext);
 * // Explicit policy: no site-zone date when the standard mapping is unavailable.
 * const fullDate = info.ianaTimeZone
 *   ? new Intl.DateTimeFormat(info.uiLocale, { timeZone: info.ianaTimeZone, dateStyle: 'full' }).format(modified)
 *   : 'Site time zone unavailable';
 * return fullDate;
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
 * Maps supplied PageContext URL substrings, legacy isOnPremises and SDK-shaped fields to a best-effort host category. Precedence is Local, Teams, Outlook, Office, SharePointOnPrem, SharePoint; individual flags can overlap. Does not initialize SDK context or inspect a provider's native host.
 *
 * @param pageContext - SPFx page context
 * @returns Environment information
 *
 * @example
 * ```ts
 * import type { PageContext } from '@microsoft/sp-page-context';
 * import { getSPFxEnvironmentInfo } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(pageContext: PageContext) {
 *   const info = getSPFxEnvironmentInfo(pageContext);
 * }
 * ```
 * @useWhen You need a synchronous best-effort host category from supplied PageContext metadata.
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
   * True when the supplied PageContext sdks.microsoftTeams field is not undefined; no SDK shape validation or acquisition occurs.
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
 * Checks modern page.type for site/web-part pages, then legacy pageType for site/web-part/list/profile/search categories, then legacy listId. A non-null, non-undefined legacy formType, including zero, selects a list form. Unknown metadata returns unknown with false category flags.
 *
 * @param pageContext - SPFx page context
 * @returns Page type information
 *
 * @example
 * ```ts
 * import type { PageContext } from '@microsoft/sp-page-context';
 * import { getSPFxPageTypeInfo } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(pageContext: PageContext) {
 *   const info = getSPFxPageTypeInfo(pageContext);
 * }
 * ```
 * @useWhen You choose page-specific presentation from modern or legacy page metadata.
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
 * import type { PageContext } from '@microsoft/sp-page-context';
 * import { getSPFxCorrelationInfo } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(pageContext: PageContext) {
 *   const info = getSPFxCorrelationInfo(pageContext);
 * }
 * ```
 * @useWhen You attach host correlation and tenant identifiers to diagnostics when available.
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

import type { PageContext } from '@microsoft/sp-page-context';

type SPFxGroupInfoResult = {
  readonly id: string;
  readonly isPublic: boolean;
};

type SPFxTimeZoneResult = {
  readonly id: number;
  readonly offset: number;
  readonly description: string;
  readonly daylightOffset: number;
  readonly standardOffset: number;
};

type SPFxEnvironmentTypeResult =
  | 'Local'
  | 'SharePoint'
  | 'SharePointOnPrem'
  | 'Teams'
  | 'Office'
  | 'Outlook';

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
 */
export function getSPFxUserInfo(pageContext: PageContext): {
  readonly loginName: string;
  readonly displayName: string;
  readonly email?: string;
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
 */
export function getSPFxSiteInfo(pageContext: PageContext): {
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
  readonly siteGroup?: SPFxGroupInfoResult;
} {
  const siteObj = pageContext.site;
  const webObj = pageContext.web;
  const legacy = (pageContext as unknown as {
    legacyPageContext?: {
      siteClassification?: string;
    };
  }).legacyPageContext;

  return {
    webId: webObj.id.toString(),
    webUrl: webObj.absoluteUrl,
    webServerRelativeUrl: webObj.serverRelativeUrl,
    title: webObj.title,
    languageId: webObj.language ?? 1033,
    logoUrl: (webObj as unknown as { logoUrl?: string }).logoUrl,
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
 */
export function getSPFxListInfo(pageContext: PageContext): {
  readonly id: string;
  readonly title: string;
  readonly serverRelativeUrl: string;
  readonly baseTemplate?: number;
  readonly isDocumentLibrary?: boolean;
} | undefined {
  const list = (pageContext as unknown as {
    list?: {
      id?: { toString: () => string };
      title?: string;
      serverRelativeUrl?: string;
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
 *
 * @param pageContext - SPFx page context
 * @returns Locale and regional information
 */
export function getSPFxLocaleInfo(pageContext: PageContext): {
  readonly locale: string;
  readonly uiLocale: string;
  readonly timeZone: SPFxTimeZoneResult | undefined;
  readonly isRtl: boolean;
} {
  const cultureInfo = pageContext.cultureInfo;
  const timeZone = (pageContext.web as { timeZoneInfo?: SPFxTimeZoneResult }).timeZoneInfo;

  return {
    locale: cultureInfo.currentCultureName,
    uiLocale: cultureInfo.currentUICultureName,
    timeZone,
    isRtl: cultureInfo.isRightToLeft,
  };
}

/**
 * Maps SPFx PageContext host data to the toolkit environment info shape.
 *
 * @param pageContext - SPFx page context
 * @returns Environment information
 */
export function getSPFxEnvironmentInfo(pageContext: PageContext): {
  readonly type: SPFxEnvironmentTypeResult;
  readonly isLocal: boolean;
  readonly isWorkbench: boolean;
  readonly isSharePoint: boolean;
  readonly isSharePointOnPrem: boolean;
  readonly isTeams: boolean;
  readonly isOffice: boolean;
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
 */
export function getSPFxPageTypeInfo(pageContext: PageContext): {
  readonly pageType: SPFxPageTypeResult;
  readonly isModernPage: boolean;
  readonly isSitePage: boolean;
  readonly isListPage: boolean;
  readonly isListFormPage: boolean;
  readonly isWebPartPage: boolean;
} {
  const modernPage = (pageContext as unknown as {
    page?: {
      type?: string;
    };
  }).page;

  const legacy = (pageContext as unknown as {
    legacyPageContext?: {
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
 */
export function getSPFxCorrelationInfo(pageContext: PageContext): {
  readonly correlationId: string | undefined;
  readonly tenantId: string | undefined;
} {
  const correlationId = pageContext.site?.correlationId?.toString();
  const aadInfo = pageContext.aadInfo as unknown as
    { tenantId?: { toString(): string } } | undefined;
  const tenantId = aadInfo?.tenantId?.toString();

  return {
    correlationId,
    tenantId,
  };
}

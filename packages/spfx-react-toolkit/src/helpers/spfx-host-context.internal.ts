/** SDK methods consumed by the private SPFx runtime bridge. */
export interface TeamsSdkLike {
  readonly app?: {
    getContext?: () => Promise<unknown>;
    registerOnThemeChangeHandler?: (handler: (theme: string) => void) => void;
  };
  getContext?: (callback: (context: unknown) => void) => void;
  registerOnThemeChangeHandler?: (handler: (theme: string) => void) => void;
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? value as Record<string, unknown>
    : undefined;
}

function isTeamsSdk(value: unknown): value is TeamsSdkLike {
  const sdk = asRecord(value);
  const app = asRecord(sdk?.app);
  return sdk !== undefined && (
    typeof app?.getContext === 'function' || typeof sdk.getContext === 'function'
  );
}

/** Resolve SPFx's SDK wrapper or a direct legacy SDK after checking callable methods. */
export function resolveSPFxTeamsSdk(context: unknown): TeamsSdkLike | undefined {
  const teams = asRecord(asRecord(asRecord(context)?.sdks)?.microsoftTeams);
  if (isTeamsSdk(teams?.teamsJs)) {
    return teams.teamsJs;
  }
  return isTeamsSdk(teams) ? teams : undefined;
}

/** Resolve the application host from the acquired SDK context, without host guesses. */
export function resolveSPFxRuntimeHost(context: unknown): 'Teams' | 'Office' | 'Outlook' | 'unknown' {
  const name = asRecord(asRecord(asRecord(context)?.app)?.host)?.name;
  switch (name) {
    case 'Teams':
    case 'TeamsModern':
      return 'Teams';
    case 'Office':
      return 'Office';
    case 'Outlook':
    case 'OutlookWin32':
      return 'Outlook';
    default:
      return 'unknown';
  }
}

/** The legacy field is optional and not a stable PageContext contract. */
export function isSingleWebPartAppPage(pageContext: unknown): true | undefined {
  const layout = asRecord(asRecord(pageContext)?.legacyPageContext)?.pageLayoutType;
  return layout === 'SingleWebPartAppPage' ? true : undefined;
}

/** Resolve layout independently of viewport width and asynchronously acquired host names. */
export function resolveWebPartShellSurface(
  requested: 'auto' | 'section' | 'application',
  hasApplicationSdk: boolean,
  indicator: true | undefined
): 'section' | 'application' {
  return requested === 'auto'
    ? hasApplicationSdk || indicator === true ? 'application' : 'section'
    : requested;
}

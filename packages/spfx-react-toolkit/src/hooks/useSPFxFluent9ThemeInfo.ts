/**
 * Fluent 9 themes prioritize acquired application themes, then native SPFx colors.
 *
 * @packageDocumentation
 * @docPage hooks/theming
 */
// useSPFxFluent9ThemeInfo.ts
// Hook for accessing Fluent UI 9 theme based on SPFx context

import { useMemo } from 'react';
import type { Theme } from '@fluentui/react-theme';
import { createFluent9ThemeFromSPFxTheme, getTeamsFluentTheme } from '../helpers/spfx-theme.helpers';
import { useSPFxThemeInfo } from './useSPFxThemeInfo';
import { useSPFxTeams } from './useSPFxTeams';

/**
 * Fluent 9 tokens derived from the acquired SDK theme or SPFx theme, with webLightTheme fallback.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { FluentProvider } from '@fluentui/react-components';
 * import { SPFxFluent9ThemeInfo, useSPFxFluent9ThemeInfo } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const info: SPFxFluent9ThemeInfo = useSPFxFluent9ThemeInfo();
 *   return <FluentProvider theme={info.theme}>Content</FluentProvider>;
 * }
 * ```
 */
export interface SPFxFluent9ThemeInfo {
  /** Fluent UI 9 theme object ready to use with FluentProvider */
  readonly theme: Theme;
  
  /** Whether SPFx application SDK context acquisition succeeded, including Office/Outlook.
   * This legacy field does not identify the actual host or guarantee an SDK theme.
   */
  readonly isTeams: boolean;
  
  /** Normalized application theme name ('default', 'dark', 'highContrast'); undefined until an SDK theme is available. */
  readonly teamsTheme?: string;
}

/**
 * Returns native Fluent Teams tokens when SDK context is supported and a normalized theme is available; otherwise converts the SPFx theme, with webLightTheme as the undefined-theme fallback. Provider and SDK theme subscriptions update this hook. The SDK dispatcher is shared per identity within this bundle; another bundle can replace its single handler. Legacy isTeams means SDK context acquisition succeeded and can also describe Office/Outlook; use useSPFxEnvironmentInfo for the actual host. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Fluent theme plus SDK availability and optional normalized application-theme metadata.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example Read Fluent9ThemeInfo metadata
 * ```tsx
 * import * as React from 'react';
 * import { FluentProvider } from '@fluentui/react-components';
 * import { useSPFxFluent9ThemeInfo } from '@apvee/spfx-react-toolkit';
 *
 * function ThemedContent() {
 *   const { theme } = useSPFxFluent9ThemeInfo();
 *   return <FluentProvider theme={theme}><span>Themed content</span></FluentProvider>;
 * }
 * ```
 *
 * @useWhen custom Fluent 9 UI must use the theme derived from this provider host.
 *
 * @example Observe the projected Fluent theme tokens
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxFluent9ThemeInfo } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const info = useSPFxFluent9ThemeInfo();
 *   const regions: DemoRegions = {
 *     children: <dl aria-label="useSPFxFluent9ThemeInfo result">
 *         <dt>SDK theme</dt><dd>{info.teamsTheme ?? 'Unavailable'}</dd>
 *         <dt>Legacy isTeams</dt><dd>{String(info.isTeams)}</dd>
 *         <dt>Neutral background</dt><dd>{info.theme.colorNeutralBackground1}</dd>
 *         <dt>Neutral foreground</dt><dd>{info.theme.colorNeutralForeground1}</dd>
 *         <dt>Brand background</dt><dd>{info.theme.colorBrandBackground}</dd>
 *       </dl>,
 *     diagnostics: <p>Legacy isTeams means SDK acquisition succeeded and also covers Office/Outlook. These are selected projected tokens; actual host Fluent theme and direction remain inherited.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Inspect the actual Fluent9ThemeInfo values and explicit unavailable fields as the native host changes.' };
 * ```
 */
export function useSPFxFluent9ThemeInfo(): SPFxFluent9ThemeInfo {
  const teamsInfo = useSPFxTeams();
  const spfxTheme = useSPFxThemeInfo();
  
  const theme = useMemo(() => {
    // Priority 1: Teams theme (native Teams themes for better integration)
    if (teamsInfo.supported && teamsInfo.theme) {
      return getTeamsFluentTheme(teamsInfo.theme);
    }
    
    // Priority 2: Convert SPFx theme to Fluent UI 9
    // Priority 3: Fallback to default light theme
    return createFluent9ThemeFromSPFxTheme(spfxTheme);
  }, [teamsInfo.supported, teamsInfo.theme, spfxTheme]);
  
  return useMemo(() => ({
    theme,
    isTeams: teamsInfo.supported,
    teamsTheme: teamsInfo.theme
  }), [theme, teamsInfo.supported, teamsInfo.theme]);
}

// useSPFxFluent9ThemeInfo.ts
// Hook for accessing Fluent UI 9 theme based on SPFx context

import { useMemo } from 'react';
import type { Theme } from '@fluentui/react-theme';
import { createFluent9ThemeFromSPFxTheme, getTeamsFluentTheme } from '../helpers/spfx-theme.helpers';
import { useSPFxThemeInfo } from './useSPFxThemeInfo';
import { useSPFxTeams } from './useSPFxTeams';

/**
 * Return type for useSPFxFluent9ThemeInfo hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const info: SPFxFluent9ThemeInfo = useSPFxFluent9ThemeInfo();
 * return <FluentProvider theme={info.theme}>Content</FluentProvider>;
 * ```
 */
export interface SPFxFluent9ThemeInfo {
  /** Fluent UI 9 theme object ready to use with FluentProvider */
  readonly theme: Theme;
  
  /** Whether SPFx application SDK context acquisition succeeded, including Office/Outlook.
   * This legacy field does not identify the actual host or guarantee an SDK theme.
   */
  readonly isTeams: boolean;
  
  /** Normalized application theme name ('default', 'dark', 'highContrast'). */
  readonly teamsTheme?: string;
}

/**
 * Returns native Fluent Teams tokens when SDK context is supported and a normalized theme is available; otherwise converts the SPFx theme, with webLightTheme as the undefined-theme fallback. Provider and SDK theme subscriptions update this hook. The SDK dispatcher is shared per identity within this bundle; another bundle can replace its single handler. Legacy isTeams means SDK context acquisition succeeded and can also describe Office/Outlook; use useSPFxEnvironmentInfo for the actual host. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Fluent theme plus SDK availability and optional normalized application-theme metadata.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example
 * ```tsx
 * function ThemedContent() {
 *   const { theme } = useSPFxFluent9ThemeInfo();
 *   return <FluentProvider theme={theme}><span>Themed content</span></FluentProvider>;
 * }
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

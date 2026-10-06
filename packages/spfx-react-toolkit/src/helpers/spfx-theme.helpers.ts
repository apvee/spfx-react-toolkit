import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { Theme } from '@fluentui/react-theme';
import {
  teamsDarkTheme,
  teamsHighContrastTheme,
  teamsLightTheme,
  webLightTheme,
} from '@fluentui/react-theme';
import { createV9Theme } from '@fluentui/react-migration-v8-v9';

/**
 * Converts an SPFx v8 theme to a Fluent UI 9 theme.
 *
 * @param spfxTheme - SPFx theme, or undefined for the default web light theme
 * @returns Fluent UI 9 theme
 */
export function createFluent9ThemeFromSPFxTheme(
  spfxTheme: IReadonlyTheme | undefined
): Theme {
  if (!spfxTheme) {
    return webLightTheme;
  }

  return createV9Theme(spfxTheme as Parameters<typeof createV9Theme>[0]);
}

/**
 * Maps a Teams theme name to the matching Fluent UI 9 Teams theme.
 *
 * @param teamsThemeName - Teams theme name
 * @returns Fluent UI 9 Teams theme
 */
export function getTeamsFluentTheme(teamsThemeName: string | undefined): Theme {
  const normalizedThemeName = teamsThemeName?.toLowerCase();

  switch (normalizedThemeName) {
    case 'dark':
      return teamsDarkTheme;
    case 'contrast':
    case 'highcontrast':
      return teamsHighContrastTheme;
    case 'default':
    default:
      return teamsLightTheme;
  }
}

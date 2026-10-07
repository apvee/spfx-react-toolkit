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
 * Preserves the migration shim's token mappings except accessible neutral
 * stroke hover/pressed colors, which use the host palette's neutralPrimary
 * and neutralDark. Missing/empty state colors retain the shim's fallback.
 * Does not mutate the SPFx theme. Undefined retains the shared webLightTheme;
 * custom palettes remain responsible for usable color/contrast differences.
 *
 * @param spfxTheme - SPFx theme, or undefined for the default web light theme
 * @returns Fluent UI 9 theme
 * @example
 * const theme = createFluent9ThemeFromSPFxTheme(spfxTheme);
 * // Pass theme to FluentProvider; host scrollbar thumb hover/pressed feedback
 * // follows neutralPrimary/neutralDark without replacing other host tokens.
 */
export function createFluent9ThemeFromSPFxTheme(
  spfxTheme: IReadonlyTheme | undefined
): Theme {
  if (!spfxTheme) {
    return webLightTheme;
  }

  const theme = createV9Theme(spfxTheme as Parameters<typeof createV9Theme>[0]);
  const hover = spfxTheme.palette?.neutralPrimary;
  const pressed = spfxTheme.palette?.neutralDark;
  return {
    ...theme,
    colorNeutralStrokeAccessibleHover: typeof hover === 'string' && hover.trim()
      ? hover : theme.colorNeutralStrokeAccessibleHover,
    colorNeutralStrokeAccessiblePressed: typeof pressed === 'string' && pressed.trim()
      ? pressed : theme.colorNeutralStrokeAccessiblePressed
  };
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

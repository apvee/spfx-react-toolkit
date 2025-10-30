// useSPFxFluent9ThemeInfo.ts
// Hook for accessing Fluent UI 9 theme based on SPFx context

import { useMemo } from 'react';
import type { Theme } from '@fluentui/react-theme';
import { 
  webLightTheme,
  teamsLightTheme, 
  teamsDarkTheme, 
  teamsHighContrastTheme 
} from '@fluentui/react-theme';
import { createV9Theme } from '@fluentui/react-migration-v8-v9';
import { useSPFxThemeInfo } from './useSPFxThemeInfo';
import { useSPFxTeams } from './useSPFxTeams';

/**
 * Return type for useSPFxFluent9ThemeInfo hook
 */
export interface SPFxFluent9ThemeInfo {
  /** Fluent UI 9 theme object ready to use with FluentProvider */
  readonly theme: Theme;
  
  /** Whether the component is running in Microsoft Teams */
  readonly isTeams: boolean;
  
  /** Teams theme name if running in Teams ('default', 'dark', 'contrast') */
  readonly teamsTheme?: string;
}

/**
 * Hook for accessing Fluent UI 9 theme based on SPFx context
 * 
 * Automatically detects the execution context and provides the appropriate
 * Fluent UI 9 theme:
 * - **In Microsoft Teams**: Returns native Teams themes (light/dark/contrast)
 * - **In SharePoint**: Converts SPFx theme (v8) to Fluent UI 9 theme
 * - **Fallback**: Returns webLightTheme if no theme available
 * 
 * The hook uses memoization to avoid expensive theme conversions on every render.
 * Theme updates are automatically handled when the user switches themes in
 * SharePoint or Teams through the SPFxProvider's theme subscription mechanism.
 * 
 * Priority order:
 * 1. Teams theme (if running in Teams context)
 * 2. SPFx theme converted to Fluent UI 9
 * 3. Default webLightTheme
 * 
 * Useful for:
 * - Wrapping components with FluentProvider
 * - Ensuring consistent theming across SharePoint and Teams
 * - Automatic theme switching when user changes Teams/SharePoint theme
 * - Building Fluent UI 9 applications in SPFx
 * 
 * @returns Theme information including Fluent UI 9 theme and context details
 * 
 * @example Basic usage with FluentProvider
 * ```tsx
 * import { FluentProvider } from '@fluentui/react-components';
 * 
 * function MyWebPart() {
 *   const { theme } = useSPFxFluent9ThemeInfo();
 *   
 *   return (
 *     <FluentProvider theme={theme}>
 *       <MyApp />
 *     </FluentProvider>
 *   );
 * }
 * ```
 * 
 * @example Accessing context information
 * ```tsx
 * function MyComponent() {
 *   const { theme, isTeams, teamsTheme } = useSPFxFluent9ThemeInfo();
 *   
 *   return (
 *     <div>
 *       <p>Running in Teams: {isTeams ? 'Yes' : 'No'}</p>
 *       {isTeams && <p>Teams theme: {teamsTheme}</p>}
 *       <FluentProvider theme={theme}>
 *         <Button>Themed Button</Button>
 *       </FluentProvider>
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Conditional rendering based on theme
 * ```tsx
 * function ThemedComponent() {
 *   const { theme, teamsTheme } = useSPFxFluent9ThemeInfo();
 *   
 *   const isDarkTheme = teamsTheme === 'dark';
 *   
 *   return (
 *     <FluentProvider theme={theme}>
 *       {isDarkTheme ? <DarkModeIcon /> : <LightModeIcon />}
 *     </FluentProvider>
 *   );
 * }
 * ```
 * 
 * @example Using with Fluent UI 9 components
 * ```tsx
 * import { FluentProvider, Button, Card } from '@fluentui/react-components';
 * 
 * function MyWebPart() {
 *   const { theme, isTeams } = useSPFxFluent9ThemeInfo();
 *   
 *   return (
 *     <FluentProvider theme={theme}>
 *       <Card>
 *         <h3>Hello from {isTeams ? 'Teams' : 'SharePoint'}!</h3>
 *         <Button appearance="primary">Click me</Button>
 *       </Card>
 *     </FluentProvider>
 *   );
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
    if (spfxTheme) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return createV9Theme(spfxTheme as any);
    }
    
    // Priority 3: Fallback to default light theme
    return webLightTheme;
  }, [teamsInfo.supported, teamsInfo.theme, spfxTheme]);
  
  return {
    theme,
    isTeams: teamsInfo.supported,
    teamsTheme: teamsInfo.theme
  };
}

/**
 * Maps Teams theme name to corresponding Fluent UI 9 theme
 * @internal
 */
function getTeamsFluentTheme(teamsThemeName: string): Theme {
  switch (teamsThemeName) {
    case 'dark':
      return teamsDarkTheme;
    case 'highContrast':
      return teamsHighContrastTheme;
    case 'default':
    default:
      return teamsLightTheme;
  }
}

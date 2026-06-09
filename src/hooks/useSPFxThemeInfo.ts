// useSPFxThemeInfo.ts
// Hook to access current SPFx theme

import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { useSPFxRuntimeSelector } from '../core/state.internal';

/**
 * Hook to access the current SPFx theme
 * 
 * Theme subscription is managed automatically by SPFx providers.
 * Updates when user switches between light/dark theme or theme settings change.
 * 
 * @returns Current theme object or undefined if not yet loaded
 * 
 * @see {@link useSPFxFluent9ThemeInfo} for Fluent UI 9 theme conversion
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const theme = useSPFxThemeInfo();
 *   
 *   return (
 *     <div style={{ 
 *       backgroundColor: theme?.semanticColors.bodyBackground,
 *       color: theme?.semanticColors.bodyText 
 *     }}>
 *       Themed content
 *     </div>
 *   );
 * }
 * ```
 */
export function useSPFxThemeInfo(): IReadonlyTheme | undefined {
  return useSPFxRuntimeSelector(state => state.theme);
}

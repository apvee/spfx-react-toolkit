// useSPFxThemeInfo.ts
// Hook to access current SPFx theme

import { useAtomValue } from 'jotai';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { useSPFxContext } from './useSPFxContext';
import { spfxAtoms } from '../core/atoms';

/**
 * Hook to access the current SPFx theme
 * 
 * Theme subscription is managed automatically by SPFxProvider.
 * Updates when user switches between light/dark theme or theme settings change.
 * 
 * @returns Current theme object or undefined if not yet loaded
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
  const { instanceId } = useSPFxContext();
  
  // Get theme atom for this instance
  const themeAtom = spfxAtoms.theme(instanceId);
  
  // Read current theme value (subscription handled by Provider)
  return useAtomValue(themeAtom);
}

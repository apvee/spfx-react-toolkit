// useSPFxThemeInfo.ts
// Hook to access current SPFx theme

import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { useSPFxRuntimeSelector } from '../core/state.internal';

/**
 * Reads the nearest provider's host SPFx theme reference. Provider subscriptions publish changes; absence remains undefined, without a local theme fallback. Use useSPFxFluent9ThemeInfo to obtain Fluent UI 9 theme tokens. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Host readonly theme, or undefined when unavailable.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example
 * ```tsx
 * function ThemeLabel() {
 *   const theme = useSPFxThemeInfo();
 *   return <span>{theme ? 'Host theme available' : 'Host theme unavailable'}</span>;
 * }
 * ```
 */
export function useSPFxThemeInfo(): IReadonlyTheme | undefined {
  return useSPFxRuntimeSelector(state => state.theme);
}

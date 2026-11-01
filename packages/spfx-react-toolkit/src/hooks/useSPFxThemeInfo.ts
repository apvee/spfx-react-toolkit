/**
 * Native SPFx themes update through the provider theme subscription.
 *
 * @packageDocumentation
 * @docPage hooks/theming
 */
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
 * import * as React from 'react';
 * import { useSPFxThemeInfo } from '@apvee/spfx-react-toolkit';
 *
 * function ThemeLabel() {
 *   const theme = useSPFxThemeInfo();
 *   return <span>{theme ? 'Host theme available' : 'Host theme unavailable'}</span>;
 * }
 * ```
 *
 * @useWhen a component needs the current native SPFx theme and theme-change updates.
 */
export function useSPFxThemeInfo(): IReadonlyTheme | undefined {
  return useSPFxRuntimeSelector(state => state.theme);
}

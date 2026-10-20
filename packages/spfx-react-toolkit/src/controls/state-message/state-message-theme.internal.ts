import * as React from 'react';
import { ThemeContext_unstable } from '@fluentui/react-shared-contexts';
import { resolveThemeMode } from '../../helpers/fluent-theme-mode.internal';
import type { ThemeMode } from '../../helpers/fluent-theme-mode.internal';
export { resolveThemeMode } from '../../helpers/fluent-theme-mode.internal';
export type { ThemeMode } from '../../helpers/fluent-theme-mode.internal';

/** Read the closest FluentProvider, including nested theme updates, without requiring SPFx. */
export function useStateMessageThemeMode(): ThemeMode {
  return resolveThemeMode(React.useContext(ThemeContext_unstable));
}

import * as React from 'react';
import { ThemeContext_unstable } from '@fluentui/react-shared-contexts';
import { useSPFxFluent9ThemeInfo } from '../../hooks/useSPFxFluent9ThemeInfo';
import { resolveThemeMode } from '../../helpers/fluent-theme-mode.internal';
/** Native surface follows the nearest effective provider; SPFx/Teams supplies the fallback. */
export function useFileFolderPickerTheme(): 'default' | 'dark' {
  const nearest = React.useContext(ThemeContext_unstable);
  const host = useSPFxFluent9ThemeInfo();
  return resolveThemeMode(nearest ?? host.theme) === 'dark' ? 'dark' : 'default';
}

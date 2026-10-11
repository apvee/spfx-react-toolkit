import * as React from 'react';
import { Button } from '@fluentui/react-components';
import { teamsDarkTheme, teamsHighContrastTheme, webLightTheme } from '@fluentui/react-theme';
import { WebPartShell, useSPFxProperties, type IWebPartShellProperties, type WebPartShellProps } from '@apvee/spfx-react-toolkit';

export interface ShellDemoOptions {
  theme: 'host' | 'light' | 'dark' | 'highContrast';
  dir: 'ltr' | 'rtl';
  surface: 'auto' | 'section' | 'application';
  padding: 'auto' | 'none' | 'small' | 'medium' | 'large';
}
const defaults: ShellDemoOptions = { theme: 'host', dir: 'ltr', surface: 'auto', padding: 'auto' };
const ShellDemoContext = React.createContext({ options: defaults, update: (_patch: Partial<ShellDemoOptions>): void => {} });
export const useShellDemoOptions = (): React.ContextType<typeof ShellDemoContext> => React.useContext(ShellDemoContext);

function ShellDemoAction(): JSX.Element {
  const { properties, setProperties } = useSPFxProperties<IWebPartShellProperties>();
  return <Button size="small" data-shell-demo="action" onClick={() => setProperties({ showTitle: !(properties?.showTitle ?? true) })}>
    {properties?.showTitle === false ? 'Show title' : 'Hide title'}
  </Button>;
}

/** Private sample host: one Shell for the web part, with observable presentation overrides. */
export function WebPartShellDemoHost(props: WebPartShellProps<IWebPartShellProperties>): JSX.Element {
  const [options, setOptions] = React.useState(defaults);
  const update = React.useCallback((patch: Partial<ShellDemoOptions>): void => setOptions(previous => ({ ...previous, ...patch })), []);
  const theme = options.theme === 'host' ? undefined : options.theme === 'dark' ? teamsDarkTheme
    : options.theme === 'highContrast' ? teamsHighContrastTheme : webLightTheme;
  return <ShellDemoContext.Provider value={{ options, update }}>
    <WebPartShell {...props} surface={props.surface ?? options.surface} padding={props.padding ?? options.padding}
      theme={props.theme ?? theme} dir={props.dir ?? options.dir} action={<ShellDemoAction />} />
  </ShellDemoContext.Provider>;
}

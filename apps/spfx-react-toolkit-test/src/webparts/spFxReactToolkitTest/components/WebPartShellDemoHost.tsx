import * as React from 'react';
import { Button } from '@fluentui/react-components';
import { WebPartShell, useSPFxProperties, type IWebPartShellProperties, type WebPartShellProps } from '@apvee/spfx-react-toolkit';

export interface ShellDemoOptions {
  appearance: WebPartShellProps<IWebPartShellProperties>['appearance'];
  size: WebPartShellProps<IWebPartShellProperties>['size'];
  titlePosition: WebPartShellProps<IWebPartShellProperties>['titlePosition'];
  actionSlot: boolean;
  customNaming: boolean;
  customEditor: boolean;
  dir: 'host' | 'ltr' | 'rtl';
  surface: 'auto' | 'section' | 'application';
  padding: 'auto' | 'none' | 'small' | 'medium' | 'large';
}
const defaults: ShellDemoOptions = { appearance: 'subtle', size: 'medium', titlePosition: 'outside', actionSlot: false, customNaming: false, customEditor: false, dir: 'host', surface: 'auto', padding: 'auto' };
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
  return <ShellDemoContext.Provider value={{ options, update }}>
    <WebPartShell {...props} surface={props.surface ?? options.surface} padding={props.padding ?? options.padding}
      appearance={props.appearance ?? options.appearance} size={props.size ?? options.size} titlePosition={props.titlePosition ?? options.titlePosition}
      strings={props.strings ?? (options.customNaming ? { titleEditorLabel: 'Project section heading', titlePlaceholder: 'Add a project heading' } : undefined)}
      titleEditorStrings={props.titleEditorStrings ?? (options.customEditor ? { bold: 'Strong', formattingControls: 'Project heading controls', colorNames: { red: 'Alert red' } } : undefined)}
      dir={props.dir ?? (options.dir === 'host' ? undefined : options.dir)} action={options.actionSlot ? { 'data-shell-demo': 'action-slot',
        children: (Component, slotProps) => <Component {...slotProps}><ShellDemoAction /></Component> } : <ShellDemoAction />} />
  </ShellDemoContext.Provider>;
}

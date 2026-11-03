import * as React from 'react';
import { Button } from '@fluentui/react-components';
import { type ExecutableShellPresentation } from './generated/jsdoc-examples/catalogue';
import { webPartExecutableGroups } from './shared/webPartExecutableGroups';
import { ExecutableHostContext } from './shared/ExecutableHostContext';
import { ExampleExplorerStateProvider } from './shared/ExampleExplorerState';
import { WebPartShell, useSPFxProperties, type IWebPartShellProperties, type WebPartShellProps } from '@apvee/spfx-react-toolkit';

const defaults: ExecutableShellPresentation = { appearance: 'subtle', size: 'medium', titlePosition: 'outside', padding: 'auto' };
const ShellPresentationContext = React.createContext({ presentation: defaults, updatePresentation: (_patch: Partial<ExecutableShellPresentation>): void => {} });
export const useShellPresentation = (): React.ContextType<typeof ShellPresentationContext> => React.useContext(ShellPresentationContext);

/** Generic native host action; source scenarios may explicitly override this slot. */
function ShellDemoAction(): JSX.Element {
  const { setProperties } = useSPFxProperties<IWebPartShellProperties>();
  return <Button size="small" data-shell-demo="action" onClick={() => setProperties({ showTitle: false })}>
    Hide title
  </Button>;
}

/** Private sample host: one Shell; source examples own presentation recipes and dictionaries. */
export function WebPartShellDemoHost(props: WebPartShellProps<IWebPartShellProperties>): JSX.Element {
  return <InstanceShellHost key={props.instance.instanceId} {...props} />;
}

function InstanceShellHost(props: WebPartShellProps<IWebPartShellProperties>): JSX.Element {
  const [overrides, setOverrides] = React.useState<Partial<ExecutableShellPresentation>>({});
  const updatePresentation = React.useCallback((patch: Partial<ExecutableShellPresentation>): void => setOverrides(previous => {
    const next = { ...previous, ...patch };
    return (Object.keys(next) as (keyof ExecutableShellPresentation)[]).every(key => next[key] === previous[key]) ? previous : next;
  }), []);
  // Readback and native forwarding share the effective policy. Undefined releases
  // a source override; explicit consumer props retain precedence.
  const appearance = props.appearance ?? overrides.appearance ?? defaults.appearance;
  const size = props.size ?? overrides.size ?? defaults.size;
  const titlePosition = props.titlePosition ?? overrides.titlePosition ?? defaults.titlePosition;
  const padding = props.padding ?? overrides.padding ?? defaults.padding;
  const surface = props.surface ?? overrides.surface;
  const dir = props.dir ?? overrides.dir;
  const strings = props.strings ?? overrides.strings;
  const titleEditorStrings = props.titleEditorStrings ?? overrides.titleEditorStrings;
  const action = props.action !== undefined ? props.action : overrides.action;
  const presentation = React.useMemo(() => ({ appearance, size, titlePosition, padding, surface, dir, strings, titleEditorStrings, action }),
    [appearance, size, titlePosition, padding, surface, dir, strings, titleEditorStrings, action]);
  const value = React.useMemo(() => ({ presentation, updatePresentation }), [presentation, updatePresentation]);
  return <ExecutableHostContext.Provider value="webPart"><ExampleExplorerStateProvider key={props.instance.instanceId} initialGroups={webPartExecutableGroups}>
    <ShellPresentationContext.Provider value={value}>
      <WebPartShell {...props} {...presentation} action={action === undefined ? <ShellDemoAction /> : action} />
    </ShellPresentationContext.Provider>
  </ExampleExplorerStateProvider></ExecutableHostContext.Provider>;
}

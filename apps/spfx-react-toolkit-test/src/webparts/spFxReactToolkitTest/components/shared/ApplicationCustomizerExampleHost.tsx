import * as React from 'react';
import type { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { FluentProvider } from '@fluentui/react-components';
import { SPFxApplicationCustomizerProvider, useSPFxFluent9ThemeInfo, useSPFxPageContext } from '@apvee/spfx-react-toolkit';
import { executableGroups } from '../generated/jsdoc-examples/catalogue';
import { ExampleExplorer } from '../ExampleExplorer';
import { ExampleExplorerStateProvider } from './ExampleExplorerState';
import { ExecutableHostContext } from './ExecutableHostContext';
import { selectExecutableGroupsForHost } from './selectExecutableGroupsForHost';

const groups = selectExecutableGroupsForHost(executableGroups, 'applicationCustomizer');

/** The native class owns Top DOM and disposal; this adapter owns exactly one matching provider. */
export function ApplicationCustomizerExampleHost({ instance }: {
  readonly instance: BaseApplicationCustomizer<{ testMessage: string }>;
}): React.ReactElement {
  return <SPFxApplicationCustomizerProvider key={instance.instanceId} instance={instance}>
    <ExecutableHostContext.Provider value="applicationCustomizer">
      <TopPresentation />
    </ExecutableHostContext.Provider>
  </SPFxApplicationCustomizerProvider>;
}

/** Project the actual native theme/direction without creating another Shell or renderer. */
function TopPresentation(): React.ReactElement {
  const { theme } = useSPFxFluent9ThemeInfo();
  const page = useSPFxPageContext();
  return <FluentProvider theme={theme} dir={page.cultureInfo.isRightToLeft ? 'rtl' : 'ltr'}>
    <ExampleExplorerStateProvider initialGroups={groups}>
      <ExampleExplorer groups={groups} />
    </ExampleExplorerStateProvider>
  </FluentProvider>;
}

/** Regression snapshot captured before single-subject consolidation; not product UI. */
import { useDemoStyles } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/demoStyles';
import { MessageBar, MessageBarBody } from '@fluentui/react-components';
import * as React from 'react';
import { SPFxApplicationCustomizerProvider, SPFxFieldCustomizerProvider, SPFxListViewCommandSetProvider, SPFxWebPartProvider, } from '@apvee/spfx-react-toolkit';
import { DemoCard, InfoGrid, StatusBadge } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared';
const ProvidersPanel: React.FC = () => {
  const demoStyles = useDemoStyles();
  const providers = [
    { label: 'SPFxWebPartProvider', value: SPFxWebPartProvider },
    { label: 'SPFxApplicationCustomizerProvider', value: SPFxApplicationCustomizerProvider },
    { label: 'SPFxFieldCustomizerProvider', value: SPFxFieldCustomizerProvider },
    { label: 'SPFxListViewCommandSetProvider', value: SPFxListViewCommandSetProvider },
  ];
  return (<div className={demoStyles.root}>
    <MessageBar intent="info">
      <MessageBarBody>
        Providers are host integration controls. This webpart verifies all public provider exports; host-specific providers are mounted only in their valid SPFx hosts.
      </MessageBarBody>
    </MessageBar>

    <DemoCard contentPlacement="configuration" title="Provider Coverage" iconName="PlugConnected" diagnostics={<InfoGrid rows={[
              {
                label: 'SPFxWebPartProvider',
                value: 'Mounted by this sample webpart render method',
                icon: 'WebAppBuilderFragment',
              },
              {
                label: 'SPFxApplicationCustomizerProvider',
                value: 'Mounted by the sample application customizer',
                icon: 'AppIconDefault',
              },
              {
                label: 'SPFxFieldCustomizerProvider',
                value: 'Export verified here; mounting requires a Field Customizer host',
                icon: 'FieldChanged',
              },
              {
                label: 'SPFxListViewCommandSetProvider',
                value: 'Export verified here; mounting requires a ListView Command Set host',
                icon: 'CommandPrompt',
              },
            ]} />}>
      <div className={demoStyles.actions}>
        <StatusBadge label="SPFxWebPartProvider active" available={true} />
        <StatusBadge label="SPFxApplicationCustomizerProvider in extension sample" available={true} />
        {providers.map(provider => (<StatusBadge key={provider.label} label={`${provider.label} export`} available={typeof provider.value === 'function'} />))}
      </div>


    </DemoCard>
  </div>);
};
export default ProvidersPanel;

import * as React from 'react';
import { MessageBar, MessageBarType, Stack } from '@fluentui/react';
import {
  SPFxApplicationCustomizerProvider,
  SPFxFieldCustomizerProvider,
  SPFxListViewCommandSetProvider,
  SPFxWebPartProvider,
} from '../../../../core';
import { DemoCard, InfoGrid, StatusBadge } from '../shared';

const ProvidersPanel: React.FC = () => {
  const providers = [
    { label: 'SPFxWebPartProvider', value: SPFxWebPartProvider },
    { label: 'SPFxApplicationCustomizerProvider', value: SPFxApplicationCustomizerProvider },
    { label: 'SPFxFieldCustomizerProvider', value: SPFxFieldCustomizerProvider },
    { label: 'SPFxListViewCommandSetProvider', value: SPFxListViewCommandSetProvider },
  ];

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <MessageBar messageBarType={MessageBarType.info}>
        Providers are host integration controls. This webpart verifies all public provider exports; host-specific providers are mounted only in their valid SPFx hosts.
      </MessageBar>

      <DemoCard title="Provider Coverage" iconName="PlugConnected">
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <StatusBadge label="SPFxWebPartProvider active" available={true} />
          <StatusBadge label="SPFxApplicationCustomizerProvider in extension sample" available={true} />
          {providers.map(provider => (
            <StatusBadge
              key={provider.label}
              label={`${provider.label} export`}
              available={typeof provider.value === 'function'}
            />
          ))}
        </Stack>

        <InfoGrid
          rows={[
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
          ]}
        />
      </DemoCard>
    </Stack>
  );
};

export default ProvidersPanel;

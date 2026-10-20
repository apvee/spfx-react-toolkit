import { useDemoStyles } from '../shared/demoStyles';
import { Field, Select } from '@fluentui/react-components';
import * as React from 'react';
import { SPFxApplicationCustomizerProvider, SPFxFieldCustomizerProvider, SPFxListViewCommandSetProvider, SPFxWebPartProvider, useSPFxContext, } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';
import { InfoGrid } from '../shared/InfoGrid';
import { StatusBadge } from '../shared/StatusBadge';
const ProvidersPanel: React.FC = () => {
  const demoStyles = useDemoStyles();
  const providers = [
    { label: 'SPFxWebPartProvider', value: SPFxWebPartProvider },
    { label: 'SPFxApplicationCustomizerProvider', value: SPFxApplicationCustomizerProvider },
    { label: 'SPFxFieldCustomizerProvider', value: SPFxFieldCustomizerProvider },
    { label: 'SPFxListViewCommandSetProvider', value: SPFxListViewCommandSetProvider },
  ];
  const context = useSPFxContext();
  const [selection, setSelection] = React.useState('SPFxWebPartProvider');
  const selected = providers.find(provider => provider.label === selection) ?? providers[0];
  const hosts: Record<string, string> = { SPFxWebPartProvider: 'WebPart', SPFxApplicationCustomizerProvider: 'AppCustomizer', SPFxFieldCustomizerProvider: 'FieldCustomizer', SPFxListViewCommandSetProvider: 'CommandSet' };
  const requiredHost = hosts[selection];
  return <DemoCard contentPlacement="configuration" title="SPFx providers" iconName="PlugConnected" description="Inspect the exported provider and the actual host context."
    controls={<Field label="Provider"><Select select={{ as: 'select', ...{ 'data-demo-variant': 'providers' } }} value={selection} onChange={(_, data) => setSelection(data.value)}>{providers.map(provider => <option key={provider.label} value={provider.label}>{provider.label}</option>)}</Select></Field>}
    diagnostics={<p>Each provider requires its own SPFx host. Selection does not mount new instances.</p>}>
    <div className={demoStyles.panel} data-providers-demo="result">
      <StatusBadge label={`${selected.label} export`} available={typeof selected.value === 'function'} />
      <InfoGrid rows={[{ label: 'Current host', value: context.kind, icon: 'CubeShape' }, { label: 'Required host', value: requiredHost, icon: 'WebAppBuilderFragment' }, { label: 'Selected provider active', value: context.kind === requiredHost ? 'Yes' : 'No: requires another host', icon: 'PlugConnected' }]} />
    </div>
  </DemoCard>;
};
export default ProvidersPanel;

import { useDemoStyles } from '../shared/demoStyles';
import * as React from 'react';
import { Field, Select } from '@fluentui/react-components';
import { PnPContextDemo } from '../demos/PnPContextDemo';
import { PnPOperationsDemo } from '../demos/PnPOperationsDemo';
import { PnPListDemo } from '../demos/PnPListDemo';
import { PnPListAccessDemo } from '../demos/PnPListAccessDemo';
const PnPPanel: React.FC = () => {
    const demoStyles = useDemoStyles();
    const [scenario, setScenario] = React.useState('context');
    const selector = (busy = false): React.ReactNode => <Field className={demoStyles.controlRow} label="Operation"><Select disabled={busy} select={{ 'data-demo-variant': 'PnPPanel' } as React.SelectHTMLAttributes<HTMLSelectElement>} value={scenario} onChange={(_, data) => setScenario(data.value)}>
    <option value="context">PnP context</option><option value="operations">PnP requests</option><option value="list">List items</option><option value="listAccess">List access</option>
  </Select></Field>;
    return <>{scenario === 'context' ? <PnPContextDemo selector={selector}/> : scenario === 'operations' ? <PnPOperationsDemo selector={selector}/> : scenario === 'list' ? <PnPListDemo selector={selector}/> : <PnPListAccessDemo selector={selector}/>}</>;
};
export default PnPPanel;

import { useDemoStyles } from '../shared/demoStyles';
import * as React from 'react';
import { Field, Select } from '@fluentui/react-components';
import { PnPListDemo } from '../demos/PnPListDemo';
import { PnPListAccessDemo } from '../demos/PnPListAccessDemo';
const PnPPanel: React.FC = () => {
    const demoStyles = useDemoStyles();
    const [scenario, setScenario] = React.useState('listAccess');
    const selector = (busy = false): React.ReactNode => <Field className={demoStyles.controlRow} label="Operation"><Select disabled={busy} select={{ 'data-demo-variant': 'PnPPanel' } as React.SelectHTMLAttributes<HTMLSelectElement>} value={scenario} onChange={(_, data) => setScenario(data.value)}>
    <option value="list">List items</option><option value="listAccess">List access</option>
  </Select></Field>;
    return <>{scenario === 'list' ? <PnPListDemo selector={selector}/> : <PnPListAccessDemo selector={selector}/>}</>;
};
export default PnPPanel;

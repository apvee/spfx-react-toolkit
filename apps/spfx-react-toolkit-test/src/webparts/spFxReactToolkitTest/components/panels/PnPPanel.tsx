import { useDemoStyles } from '../shared/demoStyles';
import * as React from 'react';
import { Field, Select } from '@fluentui/react-components';
import { PnPListDemo } from '../demos/PnPListDemo';
const PnPPanel: React.FC = () => {
    const demoStyles = useDemoStyles();
    const selector = (busy = false): React.ReactNode => <Field className={demoStyles.controlRow} label="Operation"><Select disabled={busy} select={{ 'data-demo-variant': 'PnPPanel' } as React.SelectHTMLAttributes<HTMLSelectElement>} value="list">
    <option value="list">List items</option>
  </Select></Field>;
    return <PnPListDemo selector={selector}/>;
};
export default PnPPanel;

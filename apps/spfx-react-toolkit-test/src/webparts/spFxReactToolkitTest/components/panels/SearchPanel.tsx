import { useDemoStyles } from '../shared/demoStyles';
import * as React from 'react';
import { Field, Select } from '@fluentui/react-components';
import { PnPSearchRefinersDemo } from '../demos/PnPSearchRefinersDemo';
import { PnPSearchSuggestionsDemo } from '../demos/PnPSearchSuggestionsDemo';
const SearchPanel: React.FC = () => {
    const demoStyles = useDemoStyles();
    const [scenario, setScenario] = React.useState('refiners');
    const selector = (busy = false): React.ReactNode => <Field className={demoStyles.controlRow} label="Operation"><Select disabled={busy} select={{ 'data-demo-variant': 'SearchPanel' } as React.SelectHTMLAttributes<HTMLSelectElement>} value={scenario} onChange={(_, data) => setScenario(data.value)}>
    <option value="refiners">Filtered search</option><option value="suggestions">Suggestions</option>
  </Select></Field>;
    return <>{scenario === 'refiners' ? <PnPSearchRefinersDemo selector={selector}/> : <PnPSearchSuggestionsDemo selector={selector}/>}</>;
};
export default SearchPanel;

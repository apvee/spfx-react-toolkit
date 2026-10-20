import { useDemoStyles } from '../shared/demoStyles';
import * as React from 'react';
import { Field, Select } from '@fluentui/react-components';
import { PnPSearchBasicDemo } from '../demos/PnPSearchBasicDemo';
import { PnPSearchAdvancedDemo } from '../demos/PnPSearchAdvancedDemo';
import { PnPSearchRefinersDemo } from '../demos/PnPSearchRefinersDemo';
import { PnPSearchSuggestionsDemo } from '../demos/PnPSearchSuggestionsDemo';
const SearchPanel: React.FC = () => {
    const demoStyles = useDemoStyles();
    const [scenario, setScenario] = React.useState('basic');
    const selector = (busy = false): React.ReactNode => <Field className={demoStyles.controlRow} label="Operation"><Select disabled={busy} select={{ 'data-demo-variant': 'SearchPanel' } as React.SelectHTMLAttributes<HTMLSelectElement>} value={scenario} onChange={(_, data) => setScenario(data.value)}>
    <option value="basic">Text search</option><option value="advanced">Advanced search</option><option value="refiners">Filtered search</option><option value="suggestions">Suggestions</option>
  </Select></Field>;
    return <>{scenario === 'basic' ? <PnPSearchBasicDemo selector={selector}/> : scenario === 'advanced' ? <PnPSearchAdvancedDemo selector={selector}/> : scenario === 'refiners' ? <PnPSearchRefinersDemo selector={selector}/> : <PnPSearchSuggestionsDemo selector={selector}/>}</>;
};
export default SearchPanel;

import * as React from 'react';
import { Button, Field, Input, Select } from '@fluentui/react-components';
import { SharePointGroupTrim, type SharePointGroupSelector } from '@apvee/spfx-react-toolkit';
import { useDemoStyles } from '../shared/demoStyles';
import { DemoCard } from '../shared/DemoCard';
import { TrimProtectedContent, TrimBranchStatus, useTrimObservation } from '../shared/TrimDemo';
const SharePointGroupTrimPanel: React.FC = () => {
  const demoStyles = useDemoStyles();
  const [selector, setSelector] = React.useState('name');
  const [value, setValue] = React.useState('Project Readers');
  const [match, setMatch] = React.useState<'all' | 'any'>('all');
  const [webUrl, setWebUrl] = React.useState('');
  const [applied, setApplied] = React.useState<{ groups: readonly SharePointGroupSelector[]; match: 'all' | 'any'; webUrl?: string; run: number }>();
  const observation = useTrimObservation();
  return <DemoCard title="SharePoint group trim" iconName="Group" description="Read current-user group access before mounting caller content." contentPlacement="configuration"
    controls={<>
      <Field label="Group selector"><Select select={{ as: 'select', ...{ 'data-demo-variant': 'sharepoint-group-trim', 'data-trim-demo': 'selector' } }} value={selector} onChange={(_, data) => setSelector(data.value)}><option value="name">Exact names</option><option value="id">Numeric IDs</option></Select></Field>
      <Field label="Groups (separate with semicolons)"><Input input={{ as: 'input', ...{ 'data-trim-demo': 'groups' } }} value={value} onChange={(_, data) => setValue(data.value)} placeholder="Separate multiple selectors with semicolons" /></Field>
      <Field label="Web URL (empty uses current web)"><Input input={{ as: 'input', ...{ 'data-trim-demo': 'web-url' } }} value={webUrl} onChange={(_, data) => setWebUrl(data.value)} /></Field>
      <Field label="Match"><Select select={{ as: 'select', ...{ 'data-trim-demo': 'match' } }} value={match} onChange={(_, data) => setMatch(data.value === 'any' ? 'any' : 'all')}><option>all</option><option>any</option></Select></Field>
      <Button data-trim-demo="run" onClick={() => setApplied({ groups: value.split(';').map(entry => selector === 'id' ? { id: Number(entry.trim()) } : { name: entry.trim() }), match, webUrl: webUrl.trim() || undefined, run: (applied?.run ?? 0) + 1 })}>Run / repeat check</Button>
    </>}>
    <p>Requires the matching SPFx provider and native SPHttpClient. Membership is direct in the selected web; Entra principals are not expanded. Draft edits apply only on Run; repeat remounts the check without a settled cache.</p>
    <p className={demoStyles.observations} data-trim-demo="applied-target">Applied target: {applied ? JSON.stringify(applied) : 'not run'}</p>{observation.gauge}
    <div data-trim-demo="subject">{applied ? <SharePointGroupTrim key={applied.run} groups={applied.groups} match={applied.match} webUrl={applied.webUrl}>
      <SharePointGroupTrim.Loading><TrimBranchStatus state="loading" /></SharePointGroupTrim.Loading>
      <SharePointGroupTrim.Fallback><TrimBranchStatus state="denied" /></SharePointGroupTrim.Fallback>
      <SharePointGroupTrim.ErrorFallback><TrimBranchStatus state="error" /></SharePointGroupTrim.ErrorFallback>
      <TrimProtectedContent onMount={observation.onMount} onUnmount={observation.onUnmount} />
    </SharePointGroupTrim> : <p data-trim-demo="idle">Choose groups, then Run.</p>}</div>
  </DemoCard>;
};
export default SharePointGroupTrimPanel;

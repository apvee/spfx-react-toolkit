import * as React from 'react';
import { Button, Field, Input, Select } from '@fluentui/react-components';
import { EntraGroupTrim, type EntraGroupSelector, type EntraGroupType, type EntraGroupRelationship } from '@apvee/spfx-react-toolkit';
import { useDemoStyles } from '../shared/demoStyles';
import { DemoCard } from '../shared/DemoCard';
import { TrimProtectedContent, TrimBranchStatus, useTrimObservation } from '../shared/TrimDemo';
const EntraGroupTrimPanel: React.FC = () => {
  const demoStyles = useDemoStyles();
  const [selector, setSelector] = React.useState('name');
  const [value, setValue] = React.useState('Project Readers');
  const [match, setMatch] = React.useState<'all' | 'any'>('all');
  const [groupType, setGroupType] = React.useState<EntraGroupType>('All');
  const [relationship, setRelationship] = React.useState<EntraGroupRelationship>('member');
  const [applied, setApplied] = React.useState<{ groups: readonly EntraGroupSelector[]; match: 'all' | 'any'; groupType: EntraGroupType; relationship: EntraGroupRelationship; run: number }>();
  const observation = useTrimObservation();
  return <DemoCard title="Entra group trim" iconName="Group" description="Read current-user group access before mounting caller content." contentPlacement="configuration"
    controls={<>
      <Field label="Group selector"><Select select={{ as: 'select', ...{ 'data-demo-variant': 'entra-group-trim', 'data-trim-demo': 'selector' } }} value={selector} onChange={(_, data) => setSelector(data.value)}><option value="name">Exact names</option><option value="id">GUIDs</option></Select></Field>
      <Field label="Groups (separate with semicolons)"><Input input={{ as: 'input', ...{ 'data-trim-demo': 'groups' } }} value={value} onChange={(_, data) => setValue(data.value)} placeholder="Separate multiple selectors with semicolons" /></Field>
      <Field label="Group type"><Select select={{ as: 'select', ...{ 'data-trim-demo': 'group-type' } }} value={groupType} onChange={(_, data) => setGroupType(data.value as EntraGroupType)}><option>All</option><option>M365</option><option>Security</option></Select></Field>
      <Field label="Relationship"><Select select={{ as: 'select', ...{ 'data-trim-demo': 'relationship' } }} value={relationship} onChange={(_, data) => setRelationship(data.value as EntraGroupRelationship)}><option>member</option><option>owner</option><option>memberOrOwner</option></Select></Field>
      <Field label="Match"><Select select={{ as: 'select', ...{ 'data-trim-demo': 'match' } }} value={match} onChange={(_, data) => setMatch(data.value === 'any' ? 'any' : 'all')}><option>all</option><option>any</option></Select></Field>
      <Button data-trim-demo="run" onClick={() => setApplied({ groups: value.split(';').map(entry => selector === 'id' ? { id: entry.trim() } : { name: entry.trim() }), match, groupType, relationship, run: (applied?.run ?? 0) + 1 })}>Run / repeat check</Button>
    </>}>
    <p>Requires the matching SPFx provider and native Graph client. Sample configuration requests User.Read and GroupMember.Read.All; tenant approval is separate. Hidden membership omissions do not prove absence. Unsupported ownership, missing/ambiguous names, failed metadata resolution and type mismatches select Error. For All + ID + member, an unreturned GUID selects Denied and cannot distinguish a missing group from nonmembership. Draft edits apply only on Run; repeat remounts the check without a settled cache.</p>
    <p className={demoStyles.observations} data-trim-demo="applied-target">Applied target: {applied ? JSON.stringify(applied) : 'not run'}</p>{observation.gauge}
    <div data-trim-demo="subject">{applied ? <EntraGroupTrim key={applied.run} groups={applied.groups} match={applied.match} groupType={applied.groupType} relationship={applied.relationship}>
      <EntraGroupTrim.Loading><TrimBranchStatus state="loading" /></EntraGroupTrim.Loading>
      <EntraGroupTrim.Fallback><TrimBranchStatus state="denied" /></EntraGroupTrim.Fallback>
      <EntraGroupTrim.ErrorFallback><TrimBranchStatus state="error" /></EntraGroupTrim.ErrorFallback>
      <TrimProtectedContent onMount={observation.onMount} onUnmount={observation.onUnmount} />
    </EntraGroupTrim> : <p data-trim-demo="idle">Choose groups, then Run.</p>}</div>
  </DemoCard>;
};
export default EntraGroupTrimPanel;

import * as React from 'react';
import { Button, Field, Input, Select } from '@fluentui/react-components';
import { SPPermission } from '@microsoft/sp-page-context';
import { SharePointPermissionTrim, EntraGroupTrim } from '@apvee/spfx-react-toolkit';
import { useDemoStyles } from '../shared/demoStyles';
import { DemoCard } from '../shared/DemoCard';
import { initialPermissionDraft, permissionTarget, PermissionTargetFields, TrimProtectedContent, TrimBranchStatus, useTrimObservation } from '../shared/TrimDemo';
const SharePointPermissionTrimPanel: React.FC = () => {
  const demoStyles = useDemoStyles();
  const [draft, setDraft] = React.useState(initialPermissionDraft);
  const [permission, setPermission] = React.useState('viewListItems');
  const [match, setMatch] = React.useState<'all' | 'any'>('all');
  const [variant, setVariant] = React.useState('single');
  const [group, setGroup] = React.useState('Project Readers');
  const [applied, setApplied] = React.useState<{ target: ReturnType<typeof permissionTarget>; permission: string; match: 'all' | 'any'; variant: string; group: string; run: number }>();
  const observation = useTrimObservation();
  const required = applied?.permission === 'editListItems' ? SPPermission.editListItems : applied?.permission === 'both' ? undefined : SPPermission.viewListItems;
  const protectedContent = <TrimProtectedContent onMount={observation.onMount} onUnmount={observation.onUnmount} />;
  return <DemoCard title="SharePoint permission trim" iconName="Permissions" description="Current-user effective permissions on one resource. SharePoint authorizes every operation." contentPlacement="configuration"
    controls={<>
      <Field label="Composition"><Select select={{ as: 'select', ...{ 'data-demo-variant': 'sharepoint-permission-trim' } }} value={variant} onChange={(_, data) => setVariant(data.value)}><option value="single">Permission</option><option value="and">Permission AND Entra membership</option></Select></Field>
      <PermissionTargetFields draft={draft} onChange={setDraft} />
      <Field label="Required permissions"><Select select={{ as: 'select', ...{ 'data-trim-demo': 'permissions' } }} value={permission} onChange={(_, data) => setPermission(data.value)}><option value="viewListItems">View list items</option><option value="editListItems">Edit list items</option><option value="both">View and edit list items</option></Select></Field>
      <Field label="Match"><Select select={{ as: 'select', ...{ 'data-trim-demo': 'match' } }} value={match} onChange={(_, data) => setMatch(data.value === 'any' ? 'any' : 'all')}><option>all</option><option>any</option></Select></Field>
      {variant === 'and' && <Field label="Nested exact group name"><Input input={{ as: 'input', ...{ 'data-trim-demo': 'and-group' } }} value={group} onChange={(_, data) => setGroup(data.value)} /></Field>}
      <Button data-trim-demo="run" onClick={() => setApplied({ target: permissionTarget(draft), permission, match, variant, group, run: (applied?.run ?? 0) + 1 })}>Run / repeat check</Button>
    </>}>
    <p>Requires a matching SPFx provider and native SPHttpClient; nested Entra membership also requires the native Graph client with User.Read and GroupMember.Read.All consent. Current-web masks inherit PageContext freshness. Draft edits apply only on Run.</p>
    <p className={demoStyles.observations} data-trim-demo="applied-target">Applied target: {applied ? JSON.stringify(applied) : 'not run'}</p>{observation.gauge}
    <div data-trim-demo="subject">{applied ? <SharePointPermissionTrim key={applied.run} {...applied.target} permissions={required ? [required] : [SPPermission.viewListItems, SPPermission.editListItems]} match={applied.match}>
      <SharePointPermissionTrim.Loading><TrimBranchStatus state="loading" /></SharePointPermissionTrim.Loading>
      <SharePointPermissionTrim.Fallback><TrimBranchStatus state="denied" /></SharePointPermissionTrim.Fallback>
      <SharePointPermissionTrim.ErrorFallback><TrimBranchStatus state="error" /></SharePointPermissionTrim.ErrorFallback>
      {applied.variant === 'and' ? <EntraGroupTrim groups={[{ name: applied.group }]}>
        <EntraGroupTrim.Loading><TrimBranchStatus state="loading" /></EntraGroupTrim.Loading><EntraGroupTrim.Fallback><TrimBranchStatus state="denied" /></EntraGroupTrim.Fallback><EntraGroupTrim.ErrorFallback><TrimBranchStatus state="error" /></EntraGroupTrim.ErrorFallback>{protectedContent}
      </EntraGroupTrim> : protectedContent}
    </SharePointPermissionTrim> : <p data-trim-demo="idle">Choose a target, then Run.</p>}</div>
  </DemoCard>;
};
export default SharePointPermissionTrimPanel;

import * as React from 'react';
import { Button, Field, Input, Select } from '@fluentui/react-components';
import type { SPFxScopedPermissionsOptions } from '@apvee/spfx-react-toolkit';

export interface PermissionDraft {
  scope: 'web' | 'list' | 'item' | 'folder';
  webUrl: string;
  selector: 'title' | 'id' | 'path';
  list: string;
  itemId: string;
  folderPath: string;
}
export const initialPermissionDraft: PermissionDraft = { scope: 'web', webUrl: '', selector: 'title', list: 'Tasks', itemId: '1', folderPath: '/sites/team/Shared Documents/Reports' };
export function permissionTarget(draft: PermissionDraft): SPFxScopedPermissionsOptions {
  const webUrl = draft.webUrl.trim() || undefined;
  if (draft.scope === 'web') return { scope: 'web', webUrl };
  const selector = draft.selector === 'id' ? { listId: draft.list } : draft.selector === 'path' ? { listPath: draft.list } : { listTitle: draft.list };
  if (draft.scope === 'item') return { ...selector, scope: 'item', webUrl, itemId: Number(draft.itemId) };
  if (draft.scope === 'folder') return { ...selector, scope: 'folder', webUrl, folderPath: draft.folderPath };
  return { ...selector, scope: 'list', webUrl };
}
export function PermissionTargetFields({ draft, onChange }: { draft: PermissionDraft; onChange: (draft: PermissionDraft) => void }): React.ReactElement {
  return <>
    <Field label="Resource scope"><Select select={{ as: 'select', ...{ 'data-trim-demo': 'scope' } }} value={draft.scope} onChange={(_, data) => onChange({ ...draft, scope: data.value as PermissionDraft['scope'] })}>{['web', 'list', 'item', 'folder'].map(scope => <option key={scope}>{scope}</option>)}</Select></Field>
    <Field label="Web URL (empty uses current web)"><Input input={{ as: 'input', ...{ 'data-trim-demo': 'web-url' } }} value={draft.webUrl} onChange={(_, data) => onChange({ ...draft, webUrl: data.value })} /></Field>
    {draft.scope !== 'web' && <>
      <Field label="List selector"><Select select={{ as: 'select', ...{ 'data-trim-demo': 'list-selector' } }} value={draft.selector} onChange={(_, data) => onChange({ ...draft, selector: data.value as PermissionDraft['selector'] })}><option value="title">Exact title</option><option value="id">GUID</option><option value="path">Decoded web-relative path</option></Select></Field>
      <Field label="List value"><Input input={{ as: 'input', ...{ 'data-trim-demo': 'list-value' } }} value={draft.list} onChange={(_, data) => onChange({ ...draft, list: data.value })} /></Field>
    </>}
    {draft.scope === 'item' && <Field label="Item ID"><Input input={{ as: 'input', ...{ 'data-trim-demo': 'item-id' } }} value={draft.itemId} onChange={(_, data) => onChange({ ...draft, itemId: data.value })} /></Field>}
    {draft.scope === 'folder' && <Field label="Decoded server-relative folder path"><Input input={{ as: 'input', ...{ 'data-trim-demo': 'folder-path' } }} value={draft.folderPath} onChange={(_, data) => onChange({ ...draft, folderPath: data.value })} /></Field>}
  </>;
}
export function TrimProtectedContent({ onMount, onUnmount }: { onMount: () => void; onUnmount: () => void }): React.ReactElement {
  const [count, setCount] = React.useState(0);
  React.useEffect(() => { onMount(); return onUnmount; }, [onMount, onUnmount]);
  return <div data-trim-demo="protected"><p>Allowed: protected content is mounted.</p><Button data-trim-demo="increment" onClick={() => setCount(value => value + 1)}>Protected counter: {count}</Button></div>;
}
export function useTrimObservation(): { onMount: () => void; onUnmount: () => void; gauge: React.ReactElement } {
  const [mounts, setMounts] = React.useState(0);
  const [active, setActive] = React.useState(false);
  const onMount = React.useCallback(() => { setMounts(value => value + 1); setActive(true); }, []);
  const onUnmount = React.useCallback(() => setActive(false), []);
  return { onMount, onUnmount, gauge: <p data-trim-demo="mount-gauge">Protected mounts: {mounts}; active: {String(active)}</p> };
}
export function TrimBranchStatus({ state }: { state: 'loading' | 'denied' | 'error' }): React.ReactElement {
  return <p data-trim-demo="state" data-trim-state={state}>{state === 'loading' ? 'Loading: checking current access.' : state === 'denied' ? 'Denied: the requested access was not observed.' : 'Error: access could not be checked. Verify selectors, consent and transport.'}</p>;
}

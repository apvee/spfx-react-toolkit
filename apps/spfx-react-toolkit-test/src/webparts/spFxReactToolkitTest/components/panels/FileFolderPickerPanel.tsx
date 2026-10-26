import * as React from 'react';
import { Accordion, AccordionHeader, AccordionItem, Button, Field, Input, Select, Switch } from '@fluentui/react-components';
import { FileFolderPicker, useSPFxPageContext } from '@apvee/spfx-react-toolkit';
import type { FileFolderPickerScope, FileFolderPickerValue, FileFolderPickerOption, FileFolderPickerInitialLocation, FileFolderPickerPreset, FileFolderPickerProps } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';
import { DemoAccordionPanel } from '../shared/DemoAccordionPanel';

/** One real hosted picker; configuration changes do not create content or request consent. */
const FileFolderPickerPanel: React.FC = () => {
  const page = useSPFxPageContext();
  const [kind, setKind] = React.useState<FileFolderPickerScope['kind']>('siteCollection');
  const [web, setWeb] = React.useState(''), [list, setList] = React.useState(''), [folder, setFolder] = React.useState('');
  const [startWeb, setStartWeb] = React.useState(''), [startList, setStartList] = React.useState(''), [startFolder, setStartFolder] = React.useState('');
  const [initialLocation, setInitialLocation] = React.useState<FileFolderPickerInitialLocation>();
  const [scope, setScope] = React.useState<FileFolderPickerScope>({ kind: 'siteCollection' });
  const [itemTypes, setItemTypes] = React.useState<'files' | 'folders' | 'all'>('files');
  const [fileCategory, setFileCategory] = React.useState<FileFolderPickerProps['fileCategory']>();
  const [multiple, setMultiple] = React.useState(false), [searchable, setSearchable] = React.useState(true);
  const [colorfulTagAvatars, setColorfulTagAvatars] = React.useState(false);
  const [size, setSize] = React.useState<'medium' | 'large'>('medium'), [suffixes, setSuffixes] = React.useState('');
  const [limit, setLimit] = React.useState(''), [error, setError] = React.useState('');
  const [values, setValues] = React.useState<readonly FileFolderPickerValue[]>([]), [models, setModels] = React.useState<readonly FileFolderPickerOption[]>([]);
  const [reason, setReason] = React.useState('No action');
  const [preset, setPreset] = React.useState<FileFolderPickerPreset>('focused');
  const [uploadEnabled, setUploadEnabled] = React.useState(false), [createFolderEnabled, setCreateFolderEnabled] = React.useState(false);
  const [hint, setHint] = React.useState('');
  const [ready, setReady] = React.useState(0);

  const apply = (): void => {
    const targetWeb = web.trim() || page.web.absoluteUrl;
    let next: FileFolderPickerScope;
    if (kind === 'siteCollection') next = { kind, siteUrl: web.trim() || undefined };
    else if (kind === 'library') next = { kind, webUrl: targetWeb, list: list.trim() };
    else if (kind === 'folder') next = { kind, webUrl: targetWeb, list: list.trim(), folder };
    else next = { kind, folder: folder || undefined };
    if ((kind === 'library' || kind === 'folder') && !list.trim()) { setError('Enter an existing document library.'); return; }
    if (kind === 'folder' && !folder.trim()) { setError('Enter an existing folder.'); return; }
    setInitialLocation(startWeb || startList || startFolder ? { webUrl: startWeb.trim() || undefined, list: startList.trim() || undefined, folder: startFolder || undefined } : undefined);
    setScope(next); setValues([]); setModels([]); setError(''); setReason('Source changed explicitly');
  };
  return <DemoCard title="File and folder picker" iconName="Folder" description="Microsoft hosted File Picker v8; select existing content with verified source boundaries." controls={<>
    <Field label="Source"><Select value={kind} onChange={(_, data) => { const candidate = (['siteCollection', 'library', 'folder', 'oneDrive'] as const).find(value => value === data.value); if (candidate && candidate !== kind) {
        setKind(candidate); setStartWeb(''); setStartList(''); setStartFolder('');
      } }}>
      <option value="siteCollection">Site collection</option><option value="library">Document library</option><option value="folder">Folder</option><option value="oneDrive">OneDrive Business</option>
    </Select></Field>
    {kind !== 'oneDrive' ? <Field label="Web / site collection URL (empty: current)"><Input value={web} onChange={(_, data) => setWeb(data.value)} autoComplete="off" /></Field> : <p>Uses the current authenticated user’s OneDrive Business.</p>}
    {kind === 'library' || kind === 'folder' ? <Field label="Existing document library title"><Input value={list} onChange={(_, data) => setList(data.value)} autoComplete="off" /></Field> : undefined}
    {kind === 'folder' || kind === 'oneDrive' ? <Field label="Existing folder path"><Input value={folder} onChange={(_, data) => setFolder(data.value)} autoComplete="off" /></Field> : undefined}
    <Field label="Preset"><Select value={preset} onChange={(_, data) => setPreset(data.value === 'extended' ? 'extended' : 'focused')}><option value="focused">Focused</option><option value="extended">Extended</option></Select></Field>
    <Field label="Selectable items"><Select value={itemTypes} onChange={(_, data) => { setItemTypes(data.value === 'folders' ? 'folders' : data.value === 'all' ? 'all' : 'files'); if (data.value === 'folders') setFileCategory(undefined); }}><option value="files">Files</option><option value="folders">Folders</option><option value="all">Files and folders</option></Select></Field>
    <Switch label="Multiple selection" checked={multiple} onChange={(_, data) => setMultiple(data.checked)} />
    <Accordion collapsible><AccordionItem value="options"><AccordionHeader>Options</AccordionHeader><DemoAccordionPanel>
    {kind !== 'oneDrive' ? <>
      <Field label="Initial web URL (navigation only)"><Input value={startWeb} onChange={(_, data) => setStartWeb(data.value)} /></Field>
      <Field label="Initial library title (navigation only)"><Input value={startList} onChange={(_, data) => setStartList(data.value)} /></Field>
    </> : undefined}
    <Field label="Initial folder (navigation only)"><Input value={startFolder} onChange={(_, data) => setStartFolder(data.value)} /></Field>
    <p>Apply source also applies the optional starting location. Changing source clears the starting-location fields. A SharePoint starting folder requires a starting library; OneDrive accepts a folder only.</p>
    {itemTypes !== 'folders' ? <>
      <Field label="File category"><Select value={fileCategory ?? ''} onChange={(_, data) => {
        const category = (['images', 'video', 'audio', 'media', 'documents'] as const).find(value => value === data.value);
        setFileCategory(category); if (category) setSuffixes('');
      }}><option value="">Unrestricted / custom extensions</option><option value="images">Images</option><option value="video">Video</option><option value="audio">Audio</option><option value="media">Media</option><option value="documents">Documents</option></Select></Field>
      <Field label="File extensions (comma separated)"><Input disabled={fileCategory !== undefined} value={suffixes} onChange={(_, data) => setSuffixes(data.value)} /></Field>
      <p>A category replaces custom extensions and checks file suffixes. In mixed mode folders remain selectable. Changing this policy preserves selected references.</p>
    </> : undefined}
    {multiple ? <Field label="Selection limit (empty: default)"><Input value={limit} onChange={(_, data) => setLimit(data.value)} /></Field> : undefined}
    <Field label="Tag size"><Select value={size} onChange={(_, data) => setSize(data.value === 'large' ? 'large' : 'medium')}><option>medium</option><option>large</option></Select></Field>
    <Switch label="Colorful tag avatars" checked={colorfulTagAvatars} onChange={(_, data) => setColorfulTagAvatars(data.checked)} />
    <Switch label="Hosted search" checked={searchable} onChange={(_, data) => setSearchable(data.checked)} />
    <Field label="Field description (hint)"><Input value={hint} onChange={(_, data) => setHint(data.value)} /></Field>
    <p>Optional actions write immediately, before confirmation. Cancel cannot undo those writes.</p>
    <Switch label="Enable uploads" checked={uploadEnabled} onChange={(_, data) => setUploadEnabled(data.checked)} />
    <Switch label="Enable folder creation" checked={createFolderEnabled} onChange={(_, data) => setCreateFolderEnabled(data.checked)} />
  </DemoAccordionPanel></AccordionItem></Accordion>
  </>} commands={<Button onClick={apply}>Apply source and clear demo selection</Button>} status={error || undefined} diagnostics={<Accordion collapsible><AccordionItem value="diagnostics"><AccordionHeader>Diagnostics and selection JSON</AccordionHeader><DemoAccordionPanel>
    <p>Applied source: {JSON.stringify(scope)}. Preset: {preset}. Last action: {reason}. Ready notifications: {ready}.</p>
    <pre>{JSON.stringify(values, undefined, 2)}</pre>
  </DemoAccordionPanel></AccordionItem></Accordion>}>
    <FileFolderPicker preset={preset} uploadEnabled={uploadEnabled} createFolderEnabled={createFolderEnabled} hint={hint || undefined} onReady={() => setReady(count => count + 1)} label="FileFolderPicker" scope={scope} initialLocation={initialLocation} value={values} initialOptions={models} itemTypes={itemTypes} multiselect={multiple} size={size} searchable={searchable}
      fileCategory={itemTypes === 'folders' ? undefined : fileCategory} colorfulTagAvatars={colorfulTagAvatars}
      selectionLimit={multiple && limit.trim() ? Number(limit) : undefined} extensions={itemTypes === 'folders' || fileCategory !== undefined ? undefined : suffixes.split(',').map(value => value.trim()).filter(Boolean)}
      onChange={data => { setValues(data.values); setModels(data.items); setReason(data.reason); setError(''); }} />
  </DemoCard>;
};
export default FileFolderPickerPanel;

/**
 * Hosted SharePoint and OneDrive file/folder selection with serializable metadata references.
 *
 * @packageDocumentation
 * @docPage controls/file-folder-picker
 */
import { resolveFileFolderPickerStrings } from './file-folder-picker-strings.internal';
import * as React from 'react';
import { Avatar, Button, Field, Tag, TagGroup, Tooltip, mergeClasses, useInputStyles_unstable } from '@fluentui/react-components';
import { resolveShorthand } from '@fluentui/react-utilities';
import type { FieldControlProps, InputState } from '@fluentui/react-components';
import { AttachRegular, DocumentRegular, DocumentPdfRegular, FolderDocumentRegular, FolderOpenRegular, FolderRegular } from '@fluentui/react-icons';
import { Stage } from '../stage/stage';
import { PickerControlBoundaryContext } from '../pickers/picker-runtime-context.internal';
import { fileFolderPickerIdentity } from './file-folder-picker-identity.internal';
import { useFileFolderPicker } from './useFileFolderPicker.internal';
import { useFileFolderPickerStyles } from './file-folder-picker.styles.internal';
import type { FileFolderPickerProps } from './file-folder-picker.types';

/**
 * Select existing SharePoint/OneDrive Business files or folders through Microsoft's hosted v8 UI.
 * Requires a matching SPFx provider and existing delegated permissions. Only explicit accepted actions emit onChange.
 * Uses native Fluent Field slots; the remote Microsoft surface does not inherit arbitrary host CSS.
 * No persistence or host-side file-byte download is performed. Native upload/create are explicit opt-ins with immediate effects.
 * @param props - Source restrictions, JSON selection ownership and native Field contracts. Uses a single responsive 1080×720 hosted dialog.
 * @returns Owned Field with one Fluent Input-styled tag surface, mode-specific opening action and optional Stage-hosted iframe.
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { FileFolderPicker } from '@apvee/spfx-react-toolkit';
 * <FileFolderPicker label="Documents" defaultValue={[]} multiselect extensions={['.pdf']}
 *   scope={{ kind: 'siteCollection' }} onChange={data => console.log(data.values)} />
 * ```
 *
 * @useWhen Existing SharePoint or OneDrive files/folders should be selected through Microsoft’s hosted picker as serializable metadata references.
 * @avoidWhen Use a dedicated upload workflow when the application needs upload rollback or raw local-file payloads.
 *
 *
 * @example Choose current-user OneDrive root or Projects folder
 * ```tsx
 * import * as React from 'react';
 * import { FileFolderPicker, type FileFolderPickerValue } from '@apvee/spfx-react-toolkit';
 * // Mount beneath the initialized matching SPFx provider with existing SharePoint grants.
 * function OneDriveChoices() {
 *   const [rootValues, setRootValues] = React.useState<readonly FileFolderPickerValue[]>([]);
 *   const [projectValues, setProjectValues] = React.useState<readonly FileFolderPickerValue[]>([]);
 *   return <section>
 *     <FileFolderPicker label="Current user's OneDrive root" preset="focused"
 *       scope={{ kind: 'oneDrive' }} itemTypes="all" multiselect
 *       value={rootValues} onChange={data => setRootValues(data.values)} />
 *     <output>{JSON.stringify(rootValues)}</output>
 *     <FileFolderPicker label="Files in Projects" preset="focused"
 *       scope={{ kind: 'oneDrive', folder: 'Projects' }} itemTypes="files"
 *       value={projectValues} onChange={data => setProjectValues(data.values)} />
 *     <output>{JSON.stringify(projectValues)}</output>
 *   </section>;
 * }
 * ```
 *
 *
 * @example Configure a legacy library opening without a preset
 * ```tsx
 * import * as React from 'react';
 * import { FileFolderPicker, type FileFolderPickerValue } from '@apvee/spfx-react-toolkit';
 * // Mount beneath the initialized matching SPFx provider; supply the actual owning web URL.
 * function LegacyLibraryPicker({ webUrl }: { webUrl: string }) {
 *   const [values, setValues] = React.useState<readonly FileFolderPickerValue[]>([]);
 *   const [readyOpenings, setReadyOpenings] = React.useState(0);
 *   // Deliberately omit preset: these legacy SDK-facing options cannot accompany focused/extended.
 *   return <section>
 *     <FileFolderPicker label="Documents" hint="Select a project document"
 *       scope={{ kind: 'library', webUrl, list: 'Documents' }}
 *       value={values} onChange={data => setValues(data.values)}
 *       title="Project documents" layout="details" stickyHeaders
 *       selectLabel="Use document" cancelLabel="Cancel"
 *       initialSort={{ fieldName: 'FileLeafRef', ascending: true }} navigationVisible
 *       graphEnabled={false} uploadEnabled={false} createFolderEnabled={false}
 *       onReady={() => setReadyOpenings(count => count + 1)} />
 *     <p>Hosted page ready: {readyOpenings} openings</p>
 *     <output>{JSON.stringify(values)}</output>
 *   </section>;
 * }
 * ```
 *
 *
 * @example Seed file labels and show colorful selected avatars
 * ```tsx
 * import * as React from 'react';
 * import { FileFolderPicker, type FileFolderPickerValue } from '@apvee/spfx-react-toolkit';
 * // Supply real saved references and a pure caller-owned label function beneath the matching provider.
 * function LabeledFilePicker({ initialValues, labelFor }: {
 *   initialValues: readonly FileFolderPickerValue[]; labelFor: (value: FileFolderPickerValue) => string;
 * }) {
 *   const [values, setValues] = React.useState<readonly FileFolderPickerValue[]>(
 *     // File references use name snapshots; text belongs to FileFolderPickerOption display models.
 *     () => initialValues.map((value): FileFolderPickerValue => ({ ...value, name: labelFor(value) }))
 *   );
 *   return <section>
 *     <FileFolderPicker label="Documents" preset="focused" scope={{ kind: 'siteCollection' }}
 *       colorfulTagAvatars multiselect value={values} onChange={data => setValues(data.values)} />
 *     <output>{JSON.stringify(values)}</output>
 *   </section>;
 * }
 * ```
 *
 * @example Explore hosted files and folders with explicit source policies
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Select, Switch, makeStyles } from '@fluentui/react-components';
 * import { AccordionGroup, FileFolderPicker, useSPFxContext, useSPFxPageContext, useSPFxServiceScope, getSPFxCorrelationInfo,
 *   type FileFolderPickerScope, type FileFolderPickerValue, type FileFolderPickerOption,
 *   type FileFolderPickerInitialLocation, type FileFolderPickerPreset, type FileFolderPickerProps } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * const useStyles = makeStyles({ controls: { display: 'grid', gap: '12px' }, commands: { display: 'flex', flexWrap: 'wrap', gap: '8px' }, diagnostics: { whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' } });
 * // This stable operational component is mounted only by Start. Reset retires its real runtime.
 * function Subject(props: FileFolderPickerProps): React.ReactElement { return <FileFolderPicker {...props} />; }

 *
 * // Use the existing matching SPFx provider. Start initializes clients; Browse opens the hosted UI.
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const classes = useStyles();
 *   const { spfxContext } = useSPFxContext();
 *   const page = useSPFxPageContext();
 *   const { serviceScope } = useSPFxServiceScope();
 *   const correlation = getSPFxCorrelationInfo(page);
 *   const identity = JSON.stringify([correlation.tenantId, page.user.loginName, page.web.absoluteUrl, page.site.absoluteUrl]);
 *   const boundary = React.useRef({ resetRevision, owner: spfxContext, serviceScope, identity });
 *   const session = boundary.current;
 *   const mounted = React.useRef(true);
 *   React.useEffect(() => () => { mounted.current = false; }, []);
 *   const [started, setStarted] = React.useState(false);
 *   const [customStrings, setCustomStrings] = React.useState(false);
 *   const [kind, setKind] = React.useState<FileFolderPickerScope['kind']>('siteCollection');
 *   const [web, setWeb] = React.useState(''), [list, setList] = React.useState(''), [folder, setFolder] = React.useState('');
 *   const [startWeb, setStartWeb] = React.useState(''), [startList, setStartList] = React.useState(''), [startFolder, setStartFolder] = React.useState('');
 *   const [initialLocation, setInitialLocation] = React.useState<FileFolderPickerInitialLocation>();
 *   const [scope, setScope] = React.useState<FileFolderPickerScope>({ kind: 'siteCollection' });
 *   const [itemTypes, setItemTypes] = React.useState<'files' | 'folders' | 'all'>('files');
 *   const [fileCategory, setFileCategory] = React.useState<FileFolderPickerProps['fileCategory']>();
 *   const [multiple, setMultiple] = React.useState(false), [searchable, setSearchable] = React.useState(true);
 *   const [colorfulTagAvatars, setColorfulTagAvatars] = React.useState(false);
 *   const [size, setSize] = React.useState<'medium' | 'large'>('medium'), [suffixes, setSuffixes] = React.useState('');
 *   const [limit, setLimit] = React.useState(''), [error, setError] = React.useState('');
 *   const [values, setValues] = React.useState<readonly FileFolderPickerValue[]>([]), [models, setModels] = React.useState<readonly FileFolderPickerOption[]>([]);
 *   const [reason, setReason] = React.useState('No action');
 *   const [preset, setPreset] = React.useState<FileFolderPickerPreset>('focused');
 *   const [uploadEnabled, setUploadEnabled] = React.useState(false), [createFolderEnabled, setCreateFolderEnabled] = React.useState(false);
 *   const [hint, setHint] = React.useState('');
 *   const [ready, setReady] = React.useState(0);
 *
 *   const [unresolved, setUnresolved] = React.useState<readonly FileFolderPickerValue[]>([]);
 *   const current = (): boolean => mounted.current && boundary.current === session;
 *   const active = started && session.resetRevision === resetRevision && session.owner === spfxContext && session.serviceScope === serviceScope && session.identity === identity;
 *   React.useLayoutEffect(() => {
 *     const previous = boundary.current;
 *     if (previous.resetRevision === resetRevision && previous.owner === spfxContext && previous.serviceScope === serviceScope && previous.identity === identity) return;
 *     boundary.current = { resetRevision, owner: spfxContext, serviceScope, identity };
 *     setStarted(false); setCustomStrings(false); setKind('siteCollection'); setWeb(''); setList(''); setFolder('');
 *     setStartWeb(''); setStartList(''); setStartFolder(''); setInitialLocation(undefined); setScope({ kind: 'siteCollection' });
 *     setItemTypes('files'); setFileCategory(undefined); setMultiple(false); setSearchable(true); setColorfulTagAvatars(false);
 *     setSize('medium'); setSuffixes(''); setLimit(''); setError(''); setValues([]); setModels([]); setUnresolved([]);
 *     setReason('No action'); setPreset('focused'); setUploadEnabled(false); setCreateFolderEnabled(false); setHint(''); setReady(0);
 *   }, [resetRevision, spfxContext, serviceScope, identity]);
 *   const apply = (): void => {
 *     if (!current()) return;
 *     // Apply commits local drafts only. The real picker validates policy, URLs and remote boundaries.
 *     if ((kind === 'library' || kind === 'folder') && !list.trim()) { setError('Enter an existing document library.'); return; }
 *     if (kind === 'folder' && !folder.trim()) { setError('Enter an existing folder.'); return; }
 *     const targetWeb = web.trim() || page.web.absoluteUrl;
 *     let next: FileFolderPickerScope;
 *     if (kind === 'siteCollection') next = { kind, siteUrl: web.trim() || undefined };
 *     else if (kind === 'library') next = { kind, webUrl: targetWeb, list: list.trim() };
 *     else if (kind === 'folder') next = { kind, webUrl: targetWeb, list: list.trim(), folder };
 *     else next = { kind, folder: folder || undefined };
 *     const navigation: FileFolderPickerInitialLocation | undefined = kind === 'oneDrive'
 *       ? (startFolder ? { folder: startFolder } : undefined)
 *       : (startWeb.trim() || startList.trim() || startFolder ? { webUrl: startWeb.trim() || undefined, list: startList.trim() || undefined, folder: startFolder || undefined } : undefined);
 *     setInitialLocation(navigation); setScope(next); setValues([]); setModels([]); setUnresolved([]); setError(''); setReason('Source changed explicitly');
 *   };
 *   const controls = <div className={classes.controls}>
 *     <Switch label="Custom browse text" checked={customStrings} onChange={(_, data) => setCustomStrings(data.checked)} />
 *     <Field label="Source"><Select aria-label="Source" value={kind} onChange={(_, data) => { const candidate = (['siteCollection', 'library', 'folder', 'oneDrive'] as const).find(value => value === data.value); if (candidate && candidate !== kind) {
 *         setKind(candidate); setStartWeb(''); setStartList(''); setStartFolder('');
 *       } }}>
 *       <option value="siteCollection">Site collection</option><option value="library">Document library</option><option value="folder">Folder</option><option value="oneDrive">OneDrive Business</option>
 *     </Select></Field>
 *     {kind !== 'oneDrive' ? <Field label="Web / site collection URL (empty: current)"><Input aria-label="Web / site collection URL (empty: current)" value={web} onChange={(_, data) => setWeb(data.value)} autoComplete="off" /></Field> : <p>Uses the current authenticated user’s OneDrive Business.</p>}
 *     {kind === 'library' || kind === 'folder' ? <Field label="Existing document library title"><Input aria-label="Existing document library title" value={list} onChange={(_, data) => setList(data.value)} autoComplete="off" /></Field> : undefined}
 *     {kind === 'folder' || kind === 'oneDrive' ? <Field label="Existing folder path"><Input aria-label="Existing folder path" value={folder} onChange={(_, data) => setFolder(data.value)} autoComplete="off" /></Field> : undefined}
 *     <Field label="Preset"><Select aria-label="Preset" value={preset} onChange={(_, data) => setPreset(data.value === 'extended' ? 'extended' : 'focused')}><option value="focused">Focused</option><option value="extended">Extended</option></Select></Field>
 *     <Field label="Selectable items"><Select aria-label="Selectable items" value={itemTypes} onChange={(_, data) => { setItemTypes(data.value === 'folders' ? 'folders' : data.value === 'all' ? 'all' : 'files'); if (data.value === 'folders') setFileCategory(undefined); }}><option value="files">Files</option><option value="folders">Folders</option><option value="all">Files and folders</option></Select></Field>
 *     <Switch label="Multiple selection" checked={multiple} onChange={(_, data) => setMultiple(data.checked)} />
 *     <AccordionGroup><AccordionGroup.Item value="options" title="Options">
 *     {kind !== 'oneDrive' ? <>
 *       <Field label="Initial web URL (navigation only)"><Input aria-label="Initial web URL (navigation only)" value={startWeb} onChange={(_, data) => setStartWeb(data.value)} /></Field>
 *       <Field label="Initial library title (navigation only)"><Input aria-label="Initial library title (navigation only)" value={startList} onChange={(_, data) => setStartList(data.value)} /></Field>
 *     </> : undefined}
 *     <Field label="Initial folder (navigation only)"><Input aria-label="Initial folder (navigation only)" value={startFolder} onChange={(_, data) => setStartFolder(data.value)} /></Field>
 *     <p>Apply source also applies the optional starting location. Changing source clears the starting-location fields. A SharePoint starting folder requires a starting library; OneDrive accepts a folder only.</p>
 *     {itemTypes !== 'folders' ? <>
 *       <Field label="File category"><Select aria-label="File category" value={fileCategory ?? ''} onChange={(_, data) => {
 *         const category = (['images', 'video', 'audio', 'media', 'documents'] as const).find(value => value === data.value);
 *         setFileCategory(category); if (category) setSuffixes('');
 *       }}><option value="">Unrestricted / custom extensions</option><option value="images">Images</option><option value="video">Video</option><option value="audio">Audio</option><option value="media">Media</option><option value="documents">Documents</option></Select></Field>
 *       <Field label="File extensions (comma separated)"><Input aria-label="File extensions (comma separated)" disabled={fileCategory !== undefined} value={suffixes} onChange={(_, data) => setSuffixes(data.value)} /></Field>
 *       <p>A category replaces custom extensions and checks file suffixes. In mixed mode folders remain selectable. Changing this policy preserves selected references.</p>
 *     </> : undefined}
 *     {multiple ? <Field label="Selection limit (empty: default)"><Input aria-label="Selection limit (empty: default)" value={limit} onChange={(_, data) => setLimit(data.value)} /></Field> : undefined}
 *     <Field label="Tag size"><Select aria-label="Tag size" value={size} onChange={(_, data) => setSize(data.value === 'large' ? 'large' : 'medium')}><option>medium</option><option>large</option></Select></Field>
 *     <Switch label="Colorful tag avatars" checked={colorfulTagAvatars} onChange={(_, data) => setColorfulTagAvatars(data.checked)} />
 *     <Switch label="Hosted search" checked={searchable} onChange={(_, data) => setSearchable(data.checked)} />
 *     <Field label="Field description (hint)"><Input aria-label="Field description (hint)" value={hint} onChange={(_, data) => setHint(data.value)} /></Field>
 *     <p>Optional actions write immediately, before confirmation. Cancel cannot undo those writes.</p>
 *     <Switch label="Enable uploads" checked={uploadEnabled} onChange={(_, data) => setUploadEnabled(data.checked)} />
 *     <Switch label="Enable folder creation" checked={createFolderEnabled} onChange={(_, data) => setCreateFolderEnabled(data.checked)} />
 *   </AccordionGroup.Item></AccordionGroup>
 *
 *   </div>;
 *   const commands = <div className={classes.commands}>
 *     <Button disabled={active} onClick={() => { if (current()) setStarted(true); }}>Start</Button>
 *     <Button onClick={apply}>Apply source and clear demo selection</Button>
 *   </div>;
 *   const regions: DemoRegions = {
 *     controls, commands,
 *     status: <p role="status">{error || (active ? 'Started. Browse explicitly to open the hosted picker.' : 'Not started. Start is required before the operational picker mounts.')}</p>,
 *     children: <><p>Apply commits local source and navigation drafts and clears selection. It does not verify remote existence or grant access. Starting navigation never broadens the accepted source.</p>{active ? <Subject
 *       strings={customStrings ? { browse: { files: 'Choose report attachments…', folders: 'Choose a report folder…', all: 'Choose report content…' } } : undefined}
 *       preset={preset} uploadEnabled={uploadEnabled} createFolderEnabled={createFolderEnabled} hint={hint || undefined}
 *       onReady={() => { if (current()) setReady(count => count + 1); }} label="FileFolderPicker" scope={scope} initialLocation={initialLocation}
 *       value={values} initialOptions={models} itemTypes={itemTypes} multiselect={multiple} size={size} searchable={searchable}
 *       fileCategory={itemTypes === 'folders' ? undefined : fileCategory} colorfulTagAvatars={colorfulTagAvatars}
 *       selectionLimit={multiple && limit.trim() ? Number(limit) : undefined}
 *       extensions={itemTypes === 'folders' || fileCategory !== undefined ? undefined : suffixes.split(',').map(value => value.trim()).filter(Boolean)}
 *       onChange={data => { if (current()) { setValues(data.values); setModels(data.items); setUnresolved(data.unresolvedValues); setReason(data.reason); setError(''); } }}
 *       onError={failure => { if (current()) setError(failure.message); }}
 *     /> : <p>Start is required. Reset closes the local session; already sent optional writes may still finish.</p>}</>,
 *     diagnostics: <AccordionGroup><AccordionGroup.Item value="diagnostics" title="Diagnostics and selection JSON"><pre className={classes.diagnostics}>{JSON.stringify({ started: active, scope, initialLocation, preset, values, models, unresolvedValues: unresolved, reason, ready, error }, undefined, 2)}</pre><p>Ready counts actual hosted page-loaded openings, not selection or write settlement. No history is saved.</p></AccordionGroup.Item></AccordionGroup>,
 *   };
 *   return render ? render(regions) : <section>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</section>;
 * }
 * export const demo = { host: 'webPart', kind: 'write', contentPlacement: 'configuration', observe: 'Start the hosted picker, apply an explicit source and starting location, then compare file policies and retained references; Reset closes the local session and requires Start again.' };
 * ```
 */
export function FileFolderPicker(props: FileFolderPickerProps): React.ReactElement {
  const strings = resolveFileFolderPickerStrings(props.strings);
  const styles = useFileFolderPickerStyles(), runtime = useFileFolderPicker(props);
  const Boundary = React.useContext(PickerControlBoundaryContext);
  const mode = props.itemTypes ?? 'files';
  const browseLabel = props.browseLabel ?? strings.browse[mode];
  const placeholder = strings.placeholders[mode];
  const BrowseIcon = mode === 'folders' ? FolderOpenRegular : mode === 'all' ? FolderDocumentRegular : AttachRegular;
  const disabled = props.disabled || !runtime.ready;
  const validationState = props.validationState ?? (runtime.validation !== undefined || runtime.error !== undefined || runtime.initError !== undefined ? 'error' : undefined);
  // Borrow only the native Input style state: the surface remains a group of real buttons/tags,
  // with no hidden/read-only input or editable-text semantics. Fluent owns exact border states,
  // dimensions, contrast and focus animation; only tag layout and the open-session underline extend it.
  const inputStyleState: InputState = {
    components: { root: 'span', input: 'input', contentBefore: 'span', contentAfter: 'span' },
    size: props.size ?? 'medium', appearance: 'outline',
    root: {},
    input: { disabled: props.disabled, 'aria-invalid': validationState === 'error' },
  };
  useInputStyles_unstable(inputStyleState);
  const content = (fieldProps: FieldControlProps): React.ReactElement => <div className={styles.region}>
    <div role="group" aria-label={props['aria-label'] ?? strings.labels.selection} aria-busy={runtime.busy} aria-disabled={props.disabled || undefined}
      className={mergeClasses(inputStyleState.root.className, styles.control, runtime.busy && !props.disabled && styles.open, props.controlClassName)}>
    <div className={styles.selection}>
    {runtime.values.length === 0 ? undefined : <TagGroup className={styles.tags} size={props.size === 'large' ? 'small' : 'extra-small'} disabled={props.disabled} aria-label={props['aria-label'] ?? strings.labels.selection} onDismiss={(_, data) => {
      const value = runtime.values.find(candidate => fileFolderPickerIdentity(candidate) === data.value);
      if (value !== undefined) runtime.remove(value);
    }}>
      {runtime.values.map(value => {
        const text = runtime.label(value);
        const Icon = value.kind === 'folder' ? FolderRegular : /^\.?pdf$/i.test(runtime.extension?.(value) ?? value.name?.split('.').pop() ?? '') ? DocumentPdfRegular : DocumentRegular;
        return <Tag key={fileFolderPickerIdentity(value)} value={fileFolderPickerIdentity(value)} dismissible disabled={props.disabled} className={styles.tag}
          primaryText={{ className: styles.text, title: text }} media={<Avatar icon={<Icon />} aria-label={value.kind === 'folder' ? strings.labels.folder : strings.labels.file}
            color={props.colorfulTagAvatars ? 'colorful' : undefined} idForColor={props.colorfulTagAvatars ? fileFolderPickerIdentity(value) : undefined} />} dismissIcon={{ 'aria-label': strings.labels.removeItem(text) }}>{text}</Tag>;
      })}
    </TagGroup>}
    {runtime.values.length === 0 ? <button {...fieldProps} id={undefined} type="button" className={styles.placeholder} disabled={disabled || runtime.busy} onClick={runtime.browse}
      aria-label={props['aria-label']} aria-labelledby={props['aria-labelledby'] ?? fieldProps['aria-labelledby']}>{placeholder}</button> : undefined}
    </div>
    <Tooltip content={browseLabel} relationship="description">
      <Button {...fieldProps} type="button" id={props.id ?? fieldProps.id} ref={runtime.browseButton} className={styles.action} appearance="transparent" size={props.size === 'large' ? 'medium' : 'small'} disabled={disabled || runtime.busy} onClick={runtime.browse}
        aria-haspopup="dialog" aria-label={props['aria-label'] ?? browseLabel} aria-labelledby={props['aria-labelledby'] ?? fieldProps['aria-labelledby']}
        icon={<BrowseIcon />} />
    </Tooltip>
    </div>
  </div>;
  return <>
    <Field className={mergeClasses(styles.root, props.className)} label={props.id === undefined ? props.label : resolveShorthand(props.label, { defaultProps: { htmlFor: props.id } })} hint={props.hint} required={props.required}
      validationState={validationState} validationMessage={props.validationMessage ?? runtime.validation ?? runtime.error ?? runtime.initError}>
      {fieldProps => Boundary === undefined ? content(fieldProps) : <Boundary pending={runtime.hydrating}>{content(fieldProps)}</Boundary>}
    </Field>
    <Stage presentation="dialog" size="large" className={styles.dialog} open={runtime.dialog !== undefined} onOpenChange={(_, data) => { if (!data.open) runtime.cancel(); }}>
      <Stage.Header title={browseLabel} showCloseButton={false} className={styles.hiddenHeader} />
      <Stage.Body className={styles.dialogBody}>{runtime.dialog === undefined ? undefined : <iframe key={runtime.dialog.name} name={runtime.dialog.name} ref={runtime.attachFrame} title={browseLabel} className={styles.frame} />}</Stage.Body>
    </Stage>
  </>;
}

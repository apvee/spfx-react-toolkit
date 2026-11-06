/**
 * @packageDocumentation
 * @docPage controls/pickers
 */
import * as React from 'react';
import type { ContentTypePickerProps, ContentTypePickerOption } from './picker.types';
import type { ContentTypePickerValue } from './picker-values.types';
import { createContentTypePickerSource } from './content-type-picker.source.internal';
import { usePickerSourceContext } from './usePickerSourceContext.internal';
import { usePickerRuntime } from './usePickerRuntime.internal';
import { PickerRenderer } from './picker-renderer.internal';

/**
 * Select read-only SharePoint content type references from a collection root or one list.
 * Requires an SPFx toolkit provider with SPHttpClient. Site scope (default) resolves the owning
 * collection root even for a subweb URL; list scope preserves the owning web and exact list copy ID.
 * Defaults to Dropdown (historical combobox presentation), single selection and includeHidden=true. Opening retrieves every REST
 * page, sorts by Name then ID, and caches the complete catalog. Typed search matches name prefixes
 * locally and case-insensitively; no personal history or initial 30-item cap applies.
 * Inherits PickerProps model seeds, reactive availableOptions, array selection events, errors,
 * accessibility and fixed presentation slots. ReadOnly/Sealed do not remove choices. Invalid
 * IDs/configuration throw synchronously; request failures use the inherited onError contract.
 * @param props - Catalog scope/list selector, optional group/ID/hidden filters, selection models and native presentation.
 * @returns A Fluent Field with Dropdown or TagPicker. Single and multiple selections both use scoped reference arrays.
 * @example Current collection catalog
 * ```tsx
 * import * as React from 'react';
 * import { ContentTypePicker } from '@apvee/spfx-react-toolkit';
 *
 * <ContentTypePicker
 *   label="Site content type"
 *   webUrl="https://tenant.sharepoint.com/sites/hr/sub"
 * />;
 * ```
 *
 * @example List catalog
 * ```tsx
 * import * as React from 'react';
 * import { ContentTypePicker } from '@apvee/spfx-react-toolkit';
 * <ContentTypePicker label="List content type" scope="list" list="Documents" contentTypeId="0x0101" includeDescendants multiselect />
 * ```
 *
 * @useWhen A user must select content types from a site collection, web or list catalog.
 * @avoidWhen Use FieldPicker for column definitions within a list.
 *
 * @scenario Explore site and list content type references
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Select, Switch, makeStyles } from '@fluentui/react-components';
 * import { ContentTypePicker, useSPFxContext, type ContentTypePickerOption, type PickerSelectionData } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Configuration = {
 *   presentation: 'combobox' | 'tagPicker'; multi: boolean; searchable: boolean;
 *   colorful: boolean; customStrings: boolean; customStatus: boolean; header: string;  webUrl: string; list: string; scope: 'site' | 'list'; group: string; includeHidden: boolean; contentTypeId: string; includeDescendants: boolean;
 * };
 * const baseline: Configuration = { presentation: 'combobox', multi: false, searchable: true, colorful: false, customStrings: false, customStatus: false, header: '', webUrl: '', list: '', scope: 'site', group: '', includeHidden: true, contentTypeId: '', includeDescendants: false };
 * const useStyles = makeStyles({ controls: { display: 'grid', gap: '12px' }, commands: { display: 'flex', flexWrap: 'wrap', gap: '8px' }, diagnostics: { whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' } });
 *
 * // Mount beneath the existing SPFxWebPartProvider/WebPartShell, never another provider.
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const classes = useStyles();
 *   const { spfxContext } = useSPFxContext();
 *   const scope = spfxContext.serviceScope;
 *   const boundary = React.useRef({ resetRevision, owner: spfxContext, scope });
 *   const session = boundary.current;
 *   const [config, setConfig] = React.useState<Configuration>(baseline);
 *   const [started, setStarted] = React.useState(false);
 *   const [selection, setSelection] = React.useState<PickerSelectionData<ContentTypePickerOption> | undefined>();
 *   const [available, setAvailable] = React.useState<readonly ContentTypePickerOption[] | undefined>();
 *   const [error, setError] = React.useState('');
 *   const [refresh, setRefresh] = React.useState(0);
 *   const current = (): boolean => boundary.current === session;
 *   const active = started && session.resetRevision === resetRevision && session.owner === spfxContext && session.scope === scope;
 *   React.useLayoutEffect(() => {
 *     const observed = boundary.current;
 *     if (observed.resetRevision === resetRevision && observed.owner === spfxContext && observed.scope === scope) return;
 *     boundary.current = { resetRevision, owner: spfxContext, scope };
 *     setStarted(false); setConfig(baseline); setSelection(undefined); setAvailable(undefined); setError(''); setRefresh(0);
 *   }, [resetRevision, spfxContext, scope]);
 *   const update = (patch: Partial<Configuration>): void => {
 *     if (!current()) return;
 *     const next = { ...config, ...patch };
 *     const normalizedWeb = (draft: string): string => {
 *       const web = draft.trim() || spfxContext.pageContext.web.absoluteUrl;
 *       try { return new URL(web).href.replace(/\/$/, '').toLowerCase(); }
 *       catch { return web; } // Keep unarmed partial URL drafts editable; the picker validates its actual source.
 *     };
 *     const sourceKey = (value: Configuration): string => JSON.stringify([normalizedWeb(value.webUrl.trim().replace(/\/+$/, '')), value.scope, value.scope === 'list' ? value.list.trim() : undefined, value.group.toLowerCase() || undefined, value.includeHidden, value.contentTypeId.toLowerCase() || undefined, value.includeDescendants]);
 *     // Retire old commands synchronously without remounting the operational picker.
 *     if (sourceKey(config) !== sourceKey(next)) boundary.current = { ...session };
 *     if (normalizedWeb(config.webUrl.trim().replace(/\/+$/, '')) !== normalizedWeb(next.webUrl.trim().replace(/\/+$/, '')) ||
 *       config.scope !== next.scope || next.scope === 'list' && config.list.trim() !== next.list.trim()) {
 *       setSelection(undefined); setAvailable(undefined);
 *     }
 *     setConfig(previous => ({ ...previous, ...patch }));
 *   };
 *   const target = config.webUrl.trim().replace(/\/+$/, '') || undefined;
 *   const contentTypeFilterIds = Array.from(new Set(['0x01', '0x0101',
 *     ...(selection?.selectedItems ?? []).map(option => option.value.contentTypeId),
 *     ...(config.contentTypeId ? [config.contentTypeId] : []),
 *   ]));
 *   const common = {
 *     label: 'ContentTypePicker', presentation: config.presentation, multiselect: config.multi, searchable: config.searchable,
 *     colorfulTagAvatars: config.colorful, refreshKey: refresh,
 *     selectedOptions: selection?.selectedOptions ?? [], initialOptions: selection?.selectedItems, availableOptions: available,
 *     suggestionsHeader: config.header.trim() || undefined,
 *     emptyContent: config.customStatus ? 'No candidates: type at least two characters.' : undefined,
 *     loadingContent: config.customStatus ? <span>Fetching live suggestions…</span> : undefined,
 *     errorContent: config.customStatus ? { children: 'Source unavailable: inspect diagnostics and retry.' } : undefined,
 *     strings: config.customStrings ? { labels: { selectOne: 'Choose source content', selectMany: 'Choose source content', search: 'Find a source item', retry: 'Try this source again' }, status: { noOptions: 'No matching source items' } } : undefined,
 *     onOptionSelect: (_: unknown, data: PickerSelectionData<ContentTypePickerOption>): void => { if (current()) { setSelection(data); setError(''); } },
 *     onError: (failure: Error): void => { if (current()) setError(failure.message); },
 *   };
 *   const ready = config.scope === 'site' || Boolean(config.list.trim());
 *   const controls = <div className={classes.controls}>
 *     <Field label="Presentation"><Select aria-label="Presentation" value={config.presentation} onChange={(_, data) => update({ presentation: data.value === 'tagPicker' ? 'tagPicker' : 'combobox' })}><option value="combobox">Dropdown</option><option value="tagPicker">TagPicker</option></Select></Field>
 *     <Switch label="Multiple selection" checked={config.multi} onChange={(_, data) => update({ multi: data.checked })} />
 *     <Switch label="Enable search" checked={config.searchable} onChange={(_, data) => update({ searchable: data.checked })} />
 *     <Switch label="Colorful tag avatars" checked={config.colorful} disabled={config.presentation !== 'tagPicker'} onChange={(_, data) => update({ colorful: data.checked })} />
 *     <Field label="Suggestion header (empty: default)"><Input aria-label="Suggestion header" value={config.header} onChange={(_, data) => update({ header: data.value })} /></Field>
 *     <Switch label="Customize status content" checked={config.customStatus} onChange={(_, data) => update({ customStatus: data.checked })} />
 *     <Switch label="Custom search and status text" checked={config.customStrings} onChange={(_, data) => update({ customStrings: data.checked })} />
 *     <Field label="Web URL (empty: current web)"><Input aria-label="Web URL" value={config.webUrl} onChange={(_, data) => update({ webUrl: data.value })} /></Field>
 *     <Field label="Exact list title (required for list scope)"><Input aria-label="Exact list title" value={config.list} onChange={(_, data) => update({ list: data.value })} /></Field>
 *     <Field label="Content type scope"><Select aria-label="Content type scope" value={config.scope} onChange={(_, data) => update({ scope: data.value === 'list' ? 'list' : 'site' })}><option value="site">Site collection</option><option value="list">List</option></Select></Field>
 *     <Field label="Exact content type group (empty: all)"><Input aria-label="Content type group" value={config.group} onChange={(_, data) => update({ group: data.value })} /></Field>
 *     <Switch label="Include hidden content types" checked={config.includeHidden} onChange={(_, data) => update({ includeHidden: data.checked })} />
 *     <Field label="Content type ID filter" hint="Choose a built-in parent or a currently selected metadata ID."><Select aria-label="Content type ID filter" value={config.contentTypeId} onChange={(_, data) => { if (data.value === '' || contentTypeFilterIds.includes(data.value)) update({ contentTypeId: data.value, includeDescendants: data.value ? config.includeDescendants : false }); }}><option value="">All content types</option>{contentTypeFilterIds.map(id => <option key={id} value={id}>{id === '0x01' ? 'Item (0x01)' : id === '0x0101' ? 'Document (0x0101)' : id}</option>)}</Select></Field>
 *     <Switch label="Include descendants of content type ID" checked={config.includeDescendants} disabled={!config.contentTypeId} onChange={(_, data) => update({ includeDescendants: data.checked })} />
 *   </div>;
 *   const commands = <div className={classes.commands}>
 *     <Button disabled={!ready} onClick={() => { if (current()) setStarted(true); }}>Start</Button>
 *     <Button onClick={() => { if (current()) setAvailable(selection?.selectedItems ?? []); }}>Use confirmed models as consumer options</Button>
 *     <Button onClick={() => { if (current()) setAvailable([]); }}>Empty consumer options</Button>
 *     <Button onClick={() => { if (current()) setSelection(undefined); }}>Clear controlled selection</Button>
 *     <Button onClick={() => { if (current()) { setError(''); setRefresh(value => value + 1); } }}>Refresh / retry</Button>
 *   </div>;
 *   const values = selection?.selectedValues ?? [];
 *   const regions: DemoRegions = {
 *     controls, commands,
 *     status: <p role="status">{error || (active ? 'Started: use the picker below.' : 'Not started. Start explicitly to read the source.')}</p>,
 *     children: <><p>Site scope reads the actual owning collection root, including from a subweb. List scope reads only attached definitions. Group and hierarchical ID filters apply locally over the full catalog; read-only and sealed definitions remain choices. A real effective web, scope or list change clears selection and consumer options. Presentation and equivalent source edits retain both. Selection never adds definitions, enables content types or synchronizes a hub.</p><p>Sample defaults explicitly select Dropdown, single selection, search enabled and colorful avatars off. Copy confirmed models, clear selection, then reopen to compare independent consumer choices. Empty consumer options retains selection. No history is saved. Presentation, labels, strings and equivalent models retain the live picker. Popup search keeps the selected label in the trigger. Clearing query only clears search; selection clear removes its value. Multiple selected rows and tags can be removed without searching. Escape returns focus; Tab leaves without accepting a highlighted row.</p>{!ready ? <p>Enter an exact list title to query list content types.</p> : active ? (config.scope === 'site' ? <ContentTypePicker {...common} webUrl={target} group={config.group || undefined} includeHidden={config.includeHidden} contentTypeId={config.contentTypeId || undefined} includeDescendants={config.includeDescendants} scope="site" /> : <ContentTypePicker {...common} webUrl={target} group={config.group || undefined} includeHidden={config.includeHidden} contentTypeId={config.contentTypeId || undefined} includeDescendants={config.includeDescendants} scope="list" list={config.list.trim()} />) : <p>Start is required before the operational picker mounts.</p>}</>,
 *     diagnostics: <pre className={classes.diagnostics}>{JSON.stringify({ started: active, selectedOptions: selection?.selectedOptions ?? [], selectedItems: selection?.selectedItems ?? [], selectedValues: values, unresolvedOptions: selection?.unresolvedOptions ?? [], reason: selection?.reason ?? 'No action', optionValue: selection?.optionValue, optionText: selection?.optionText, roundTrip: JSON.stringify(JSON.parse(JSON.stringify(values))) === JSON.stringify(values), consumerOptions: available, error }, undefined, 2)}</pre>,
 *   };
 *   return render ? render(regions) : <section>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</section>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Start a site or list content-type picker, apply group and parent-ID filters, and observe that source changes clear demo selection while presentation changes retain it.' };
 * ```
 */
export function ContentTypePicker(props: ContentTypePickerProps): React.ReactElement {
  const context = usePickerSourceContext();
  const source = createContentTypePickerSource(props, context);
  const runtime = usePickerRuntime<ContentTypePickerValue, ContentTypePickerOption>(props, source, { multiselect: false });
  return <PickerRenderer<ContentTypePickerValue, ContentTypePickerOption> props={props} runtime={runtime} defaultPresentation="combobox" defaultMultiselect={false} />;
}

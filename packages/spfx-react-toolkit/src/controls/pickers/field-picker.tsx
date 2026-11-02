/**
 * @packageDocumentation
 * @docPage controls/pickers
 */
import * as React from 'react';
import type { FieldPickerProps, FieldPickerOption } from './picker.types';
import { createFieldPickerSource } from './field-picker.source.internal';
import { usePickerSourceContext } from './usePickerSourceContext.internal';
import type { FieldPickerValue } from './picker-values.types';
import { usePickerRuntime } from './usePickerRuntime.internal';
import { PickerRenderer } from './picker-renderer.internal';

/**
 * Search list fields or site columns beneath an SPFx provider, with icons based on verified field types.
 * Defaults: Dropdown (historical combobox presentation), multiple selection, hidden/read-only included, Title ascending.
 * Opening the menu loads all matching fields, following native paging. Typed search filters the complete memory-cached catalog locally, case-insensitively; no initial display limit.
 * @param props - List selector (omitted for site columns), references and Fluent options.
 * @returns A native Fluent Field and picker; blank clearing emits an empty array.
 * @example Select writable list fields
 * ```tsx
 * import * as React from 'react';
 * import { FieldPicker } from '@apvee/spfx-react-toolkit';
 * <FieldPicker label="Campi" list="Projects" includeReadOnly={false} />
 * ```
 *
 * @useWhen A user must select fields from a specified SharePoint list.
 * @avoidWhen Use ContentTypePicker when the configuration requires content-type identities.
 *
 * @example Explore site columns and list fields
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Select, Switch, makeStyles } from '@fluentui/react-components';
 * import { FieldPicker, useSPFxContext, type FieldPickerOption, type PickerSelectionData } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Configuration = {
 *   presentation: 'combobox' | 'tagPicker'; multi: boolean; searchable: boolean;
 *   colorful: boolean; customStrings: boolean; customStatus: boolean; header: string;  webUrl: string; list: string;
 * };
 * const baseline: Configuration = { presentation: 'combobox', multi: false, searchable: true, colorful: false, customStrings: false, customStatus: false, header: '', webUrl: '', list: '' };
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
 *   const [selection, setSelection] = React.useState<PickerSelectionData<FieldPickerOption> | undefined>();
 *   const [available, setAvailable] = React.useState<readonly FieldPickerOption[] | undefined>();
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
 *     const sourceKey = (value: Configuration): string => JSON.stringify([normalizedWeb(value.webUrl), value.list.trim() || undefined]);
 *     // Retire old commands synchronously without remounting the operational picker.
 *     if (sourceKey(config) !== sourceKey(next)) boundary.current = { ...session };
 *     setConfig(previous => ({ ...previous, ...patch }));
 *   };
 *   const target = config.webUrl.trim() || undefined;
 *   const common = {
 *     label: 'FieldPicker', presentation: config.presentation, multiselect: config.multi, searchable: config.searchable,
 *     colorfulTagAvatars: config.colorful, refreshKey: refresh,
 *     selectedOptions: selection?.selectedOptions ?? [], initialOptions: selection?.selectedItems, availableOptions: available,
 *     suggestionsHeader: config.header.trim() || undefined,
 *     emptyContent: config.customStatus ? 'No candidates: type at least two characters.' : undefined,
 *     loadingContent: config.customStatus ? <span>Fetching live suggestions…</span> : undefined,
 *     errorContent: config.customStatus ? { children: 'Source unavailable: inspect diagnostics and retry.' } : undefined,
 *     strings: config.customStrings ? { labels: { selectOne: 'Choose source content', selectMany: 'Choose source content', search: 'Find a source item', retry: 'Try this source again' }, status: { noOptions: 'No matching source items' } } : undefined,
 *     onOptionSelect: (_: unknown, data: PickerSelectionData<FieldPickerOption>): void => { if (current()) { setSelection(data); setError(''); } },
 *     onError: (failure: Error): void => { if (current()) setError(failure.message); },
 *   };
 *   const ready = true;
 *   const controls = <div className={classes.controls}>
 *     <Field label="Presentation"><Select aria-label="Presentation" value={config.presentation} onChange={(_, data) => update({ presentation: data.value === 'tagPicker' ? 'tagPicker' : 'combobox' })}><option value="combobox">Dropdown</option><option value="tagPicker">TagPicker</option></Select></Field>
 *     <Switch label="Multiple selection" checked={config.multi} onChange={(_, data) => update({ multi: data.checked })} />
 *     <Switch label="Enable search" checked={config.searchable} onChange={(_, data) => update({ searchable: data.checked })} />
 *     <Switch label="Colorful tag avatars" checked={config.colorful} disabled={config.presentation !== 'tagPicker'} onChange={(_, data) => update({ colorful: data.checked })} />
 *     <Field label="Suggestion header (empty: default)"><Input aria-label="Suggestion header" value={config.header} onChange={(_, data) => update({ header: data.value })} /></Field>
 *     <Switch label="Customize status content" checked={config.customStatus} onChange={(_, data) => update({ customStatus: data.checked })} />
 *     <Switch label="Custom search and status text" checked={config.customStrings} onChange={(_, data) => update({ customStrings: data.checked })} />
 *     <Field label="Web URL (empty: current web)"><Input aria-label="Web URL" value={config.webUrl} onChange={(_, data) => update({ webUrl: data.value })} /></Field>
 *     <Field label="Exact list title (Field: empty means site columns)"><Input aria-label="Exact list title" value={config.list} onChange={(_, data) => update({ list: data.value })} /></Field>
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
 *     children: <><p>An empty list means site columns; a list string is an exact title. The complete paged metadata catalog is cached, then filtered locally by a case-insensitive name prefix. There is no initial-catalog cap.</p><p>Sample defaults explicitly select Dropdown, single selection, search enabled and colorful avatars off. Copy confirmed models, clear selection, then reopen to compare independent consumer choices. Empty consumer options retains selection. No history is saved. Presentation, labels, strings and equivalent models retain the live picker. Popup search keeps the selected label in the trigger. Clearing query only clears search; selection clear removes its value. Multiple selected rows and tags can be removed without searching. Escape returns focus; Tab leaves without accepting a highlighted row.</p>{!ready ? <p></p> : active ? <FieldPicker {...common} webUrl={target} list={config.list.trim() || undefined} /> : <p>Start is required before the operational picker mounts.</p>}</>,
 *     diagnostics: <pre className={classes.diagnostics}>{JSON.stringify({ started: active, selectedOptions: selection?.selectedOptions ?? [], selectedItems: selection?.selectedItems ?? [], selectedValues: values, unresolvedOptions: selection?.unresolvedOptions ?? [], reason: selection?.reason ?? 'No action', optionValue: selection?.optionValue, optionText: selection?.optionText, roundTrip: JSON.stringify(JSON.parse(JSON.stringify(values))) === JSON.stringify(values), consumerOptions: available, error }, undefined, 2)}</pre>,
 *   };
 *   return render ? render(regions) : <section>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</section>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Start the field picker for site columns or an existing list, then search its complete catalog and compare Dropdown with removable selected tags.' };
 * ```
 */
export function FieldPicker(props: FieldPickerProps): React.ReactElement {
  const context = usePickerSourceContext(), source = createFieldPickerSource(props, context);
  const runtime = usePickerRuntime<FieldPickerValue, FieldPickerOption>(props, source, { multiselect: true });
  return <PickerRenderer<FieldPickerValue, FieldPickerOption> props={props} runtime={runtime} defaultPresentation="combobox" defaultMultiselect showBlankOption={props.showBlankOption} />;
}

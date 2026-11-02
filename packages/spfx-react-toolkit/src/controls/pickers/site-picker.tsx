/**
 * SharePoint metadata and principal discovery, scoped references and explicit selection.
 *
 * @packageDocumentation
 * @docPage controls/pickers
 */
import * as React from 'react';
import { useSPFxPnPContext } from '../../hooks/useSPFxPnPContext';
import type { SitePickerProps, SitePickerOption } from './picker.types';
import { createSitePickerSource } from './site-picker.source.internal';
import type { SiteSearchTransport } from './site-picker.source.internal';
import { usePickerSourceContext } from './usePickerSourceContext.internal';
import type { SitePickerValue } from './picker-values.types';
import { usePickerRuntime } from './usePickerRuntime.internal';
import { PickerRenderer } from './picker-renderer.internal';

/**
 * Search site collections or hubs with the provider's authenticated clients and memory-only query caching.
 * Defaults: site collection mode, Dropdown (historical combobox presentation), single selection, Title ascending. Initial menu includes consumer options, current collection and associated hub; includeCurrentSite/includeHubSite both default true.
 * @param props - Site mode/scope, models, selection and presentation options. KQL is supported only in mode="site".
 * @returns A native Fluent Field and site picker; changing display props retains its runtime.
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { SitePicker } from '@apvee/spfx-react-toolkit';
 * <SitePicker label="Sito" mode="site" includeCurrentSite includeHubSite />
 * ```
 *
 * @useWhen A user must select site collection references, registered hubs or associated sites.
 * @avoidWhen Use ListPicker to select a list within a known source web.
 *
 *
 * @example Compare contextual search and custom feedback
 * ```tsx
 * import * as React from 'react';
 * import { SitePicker, ListPicker, PrincipalPicker } from '@apvee/spfx-react-toolkit';
 * // Render beneath the initialized matching SPFx provider. Search does not grant access.
 * function SearchFeedbackChoices() {
 *   return <section>
 *     <SitePicker label="Site" minimumSearchLength={3} includeCurrentSite includeHubSite />
 *     <ListPicker label="List" emptyContent="Nessuna lista disponibile per i criteri correnti" />
 *     <PrincipalPicker label="Principals" loadingContent="Ricerca persone e gruppi…" />
 *   </section>;
 * }
 * ```
 *
 * @example Explore site selection and contextual search
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Select, Switch, makeStyles } from '@fluentui/react-components';
 * import { SitePicker, useSPFxContext, type SitePickerOption, type PickerSelectionData } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Configuration = {
 *   presentation: 'combobox' | 'tagPicker'; multi: boolean; searchable: boolean;
 *   colorful: boolean; customStrings: boolean; customStatus: boolean; header: string;  includeCurrentSite: boolean; includeHubSite: boolean;
 * };
 * const baseline: Configuration = { presentation: 'combobox', multi: false, searchable: true, colorful: false, customStrings: false, customStatus: false, header: '', includeCurrentSite: true, includeHubSite: true };
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
 *   const [selection, setSelection] = React.useState<PickerSelectionData<SitePickerOption> | undefined>();
 *   const [available, setAvailable] = React.useState<readonly SitePickerOption[] | undefined>();
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
 *     const sourceKey = (value: Configuration): string => JSON.stringify([value.includeCurrentSite, value.includeHubSite]);
 *     // Retire old commands synchronously without remounting the operational picker.
 *     if (sourceKey(config) !== sourceKey(next)) boundary.current = { ...session };
 *     setConfig(previous => ({ ...previous, ...patch }));
 *   };
 *   const common = {
 *     label: 'SitePicker', presentation: config.presentation, multiselect: config.multi, searchable: config.searchable,
 *     colorfulTagAvatars: config.colorful, refreshKey: refresh,
 *     selectedOptions: selection?.selectedOptions ?? [], initialOptions: selection?.selectedItems, availableOptions: available,
 *     suggestionsHeader: config.header.trim() || undefined,
 *     emptyContent: config.customStatus ? 'No candidates: type at least two characters.' : undefined,
 *     loadingContent: config.customStatus ? <span>Fetching live suggestions…</span> : undefined,
 *     errorContent: config.customStatus ? { children: 'Source unavailable: inspect diagnostics and retry.' } : undefined,
 *     strings: config.customStrings ? { labels: { selectOne: 'Choose source content', selectMany: 'Choose source content', search: 'Find a source item', retry: 'Try this source again' }, status: { noOptions: 'No matching source items' } } : undefined,
 *     onOptionSelect: (_: unknown, data: PickerSelectionData<SitePickerOption>): void => { if (current()) { setSelection(data); setError(''); } },
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
 *     <Switch label="Offer current site collection" checked={config.includeCurrentSite} onChange={(_, data) => update({ includeCurrentSite: data.checked })} />
 *     <Switch label="Offer associated hub" checked={config.includeHubSite} onChange={(_, data) => update({ includeHubSite: data.checked })} />
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
 *     children: <><p>Current collection and actually associated hub remain pinned during typed search, outside the ordinary five-item suggestion cap. There is no SitePicker webUrl prop. Search uses the real authenticated SharePoint search source.</p><p>Sample defaults explicitly select Dropdown, single selection, search enabled and colorful avatars off. Copy confirmed models, clear selection, then reopen to compare independent consumer choices. Empty consumer options retains selection. No history is saved. Presentation, labels, strings and equivalent models retain the live picker. Popup search keeps the selected label in the trigger. Clearing query only clears search; selection clear removes its value. Multiple selected rows and tags can be removed without searching. Escape returns focus; Tab leaves without accepting a highlighted row.</p>{!ready ? <p></p> : active ? <SitePicker {...common} includeCurrentSite={config.includeCurrentSite} includeHubSite={config.includeHubSite} /> : <p>Start is required before the operational picker mounts.</p>}</>,
 *     diagnostics: <pre className={classes.diagnostics}>{JSON.stringify({ started: active, selectedOptions: selection?.selectedOptions ?? [], selectedItems: selection?.selectedItems ?? [], selectedValues: values, unresolvedOptions: selection?.unresolvedOptions ?? [], reason: selection?.reason ?? 'No action', optionValue: selection?.optionValue, optionText: selection?.optionText, roundTrip: JSON.stringify(JSON.parse(JSON.stringify(values))) === JSON.stringify(values), consumerOptions: available, error }, undefined, 2)}</pre>,
 *   };
 *   return render ? render(regions) : <section>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</section>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Start the site picker, then compare its current-site and associated-hub choices with typed search while the selected value remains visible.' };
 * ```
 */
export function SitePicker(props: SitePickerProps): React.ReactElement {
  const context = usePickerSourceContext(), pnp = useSPFxPnPContext(undefined, { cache: { enabled: false } });
  const search = React.useMemo<SiteSearchTransport | undefined>(() => {
    const sp = pnp.sp;
    if (sp === undefined) return undefined;
    return async (query, options) => {
      // The owning existing search service registers PnP search lazily when this feature is requested.
      const { createSPFxPnPSearchService } = await import(/* webpackChunkName: 'spfx-picker-search' */ '../../services/spfx-pnp-search.service');
      return createSPFxPnPSearchService(sp, { selectProperties: ['Title', 'Path', 'SiteId', 'DepartmentId'] }).search(builder => {
        const sorted = builder.text(query);
        sorted.toSearchQuery().TrimDuplicates = options.trimDuplicates;
        return sorted;
      }, { pageSize: options.pageSize, startRow: options.startRow });
    };
  }, [pnp.sp]);
  const source = createSitePickerSource(props, context, search), runtime = usePickerRuntime<SitePickerValue, SitePickerOption>(props, source, { multiselect: false });
  return <PickerRenderer<SitePickerValue, SitePickerOption> props={{ ...props, placeholder: props.searchPlaceholder ?? props.placeholder }} runtime={runtime} defaultPresentation="combobox" defaultMultiselect={false} />;
}

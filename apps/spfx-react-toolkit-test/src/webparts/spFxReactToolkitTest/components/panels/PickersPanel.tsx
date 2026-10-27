import * as React from 'react';
import { Accordion, AccordionHeader, AccordionItem, Button, Field, Input, Select, Switch } from '@fluentui/react-components';
import {
  SitePicker, ListPicker, FieldPicker, ViewPicker, ContentTypePicker, ListItemPicker, PrincipalPicker,
  type PickerOption, type PickerValue, type PickerSelectionData,
  type SitePickerOption, type ListPickerOption, type FieldPickerOption, type ViewPickerOption, type ContentTypePickerOption,
  type ListItemPickerOption, type PrincipalPickerOption, type PrincipalType,
} from '@apvee/spfx-react-toolkit';
import { DemoAccordionPanel } from '../shared/DemoAccordionPanel';
import { DemoCard } from '../shared/DemoCard';

const names = ['SitePicker', 'ListPicker', 'FieldPicker', 'ViewPicker', 'ContentTypePicker', 'ListItemPicker', 'PrincipalPicker'] as const;
type Name = typeof names[number];
/** One active live example. Diagnostics are consumer-owned state, never picker-managed persistence. */
const PickersPanel: React.FC = () => {
  const [customStrings, setCustomStrings] = React.useState(false);
  const [name, setName] = React.useState<Name>('SitePicker');
  const [presentation, setPresentation] = React.useState<'combobox' | 'tagPicker'>('combobox');
  const [colorfulTagAvatars, setColorfulTagAvatars] = React.useState(false);
  const [multi, setMulti] = React.useState(false), [searchable, setSearchable] = React.useState(true);
  const [webUrl, setWebUrl] = React.useState(''), [list, setList] = React.useState(''), [textField, setTextField] = React.useState('Title');
  const [includeCurrentSite, setIncludeCurrentSite] = React.useState(true), [includeHubSite, setIncludeHubSite] = React.useState(true);
  const [contentTypeScope, setContentTypeScope] = React.useState<'site' | 'list'>('site');
  const [contentTypeGroup, setContentTypeGroup] = React.useState(''), [includeHidden, setIncludeHidden] = React.useState(true);
  const [contentTypeId, setContentTypeId] = React.useState(''), [includeDescendants, setIncludeDescendants] = React.useState(false);
  const [personal, setPersonal] = React.useState(false), [currentUser, setCurrentUser] = React.useState(false);
  const [types, setTypes] = React.useState<'users' | 'groups' | 'all'>('users');
  const [idMode, setIdMode] = React.useState<'available' | 'entra' | 'sharePoint' | 'both'>('available'), [ensureUser, setEnsureUser] = React.useState(false);
  const [selections, setSelections] = React.useState<Partial<Record<Name, PickerSelectionData<PickerOption>>>>({});
  const [available, setAvailable] = React.useState<Partial<Record<Name, readonly PickerOption[]>>>({});
  const [catalogLimit, setCatalogLimit] = React.useState(30);
  const [header, setHeader] = React.useState(''), [customStatus, setCustomStatus] = React.useState(false);
  const [error, setError] = React.useState(''), [revision, setRevision] = React.useState(0);
  const clearContentTypes = (): void => { setSelections(previous => ({ ...previous, ContentTypePicker: undefined })); setAvailable(previous => ({ ...previous, ContentTypePicker: undefined })); };
  const selection = selections[name];
  const contentTypeFilterIds = Array.from(new Set(['0x01', '0x0101',
    ...(selections.ContentTypePicker?.selectedItems ?? []).flatMap(option => option.value.kind === 'contentType' ? [option.value.contentTypeId] : []),
    ...(contentTypeId ? [contentTypeId] : []),
  ]));
  const update = <T extends PickerOption>(data: PickerSelectionData<T>): void => { setSelections(previous => ({ ...previous, [name]: data })); };
  const common = {
    strings: customStrings ? { labels: { selectOne: 'Choose source content', selectMany: 'Choose source content', search: 'Find a source item', retry: 'Try this source again' }, status: { noOptions: 'No matching source items' } } : undefined,
    label: name, presentation, multiselect: multi, searchable, refreshKey: revision, colorfulTagAvatars,
    selectedOptions: selection?.selectedOptions ?? [],
    suggestionsHeader: header.trim() || undefined,
    emptyContent: customStatus ? 'No candidates: type at least two characters.' : undefined,
    loadingContent: customStatus ? <span>Fetching live suggestions…</span> : undefined,
    errorContent: customStatus ? { children: 'Source unavailable: inspect diagnostics and retry.' } : undefined,
    onError: (failure: Error) => setError(failure.message),
  };
  const principalTypes: readonly PrincipalType[] = types === 'users' ? ['user'] : types === 'groups' ? ['sharePointGroup', 'microsoft365Group', 'securityGroup', 'distributionList'] : ['user', 'sharePointGroup', 'microsoft365Group', 'securityGroup', 'distributionList'];
  const target = webUrl.trim() || undefined;
  const userIdentity = idMode === 'sharePoint' || idMode === 'both' ? { mode: idMode, webUrl: target, ensureUser } : { mode: idMode };
  let subject: React.ReactNode;
  switch (name) {
    case 'SitePicker': subject = <SitePicker {...common} includeCurrentSite={includeCurrentSite} includeHubSite={includeHubSite} initialOptions={selection?.selectedItems as readonly SitePickerOption[] | undefined} availableOptions={available[name] as readonly SitePickerOption[] | undefined} onOptionSelect={(_, data) => update(data)} />; break;
    case 'ListPicker': subject = <ListPicker {...common} initialCatalogLimit={catalogLimit} webUrl={target} initialOptions={selection?.selectedItems as readonly ListPickerOption[] | undefined} availableOptions={available[name] as readonly ListPickerOption[] | undefined} onOptionSelect={(_, data) => update(data)} />; break;
    case 'FieldPicker': subject = <FieldPicker {...common} webUrl={target} list={list.trim() || undefined} initialOptions={selection?.selectedItems as readonly FieldPickerOption[] | undefined} availableOptions={available[name] as readonly FieldPickerOption[] | undefined} onOptionSelect={(_, data) => update(data)} />; break;
    case 'ContentTypePicker': {
      const contentTypeProps = { ...common, webUrl: target, group: contentTypeGroup || undefined, includeHidden, contentTypeId: contentTypeId.trim() || undefined, includeDescendants, initialOptions: selection?.selectedItems as readonly ContentTypePickerOption[] | undefined, availableOptions: available[name] as readonly ContentTypePickerOption[] | undefined, onOptionSelect: (_: unknown, data: PickerSelectionData<ContentTypePickerOption>) => update(data) };
      subject = contentTypeScope === 'site' ? <ContentTypePicker {...contentTypeProps} scope="site" /> : list.trim() ? <ContentTypePicker {...contentTypeProps} scope="list" list={list.trim()} /> : <p>Enter an exact list title to query list content types.</p>;
      break;
    }
    case 'ViewPicker': subject = list.trim() ? <ViewPicker {...common} webUrl={target} list={list} initialOptions={selection?.selectedItems as readonly ViewPickerOption[] | undefined} availableOptions={available[name] as readonly ViewPickerOption[] | undefined} onOptionSelect={(_, data) => update(data)} /> : <p>Enter an exact list title to query views.</p>; break;
    case 'ListItemPicker': subject = list.trim() ? <ListItemPicker {...common} webUrl={target} list={list} textField={textField} initialOptions={selection?.selectedItems as readonly ListItemPickerOption[] | undefined} availableOptions={available[name] as readonly ListItemPickerOption[] | undefined} onOptionSelect={(_, data) => update(data)} /> : <p>Enter an exact list title and display field to query items.</p>; break;
    case 'PrincipalPicker': subject = <PrincipalPicker {...common} webUrl={target} principalTypes={principalTypes} userIdentity={userIdentity} loadPersonalSuggestions={personal} initialSuggestions={{ includeCurrentUser: currentUser }} initialOptions={selection?.selectedItems as readonly PrincipalPickerOption[] | undefined} availableOptions={available[name] as readonly PrincipalPickerOption[] | undefined} onOptionSelect={(_, data) => update(data)} />; break;
  }
  const values: readonly PickerValue[] = selection?.selectedValues ?? [];
  return <DemoCard title="SharePoint pickers" iconName="Search" description="One live authenticated picker; the Property Pane exposes the seven corresponding native fields." contentPlacement="configuration" controls={<>
    <Switch data-picker-demo="strings" label="Custom search and status text" checked={customStrings} onChange={(_, data) => setCustomStrings(data.checked)} />
    <Field label="Control"><Select value={name} onChange={(_, data) => { const candidate = names.find(value => value === data.value); if (candidate) setName(candidate); }}>{names.map(value => <option key={value}>{value}</option>)}</Select></Field>
    <Field label="Presentation"><Select value={presentation} onChange={(_, data) => setPresentation(data.value === 'tagPicker' ? 'tagPicker' : 'combobox')}><option value="combobox">Dropdown</option><option value="tagPicker">TagPicker</option></Select></Field>
    <Switch label="Multiple selection" checked={multi} onChange={(_, data) => setMulti(data.checked)} />
    <Switch label="Colorful tag avatars" checked={colorfulTagAvatars} disabled={presentation !== 'tagPicker'} onChange={(_, data) => setColorfulTagAvatars(data.checked)} />
    <Switch label="Enable search" checked={searchable} onChange={(_, data) => setSearchable(data.checked)} />
    <Field label="Suggestion header (empty: default)"><Input value={header} onChange={(_, data) => setHeader(data.value)} /></Field>
    <Switch label="Customize status content" checked={customStatus} onChange={(_, data) => setCustomStatus(data.checked)} />
    <Field label="Web URL (empty: current web)"><Input value={webUrl} onChange={(_, data) => { if (webUrl.trim().toLowerCase().replace(/\/+$/, '') !== data.value.trim().toLowerCase().replace(/\/+$/, '')) clearContentTypes(); setWebUrl(data.value); }} /></Field>
    <Field label="Exact list title (Field: empty means site columns)"><Input value={list} onChange={(_, data) => { if (contentTypeScope === 'list' && list.trim() !== data.value.trim()) clearContentTypes(); setList(data.value); }} /></Field>
    {name === 'ContentTypePicker' && <>
      <Field label="Content type scope"><Select value={contentTypeScope} onChange={(_, data) => { const next = data.value === 'list' ? 'list' : 'site'; if (next !== contentTypeScope) clearContentTypes(); setContentTypeScope(next); }}><option value="site">Site collection</option><option value="list">List</option></Select></Field>
      <Field label="Exact content type group (empty: all)"><Input value={contentTypeGroup} onChange={(_, data) => setContentTypeGroup(data.value)} /></Field>
      <Switch label="Include hidden content types" checked={includeHidden} onChange={(_, data) => setIncludeHidden(data.checked)} />
      <Field label="Content type ID filter" hint="Choose a built-in parent or a currently selected content type."><Select value={contentTypeId} onChange={(_, data) => { if (data.value !== '' && !contentTypeFilterIds.includes(data.value)) return; setContentTypeId(data.value); if (!data.value) setIncludeDescendants(false); }}><option value="">All content types</option>{contentTypeFilterIds.map(id => <option key={id} value={id}>{id === '0x01' ? 'Item (0x01)' : id === '0x0101' ? 'Document (0x0101)' : id}</option>)}</Select></Field>
      <Switch label="Include descendants of content type ID" checked={includeDescendants} onChange={(_, data) => setIncludeDescendants(data.checked)} disabled={!contentTypeId.trim()} />
    </>}
    {name === 'ListPicker' && <Field label="Initial catalog limit (1–200, default 30)"><Select value={String(catalogLimit)} onChange={(_, data) => setCatalogLimit(Number(data.value))}>{[10, 20, 30, 50, 100, 200].map(value => <option key={value} value={value}>{value}</option>)}</Select></Field>}
    <Field label="Item display field" hint="For files, search also matches the filename; an empty display field uses the filename as its label."><Input value={textField} onChange={(_, data) => setTextField(data.value)} /></Field>
    {name === 'SitePicker' && <>
      <Switch label="Offer current site collection" checked={includeCurrentSite} onChange={(_, data) => setIncludeCurrentSite(data.checked)} />
      <Switch label="Offer associated hub" checked={includeHubSite} onChange={(_, data) => setIncludeHubSite(data.checked)} />
    </>}
    {name === 'PrincipalPicker' && <>
      <Field label="Principal types"><Select value={types} onChange={(_, data) => setTypes(data.value === 'groups' ? 'groups' : data.value === 'all' ? 'all' : 'users')}><option value="users">Users</option><option value="groups">Groups</option><option value="all">Users and groups</option></Select></Field>
      <Field label="User IDs"><Select value={idMode} onChange={(_, data) => setIdMode(data.value === 'both' ? 'both' : data.value === 'sharePoint' ? 'sharePoint' : data.value === 'entra' ? 'entra' : 'available')}><option>available</option><option>entra</option><option>sharePoint</option><option>both</option></Select></Field>
      <Switch label="Allow ensureUser on explicit confirmation" checked={ensureUser} onChange={(_, data) => setEnsureUser(data.checked)} disabled={idMode !== 'sharePoint' && idMode !== 'both'} />
      <Switch label="Load Graph personal suggestions" checked={personal} onChange={(_, data) => setPersonal(data.checked)} />
      <Switch label="Offer current user" checked={currentUser} onChange={(_, data) => setCurrentUser(data.checked)} />
    </>}
    <Button onClick={() => setAvailable(previous => ({ ...previous, [name]: selection?.selectedItems ?? [] }))}>Use confirmed models as consumer options</Button>
    <Button onClick={() => setAvailable(previous => ({ ...previous, [name]: [] }))}>Empty consumer options</Button>
    <Button onClick={() => setSelections(previous => ({ ...previous, [name]: undefined }))}>Clear controlled selection</Button>
    <Button onClick={() => setRevision(value => value + 1)}>Refresh / retry</Button>
  </>}>
    <p>Uses real SharePoint/Graph REST through the provider. Personal suggestions are off by default. The consumer option set is independent of selection: copy confirmed models, clear selection, then reopen. Changing options retains the mounted control. No history is saved. List loads an initial scoped catalog with a configurable limit. Field/View/ContentType load and show all matching metadata; their case insensitive search filters the cached catalog locally. Other typed searches continue through their configured source. SitePicker keeps the current collection and actually associated hub visible during typed search, outside the ordinary suggestion cap. Dropdown menus keep the checked current value in Selezionato during search; click another value to replace it or use the clear icon. Multiple selections can be removed from Selezionati without searching. Type a nonmatching query: fixed/selected choices remain visible alongside the no-results message. Waiting for debounce shows search in progress; matching only an already visible choice shows no other results.</p>
    <p>Open Dropdown to search inside its popup: the selected name remains in the field while the query changes. Query clear only clears search; the single-selection clear removes the value. With Multiple selection enabled, choose two values to see the count, then remove a checked row from Selezionati. Switch to TagPicker to compare removable tags with no duplicate selected rows. Turn off Enable search to use the same choices without a search box. Escape returns to the trigger; Tab leaves without choosing the highlighted row.</p>
    <p>ContentTypePicker site scope reads the owning site collection root, including when Web URL targets a subweb. List scope reads only content types already attached to the selected list; selection never adds a definition, enables content types or synchronizes the hub. Group and hierarchical ID filters apply locally. Change effective web/list/scope to clear the demo content type selection; presentation changes retain it.</p>
    <p>To check lifecycle stability, reopen a completed search and change the suggestion header or reapply equivalent consumer models. Selection labels and accumulated pages should remain stable; slow initial reads show loading. Principal operations that require Graph wait for its initialization; SharePoint-only selections remain usable if Graph initialization fails. Reopen the native Property Pane with saved selections: labels already provided remain visible, while initial selections without display models use an selection-control skeleton until hydration settles. Typing, refreshing and updating mounted fields must not restore that skeleton.</p>
    {subject}
    <Accordion collapsible><AccordionItem value="diagnostics"><AccordionHeader>Selection and diagnostics</AccordionHeader><DemoAccordionPanel><pre>{JSON.stringify({ values, contentTypeScope: name === 'ContentTypePicker' ? contentTypeScope : undefined, roundTrip: JSON.stringify(JSON.parse(JSON.stringify(values))) === JSON.stringify(values), unresolved: selection?.unresolvedOptions ?? [], error }, undefined, 2)}</pre></DemoAccordionPanel></AccordionItem></Accordion>
  </DemoCard>;
};
export default PickersPanel;

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { FluentProvider, webLightTheme, webDarkTheme, Button, makeStyles } from '@fluentui/react-components';
import { createDOMRenderer } from '@griffel/core';
import { RendererProvider } from '@griffel/react';
import { SPFxWebPartProvider } from '../../../packages/spfx-react-toolkit/lib/core/provider-webpart';
import { SitePicker, ListPicker, FieldPicker, ViewPicker, ContentTypePicker, ListItemPicker, PrincipalPicker } from '@apvee/spfx-react-toolkit/controls/pickers';
import type { PickerOption, PickerSelectionData, SitePickerOption, ListPickerOption, FieldPickerOption, ViewPickerOption, ContentTypePickerOption, ListItemPickerOption, PrincipalPickerOption } from '@apvee/spfx-react-toolkit/controls/pickers';
import { datasets, contentTypes, listContentTypes, names, type Name } from './data';
import { host, stats, state } from './transport';
import { createNativeFixture } from './native';
import { TagMediaMatrix } from './tag-media';

const confirmation: { settle?: (message?: string) => void; pending: boolean } = { pending: false };
const useStyles = makeStyles({ gallery: { display: 'flex', flexWrap: 'wrap', gap: '16px' }, wide: { width: '460px', padding: '16px', border: '1px solid #ccc' }, narrow: { width: '280px', padding: '16px', border: '1px solid #ccc' } });
function Case({ name, presentation, multi, width }: { name: Name; presentation: 'combobox' | 'tagPicker'; multi: boolean; width: 'wide' | 'narrow' }) {
  const styles = useStyles(), id = `${name}-${presentation}-${multi ? 'multi' : 'single'}-${width}`;
  const [options, setOptions] = React.useState(datasets[name]), [selected, setSelected] = React.useState<PickerSelectionData<PickerOption>>();
  const [personal, setPersonal] = React.useState(false), [allTypes, setAllTypes] = React.useState(false), [refreshKey, setRefreshKey] = React.useState(0);
  const [idMode, setIdMode] = React.useState<'available' | 'entra' | 'sharePoint' | 'both'>('available');
  const [contentTypeScope, setContentTypeScope] = React.useState<'site' | 'list'>('site');
  const [cosmetic, setCosmetic] = React.useState(false);
  const [contextSites, setContextSites] = React.useState(false);
  const [searchable, setSearchable] = React.useState(true), [deferred, setDeferred] = React.useState(false), [blur, setBlur] = React.useState(false), [limited, setLimited] = React.useState(false);
  const [paged, setPaged] = React.useState(false);
  const [events, setEvents] = React.useState(0), [validations, setValidations] = React.useState(0);
  const [errors, setErrors] = React.useState<string[]>([]);
  const currentDataset = name === 'ContentTypePicker' && contentTypeScope === 'list' ? listContentTypes : datasets[name];
  const validateSelection = React.useCallback(() => {
    setValidations(value => value + 1);
    if (!deferred) return undefined;
    confirmation.pending = true;
    return new Promise<string | undefined>(resolve => {
      confirmation.settle = message => { confirmation.pending = false; confirmation.settle = undefined; resolve(message); };
    });
  }, [deferred]);
  const common = { label: id, id: name === 'SitePicker' && width === 'narrow' ? `${id}-selection` : undefined, hint: cosmetic ? 'Cosmetic hint' : undefined, presentation, multiselect: multi, searchable, suggestionsLimit: paged ? 2 : 50, size: width === 'narrow' ? 'large' as const : 'medium' as const, selectionLimit: limited ? 1 : undefined, validationTrigger: blur ? 'blur' as const : 'change' as const,
    validateSelection, debounceMs: 0, refreshKey, selectedOptions: selected?.selectedOptions ?? [],
    onOptionSelect: (_event: unknown, data: PickerSelectionData<PickerOption>) => { setSelected(data); setEvents(value => value + 1); }, onError: (error: Error) => setErrors(previous => [...previous, error.message]) };
  let control: React.ReactNode;
  switch (name) {
    case 'SitePicker': control = <SitePicker {...common} mode="hub" includeCurrentSite={contextSites} includeHubSite={contextSites} initialOptions={selected?.selectedItems as readonly SitePickerOption[]} availableOptions={options as readonly SitePickerOption[]} />; break;
    case 'ListPicker': control = <ListPicker {...common} initialOptions={selected?.selectedItems as readonly ListPickerOption[]} availableOptions={options as readonly ListPickerOption[]} />; break;
    case 'FieldPicker': control = <FieldPicker {...common} list="Projects" initialOptions={selected?.selectedItems as readonly FieldPickerOption[]} availableOptions={options as readonly FieldPickerOption[]} />; break;
    case 'ContentTypePicker': {
      const contentTypeProps = { ...common, initialOptions: selected?.selectedItems as readonly ContentTypePickerOption[], availableOptions: options as readonly ContentTypePickerOption[] };
      control = contentTypeScope === 'site' ? <ContentTypePicker {...contentTypeProps} scope="site" /> : <ContentTypePicker {...contentTypeProps} scope="list" list="Projects" />; break;
    }
    case 'ViewPicker': control = <ViewPicker {...common} list="Projects" initialOptions={selected?.selectedItems as readonly ViewPickerOption[]} availableOptions={options as readonly ViewPickerOption[]} />; break;
    case 'ListItemPicker': control = <ListItemPicker {...common} list="Projects" textField="Title" initialOptions={selected?.selectedItems as readonly ListItemPickerOption[]} availableOptions={options as readonly ListItemPickerOption[]} />; break;
    case 'PrincipalPicker': control = <PrincipalPicker {...common} principalTypes={allTypes ? ['user', 'microsoft365Group'] : ['user']} loadPersonalSuggestions={personal} initialSuggestions={{ includeCurrentUser: personal }} userIdentity={idMode === 'sharePoint' || idMode === 'both' ? { mode: idMode, ensureUser: false } : { mode: idMode }} initialOptions={selected?.selectedItems as readonly PrincipalPickerOption[]} availableOptions={options as readonly PrincipalPickerOption[]} />; break;
  }
  return <section data-picker-case={id} className={width === 'wide' ? styles.wide : styles.narrow}>
    <h2>{id}</h2><Button data-case-before>Before picker</Button>{control}<Button data-case-after>After picker</Button>
    {name === 'ContentTypePicker' && <Button data-case-action="scope" onClick={() => { const next = contentTypeScope === 'site' ? 'list' : 'site'; setContentTypeScope(next); setSelected(undefined); setOptions(next === 'site' ? contentTypes : listContentTypes); }}>Content type scope: {contentTypeScope}</Button>}
    {name === 'SitePicker' && <Button data-case-action="context-sites" onClick={() => setContextSites(value => !value)}>Current and hub suggestions</Button>}
    <Button data-case-action="long" onClick={() => setOptions(currentDataset.map(option => ({ ...option, ...(option.value.kind === 'user' ? { email: `${"indirizzo.molto.lungo.".repeat(30)}@tenant.test` } : {}), text: `Nome molto lungo ${"descrizione senza interruzioni ".repeat(20)}${option.text}` })))}>Long labels</Button>
    <Button data-case-action="cosmetic" onClick={() => setCosmetic(value => !value)}>Cosmetic update</Button>
    <Button data-case-action="equivalent" onClick={() => setOptions(previous => previous.map(option => ({ ...option, value: { ...option.value } })))}>Equivalent options</Button>
    <Button data-case-action="empty" onClick={() => setOptions([])}>Empty options</Button>
    <Button data-case-action="update" onClick={() => setOptions(currentDataset.map(o => ({ ...o, text: `Renamed ${o.text}` })))}>Update options</Button>
    <Button data-case-action="clear" onClick={() => setSelected(undefined)}>External clear</Button>
    <Button data-case-action="personal" onClick={() => { setPersonal(v => !v); setAllTypes(true); }}>Personal suggestions simulation</Button>
    <Button data-case-action="refresh" onClick={() => setRefreshKey(v => v + 1)}>Refresh</Button>
    <Button data-case-action="both" onClick={() => setIdMode('both')}>Both IDs</Button>
    <Button data-case-action="pages" onClick={() => setPaged(value => !value)}>Toggle small query pages</Button>
    <Button data-case-action="searchable" onClick={() => setSearchable(value => !value)}>Toggle search</Button>
    <Button data-case-action="defer" onClick={() => setDeferred(value => !value)}>Toggle deferred validation</Button>
    <Button data-case-action="blur" onClick={() => setBlur(value => !value)}>Toggle blur validation</Button>
    <Button data-case-action="limit" onClick={() => setLimited(value => !value)}>Toggle one-selection limit</Button>
    <output data-case-events>{events}</output><output data-case-validations>{validations}</output>
    <output data-case-selection>{JSON.stringify(selected?.selectedValues ?? [])}</output>
    <output data-case-errors>{JSON.stringify(errors)}</output>
  </section>;
}
function Gallery() {
  const styles = useStyles();
  const [dark, setDark] = React.useState(false), [rtl, setRtl] = React.useState(false);
  return <FluentProvider theme={dark ? webDarkTheme : webLightTheme} dir={rtl ? 'rtl' : 'ltr'}><h1>Compiled SharePoint pickers</h1><p>SDK service identifiers and all REST responses are simulated. Actual compiled controls, provider, hooks and Fluent components run; this is not tenant validation.</p>
    <Button data-global-action="fail" onClick={() => { state.fail = !state.fail; }}>Toggle simulated SharePoint failure</Button>
    <Button data-global-action="personal-fail" onClick={() => { state.personalFail = !state.personalFail; }}>Toggle simulated personal failure</Button>
    <Button data-global-action="outside">Outside focus target</Button>
    <Button data-global-action="theme" onClick={() => setDark(value => !value)}>Toggle gallery theme</Button>
    <Button data-global-action="rtl" onClick={() => setRtl(value => !value)}>Toggle gallery direction</Button>
    <main className={styles.gallery}>{names.flatMap(name => (['combobox', 'tagPicker'] as const).flatMap(presentation => [false, true].flatMap(multi => (['wide', 'narrow'] as const).map(width => <Case key={`${name}-${presentation}-${multi}-${width}`} name={name} presentation={presentation} multi={multi} width={width} />))))}</main>
  </FluentProvider>;
}
// Native factories own their React roots, just as SPFx calls onRender/onDispose.
const nativeSection = document.createElement('section');
nativeSection.dataset.nativePickerFixture = '';
document.body.appendChild(nativeSection);
const nativeFixture = createNativeFixture(nativeSection);
// Opt-in real DOM/token coverage: independent native React root beneath a host
// provider, contrasted with a dark ordinary React picker. No production API.
const paneThemeFixture = new URLSearchParams(window.location.search).has('paneTheme');
if (paneThemeFixture) {
  const nativeHost = document.createElement('section');
  document.body.appendChild(nativeHost);
  function NativeThemeFixture() {
    const [dark, setDark] = React.useState(false);
    const [name, setName] = React.useState<Name>('SitePicker');
    const [presentation, setPresentation] = React.useState<'combobox' | 'tagPicker'>('combobox');
    const [counters, setCounters] = React.useState('');
    const target = React.useRef<HTMLDivElement>(null); // Native React DOM ref contract.
    React.useLayoutEffect(() => { target.current?.appendChild(nativeSection); }, []);
    React.useEffect(() => { nativeFixture.mount(name, 'model', presentation); return () => nativeFixture.dispose(); }, [name, presentation]);
    return <FluentProvider theme={dark ? webDarkTheme : webLightTheme}>
      <h2>Native pane theme fixture</h2>
      <label>Native picker<select aria-label="Native picker" value={name} onChange={event => setName(event.target.value as Name)}>{names.map(item => <option key={item}>{item}</option>)}</select></label>
      <label>Native presentation<select aria-label="Native presentation" value={presentation} onChange={event => setPresentation(event.target.value as 'combobox' | 'tagPicker')}><option value="combobox">Dropdown</option><option value="tagPicker">TagPicker</option></select></label>
      <Button onClick={() => setDark(value => !value)}>Toggle native pane theme</Button>
      <Button onClick={() => setCounters(JSON.stringify(stats))}>Read native counters</Button>
      <output aria-label="Native pane theme">{dark ? 'dark' : 'light'}</output>
      <output aria-label="Native counters">{counters}</output>
      <div ref={target} />
    </FluentProvider>;
  }
  ReactDOM.render(<RendererProvider renderer={createDOMRenderer(document)}><NativeThemeFixture /></RendererProvider>, nativeHost);
}

const root = document.getElementById('root');
if (!root) throw new Error('Missing fixture root');
// Observable transport counters are fixture-only evidence, not public picker API.
Object.assign(window, { pickerFixture: { stats, state, confirmation, native: nativeFixture } });
ReactDOM.render(<RendererProvider renderer={createDOMRenderer(document)}><FluentProvider theme={paneThemeFixture ? webDarkTheme : webLightTheme}><SPFxWebPartProvider instance={host}>{new URLSearchParams(window.location.search).has('tagMedia') ? <TagMediaMatrix /> : paneThemeFixture ? <Case name="SitePicker" presentation="combobox" multi={false} width="wide" /> : <Gallery />}</SPFxWebPartProvider></FluentProvider></RendererProvider>, root);

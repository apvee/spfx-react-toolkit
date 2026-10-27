import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Button, Input, FluentProvider, makeStyles, mergeClasses, webLightTheme, webDarkTheme } from '@fluentui/react-components';
import { createDOMRenderer } from '@griffel/core';
import { RendererProvider } from '@griffel/react';
import { SPFxWebPartProvider } from '../../../packages/spfx-react-toolkit/lib/core/provider-webpart';
import { FileFolderPicker } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';
import type { FileFolderPickerValue, FileFolderPickerChangeData } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';
import { PropertyPaneFileFolderPicker } from '@apvee/spfx-react-toolkit/property-pane/file-folder-picker';
import type { IPropertyPaneCustomFieldProps, IPropertyPaneField } from '@microsoft/sp-property-pane';
import { host, values, options, observations } from './transport';

const useStyles = makeStyles({
  page: { padding: '20px', fontFamily: 'system-ui' },
  gallery: { display: 'flex', flexWrap: 'wrap', alignItems: 'flex-start', gap: '20px' },
  card: { boxSizing: 'border-box', padding: '12px', border: '1px solid #aaa', minWidth: 0, maxWidth: '100%' },
  wide: { width: '600px' }, narrow: { width: '280px' },
  reference: { width: '100%' },
  actions: { display: 'flex', flexWrap: 'wrap', gap: '6px', marginTop: '12px' },
  json: { display: 'block', fontFamily: 'monospace', maxHeight: '160px', overflowY: 'auto', fontSize: '11px', marginTop: '8px' },
});
function Case({ size, narrow, rtl, dark }: { size: 'medium' | 'large'; narrow: boolean; rtl: boolean; dark: boolean }) {
  const styles = useStyles();
  const [selected, setSelected] = React.useState<readonly FileFolderPickerValue[]>(values);
  const [changes, setChanges] = React.useState<string[]>([]);
  const [disabled, setDisabled] = React.useState(false), [invalid, setInvalid] = React.useState(false);
  const [errors, setErrors] = React.useState<string[]>([]);
  const id = `${size}-${narrow ? 'narrow' : 'wide'}-${rtl ? 'rtl' : 'ltr'}`;
  function changed(data: FileFolderPickerChangeData): void { setSelected(data.values); setChanges(previous => [...previous, `${data.reason}: ${data.values.length} values, ${data.items.length} models, ${data.unresolvedValues.length} unresolved`]); }
  return <FluentProvider theme={dark ? webDarkTheme : webLightTheme} dir={rtl ? 'rtl' : 'ltr'} className={mergeClasses(styles.card, narrow ? styles.narrow : styles.wide)}>
    <section aria-label={id} data-fixture-case={id}>
      <h2>{id}</h2>
      <Input className={styles.reference} size={size} readOnly aria-label={`Fluent reference ${id}`} disabled={disabled} aria-invalid={invalid} />
      <FileFolderPicker preset="focused" disabled={disabled} validationState={invalid ? 'error' : undefined} validationMessage={invalid ? 'Fixture validation' : undefined} size={size} label={`Documents ${id}`} browseLabel={`Browse ${id}`} value={selected} initialOptions={options}
        multiselect itemTypes="all" onChange={changed} onError={error => setErrors(previous => [...previous, error.message])} />
      <div className={styles.actions}><Button onClick={() => setSelected(values)}>Restore three references</Button><Button onClick={() => setDisabled(value => !value)}>Toggle disabled</Button><Button onClick={() => setInvalid(value => !value)}>Toggle invalid</Button><Button onClick={() => setSelected([])}>External empty</Button></div>
      <p role="status">Confirmed values: {selected.length}. Change events: {changes.length}.</p>
      <output className={styles.json} aria-label={`Selection JSON ${id}`}>{JSON.stringify(selected)}</output>
      <output className={styles.json} aria-label={`Changes ${id}`}>{JSON.stringify(changes)}</output>
      <output className={styles.json} aria-label={`Errors ${id}`}>{JSON.stringify(errors)}</output>
    </section>
  </FluentProvider>;
}
function NativeCase() {
  const [nativeDark, setNativeDark] = React.useState(false);
  const styles = useStyles();
  const target = React.useRef<HTMLDivElement>(null); // React DOM ref interoperability.
  const field = React.useRef<IPropertyPaneField<IPropertyPaneCustomFieldProps>>();
  const saved = React.useRef<readonly FileFolderPickerValue[]>(values);
  const [json, setJson] = React.useState(JSON.stringify(values));
  const [open, setOpen] = React.useState(false);
  const [events, setEvents] = React.useState(0);
  const [reopenEvidence, setReopenEvidence] = React.useState('Not reopened');
  const previousClose = React.useRef<string>();
  function render(): void {
    if (target.current === null) return;
    field.current = PropertyPaneFileFolderPicker('documents', {
      instance: host, key: 'native-file-folder-fixture', preset: 'focused', value: saved.current,
      label: 'Native documents', browseLabel: 'Browse native documents', initialOptions: options,
      multiselect: true, itemTypes: 'all', size: 'large',
    });
    field.current.properties.onRender(target.current, undefined, (_property, next) => {
      saved.current = next as FileFolderPickerValue[]; // SDK callback is untyped; actual native factory validates JSON before emission.
      setJson(JSON.stringify(saved.current)); setEvents(count => count + 1);
    });
    setOpen(true);
  }
  function close(): void {
    if (target.current !== null) field.current?.properties.onDispose?.(target.current);
    field.current = undefined; previousClose.current = JSON.stringify(saved.current); setOpen(false);
  }
  function reopen(): void {
    const current = JSON.stringify(saved.current);
    setReopenEvidence(previousClose.current === undefined ? 'First opening' : previousClose.current === current ? 'PASS: complete ordered JSON equals close snapshot' : 'FAIL: references changed');
    render();
  }
  React.useEffect(() => () => { if (target.current !== null) field.current?.properties.onDispose?.(target.current); }, []);
  return <FluentProvider theme={nativeDark ? webDarkTheme : webLightTheme}><section className={mergeClasses(styles.card, styles.narrow)} aria-label="Native property pane simulation">
    <h2>Native property pane lifecycle</h2><Button onClick={() => setNativeDark(value => !value)}>Toggle native pane theme</Button><output aria-label="Native pane theme">{nativeDark ? 'dark' : 'light'}</output><p>Actual native factory. SPFx pane open/close and property callback are simulated; accepted JSON is retained in this fixture only.</p>
    <div className={styles.actions}><Button disabled={open} onClick={reopen}>Open native pane</Button><Button disabled={!open} onClick={close}>Close native pane</Button><Button disabled={!open} onClick={render}>Equivalent native onRender</Button></div>
    <div ref={target} />
    <p role="status">Native pane: {open ? 'open' : 'closed'}. Native events: {events}. {reopenEvidence}</p>
    <output className={styles.json} aria-label="Native complete saved JSON">{json}</output>
  </section></FluentProvider>;
}
function Fixture() {
  const styles = useStyles();
  const [dark, setDark] = React.useState(false);
  const [stats, setStats] = React.useState('');
  return <FluentProvider theme={dark ? webDarkTheme : webLightTheme} className={styles.page}>
    <h1>Compiled FileFolderPicker browser fixture</h1>
    <p>Production compiled FileFolderPicker, provider, hooks, scope/result validation, Fluent tags, Stage dialog and native property-pane factory. SPFx SDK, metadata responses and hosted Microsoft session/navigation are simulated. No real tokens, credentials, tenant calls or property persistence.</p>
    <div className={styles.actions}><Button onClick={() => setDark(value => !value)}>Toggle host theme</Button><Button onClick={() => setStats(JSON.stringify(observations))}>Read simulation counters</Button></div><output aria-label="Simulation counters">{stats}</output>
    <main className={styles.gallery}>
      <SPFxWebPartProvider instance={host}>
        <Case dark={dark} size="medium" narrow={false} rtl={false} /><Case dark={dark} size="large" narrow={false} rtl={false} />
        <Case dark={dark} size="medium" narrow rtl /><Case dark={dark} size="large" narrow rtl />
      </SPFxWebPartProvider>
      <NativeCase />
    </main>
  </FluentProvider>;
}
const root = document.getElementById('root');
if (root === null) throw new Error('Missing fixture root');
ReactDOM.render(<RendererProvider renderer={createDOMRenderer(document)}><Fixture /></RendererProvider>, root);

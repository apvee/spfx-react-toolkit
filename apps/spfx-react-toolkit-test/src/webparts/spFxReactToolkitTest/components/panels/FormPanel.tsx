import * as React from 'react';
import { DemoCard, useDemoStyles } from '../shared';
import { Button, Checkbox, Field, Select, ColorArea, ColorSlider, ColorSwatch, FluentProvider, makeStyles } from '@fluentui/react-components';
import { Form, createForm, FormDateTimeValue, FormTimePickerParseResult, FormHsvColor, FormMethods, flex, gap, useSx, width } from '@apvee/spfx-react-toolkit';

interface DemoRow { code: string; note: string; }
interface DemoRecord {
  reference: string; notes: string; query: string; approved: boolean; enabled: boolean;
  // Fluent SpinButton's empty committed domain explicitly uses null.
  // eslint-disable-next-line @rushstack/no-new-null
  quantity: number | null;
  progress: number; country: string; priority: string; category: string; categories: string[];
  checks: string[]; combo: string; combos: string[]; tags: string[]; rating: number;
  accent: FormHsvColor; swatch: string;
  // Compat DatePicker and TimePicker explicitly use null for an empty selection.
  // eslint-disable-next-line @rushstack/no-new-null
  date: Date | null;
  // Compat TimePicker's complete live value includes null for empty selection.
  // eslint-disable-next-line @rushstack/no-new-null
  time: Date | null;
  appointment: FormDateTimeValue;
  // HTMLInputElement.files explicitly returns null when there is no native selection.
  // eslint-disable-next-line @rushstack/no-new-null
  files: FileList | null;
  custom: string; customer: { name: string }; address: { city: string }; rows: DemoRow[]; lines: DemoRow[];
}
const Order = createForm<DemoRecord>();
const Numeric = createForm<{ quantity: string }, { quantity: number }>({ schema: {
  '~standard': {
    version: 1, vendor: 'Form sample',
    validate: (input: unknown) => {
      if (typeof input !== 'object' || !input || !('quantity' in input) || typeof input.quantity !== 'string') {
        return { issues: [{ message: 'Expected quantity text', path: ['quantity'] }] };
      }
      if (input.quantity === 'throw') throw new Error('Simulated validator failure');
      const quantity = Number(input.quantity);
      return input.quantity.trim() && Number.isFinite(quantity)
        ? { value: { quantity } }
        : { issues: [{ message: 'Enter a finite number', path: ['quantity'] }] };
    },
  },
} });
const options = [{ value: '', label: 'Choose an identity' }, { value: 'it', label: 'Italy' }, { value: 'pt', label: 'Portugal' }];
const choices = options.slice(1);
const formatDate = (date?: Date): string => date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}` : '';
// Compat DatePicker's parser explicitly requires null for invalid/empty selection.
// eslint-disable-next-line @rushstack/no-new-null
function parseDate(text: string): Date | null {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (match) {
    const year = Number(match[1]); const month = Number(match[2]) - 1; const day = Number(match[3]);
    const date = new Date(year, month, day);
    if (date.getFullYear() === year && date.getMonth() === month && date.getDate() === day) return date;
  }
  // Compat DatePicker rejects a draft by returning its documented nullable sentinel.
  // eslint-disable-next-line @rushstack/no-new-null
  return null;
}
const formatTime = (date: Date): string => `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
function parseTime(text: string | undefined): FormTimePickerParseResult {
  const match = /^(\d{2}):(\d{2})$/.exec(text ?? '');
  if (match && Number(match[1]) < 24 && Number(match[2]) < 60) return { date: new Date(1970, 0, 1, Number(match[1]), Number(match[2])) };
  // Compat TimePicker's parse result explicitly uses null for invalid/empty time.
  // eslint-disable-next-line @rushstack/no-new-null
  return { date: null, error: text ? 'invalid-input' : undefined };
}
const createRow = (): DemoRow => ({ code: 'New row', note: '' });
const createRecord = (reference = 'Quarterly report'): DemoRecord => ({
  reference, notes: 'Team activity summary and next steps.', query: '', approved: false, enabled: true, quantity: 2, progress: 25,
  country: 'it', priority: 'pt', category: 'it', categories: ['pt'], checks: ['it'], combo: 'it', combos: ['pt'], tags: ['it'],
  rating: 3, accent: { h: 210, s: 0.7, v: 0.8, a: 1 }, swatch: 'blue', date: new Date(2026, 9, 8),
  time: new Date(1970, 0, 1, 9, 30), appointment: { date: new Date(2026, 9, 8), time: new Date(1970, 0, 1, 10) },
  // Native file inputs cannot restore a selection from defaults; null is their empty domain.
  // eslint-disable-next-line @rushstack/no-new-null
  files: null,
  custom: 'Custom value', customer: { name: 'Ada' }, address: { city: 'Rome' },
  rows: [{ code: 'Headless A', note: 'First' }, { code: 'Headless B', note: 'Second' }],
  lines: [{ code: 'Section A', note: 'First' }, { code: 'Section B', note: 'Second' }],
});

const adapterFields = {"Input":"custom","Textarea":"notes","SearchBox":"query","Checkbox":"approved","Switch":"enabled","SpinButton":"quantity","Slider":"progress","Select":"country","RadioGroup":"priority","Dropdown":"category","DropdownMulti":"categories","CheckboxGroup":"checks","Combobox":"combo","ComboboxMulti":"combos","TagPicker":"tags","Rating":"rating","ColorPicker":"accent","SwatchPicker":"swatch","DatePicker":"date","TimePicker":"time","DateTime":"appointment","FileInput":"files"} as const;
type Adapter = keyof typeof adapterFields;
type Family = 'rules' | 'object' | 'array' | 'schema' | 'dynamic';
const useStyles = makeStyles({ frame: { maxInlineSize: '100%', minInlineSize: 0 }, output: { whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' } });
/** Native values are inspected, not transformed into a hidden field model. */
function inspectValue(value: unknown): unknown {
  if (value instanceof Date) return { year: value.getFullYear(), month: value.getMonth() + 1, day: value.getDate(), hour: value.getHours(), minute: value.getMinutes() };
  if (typeof FileList !== 'undefined' && value instanceof FileList) return { count: value.length, names: Array.from(value).map(file => file.name) };
  if (Array.isArray(value)) return value.map(inspectValue);
  if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, inspectValue(item)]));
  return value;
}
function errorMessages(value: unknown, prefix = ''): string[] {
  if (!value || typeof value !== 'object') return [];
  return Object.entries(value).reduce<string[]>((messages, [key, item]) => {
    if (key === 'ref') return messages;
    if (key === 'message' && typeof item === 'string') return [...messages, `${prefix}: ${item}`];
    return [...messages, ...errorMessages(item, prefix ? `${prefix}.${key}` : key)];
  }, []);
}

const SchemaExample: React.FC<{ methods: FormMethods<{ quantity: string }, unknown, { quantity: number }>; onSubmit: (value: { quantity: number }) => Promise<void>; disabled: boolean }> = ({ methods, onSubmit, disabled }) => <Numeric.Root disabled={disabled} methods={methods} onSubmit={onSubmit} data-form-demo="root">
  <Numeric.Input name="quantity" label="Quantity" /><Numeric.Error /><Numeric.Error name="quantity" />
  <Numeric.Submit data-form-demo="submit">Save quantity</Numeric.Submit><Numeric.Reset data-form-demo="reset">Reset</Numeric.Reset>
</Numeric.Root>;
const FormPanel: React.FC = () => {
  const sx = useSx(); const styles = useStyles(); const demo = useDemoStyles();
  const [family, setFamily] = React.useState<Family>('rules');
  const [adapter, setAdapter] = React.useState<Adapter>('Textarea');
  const [colorfulTagAvatars, setColorfulTagAvatars] = React.useState(false);
  const [size, setSize] = React.useState(640); const [disabled, setDisabled] = React.useState(false);
  const [customStrings, setCustomStrings] = React.useState(false);
  const formStrings = customStrings ? { unexpectedFailure: 'The report could not be saved. Please retry.', fileInput: { browseLabel: 'Attach report', clearLabel: 'Remove attachment' }, arraySection: { add: 'Add report line' } } : undefined;
  const [rtl, setRtl] = React.useState(false); const [fail, setFail] = React.useState(false);
  const [delay, setDelay] = React.useState(false); const pendingSave = React.useRef<(() => void) | undefined>(undefined);
  const [count, setCount] = React.useState(0); const [status, setStatus] = React.useState('Ready to edit.');
  const [saved, setSaved] = React.useState<unknown>(); const [slotClicks, setSlotClicks] = React.useState(0);
  const [slotObservation, setSlotObservation] = React.useState('Not inspected.');
  const slotLabelRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
  const slotHintRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const formRef = React.useMemo(() => React.createRef<HTMLFormElement>(), []);
  const methods = Order.useForm({ defaultValues: createRecord('Quarterly report'), mode: 'onBlur', strings: formStrings });
  const schemaMethods = Numeric.useForm({ defaultValues: { quantity: '12' }, strings: formStrings });
  const dynamicMethods = Form.useForm({ defaultValues: { dynamic: 'Quarterly report' }, strings: formStrings });
  const values = methods.watch(); const { dirtyFields, touchedFields, errors } = methods.formState;
  const recordDirty = methods.formState.isDirty;
  const schemaDirty = schemaMethods.formState.isDirty;
  const dynamicDirty = dynamicMethods.formState.isDirty;
  const isDirty = family === 'schema' ? schemaDirty : family === 'dynamic' ? dynamicDirty : recordDirty;
  const loadCurrentModel = (): void => {
    if (family === 'schema') { schemaMethods.reset({ quantity: '12' }); setStatus('Initial quantity restored.'); }
    else if (family === 'dynamic') { dynamicMethods.reset({ dynamic: 'Annual report' }); setStatus('New dynamic report loaded.'); }
    else { methods.reset(createRecord('Annual report')); setStatus('New report loaded.'); }
  };
  const recordSubmitting = methods.formState.isSubmitting;
  const schemaSubmitting = schemaMethods.formState.isSubmitting;
  const dynamicSubmitting = dynamicMethods.formState.isSubmitting;
  const isSubmitting = family === 'schema' ? schemaSubmitting : family === 'dynamic' ? dynamicSubmitting : recordSubmitting;
  const changeAdapter = (value: string): void => {
    const next = Object.keys(adapterFields).find(key => key === value) as Adapter | undefined;
    if (!next || next === adapter) return;
    // Each adapter retains its native value domain. Only the newly selected variable field resets;
    // the common reference and other caller-owned draft fields remain unchanged.
    const previousValue = methods.getValues(adapterFields[adapter]);
    const textAdapters: readonly Adapter[] = ['Input', 'Textarea', 'SearchBox'];
    if (textAdapters.includes(next) && textAdapters.includes(adapter) && typeof previousValue === 'string') {
      methods.setValue(adapterFields[next], previousValue);
      setStatus('Compatible adapter: text preserved.');
    } else if (adapterFields[next] !== adapterFields[adapter]) {
      methods.resetField(adapterFields[next]);
      methods.setValue(adapterFields[next], createRecord()[adapterFields[next]]);
      setStatus('Variable field reset for the new domain; title preserved.');
    } else setStatus('Compatible adapter: value preserved.');
    setAdapter(next);
  };
  const save = async (value: unknown): Promise<void> => {
    setCount(current => current + 1); setStatus('Saving');
    if (delay) await new Promise<void>(resolve => { pendingSave.current = resolve; });
    pendingSave.current = undefined;
    if (fail) { setStatus('Save failed: disable the simulation and try again.'); throw new Error('Simulated save failure'); }
    setSaved(value); setStatus('Saved');
  };
  const variableField = (): React.ReactNode => {
    switch (adapter) {
      case 'Textarea': return <div data-form-control="Textarea"><Order.Textarea name="notes" label={{ ref: slotLabelRef,  'data-form-demo': 'slot-label', onClick: () => setSlotClicks(value => value + 1), children: 'Report description' }} hint={{ ref: slotHintRef, 'data-form-demo': 'slot-hint', children: 'Summarize team results, decisions and next steps.' }} resize="vertical" /></div>;
      case 'SearchBox': return <div data-form-control="SearchBox"><Order.SearchBox name="query" label="Search query" /></div>;
      case 'Checkbox': return <div data-form-control="Checkbox"><Order.Checkbox name="approved" label="Approved" /></div>;
      case 'SpinButton': return <div data-form-control="SpinButton"><Order.SpinButton name="quantity" label="Quantity"  min={0} max={20} /></div>;
      case 'Slider': return <div data-form-control="Slider"><Order.Slider name="progress" label="Progress" min={0} max={100} step={5} /></div>;
      case 'Select': return <div data-form-control="Select"><Order.Select name="country" label="Country" options={options} /></div>;
      case 'RadioGroup': return <div data-form-control="RadioGroup"><Order.RadioGroup name="priority" label="Priority identity" options={choices} /></div>;
      case 'Dropdown': return <div data-form-control="Dropdown"><Order.Dropdown name="category" label="Category" options={choices} /></div>;
      case 'DropdownMulti': return <div data-form-control="DropdownMulti"><Order.DropdownMulti name="categories" label="Categories" options={choices} /></div>;
      case 'CheckboxGroup': return <div data-form-control="CheckboxGroup"><Order.CheckboxGroup name="checks" label="Checkbox identities" options={choices} /></div>;
      case 'Combobox': return <div data-form-control="Combobox"><Order.Combobox name="combo" label="Editable identity" options={choices} /></div>;
      case 'ComboboxMulti': return <div data-form-control="ComboboxMulti"><Order.ComboboxMulti name="combos" label="Editable identities" options={choices} /></div>;
      case 'TagPicker': return <div data-form-control="TagPicker"><Order.TagPicker name="tags" label="Identity tags" options={choices} colorfulTagAvatars={colorfulTagAvatars} strings={{ noOptions: { 'data-form-demo': 'slot-no-options', children: 'No available identities' } }} /></div>;
      case 'Rating': return <div data-form-control="Rating"><Order.Rating name="rating" label="Rating" max={5} /></div>;
      case 'ColorPicker': return <div data-form-control="ColorPicker"><Order.ColorPicker name="accent" label="Accent HSV"><ColorArea aria-label="Saturation and value" inputX={{ 'aria-label': 'Accent saturation' }} inputY={{ 'aria-label': 'Accent value' }} /><ColorSlider channel="hue" aria-label="Hue" /></Order.ColorPicker></div>;
      case 'SwatchPicker': return <div data-form-control="SwatchPicker"><Order.SwatchPicker name="swatch" label="Swatch identity"><ColorSwatch value="blue" color="#0078d4" aria-label="Blue" /><ColorSwatch value="red" color="#d13438" aria-label="Red" /></Order.SwatchPicker></div>;
      case 'DatePicker': return <div data-form-control="DatePicker"><Order.DatePicker name="date" label="Calendar date" allowTextInput formatDate={formatDate} parseDateFromString={parseDate} /></div>;
      case 'TimePicker': return <div data-form-control="TimePicker"><Order.TimePicker name="time" label="Wall-clock time" freeform hourCycle="h23" formatDateToTimeString={formatTime} parseTimeStringToDate={parseTime} dateAnchor={new Date(1970, 0, 1)} /></div>;
      case 'DateTime': return <div data-form-control="DateTime"><Order.DateTime name="appointment" label="Separate date and time" dateLabel={{ children: 'Appointment date' }} timeLabel={{ children: (Component, props) => <Component {...props}>Appointment time</Component> }} datePickerProps={{ allowTextInput: true, formatDate, parseDateFromString: parseDate }} timePickerProps={{ freeform: true, hourCycle: 'h23', formatDateToTimeString: formatTime, parseTimeStringToDate: parseTime, dateAnchor: new Date(1970, 0, 1) }} /></div>;
      case 'FileInput': return <div data-form-control="FileInput"><Order.FileInput name="files" label="Native attachments" multiple browseLabel="Choose attachments" clearLabel="Clear native attachments" formatFiles={files => `${files.length} actual file(s): ${Array.from(files).map(file => file.name).join(', ')}`} /></div>;
      case 'Input': return <div data-form-control="Input"><Order.Input name="custom" label="Detail" rules={{ required: 'Enter a detail' }} /></div>;
      case 'Switch': return <div data-form-control="Switch"><Order.Switch name="enabled" label="Enabled" /></div>;
    }
  };
  return <section data-form-demo="panel" className={demo.panel}><FluentProvider dir={rtl ? 'rtl' : undefined}>
    <DemoCard title="Form" iconName="Edit" description="Edit the report and try one field at a time in the same form."
      controls={<>
        <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'strings' } }} label="Custom labels and failure text" checked={customStrings} onChange={(_, data) => setCustomStrings(data.checked === true)} />
        <Field label="Example"><Select data-demo-variant="Family" data-form-demo="mode" value={family} disabled={isSubmitting} onChange={(_, data) => {
          if (data.value === 'rules' || data.value === 'object' || data.value === 'array' || data.value === 'schema' || data.value === 'dynamic') { if (data.value === 'schema') schemaMethods.reset({ quantity: '12' }); if (data.value === 'dynamic') dynamicMethods.reset({ dynamic: 'Quarterly report' }); setFamily(data.value); setStatus(data.value === 'schema' || data.value === 'dynamic' ? 'New model for the selected family.' : 'Report model preserved.'); }
        }}><option value="rules">Variable field</option><option value="object">Object</option><option value="array">Ordered rows</option><option value="schema">Numeric schema</option><option value="dynamic">Dynamic form</option></Select></Field>
        {family === 'rules' && <Field label="Field"><Select data-demo-variant="FormAdapter" data-form-demo="adapter" value={adapter} disabled={isSubmitting} onChange={(_, data) => changeAdapter(data.value)}>{Object.keys(adapterFields).map(value => <option key={value} value={value}>{value}</option>)}</Select></Field>}
        <Field label="Width"><Select data-form-demo="resize" value={size} onChange={(_, data) => setSize(Number(data.value))}>{[320, 500, 640, 900].map(value => <option key={value} value={value}>{value} px</option>)}</Select></Field>
        <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'disabled' } }} label="Disabled" checked={disabled} onChange={(_, data) => setDisabled(data.checked === true)} />
        {family === 'rules' && adapter === 'TagPicker' && <Checkbox label="Colorful tag avatars" checked={colorfulTagAvatars} onChange={(_, data) => setColorfulTagAvatars(data.checked === true)} />}
      </>}
      commands={<><Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'save-fail' } }} label="Simulate error" checked={fail} onChange={(_, data) => setFail(data.checked === true)} /><Button data-form-demo="new-record" disabled={isSubmitting} onClick={loadCurrentModel}>{family === 'schema' ? 'Reset quantity' : 'Load new report'}</Button>{isSubmitting && <Button data-form-demo="complete-save" onClick={() => pendingSave.current?.()}>Complete save</Button>}</>}
      status={<output data-form-demo="status" role="status">{status} · Saves: {count} · Dirty: {String(isDirty)}</output>}
      diagnostics={<>
        <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'direction' } }} label="Override RTL" checked={rtl} onChange={(_, data) => setRtl(data.checked === true)} />
        <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'save-delay' } }} label="Complete save manually" checked={delay} onChange={(_, data) => setDelay(data.checked === true)} />
        <Button data-form-demo="inspect-slots" onClick={() => setSlotObservation(`label:${slotLabelRef.current?.tagName} hint:${slotHintRef.current?.tagName} label-clicks:${slotClicks}`)}>Inspect slots</Button><output data-form-demo="slot-observation">{slotObservation}</output>
        <pre data-form-demo="values" className={styles.output}>{JSON.stringify(inspectValue(values), undefined, 2)}</pre>
        <pre data-form-demo="dirty">{JSON.stringify(dirtyFields)}</pre><pre data-form-demo="touched">{JSON.stringify(touchedFields)}</pre><pre data-form-demo="errors">{JSON.stringify(errorMessages(errors))}</pre>
        <pre data-form-demo="saved" className={styles.output}>{JSON.stringify(inspectValue(saved), undefined, 2)}</pre>
        <pre data-form-demo="schema-values">{JSON.stringify(schemaMethods.watch())}</pre><pre data-form-demo="schema-errors">{JSON.stringify(errorMessages(schemaMethods.formState.errors))}</pre>
      </>}>
      <div data-form-demo="frame" className={sx(styles.frame, flex.column, gap.medium, width.px(size))}>
        {family === 'schema' ? <SchemaExample methods={schemaMethods} onSubmit={save} disabled={disabled} /> : family === 'dynamic' ? <Form.Root methods={dynamicMethods} onSubmit={save} disabled={disabled} data-form-demo="root"><Form.Input name="dynamic" label="Title" /><Form.Submit data-form-demo="submit">Save</Form.Submit><Form.Reset data-form-demo="reset">Reset</Form.Reset></Form.Root> :
        <Order.Root ref={formRef} methods={methods} disabled={disabled} onSubmit={save} onInvalid={() => setStatus('Check the required fields.')} data-form-demo="root" className={sx(flex.column, gap.medium)}>
          <Order.Input name="reference" label="Title" required rules={{ required: 'Title is required' }} data-form-demo="reference" />
          {family === 'rules' ? variableField() : family === 'object' ? <><Order.Object name="customer">{({ Fields }) => <Fields.Input name="name" label="Owner" />}</Order.Object><Order.ObjectSection name="address" title="Location" description="Team address.">{({ Fields }) => <Fields.Input name="city" label="City" />}</Order.ObjectSection></> :
          <Order.ArraySection data-form-demo="array-section" name="lines" title="Activities" createItem={createRow} minItems={1} maxItems={3} reorderable itemLabel={item => `Activities ${item.index + 1}`} renderItem={item => <item.Fields.Input name="code" label={`Activity name ${item.index + 1}`} />} />}
          <Order.Error /><div className={sx(flex.row, gap.small)}><Order.Submit data-form-demo="submit" appearance="primary">Save report</Order.Submit><Order.Reset data-form-demo="reset">Reset</Order.Reset></div>
        </Order.Root>}
      </div>
    </DemoCard>
  </FluentProvider></section>;
};
export default FormPanel;

import * as React from 'react';
import { DemoCard, DemoAccordionPanel, useDemoStyles } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared';
import { Accordion, AccordionItem, AccordionHeader, Button, Checkbox, Field, Input, Select, ColorArea, ColorSlider, ColorSwatch, FluentProvider, makeStyles, tokens } from '@fluentui/react-components';
import { Form, createForm, FormDateTimeValue, FormTimePickerParseResult, FormHsvColor, FormMethods, SettingsGroup, flex, gap, useSx, width } from '@apvee/spfx-react-toolkit';

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
const createRecord = (reference = 'Order A'): DemoRecord => ({
  reference, notes: 'Editable notes', query: '', approved: false, enabled: true, quantity: 2, progress: 25,
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
const useStyles = makeStyles({
  frame: { maxInlineSize: '100%', minInlineSize: '0' },
  // Match Fluent Dropdown/Combobox's default minimum before adding another column.
  grid: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(min(100%, 250px), 1fr))' },
  output: { whiteSpace: 'pre-wrap', overflowWrap: 'anywhere', maxInlineSize: '100%' },
  row: { borderBlockEndWidth: '1px', borderBlockEndStyle: 'solid', borderBlockEndColor: tokens.colorNeutralStroke2, paddingBlock: tokens.spacingVerticalM },
});
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
const CustomField: React.FC = () => {
  const field = Order.useField<string>({ name: 'custom' });
  return <Field label="Custom hook input"><Input data-form-demo="custom" input={{ ref: field.ref }} value={field.value} disabled={field.disabled}
    onChange={(_, data) => field.onChange(data.value)} onBlur={field.onBlur} aria-invalid={field.invalid} /></Field>;
};
const SchemaExample: React.FC<{ methods: FormMethods<{ quantity: string }, unknown, { quantity: number }>; onSubmit: (value: { quantity: number }) => void }> = ({ methods, onSubmit }) => <section data-form-demo="schema"><h3>Schema transform and thrown validator</h3>
  <p>Live quantity is a string; successful submitted quantity is a number. The value “throw” simulates an unexpected validator exception.</p>
  <Numeric.Root methods={methods} onSubmit={onSubmit}>
    <Numeric.Input name="quantity" label="Transform quantity" />
    <Numeric.Error /><Numeric.Error name="quantity" />
    <Numeric.Submit>Submit transformed</Numeric.Submit><Numeric.Reset>Reset transform</Numeric.Reset>
  </Numeric.Root>
</section>;
const DynamicExample: React.FC = () => <section data-form-demo="dynamic"><h3>Dynamic Form family</h3>
  <Form.Root defaultValues={{ dynamic: 'Dynamic name' }} onSubmit={() => undefined}>
    <Form.Input name="dynamic" label="Dynamic input" /><Form.Submit>Submit dynamic</Form.Submit><Form.Reset>Reset dynamic</Form.Reset>
  </Form.Root>
</section>;
type ValidationMode = 'onBlur' | 'onChange' | 'onSubmit';
const RecordExample: React.FC<{ disabled: boolean; validationMode: ValidationMode; controls: React.ReactNode; mode: string; size: number }> = ({ disabled, validationMode, controls, mode, size }) => {
  const sx = useSx();
  const styles = useStyles();
  const schemaMethods = Numeric.useForm({ defaultValues: { quantity: '12' } });
  const [schemaOutput, setSchemaOutput] = React.useState<{ quantity: number }>();
  const schemaValues = schemaMethods.watch();
  const methods: FormMethods<DemoRecord> = Order.useForm({ defaultValues: createRecord(), mode: validationMode, criteriaMode: 'all' });
  const formRef = React.useMemo(() => React.createRef<HTMLFormElement>(), []);
  const slotLabelRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
  const slotHintRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const [slotClicks, setSlotClicks] = React.useState(0);
  const [slotObservation, setSlotObservation] = React.useState('Slot refs not inspected');
  const [delay, setDelay] = React.useState(true);
  const [quantityReadOnly, setQuantityReadOnly] = React.useState(true);
  const [scopeDisabled, setScopeDisabled] = React.useState(false);
  const [fail, setFail] = React.useState(false);
  const [count, setCount] = React.useState(0);
  const [status, setStatus] = React.useState('Not saved');
  const [saved, setSaved] = React.useState<DemoRecord>();
  const values = methods.watch();
  const { dirtyFields, touchedFields, errors, isDirty, isSubmitting } = methods.formState;
  const save = async (snapshot: DemoRecord): Promise<void> => {
    setCount(value => value + 1); setStatus('Saving');
    if (delay) await new Promise<void>(resolve => setTimeout(resolve, 900));
    if (fail) { setStatus('Save failed; disable failure and retry'); throw new Error('Simulated save failure'); }
    setSaved(snapshot); setStatus('Saved');
  };
  return <DemoCard diagnosticsLabel="Advanced diagnostics" title="Form — typed bindings and composition" iconName="Edit" description="Required is presentation; Reference has its own rule. Disabled fields are omitted on submit; delayed saves own a snapshot." controls={<>
    {controls}
    <div className={sx(flex.column, gap.medium)}>
    <h4>Record field props</h4>
    <div role="group" aria-label="Record field props" className={sx(flex.row, flex.wrap, gap.medium)}>
      <Checkbox data-form-demo="quantity-read-only" label="Read-only quantity" checked={quantityReadOnly} onChange={(_, data) => setQuantityReadOnly(data.checked === true)} />
      <Checkbox data-form-demo="scope-disabled" label="Disable address scope" checked={scopeDisabled} onChange={(_, data) => setScopeDisabled(data.checked === true)} />
    </div>
    </div>
    </>} commands={<>
    {mode === 'schema' && <Checkbox data-form-demo="throw-validator" label="Throw validator" checked={schemaValues.quantity === 'throw'}
      onChange={(_, data) => schemaMethods.setValue('quantity', data.checked === true ? 'throw' : '12', { shouldDirty: true })} />}
    <div hidden={mode !== 'rules'}>
    <h4>Save and data simulations</h4>
    <div role="group" aria-label="Save and record data simulations" className={sx(flex.row, flex.wrap, gap.medium)}>
      <Checkbox data-form-demo="save-delay" label="Delay save (900 ms)" checked={delay} onChange={(_, data) => setDelay(data.checked === true)} />
      <Checkbox data-form-demo="save-fail" label="Fail save" checked={fail} onChange={(_, data) => setFail(data.checked === true)} />
      <Button type="button" data-form-demo="new-record" disabled={isSubmitting} onClick={() => { methods.reset(createRecord('Order B')); setStatus('New record loaded'); }}>Load new record</Button>
      <Button type="button" data-form-demo="dirty-refresh" disabled={isSubmitting} onClick={() => { methods.reset(createRecord('Server refresh'), { keepDirtyValues: true }); setStatus('Same record refreshed; dirty values retained'); }}>Refresh retaining dirty values</Button>
      <Button type="button" data-form-demo="clear-file" disabled={disabled || isSubmitting} onClick={() => methods.resetField('files')}>Clear files through resetField</Button>
      <Button type="button" data-form-demo="request-submit" disabled={disabled} onClick={() => { formRef.current?.requestSubmit(); formRef.current?.requestSubmit(); }}>Request submit twice</Button>
    </div>
    </div>
    </>} status={<>
    <div hidden={mode !== 'rules'}>
    <p data-form-demo="save-count">Save callback count: {count}</p>
    <p role="status" data-form-demo="status">{status}; pending: {String(isSubmitting)}</p>
    </div>
    <p hidden={mode !== 'schema'} role="status">{schemaOutput ? `Submitted numeric quantity: ${schemaOutput.quantity}` : 'Schema value not submitted'}</p>
    </>} diagnostics={<>
    <div hidden={mode !== 'schema'}>
      <h4>Schema transform</h4>
      <pre data-form-demo="schema-values">{JSON.stringify({ live: schemaValues.quantity, type: typeof schemaValues.quantity })}</pre>
      <pre data-form-demo="schema-output">{JSON.stringify(schemaOutput ? { ...schemaOutput, type: typeof schemaOutput.quantity } : { status: 'Not submitted' })}</pre>
      <pre data-form-demo="schema-errors">{JSON.stringify(errorMessages(schemaMethods.formState.errors))}</pre>
    </div>
    <Button type="button" data-form-demo="inspect-slots" onClick={() => setSlotObservation(`Label: ${slotLabelRef.current?.tagName}; hint: ${slotHintRef.current?.tagName}; native target: ${slotLabelRef.current?.closest('label')?.htmlFor}`)}>Inspect presentation refs</Button>
    <output data-form-demo="slot-observation">{slotObservation}; label clicks: {slotClicks}</output>
    <h4>Live RHF values — local date parts, native files and custom value</h4><pre className={styles.output} data-form-demo="values">{JSON.stringify(inspectValue(values), undefined, 2)}</pre>
    <h4>Dirty fields</h4><pre data-form-demo="dirty">{JSON.stringify({ isDirty, fields: dirtyFields })}</pre>
    <h4>Touched fields</h4><pre data-form-demo="touched">{JSON.stringify(touchedFields)}</pre>
    <h4>Errors</h4><pre className={styles.output} data-form-demo="errors">{JSON.stringify(errorMessages(errors))}</pre>
    <h4>Saved snapshot</h4><pre className={styles.output} data-form-demo="saved">{JSON.stringify(inspectValue(saved), undefined, 2)}</pre>
    </>}>
    <div data-form-demo="frame" className={sx(styles.frame, flex.column, gap.large, width.px(size))}>
    <div hidden={mode !== 'rules'}>
    <Order.Root ref={formRef} methods={methods} disabled={disabled} onSubmit={save} onInvalid={() => setStatus('Validation failed')} data-form-demo="root" className={sx(flex.column, gap.large)}>
      <SettingsGroup data-form-demo="settings-group" title="Record settings" description="Click a row title to focus Reference or toggle Enabled. Clear Reference and save to observe validation; use Reset record to restore both values.">
        <SettingsGroup.Item field title="Reference" description="Required record identity. This row supplies the field label and hint." data-form-control="Input">
          <Order.Input data-form-demo="settings-reference" name="reference" required rules={{ required: 'Reference is required' }} />
        </SettingsGroup.Item>
        <SettingsGroup.Item field title="Enabled" description="Changes appear in live values and are persisted only by Save record." data-form-control="Switch">
          <Order.Switch data-form-demo="settings-enabled" name="enabled" />
        </SettingsGroup.Item>
      </SettingsGroup>
      <div className={sx(styles.grid, gap.medium)}>
        <div data-form-control="Textarea"><Order.Textarea name="notes" label={{ ref: slotLabelRef, className: styles.output, 'data-form-demo': 'slot-label', onClick: () => setSlotClicks(value => value + 1), children: 'Notes' }} hint={{ ref: slotHintRef, 'data-form-demo': 'slot-hint', children: (Component, props) => <Component {...props}>Callback hint: label clicked {slotClicks} times.</Component> }} resize="vertical" /></div>
        <div data-form-control="SearchBox"><Order.SearchBox name="query" label="Search query" /></div>
        <div data-form-control="Checkbox"><Order.Checkbox name="approved" label="Approved" /></div>
        <div data-form-control="SpinButton"><Order.SpinButton name="quantity" label="Quantity" readOnly={quantityReadOnly} min={0} max={20} /></div>
        <div data-form-control="Slider"><Order.Slider name="progress" label="Progress" min={0} max={100} step={5} /></div>
        <div data-form-control="Select"><Order.Select name="country" label="Country" options={options} /></div>
        <div data-form-control="RadioGroup"><Order.RadioGroup name="priority" label="Priority identity" options={choices} /></div>
        <div data-form-control="Dropdown"><Order.Dropdown name="category" label="Category" options={choices} /></div>
        <div data-form-control="DropdownMulti"><Order.DropdownMulti name="categories" label="Categories" options={choices} /></div>
        <div data-form-control="CheckboxGroup"><Order.CheckboxGroup name="checks" label="Checkbox identities" options={choices} /></div>
        <div data-form-control="Combobox"><Order.Combobox name="combo" label="Editable identity" options={choices} /></div>
        <div data-form-control="ComboboxMulti"><Order.ComboboxMulti name="combos" label="Editable identities" options={choices} /></div>
        <div data-form-control="TagPicker"><Order.TagPicker name="tags" label="Identity tags" options={choices} strings={{ noOptions: { 'data-form-demo': 'slot-no-options', children: 'No available identities' } }} /></div>
        <div data-form-control="Rating"><Order.Rating name="rating" label="Rating" max={5} /></div>
        <div data-form-control="ColorPicker"><Order.ColorPicker name="accent" label="Accent HSV"><ColorArea aria-label="Saturation and value" inputX={{ 'aria-label': 'Accent saturation' }} inputY={{ 'aria-label': 'Accent value' }} /><ColorSlider channel="hue" aria-label="Hue" /></Order.ColorPicker></div>
        <div data-form-control="SwatchPicker"><Order.SwatchPicker name="swatch" label="Swatch identity"><ColorSwatch value="blue" color="#0078d4" aria-label="Blue" /><ColorSwatch value="red" color="#d13438" aria-label="Red" /></Order.SwatchPicker></div>
        <div data-form-control="DatePicker"><Order.DatePicker name="date" label="Calendar date" allowTextInput formatDate={formatDate} parseDateFromString={parseDate} /></div>
        <div data-form-control="TimePicker"><Order.TimePicker name="time" label="Wall-clock time" freeform hourCycle="h23" formatDateToTimeString={formatTime} parseTimeStringToDate={parseTime} dateAnchor={new Date(1970, 0, 1)} /></div>
        <div data-form-control="DateTime"><Order.DateTime name="appointment" label="Separate date and time" dateLabel={{ children: 'Appointment date' }} timeLabel={{ children: (Component, props) => <Component {...props}>Appointment time</Component> }} datePickerProps={{ allowTextInput: true, formatDate, parseDateFromString: parseDate }} timePickerProps={{ freeform: true, hourCycle: 'h23', formatDateToTimeString: formatTime, parseTimeStringToDate: parseTime, dateAnchor: new Date(1970, 0, 1) }} /></div>
        <div data-form-control="FileInput"><Order.FileInput name="files" label="Native attachments" multiple browseLabel="Choose attachments" clearLabel="Clear native attachments" formatFiles={files => `${files.length} actual file(s): ${Array.from(files).map(file => file.name).join(', ')}`} /></div>
      </div>
      <Order.Input name="reference" label="Read-only reference preview" readOnly />
      <CustomField />
      <div data-form-demo="object"><Order.Object name="customer">{({ Fields }) => <Fields.Input name="name" label="Headless customer name" />}</Order.Object></div>
      <Order.ObjectSection data-form-demo="object-section" name="address" disabled={scopeDisabled} title={{ 'data-form-demo': 'slot-section-title', children: 'Address' }} description={{ 'data-form-demo': 'slot-section-description', children: (Component, props) => <Component {...props}>Relative field names inside a named fieldset.</Component> }}>{({ Fields }) => <Fields.Input name="city" label="Address city" />}</Order.ObjectSection>
      <Accordion collapsible multiple><AccordionItem value="composition"><AccordionHeader>Advanced composition and array operations</AccordionHeader><DemoAccordionPanel>
      <div data-form-demo="headless-array"><h4>Headless rows — explicit operations</h4>
        <Order.Array name="rows">{array => <>
          <p data-form-demo="headless-count">Headless row count: {array.items.length}</p>
          <div className={sx(flex.row, flex.wrap, gap.small)}>
            <Button type="button" disabled={disabled} data-form-array-action="append" onClick={() => array.append(createRow())}>Append row</Button>
            <Button type="button" disabled={disabled} data-form-array-action="prepend" onClick={() => array.prepend({ code: 'Prepended', note: '' })}>Prepend row</Button>
            <Button type="button" disabled={disabled} data-form-array-action="insert" onClick={() => array.insert(1, { code: 'Inserted', note: '' })}>Insert at 2</Button>
            <Button type="button" disabled={disabled} data-form-array-action="remove" onClick={() => array.remove(0)}>Remove first</Button>
            <Button type="button" disabled={disabled} data-form-array-action="update" onClick={() => array.update(0, { code: 'Updated', note: 'Replacement' })}>Update first</Button>
            <Button type="button" disabled={disabled} data-form-array-action="replace" onClick={() => array.replace([{ code: 'Replaced', note: 'Only row' }])}>Replace rows</Button>
            <Button type="button" disabled={disabled} data-form-array-action="move" onClick={() => array.move(0, array.items.length - 1)}>Move first to last</Button>
            <Button type="button" disabled={disabled} data-form-array-action="swap" onClick={() => array.swap(0, 1)}>Swap first two</Button>
          </div>
          {array.items.map(item => <div key={item.key} className={styles.row} data-form-headless-key={item.key}>
            <p>Row {item.index + 1} / {item.total}</p><item.Fields.Input name="code" label={`Headless code ${item.index + 1}`} /><item.Fields.Input name="note" label={`Headless note ${item.index + 1}`} />
          </div>)}
          <Order.Error name="rows" />
        </>}</Order.Array>
      </div>
      <Order.ArraySection data-form-demo="array-section" name="lines" title="Ordered line items" description="Add/Remove/Move bounds constrain UI actions only; record resets and validation are separate."
        createItem={createRow} minItems={1} maxItems={3} reorderable itemLabel={item => `Line ${item.index + 1}`}
        strings={{ add: 'Add line', empty: 'No lines', remove: label => `Remove ${label}`, moveUp: label => `Move ${label} up`, moveDown: label => `Move ${label} down`, added: label => `${label} added`, removed: label => `${label} removed`, moved: (label, position, total) => `${label} moved to ${position} of ${total}` }}
        renderItem={item => <><item.Fields.Input name="code" label={`Line code ${item.index + 1}`} /><item.Fields.Input name="note" label={`Line note ${item.index + 1}`} /></>} />
      </DemoAccordionPanel></AccordionItem></Accordion>
      <Order.Error /><Order.Error name="reference" />
      <div className={sx(flex.row, gap.small)}><Order.Submit data-form-demo="submit" appearance="primary">Save record</Order.Submit><Order.Reset data-form-demo="reset">Reset record</Order.Reset></div>
    </Order.Root>
    </div>
    <div hidden={mode !== 'schema'}><SchemaExample methods={schemaMethods} onSubmit={value => setSchemaOutput(value)} /></div>
    <Accordion collapsible><AccordionItem value="dynamic"><AccordionHeader>Dynamic Form family</AccordionHeader><DemoAccordionPanel><DynamicExample /></DemoAccordionPanel></AccordionItem></Accordion>
    </div>
  </DemoCard>;
};
const FormPanel: React.FC = () => {
  const demo = useDemoStyles();
  const [rtl, setRtl] = React.useState(false);
  const [size, setSize] = React.useState(900); const [disabled, setDisabled] = React.useState(false);
  const [mode, setMode] = React.useState('rules');
  const [validationMode, setValidationMode] = React.useState<ValidationMode>('onBlur');
  return <div data-form-demo="panel" className={demo.root}>
    <FluentProvider dir={rtl ? 'rtl' : undefined}>
    <RecordExample disabled={disabled} validationMode={validationMode} mode={mode} size={size} controls={<div className={demo.controls}>
      <Checkbox data-form-demo="direction" label="Override host direction to RTL" checked={rtl} onChange={(_, data) => setRtl(data.checked === true)} />
      <Field label="Frame width"><Select data-form-demo="resize" value={size} onChange={event => setSize(Number(event.target.value))}>{[320, 500, 640, 900].map(value => <option key={value} value={value}>{value}px</option>)}</Select></Field>
      <Checkbox data-form-demo="disabled" label="Disable record form" checked={disabled} onChange={(_, data) => setDisabled(data.checked === true)} />
      <Field label="Record validation mode"><Select data-form-demo="validation-mode" value={validationMode} onChange={(_, data) => {
        if (data.value === 'onBlur' || data.value === 'onChange' || data.value === 'onSubmit') setValidationMode(data.value);
      }}><option value="onBlur">On blur</option><option value="onChange">On change</option><option value="onSubmit">On submit</option></Select></Field>
      <Field label="Validation example"><Select data-form-demo="mode" value={mode} onChange={event => setMode(event.target.value)}><option value="rules">Rules only</option><option value="schema">Typed schema transform</option></Select></Field>
    </div>} />
    </FluentProvider>
  </div>;
};
export default FormPanel;

/**
 * Typed Form families, RHF value ownership and native Fluent bindings.
 *
 * @packageDocumentation
 * @docPage controls/form
 */
import type { StandardSchemaV1 } from '@standard-schema/spec';
import type { FieldValues, Resolver } from 'react-hook-form';
import type { CreateForm, FormSuite } from './form.types';
import { createFormSuiteRuntime } from './form';
import { createStandardSchemaResolver } from './form-validation.internal';
/**
 * Creates a stable typed Form family, optionally binding one validation owner.
 * Live methods read TInput; a schema/resolver may submit a distinct transformed TOutput.
 * Create the factory outside render. Input binds the entire string domain and has no codec.
 * @returns A family whose methods and fields share one ownership identity.
 * Example: `const Order = createForm<{ reference: string }>();`
 * @example Typed order editor
 * ```tsx
 * import * as React from 'react';
 * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ reference: string }>();
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ reference: '' }} onSubmit={values => { console.log(values.reference); }}>
 *     <Order.Input name="reference" /><Order.Submit>Save</Order.Submit>
 *   </Order.Root>;
 * }
 * ```
 * @useWhen A stable form family should type-check live paths and validated output with one validation owner.
 * @avoidWhen Existing raw RHF methods or methods from another factory must be reused unchanged; the returned Root accepts only its own suite’s methods.
 * @throws TypeError - Untyped configuration supplies both owners or neither valid schema/resolver owner.
 * @example Rules-only profile editor
 * ```tsx
 * import * as React from 'react';
 * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
 *
 * type ProfileValues = { name: string; newsletter: boolean };
 * const Profile = createForm<ProfileValues>();
 * export function ProfileEditor({ saveProfile }: { saveProfile: (values: ProfileValues) => Promise<void> }): React.ReactElement {
 *   return <Profile.Root defaultValues={{ name: '', newsletter: false }}
 *     onSubmit={async values => { await saveProfile(values); }}>
 *     <Profile.Input name="name" label="Name" required rules={{ required: 'Enter your name.' }} />
 *     <Profile.Switch name="newsletter" label="Newsletter" />
 *     <Profile.Error />
 *     <Profile.Submit>Save</Profile.Submit>
 *     <Profile.Reset>Reset</Profile.Reset>
 *   </Profile.Root>;
 * }
 * ```
 * @example Settings rows as Form fields
 * ```tsx
 * import * as React from 'react';
 * import { createForm, SettingsGroup } from '@apvee/spfx-react-toolkit';
 *
 * type AccountValues = { profile: { name: string; enabled: boolean } };
 * const Account = createForm<AccountValues>();
 * export function AccountEditor({ saveAccount }: { saveAccount: (values: AccountValues) => Promise<void> }): React.ReactElement {
 *   return <Account.Root defaultValues={{ profile: { name: '', enabled: true } }}
 *     onSubmit={async values => { await saveAccount(values); }}>
 *     <Account.Object name="profile">{({ Fields }) =>
 *       <SettingsGroup title="Profile">
 *         <SettingsGroup.Item field title="Name" description="The name shown to your colleagues.">
 *           <Fields.Input name="name" required rules={{ required: 'Enter a name.' }} />
 *         </SettingsGroup.Item>
 *         <SettingsGroup.Item field title="Enabled" description="Allow this account to receive updates.">
 *           <Fields.Switch name="enabled" />
 *         </SettingsGroup.Item>
 *       </SettingsGroup>
 *     }</Account.Object>
 *     <Account.Error />
 *     <Account.Submit>Save</Account.Submit>
 *     <Account.Reset>Reset</Account.Reset>
 *   </Account.Root>;
 * }
 * ```
 * @example Custom native input binding
 * ```tsx
 * import * as React from 'react';
 * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
 *
 * const Notes = createForm<{ note: string }>();
 *
 * function CustomNote(): React.ReactElement {
 *   const field = Notes.useField<string>({ name: 'note', rules: { required: 'Enter a note.' } });
 *   return <label>
 *     Note
 *     <input ref={field.ref} value={field.value} disabled={field.disabled}
 *       aria-invalid={field.invalid}
 *       onChange={event => field.onChange(event.currentTarget.value)}
 *       onBlur={field.onBlur} />
 *   </label>;
 * }
 *
 * export function NotesEditor(): React.ReactElement {
 *   return <Notes.Root defaultValues={{ note: '' }} onSubmit={values => { console.log(values.note); }}>
 *     <CustomNote />
 *     <Notes.Error name="note" />
 *     <Notes.Submit>Save</Notes.Submit>
 *   </Notes.Root>;
 * }
 * ```
 * @remarks
 * ### Ownership and dependencies
 * The no-argument overload leaves rules or Root validation unbound. Passing a Standard Schema directly infers live input and submit output; passing one schema/resolver option binds that owner permanently. Validation never rewrites live values into output values. Fields require the same suite’s Root, with compatible React Hook Form, shared Fluent/Griffel and date/time compat peers. No SPFx provider is required. Native Form resource mappings must be configured in the SPFx host even when supplying complete UI overrides.
 * ### Errors and asynchronous work
 * Root awaits validation, onSubmit and returned onInvalid promises and gates duplicate native submissions. Submit and Reset follow pending state; fields remain editable. Direct methods.handleSubmit retains RHF behavior and bypasses the DOM gate. Expected service failures belong in methods.setError; unexpected validation failures use __formValidation and unexpected submit/invalid-callback failures use root.submit. Pending clears so later attempts can succeed.
 * ### Records and reset
 * Default values initialize the baseline; load another record with reset(newRecord), rather than relying on object identity. Reset/resetField synchronize adapter drafts even for same-valued records. keepDirtyValues is explicit and preserves eligible dirty drafts; primitive remounts can change native ref identity. Browser files cannot be restored by defaults or reset.
 * ### Required metadata and validation
 * The required prop supplies visible and accessible metadata; it does not create a validation rule. Choose field rules, a schema or a resolver as the validation owner. Callers own persistence and domain messages; expected service failures can use same-suite `methods.setError('root.serverError', { message: 'Please try again.' })`.
 * ### Custom control ownership
 * Invoke suite useField inside its owning Root or scope. Supply a truthful native/RHF ref and your own accessible name, hint/error associations and focus target. A custom binding does not consume SettingsGroup field-row naming; use an ordinary row when composing one.
 *
 *
 * @example Configure local appointment date and wall-clock time
 * ```tsx
 * import * as React from 'react';
 * import { createForm } from '@apvee/spfx-react-toolkit';
 * // Fluent compat DatePicker and TimePicker require null for their empty domains.
 * const Appointment = createForm<{ when: { date: Date | null; time: Date | null } }>();
 * function AppointmentEditor({ save }: { save: (when: { date: Date | null; time: Date | null }) => Promise<void> }) {
 *   return <Appointment.Root defaultValues={{ when: { date: null, time: null } }} onSubmit={values => save(values.when)}>
 *     <Appointment.DateTime name="when" label="Appointment" dateLabel="Date" timeLabel="Time"
 *       datePickerProps={{ allowTextInput: true, firstDayOfWeek: 1 }}
 *       timePickerProps={{ freeform: true, hourCycle: 'h23', increment: 15 }} />
 *     <Appointment.Error name="when" /><Appointment.Submit>Save</Appointment.Submit>
 *   </Appointment.Root>;
 * }
 * ```
 *
 *
 * @example Localize controls and one form instance
 * ```tsx
 * import * as React from 'react';
 * import { EditableText, Pagination, Stage, createForm } from '@apvee/spfx-react-toolkit';
 * const ReportForm = createForm<{ title: string }>();
 * function LocalizedReport() {
 *   const methods = ReportForm.useForm({ defaultValues: { title: '' }, strings: {
 *     unexpectedFailure: 'The report could not be saved.', fileInput: { browseLabel: 'Attach report' },
 *   } });
 *   return <section>
 *     <Pagination pagination={{ type: 'pages', pageIndex: 0, pageSize: 10, totalItems: 25 }} strings={{ nextPage: 'Next report page' }} />
 *     <Stage presentation="dialog" trigger={<button>Open report</button>}>
 *       <Stage.Header title="Report" strings={{ closeLabel: 'Dismiss report' }} /><Stage.Body>Report details</Stage.Body>
 *     </Stage>
 *     <EditableText value="<h2>Report</h2>" mode="read" strings={{ formattingControls: 'Report formatting' }} />
 *     <ReportForm.Root methods={methods} onSubmit={() => { throw new Error('Service unavailable'); }}>
 *       <ReportForm.Input name="title" label="Title" /><ReportForm.Error /><ReportForm.Submit>Save</ReportForm.Submit>
 *     </ReportForm.Root>
 *   </section>;
 * }
 * ```
 * @example Edit typed reports across native fields and composition
 * ```tsx
 * import * as React from 'react';
 * import { Button, Checkbox, Field, Input, Select, ColorArea, ColorSlider, ColorSwatch, FluentProvider, makeStyles } from '@fluentui/react-components';
 * import { createForm, FormDateTimeValue, FormTimePickerParseResult, FormHsvColor, SettingsGroup, flex, gap, useSx, width } from '@apvee/spfx-react-toolkit';
 * interface DemoRegions { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode; }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <section>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</section>;
 * }
 * interface DemoRow { code: string; note: string; }
 * interface DemoRecord {
 *   reference: string; notes: string; query: string; approved: boolean; enabled: boolean;
 *   // Fluent SpinButton's empty committed domain explicitly uses null.
 *   // eslint-disable-next-line @rushstack/no-new-null
 *   quantity: number | null;
 *   progress: number; country: string; priority: string; category: string; categories: string[];
 *   checks: string[]; combo: string; combos: string[]; tags: string[]; rating: number;
 *   accent: FormHsvColor; swatch: string;
 *   // Compat DatePicker and TimePicker explicitly use null for an empty selection.
 *   // eslint-disable-next-line @rushstack/no-new-null
 *   date: Date | null;
 *   // Compat TimePicker's complete live value includes null for empty selection.
 *   // eslint-disable-next-line @rushstack/no-new-null
 *   time: Date | null;
 *   appointment: FormDateTimeValue;
 *   // HTMLInputElement.files explicitly returns null when there is no native selection.
 *   // eslint-disable-next-line @rushstack/no-new-null
 *   files: FileList | null;
 *   custom: string; customer: { name: string }; address: { city: string }; rows: DemoRow[]; lines: DemoRow[];
 * }
 * const Order = createForm<DemoRecord>();
 * const options = [{ value: '', label: 'Choose an identity' }, { value: 'it', label: 'Italy' }, { value: 'pt', label: 'Portugal' }];
 * const choices = options.slice(1);
 * const formatDate = (date?: Date): string => date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}` : '';
 * // Compat DatePicker's parser explicitly requires null for invalid/empty selection.
 * // eslint-disable-next-line @rushstack/no-new-null
 * function parseDate(text: string): Date | null {
 *   const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
 *   if (match) {
 *     const year = Number(match[1]); const month = Number(match[2]) - 1; const day = Number(match[3]);
 *     const date = new Date(year, month, day);
 *     if (date.getFullYear() === year && date.getMonth() === month && date.getDate() === day) return date;
 *   }
 *   // Compat DatePicker rejects a draft by returning its documented nullable sentinel.
 *   // eslint-disable-next-line @rushstack/no-new-null
 *   return null;
 * }
 * const formatTime = (date: Date): string => `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
 * function parseTime(text: string | undefined): FormTimePickerParseResult {
 *   const match = /^(\d{2}):(\d{2})$/.exec(text ?? '');
 *   if (match && Number(match[1]) < 24 && Number(match[2]) < 60) return { date: new Date(1970, 0, 1, Number(match[1]), Number(match[2])) };
 *   // Compat TimePicker's parse result explicitly uses null for invalid/empty time.
 *   // eslint-disable-next-line @rushstack/no-new-null
 *   return { date: null, error: text ? 'invalid-input' : undefined };
 * }
 * const createRow = (): DemoRow => ({ code: 'New row', note: '' });
 * const createRecord = (reference = 'Quarterly report'): DemoRecord => ({
 *   reference, notes: 'Team activity summary and next steps.', query: '', approved: false, enabled: true, quantity: 2, progress: 25,
 *   country: 'it', priority: 'pt', category: 'it', categories: ['pt'], checks: ['it'], combo: 'it', combos: ['pt'], tags: ['it'],
 *   rating: 3, accent: { h: 210, s: 0.7, v: 0.8, a: 1 }, swatch: 'blue', date: new Date(2026, 9, 8),
 *   time: new Date(1970, 0, 1, 9, 30), appointment: { date: new Date(2026, 9, 8), time: new Date(1970, 0, 1, 10) },
 *   // Native file inputs cannot restore a selection from defaults; null is their empty domain.
 *   // eslint-disable-next-line @rushstack/no-new-null
 *   files: null,
 *   custom: 'Custom value', customer: { name: 'Ada' }, address: { city: 'Rome' },
 *   rows: [{ code: 'Headless A', note: 'First' }, { code: 'Headless B', note: 'Second' }],
 *   lines: [{ code: 'Section A', note: 'First' }, { code: 'Section B', note: 'Second' }],
 * });
 *
 * const adapterFields = {"Input":"custom","Textarea":"notes","SearchBox":"query","Checkbox":"approved","Switch":"enabled","SpinButton":"quantity","Slider":"progress","Select":"country","RadioGroup":"priority","Dropdown":"category","DropdownMulti":"categories","CheckboxGroup":"checks","Combobox":"combo","ComboboxMulti":"combos","TagPicker":"tags","Rating":"rating","ColorPicker":"accent","SwatchPicker":"swatch","DatePicker":"date","TimePicker":"time","DateTime":"appointment","FileInput":"files"} as const;
 * type Adapter = keyof typeof adapterFields;
 * type Family = 'rules' | 'object' | 'array';
 * const useStyles = makeStyles({ frame: { maxInlineSize: '100%', minInlineSize: 0 }, output: { whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' } });
 * // Inspect native values without transforming the model.
 * function inspectValue(value: unknown): unknown {
 *   if (value instanceof Date) return { year: value.getFullYear(), month: value.getMonth() + 1, day: value.getDate(), hour: value.getHours(), minute: value.getMinutes() };
 *   if (typeof FileList !== 'undefined' && value instanceof FileList) return { count: value.length, names: Array.from(value).map(file => file.name) };
 *   if (Array.isArray(value)) return value.map(inspectValue);
 *   if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, inspectValue(item)]));
 *   return value;
 * }
 * function errorMessages(value: unknown, prefix = ''): string[] {
 *   if (!value || typeof value !== 'object') return [];
 *   return Object.entries(value).reduce<string[]>((messages, [key, item]) => {
 *     if (key === 'ref') return messages;
 *     if (key === 'message' && typeof item === 'string') return [...messages, `${prefix}: ${item}`];
 *     return [...messages, ...errorMessages(item, prefix ? `${prefix}.${key}` : key)];
 *   }, []);
 * }
 *
 * function isAdapter(value: string): value is Adapter { return Object.prototype.hasOwnProperty.call(adapterFields, value); }
 * function CustomField(): React.ReactElement {
 *   const field = Order.useField<string>({ name: 'custom' });
 *   return <Field label="Custom hook input" validationMessage={field.error} validationState={field.invalid ? 'error' : undefined}>
 *     <Input input={{ ref: field.ref }} value={field.value} disabled={field.disabled} onBlur={field.onBlur}
 *       onChange={(_, data) => field.onChange(data.value)} aria-invalid={field.invalid} />
 *   </Field>;
 * }
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx(); const styles = useStyles();
 *   const [family, setFamily] = React.useState<Family>('rules');
 *   const [adapter, setAdapter] = React.useState<Adapter>('Textarea');
 *   const [colorfulTagAvatars, setColorfulTagAvatars] = React.useState(false);
 *   const [quantityReadOnly, setQuantityReadOnly] = React.useState(false); const [scopeDisabled, setScopeDisabled] = React.useState(false);
 *   const [size, setSize] = React.useState(640); const [disabled, setDisabled] = React.useState(false);
 *   const [customStrings, setCustomStrings] = React.useState(false); const [rtl, setRtl] = React.useState(false);
 *   const [fail, setFail] = React.useState(false); const [delay, setDelay] = React.useState(false);
 *   const [count, setCount] = React.useState(0); const [status, setStatus] = React.useState('Ready to edit.');
 *   const [saved, setSaved] = React.useState<DemoRecord>(); const [slotClicks, setSlotClicks] = React.useState(0);
 *   const [slotObservation, setSlotObservation] = React.useState('Not inspected.');
 *   const slotLabelRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
 *   const slotHintRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
 *   const formRef = React.useMemo(() => React.createRef<HTMLFormElement>(), []);
 *   const methods = Order.useForm({ defaultValues: createRecord(), mode: 'onBlur', strings: customStrings ? {
 *     unexpectedFailure: 'The report could not be saved. Please retry.',
 *     fileInput: { browseLabel: 'Attach report', clearLabel: 'Remove attachment' }, arraySection: { add: 'Add report line' },
 *   } : undefined });
 *   const values = methods.watch(); const { dirtyFields, touchedFields, errors, isDirty, isSubmitting } = methods.formState;
 *   const pendingSave = React.useRef<(() => void) | undefined>(undefined);
 *   const [held, setHeld] = React.useState(false);
 *   const generation = React.useRef(0); const active = React.useRef(true);
 *   const previousReset = React.useRef(resetRevision); const [resetPending, setResetPending] = React.useState(false);
 *   const requestReset = React.useCallback((): void => {
 *     generation.current += 1; setResetPending(true);
 *     pendingSave.current?.(); pendingSave.current = undefined;
 *   }, []);
 *   React.useEffect(() => () => {
 *     active.current = false; generation.current += 1;
 *     pendingSave.current?.(); pendingSave.current = undefined;
 *   }, []);
 *   React.useEffect(() => {
 *     if (previousReset.current === resetRevision) return;
 *     previousReset.current = resetRevision; requestReset();
 *   }, [resetRevision, requestReset]);
 *
 *   // Drain the actual Root submission before RHF reset and the next session.
 *   React.useEffect(() => {
 *     if (!resetPending || isSubmitting) return;
 *     methods.reset(createRecord()); setFamily('rules'); setAdapter('Textarea'); setColorfulTagAvatars(false);
 *     setQuantityReadOnly(false); setScopeDisabled(false); setSize(640); setDisabled(false); setCustomStrings(false); setRtl(false);
 *     setFail(false); setDelay(false); setCount(0); setStatus('Ready to edit.'); setSaved(undefined); setHeld(false);
 *     setSlotClicks(0); setSlotObservation('Not inspected.'); setResetPending(false);
 *   }, [resetPending, isSubmitting, methods]);
 *   const busy = isSubmitting || resetPending;
 *   const changeAdapter = (value: string): void => {
 *     if (busy || !isAdapter(value) || value === adapter) return;
 *     const next = value; const previousValue = methods.getValues(adapterFields[adapter]);
 *     const textAdapters: readonly Adapter[] = ['Input', 'Textarea', 'SearchBox'];
 *     if (textAdapters.includes(next) && textAdapters.includes(adapter) && typeof previousValue === 'string') {
 *       methods.setValue(adapterFields[next], previousValue); setStatus('Compatible adapter: text preserved.');
 *     } else {
 *       methods.resetField(adapterFields[next]); methods.setValue(adapterFields[next], createRecord()[adapterFields[next]]);
 *       setStatus('Variable field reset for the new domain; title preserved.');
 *     }
 *     setAdapter(next);
 *   };
 *   // RHF can defer the callback; capture its source generation during render.
 *   const submissionGeneration = generation.current;
 *   const save = async (value: DemoRecord): Promise<void> => {
 *     const currentGeneration = submissionGeneration;
 *     if (!active.current || currentGeneration !== generation.current) return;
 *     const shouldFail = fail; const shouldDelay = delay;
 *     setCount(current => current + 1); setStatus('Saving');
 *     if (shouldDelay) {
 *       await new Promise<void>(resolve => { pendingSave.current = resolve; setHeld(true); });
 *       if (currentGeneration === generation.current) pendingSave.current = undefined;
 *     }
 *     if (!active.current || currentGeneration !== generation.current) return;
 *     setHeld(false);
 *     if (shouldFail) { setStatus('Save failed: disable the simulation and try again.'); throw new Error('Simulated save failure'); }
 *     setSaved(value); setStatus('Saved');
 *   };
 *   const notifyInvalid = (): void => {
 *     if (active.current && submissionGeneration === generation.current) setStatus('Check the required fields.');
 *   };
 *   // Capture the owned completion, so a retained old button cannot release a later save.
 *   const completion = pendingSave.current;
 *
 *   const variableField = (): React.ReactNode => {
 *     switch (adapter) {
 *       case 'Textarea': return <div data-form-control="Textarea"><Order.Textarea name="notes" label={{ ref: slotLabelRef,  'data-form-demo': 'slot-label', onClick: () => setSlotClicks(value => value + 1), children: 'Report description' }} hint={{ ref: slotHintRef, 'data-form-demo': 'slot-hint', children: (Component, props) => <Component {...props}>Summarize team results, decisions and next steps.{slotClicks > 0 && ` label clicked ${slotClicks} times`}</Component> }} resize="vertical" /></div>;
 *       case 'SearchBox': return <div data-form-control="SearchBox"><Order.SearchBox name="query" label="Search query" /></div>;
 *       case 'Checkbox': return <div data-form-control="Checkbox"><Order.Checkbox name="approved" label="Approved" /></div>;
 *       case 'SpinButton': return <div data-form-control="SpinButton"><Order.SpinButton name="quantity" label="Quantity"  readOnly={quantityReadOnly} min={0} max={20} /></div>;
 *       case 'Slider': return <div data-form-control="Slider"><Order.Slider name="progress" label="Progress" min={0} max={100} step={5} /></div>;
 *       case 'Select': return <div data-form-control="Select"><Order.Select name="country" label="Country" options={options} /></div>;
 *       case 'RadioGroup': return <div data-form-control="RadioGroup"><Order.RadioGroup name="priority" label="Priority identity" options={choices} /></div>;
 *       case 'Dropdown': return <div data-form-control="Dropdown"><Order.Dropdown name="category" label="Category" options={choices} /></div>;
 *       case 'DropdownMulti': return <div data-form-control="DropdownMulti"><Order.DropdownMulti name="categories" label="Categories" options={choices} /></div>;
 *       case 'CheckboxGroup': return <div data-form-control="CheckboxGroup"><Order.CheckboxGroup name="checks" label="Checkbox identities" options={choices} /></div>;
 *       case 'Combobox': return <div data-form-control="Combobox"><Order.Combobox name="combo" label="Editable identity" options={choices} /></div>;
 *       case 'ComboboxMulti': return <div data-form-control="ComboboxMulti"><Order.ComboboxMulti name="combos" label="Editable identities" options={choices} /></div>;
 *       case 'TagPicker': return <div data-form-control="TagPicker"><Order.TagPicker name="tags" label="Identity tags" options={choices} colorfulTagAvatars={colorfulTagAvatars} strings={{ noOptions: { 'data-form-demo': 'slot-no-options', children: 'No available identities' } }} /></div>;
 *       case 'Rating': return <div data-form-control="Rating"><Order.Rating name="rating" label="Rating" max={5} /></div>;
 *       case 'ColorPicker': return <div data-form-control="ColorPicker"><Order.ColorPicker name="accent" label="Accent HSV"><ColorArea aria-label="Saturation and value" inputX={{ 'aria-label': 'Accent saturation' }} inputY={{ 'aria-label': 'Accent value' }} /><ColorSlider channel="hue" aria-label="Hue" /></Order.ColorPicker></div>;
 *       case 'SwatchPicker': return <div data-form-control="SwatchPicker"><Order.SwatchPicker name="swatch" label="Swatch identity"><ColorSwatch value="blue" color="#0078d4" aria-label="Blue" /><ColorSwatch value="red" color="#d13438" aria-label="Red" /></Order.SwatchPicker></div>;
 *       case 'DatePicker': return <div data-form-control="DatePicker"><Order.DatePicker name="date" label="Calendar date" allowTextInput formatDate={formatDate} parseDateFromString={parseDate} /></div>;
 *       case 'TimePicker': return <div data-form-control="TimePicker"><Order.TimePicker name="time" label="Wall-clock time" freeform hourCycle="h23" formatDateToTimeString={formatTime} parseTimeStringToDate={parseTime} dateAnchor={new Date(1970, 0, 1)} /></div>;
 *       case 'DateTime': return <div data-form-control="DateTime"><Order.DateTime name="appointment" label="Separate date and time" dateLabel={{ children: 'Appointment date' }} timeLabel={{ children: (Component, props) => <Component {...props}>Appointment time</Component> }} datePickerProps={{ allowTextInput: true, formatDate, parseDateFromString: parseDate }} timePickerProps={{ freeform: true, hourCycle: 'h23', formatDateToTimeString: formatTime, parseTimeStringToDate: parseTime, dateAnchor: new Date(1970, 0, 1) }} /></div>;
 *       case 'FileInput': return <div data-form-control="FileInput"><Order.FileInput name="files" label="Native attachments" multiple browseLabel={customStrings ? undefined : "Choose attachments"} clearLabel={customStrings ? undefined : "Clear native attachments"} formatFiles={files => `${files.length} actual file(s): ${Array.from(files).map(file => file.name).join(', ')}`} /></div>;
 *       case 'Input': return <div data-form-control="Input"><Order.Input name="custom" label="Detail" rules={{ required: 'Enter a detail' }} /></div>;
 *       case 'Switch': return <div data-form-control="Switch"><Order.Switch name="enabled" label="Enabled" /></div>;
 *     }
 *   };
 *
 *   return <section data-form-demo="panel"><FluentProvider dir={rtl ? 'rtl' : undefined}>{render({
 *     controls: <>
 *       <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'strings' } }} label="Custom labels and failure text" checked={customStrings} onChange={(_, data) => setCustomStrings(data.checked === true)} />
 *       <Field label="Example"><Select data-demo-variant="Family" data-form-demo="mode" value={family} disabled={busy} onChange={(_, data) => {
 *         if (data.value === 'rules' || data.value === 'object' || data.value === 'array') { setFamily(data.value); setStatus('Report model preserved.'); }
 *       }}><option value="rules">Variable field</option><option value="object">Object</option><option value="array">Ordered rows</option></Select></Field>
 *       {family === 'rules' && <Field label="Field"><Select data-demo-variant="FormAdapter" data-form-demo="adapter" value={adapter} disabled={busy} onChange={(_, data) => changeAdapter(data.value)}>{Object.keys(adapterFields).map(value => <option key={value} value={value}>{value}</option>)}</Select></Field>}
 *       <Field label="Width"><Select data-form-demo="resize" value={size} onChange={(_, data) => setSize(Number(data.value))}>{[320, 500, 640, 900].map(value => <option key={value} value={value}>{value} px</option>)}</Select></Field>
 *       <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'disabled' } }} label="Disabled" checked={disabled} onChange={(_, data) => setDisabled(data.checked === true)} />
 *       {family === 'rules' && adapter === 'TagPicker' && <Checkbox label="Colorful tag avatars" checked={colorfulTagAvatars} onChange={(_, data) => setColorfulTagAvatars(data.checked === true)} />}
 *       {family === 'rules' && adapter === 'SpinButton' && <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'quantity-read-only' } }} label="Read-only quantity" checked={quantityReadOnly} onChange={(_, data) => setQuantityReadOnly(data.checked === true)} />}
 *       {family === 'object' && <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'scope-disabled' } }} label="Disable address scope" checked={scopeDisabled} onChange={(_, data) => setScopeDisabled(data.checked === true)} />}
 *     </>,
 *     commands: <>
 *       <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'save-fail' } }} label="Simulate error" checked={fail} onChange={(_, data) => setFail(data.checked === true)} />
 *       <Button data-form-demo="new-record" disabled={busy} onClick={() => { methods.reset(createRecord('Annual report')); setStatus('New report loaded.'); }}>Load new report</Button>
 *       <Button data-form-demo="dirty-refresh" disabled={busy} onClick={() => { methods.reset(createRecord('Server refresh'), { keepDirtyValues: true }); setStatus('Dirty values retained.'); }}>Refresh retaining dirty values</Button>
 *       <Button data-form-demo="clear-file" disabled={busy || disabled} onClick={() => methods.resetField('files')}>Clear files through resetField</Button>
 *       <Button data-form-demo="request-submit" disabled={disabled || resetPending} onClick={() => { formRef.current?.requestSubmit(); formRef.current?.requestSubmit(); }}>Request submit twice</Button>
 *       {isSubmitting && held && <Button data-form-demo="complete-save" onClick={() => completion?.()}>Complete save</Button>}
 *       <Button data-form-demo="reset-scenario" onClick={requestReset}>Restore example baseline</Button>
 *     </>,
 *     status: <><output data-form-demo="status" role="status">{status} · Saves: {count} · Dirty: {String(isDirty)}</output><output data-form-demo="pending">Pending: {String(busy)}</output></>,
 *     diagnostics: <>
 *       <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'direction' } }} label="Override RTL" checked={rtl} onChange={(_, data) => setRtl(data.checked === true)} />
 *       <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'save-delay' } }} label="Complete save manually" checked={delay} onChange={(_, data) => setDelay(data.checked === true)} />
 *       <Button data-form-demo="inspect-slots" onClick={() => setSlotObservation(`label:${slotLabelRef.current?.tagName} hint:${slotHintRef.current?.tagName} label-clicks:${slotClicks}`)}>Inspect slots</Button><output data-form-demo="slot-observation">{slotObservation}</output>
 *       <pre data-form-demo="values" className={styles.output}>{JSON.stringify(inspectValue(values), undefined, 2)}</pre>
 *       <pre data-form-demo="dirty">{JSON.stringify(dirtyFields)}</pre><pre data-form-demo="touched">{JSON.stringify(touchedFields)}</pre><pre data-form-demo="errors">{JSON.stringify(errorMessages(errors))}</pre>
 *       <pre data-form-demo="saved" className={styles.output}>{JSON.stringify(inspectValue(saved), undefined, 2)}</pre>
 *     </>,
 *     children: <div data-form-demo="frame" className={sx(styles.frame, flex.column, gap.medium, width.px(size))}>
 *       <Order.Root ref={formRef} methods={methods} disabled={disabled} onSubmit={save} onInvalid={notifyInvalid} data-form-demo="root" className={sx(flex.column, gap.medium)}>
 *         <SettingsGroup data-form-demo="settings-group" title="Report settings">
 *           <SettingsGroup.Item field title="Title" description="Required record identity." data-form-demo="settings-reference">
 *             <Order.Input name="reference" required rules={{ required: 'Title is required' }} data-form-demo="reference" />
 *           </SettingsGroup.Item>
 *           {family !== 'rules' && <SettingsGroup.Item field title="Enabled" description="Changes are persisted only by Save report." data-form-demo="settings-enabled"><Order.Switch name="enabled" /></SettingsGroup.Item>}
 *         </SettingsGroup>
 *         {family === 'rules' ? variableField() : family === 'object' ? <>
 *           <Field label="Read-only reference preview"><Input value={values.reference} readOnly disabled={disabled} /></Field><CustomField />
 *           <Order.Object name="customer">{({ Fields }) => <Fields.Input name="name" label="Owner" />}</Order.Object>
 *           <Order.ObjectSection name="address" disabled={scopeDisabled} title={{ 'data-form-demo': 'slot-section-title', children: 'Location' }} description={{ 'data-form-demo': 'slot-section-description', children: (Component, props) => <Component {...props}>Team address.</Component> }}>{({ Fields }) => <Fields.Input name="city" label="City" />}</Order.ObjectSection>
 *         </> : <>
 *           <Order.Array name="rows">{array => <section data-form-demo="headless-array">
 *             <p>Headless rows: {array.items.length}</p><div className={sx(flex.row, flex.wrap, gap.small)}>
 *               <Button type="button" disabled={disabled} data-form-array-action="append" onClick={() => array.append(createRow())}>Append row</Button>
 *               <Button type="button" disabled={disabled} data-form-array-action="prepend" onClick={() => array.prepend({ code: 'Prepended', note: '' })}>Prepend row</Button>
 *               <Button type="button" disabled={disabled} data-form-array-action="insert" onClick={() => array.insert(1, { code: 'Inserted', note: '' })}>Insert at 2</Button>
 *               <Button type="button" disabled={disabled} data-form-array-action="remove" onClick={() => array.remove(0)}>Remove first</Button>
 *               <Button type="button" disabled={disabled} data-form-array-action="update" onClick={() => array.update(0, { code: 'Updated', note: 'Replacement' })}>Update first</Button>
 *               <Button type="button" disabled={disabled} data-form-array-action="replace" onClick={() => array.replace([{ code: 'Replaced', note: 'Only row' }])}>Replace rows</Button>
 *               <Button type="button" disabled={disabled} data-form-array-action="move" onClick={() => array.move(0, array.items.length - 1)}>Move first to last</Button>
 *               <Button type="button" disabled={disabled} data-form-array-action="swap" onClick={() => array.swap(0, 1)}>Swap first two</Button>
 *             </div>
 *             {array.items.map(item => <div key={item.key} data-form-headless-key={item.key}><item.Fields.Input name="code" label={`Headless code ${item.index + 1}`} /><item.Fields.Input name="note" label={`Headless note ${item.index + 1}`} /></div>)}
 *             <Order.Error name="rows" />
 *           </section>}</Order.Array>
 *           <Order.ArraySection data-form-demo="array-section" name="lines" title="Activities" description="Bounds constrain UI actions, not validation or resets."
 *             createItem={createRow} minItems={1} maxItems={3} reorderable itemLabel={item => `Activities ${item.index + 1}`}
 *             strings={{ add: customStrings ? 'Add report line' : 'Add activity', empty: 'No activities', remove: label => `Remove ${label}`, moveUp: label => `Move ${label} up`, moveDown: label => `Move ${label} down`, added: label => `${label} added`, removed: label => `${label} removed`, moved: (label, position, total) => `${label} moved to ${position} of ${total}` }}
 *             renderItem={item => <><item.Fields.Input name="code" label={`Activity name ${item.index + 1}`} /><item.Fields.Input name="note" label={`Activity note ${item.index + 1}`} /></>} />
 *         </>}
 *         <Order.Error /><Order.Error name="reference" />
 *         <div className={sx(flex.row, gap.small)}><Order.Submit data-form-demo="submit" appearance="primary" disabled={resetPending}>Save report</Order.Submit><Order.Reset data-form-demo="reset" disabled={resetPending}>Reset</Order.Reset></div>
 *       </Order.Root>
 *     </div>,
 *   })}</FluentProvider></section>;
 * }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'Edit native values across 22 adapters and shared report compositions; save a snapshot, recover from failure, and reset pending work.' };
 * ```
 * @example Validate and transform quantity with recoverable failures
 * ```tsx
 * import * as React from 'react';
 * import { Button, Checkbox, Field, Select, FluentProvider, makeStyles } from '@fluentui/react-components';
 * import { createForm, flex, gap, useSx, width } from '@apvee/spfx-react-toolkit';
 * interface DemoRegions { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode; }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <section>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</section>;
 * }
 * const Numeric = createForm<{ quantity: string }, { quantity: number }>({ schema: {
 *   '~standard': {
 *     version: 1, vendor: 'Form sample',
 *     validate: (input: unknown) => {
 *       if (typeof input !== 'object' || !input || !('quantity' in input) || typeof input.quantity !== 'string') {
 *         return { issues: [{ message: 'Expected quantity text', path: ['quantity'] }] };
 *       }
 *       if (input.quantity === 'throw') throw new Error('Simulated validator failure');
 *       const quantity = Number(input.quantity);
 *       return input.quantity.trim() && Number.isFinite(quantity)
 *         ? { value: { quantity } }
 *         : { issues: [{ message: 'Enter a finite number', path: ['quantity'] }] };
 *     },
 *   },
 * } });
 * function errorMessages(value: unknown, prefix = ''): string[] {
 *   if (!value || typeof value !== 'object') return [];
 *   return Object.entries(value).reduce<string[]>((messages, [key, item]) => {
 *     if (key === 'ref') return messages;
 *     if (key === 'message' && typeof item === 'string') return [...messages, `${prefix}: ${item}`];
 *     return [...messages, ...errorMessages(item, prefix ? `${prefix}.${key}` : key)];
 *   }, []);
 * }
 *
 * const useStyles = makeStyles({ frame: { maxInlineSize: '100%', minInlineSize: 0 }, output: { whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' } });
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx(); const styles = useStyles();
 *   const [size, setSize] = React.useState(640); const [disabled, setDisabled] = React.useState(false);
 *   const [customStrings, setCustomStrings] = React.useState(false); const [rtl, setRtl] = React.useState(false);
 *   const [fail, setFail] = React.useState(false); const [delay, setDelay] = React.useState(false);
 *   const [count, setCount] = React.useState(0); const [status, setStatus] = React.useState('Ready to edit.');
 *   const [saved, setSaved] = React.useState<{ quantity: number }>();
 *   const formRef = React.useMemo(() => React.createRef<HTMLFormElement>(), []);
 *   const methods = Numeric.useForm({ defaultValues: { quantity: '12' }, strings: customStrings ? { unexpectedFailure: 'The report could not be saved. Please retry.' } : undefined });
 *   const values = methods.watch(); const { dirtyFields, touchedFields, errors, isDirty, isSubmitting } = methods.formState;
 *   const pendingSave = React.useRef<(() => void) | undefined>(undefined);
 *   const [held, setHeld] = React.useState(false);
 *   const generation = React.useRef(0); const active = React.useRef(true);
 *   const previousReset = React.useRef(resetRevision); const [resetPending, setResetPending] = React.useState(false);
 *   const requestReset = React.useCallback((): void => {
 *     generation.current += 1; setResetPending(true);
 *     pendingSave.current?.(); pendingSave.current = undefined;
 *   }, []);
 *   React.useEffect(() => () => {
 *     active.current = false; generation.current += 1;
 *     pendingSave.current?.(); pendingSave.current = undefined;
 *   }, []);
 *   React.useEffect(() => {
 *     if (previousReset.current === resetRevision) return;
 *     previousReset.current = resetRevision; requestReset();
 *   }, [resetRevision, requestReset]);
 *
 *   React.useEffect(() => {
 *     if (!resetPending || isSubmitting) return;
 *     methods.reset({ quantity: '12' }); setSize(640); setDisabled(false); setCustomStrings(false); setRtl(false);
 *     setFail(false); setDelay(false); setCount(0); setStatus('Ready to edit.'); setSaved(undefined); setHeld(false); setResetPending(false);
 *   }, [resetPending, isSubmitting, methods]);
 *   const busy = isSubmitting || resetPending;
 *   // RHF can defer the callback; capture its source generation during render.
 *   const submissionGeneration = generation.current;
 *   const save = async (value: { quantity: number }): Promise<void> => {
 *     const currentGeneration = submissionGeneration;
 *     if (!active.current || currentGeneration !== generation.current) return;
 *     const shouldFail = fail; const shouldDelay = delay;
 *     setCount(current => current + 1); setStatus('Saving');
 *     if (shouldDelay) {
 *       await new Promise<void>(resolve => { pendingSave.current = resolve; setHeld(true); });
 *       if (currentGeneration === generation.current) pendingSave.current = undefined;
 *     }
 *     if (!active.current || currentGeneration !== generation.current) return;
 *     setHeld(false);
 *     if (shouldFail) { setStatus('Save failed: disable the simulation and try again.'); throw new Error('Simulated save failure'); }
 *     setSaved(value); setStatus('Saved');
 *   };
 *   const notifyInvalid = (): void => {
 *     if (active.current && submissionGeneration === generation.current) setStatus('Check the required fields.');
 *   };
 *   // Capture the owned completion, so a retained old button cannot release a later save.
 *   const completion = pendingSave.current;
 *
 *   return <section data-form-demo="panel"><FluentProvider dir={rtl ? 'rtl' : undefined}>{render({
 *     controls: <>
 *       <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'strings' } }} label="Custom labels and failure text" checked={customStrings} onChange={(_, data) => setCustomStrings(data.checked === true)} />
 *       <Field label="Width"><Select data-form-demo="resize" value={size} onChange={(_, data) => setSize(Number(data.value))}>{[320, 500, 640, 900].map(value => <option key={value} value={value}>{value} px</option>)}</Select></Field>
 *       <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'disabled' } }} label="Disabled" checked={disabled} onChange={(_, data) => setDisabled(data.checked === true)} />
 *     </>,
 *     commands: <>
 *       <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'save-fail' } }} label="Simulate error" checked={fail} onChange={(_, data) => setFail(data.checked === true)} />
 *       <Button data-form-demo="new-record" disabled={busy} onClick={() => { methods.reset({ quantity: '12' }); setStatus('New model loaded.'); }}>Load new report</Button>
 *       <Button data-form-demo="request-submit" disabled={disabled || resetPending} onClick={() => { formRef.current?.requestSubmit(); formRef.current?.requestSubmit(); }}>Request submit twice</Button>
 *       {isSubmitting && held && <Button data-form-demo="complete-save" onClick={() => completion?.()}>Complete save</Button>}
 *       <Button data-form-demo="reset-scenario" onClick={requestReset}>Restore example baseline</Button>
 *     </>,
 *     status: <><output data-form-demo="status" role="status">{status} · Saves: {count} · Dirty: {String(isDirty)}</output><output data-form-demo="pending">Pending: {String(busy)}</output></>,
 *     diagnostics: <>
 *       <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'direction' } }} label="Override RTL" checked={rtl} onChange={(_, data) => setRtl(data.checked === true)} />
 *       <Checkbox input={{ as: 'input', ...{ 'data-form-demo': 'save-delay' } }} label="Complete save manually" checked={delay} onChange={(_, data) => setDelay(data.checked === true)} />
 *       <pre data-form-demo="values" className={styles.output}>{JSON.stringify(values, undefined, 2)}</pre>
 *       <pre data-form-demo="dirty">{JSON.stringify(dirtyFields)}</pre><pre data-form-demo="touched">{JSON.stringify(touchedFields)}</pre><pre data-form-demo="errors">{JSON.stringify(errorMessages(errors))}</pre>
 *       <pre data-form-demo="saved" className={styles.output}>{JSON.stringify(saved, undefined, 2)}</pre>
 *     </>,
 *     children: <div data-form-demo="frame" className={sx(styles.frame, flex.column, gap.medium, width.px(size))}>
 *       <Numeric.Root ref={formRef} methods={methods} disabled={disabled} onSubmit={save} onInvalid={notifyInvalid} data-form-demo="root">
 *         <Numeric.Input name="quantity" label="Quantity" /><Numeric.Error /><Numeric.Error name="quantity" />
 *         <Numeric.Submit data-form-demo="submit" disabled={resetPending}>Save report</Numeric.Submit><Numeric.Reset data-form-demo="reset" disabled={resetPending}>Reset</Numeric.Reset>
 *       </Numeric.Root>
 *     </div>,
 *   })}</FluentProvider></section>;
 * }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'Submit numeric output while live input stays text; blank, nonfinite, thrown validation and failed saves recover on retry and reset.' };
 * ```
 */
export const createForm: CreateForm = (<TInput extends FieldValues, TOutput = TInput, TContext = unknown>(configuration?: StandardSchemaV1<TInput, TOutput> | {
    /** Standard Schema v1 owner; submitted output follows its transformation. */
    schema: StandardSchemaV1<TInput, TOutput>;
    /**
     * Excluded when schema owns validation; supplying both is rejected.
     * Omit this member for the other validation-owner branches; it cannot replace a factory-bound owner.
     */
    resolver?: never;
} | {
    /** React Hook Form resolver owner for validation and output transformation. */
    resolver: Resolver<TInput, TContext, TOutput>;
    /**
     * Excluded when resolver owns validation; supplying both is rejected.
     * Omit this member for the other validation-owner branches; it cannot replace a factory-bound owner.
     */
    schema?: never;
}): FormSuite<TInput, TContext, TOutput, boolean> => {
    if (configuration === undefined)
        return createFormSuiteRuntime<TInput, TContext, TOutput, false>();
    if ('~standard' in configuration)
        return createFormSuiteRuntime<TInput, TContext, TOutput, true>(createStandardSchemaResolver(configuration));
    if (configuration.schema !== undefined && configuration.resolver !== undefined)
        throw new TypeError('createForm accepts either schema or resolver, not both.');
    if (configuration.schema !== undefined)
        return createFormSuiteRuntime<TInput, TContext, TOutput, true>(createStandardSchemaResolver(configuration.schema));
    if (typeof configuration.resolver === 'function')
        return createFormSuiteRuntime<TInput, TContext, TOutput, true>(configuration.resolver);
    throw new TypeError('createForm requires a Standard Schema v1 object or schema/resolver configuration.');
}) as CreateForm;

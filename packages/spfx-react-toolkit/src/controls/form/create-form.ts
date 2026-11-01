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

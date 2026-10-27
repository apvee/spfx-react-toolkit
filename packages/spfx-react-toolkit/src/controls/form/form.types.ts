import type * as React from 'react';
import type { FormStringsOverride } from './form-strings.types';
import type { ButtonProps } from '@fluentui/react-components';
import type { StandardSchemaV1 } from '@standard-schema/spec';
import type { FieldError, FieldErrors, FieldPath, FieldValues, FormState, RegisterOptions, Resolver, UseFormProps, UseFormReturn } from 'react-hook-form';
import type { CompatibleFieldPath, FormInputProps, FormTextareaProps, FormSearchBoxProps, FormCheckboxProps, FormSwitchProps, FormSpinButtonProps, FormSliderProps } from './form-control.types';
import type { FormSelectProps, FormRadioGroupProps, FormDropdownProps, FormDropdownMultiProps, FormCheckboxGroupProps, FormComboboxProps, FormComboboxMultiProps, FormTagPickerProps } from './form-control.types';
import type { FormRatingProps, FormColorPickerProps, FormSwatchPickerProps } from './form-control.types';
import type { FormDatePickerProps, FormTimePickerProps, FormDateTimeProps } from './form-control.types';
import type { FormFileInputProps } from './form-control.types';
import type { FormObjectComponent, FormArrayComponent, FormObjectSectionComponent, FormArraySectionComponent } from './form-composition.types';
declare const formMethodsBrand: unique symbol;
/** RHF errors plus the reserved non-root validation failure that blocks submission.
 * @example `const blocker = methods.formState.errors.__formValidation?.message;`
 */
export type FormErrors<TInput extends FieldValues> = FieldErrors<TInput> & {
    /** Internal unexpected validation error; never use this name for domain input. */
    __formValidation?: FieldError;
};
/** RHF methods belonging to one Form suite and its private reset coordinator.
 * Use the suite hook to create these methods; raw RHF methods are not interchangeable.
 * @typeParam TInput - Actual live values read by getValues/watch and written by setValue/reset.
 * @typeParam TContext - Context passed through to the validation owner.
 * @typeParam TOutput - Validated submit values; defaults to TInput.
 * @example `const methods = Order.useForm({ defaultValues: { reference: '' } });`
 */
export type FormMethods<TInput extends FieldValues = FieldValues, TContext = unknown, TOutput = TInput> = Omit<UseFormReturn<TInput, TContext, TOutput>, 'formState'> & {
    /** React Hook Form state; validation failure errors include the reserved entry. */
    formState: Omit<FormState<TInput>, 'errors'> & {
        /** Domain errors and the reserved validation failure that blocks submission. */
        errors: FormErrors<TInput>;
    };
    /** @internal Nominal suite/reset ownership marker. */
    readonly [formMethodsBrand]: true;
};
/** Standard Schema owner: transforms TInput into submitted TOutput.
 * @example `createForm({ schema })`
 */
export interface FormSchemaValidation<TInput extends FieldValues, TOutput> {
    /** Standard Schema v1 validator; its output reaches onSubmit. */
    schema: StandardSchemaV1<TInput, TOutput>;
    /** A schema cannot accompany another resolver owner. */
    resolver?: never;
}
/** RHF resolver owner; reads live TInput and submits its validated TOutput.
 * @example `createForm<{ text: string }>({ resolver })`
 */
export interface FormResolverValidation<TInput extends FieldValues, TContext, TOutput> {
    /** Resolver invoked by RHF at the configured validation lifecycle. */
    resolver: Resolver<TInput, TContext, TOutput>;
    /** A resolver cannot accompany a schema owner. */
    schema?: never;
}
/** Rules-only form initialization without a schema/resolver owner.
 * @example
 * ```tsx
 * const validation: FormWithoutValidation = {};
 * const Profile = createForm<{ name: string }>();
 * <Profile.Root {...validation} defaultValues={{ name: '' }} onSubmit={save}>
 *   <Profile.Input name="name" rules={{ required: 'Enter a name.' }} />
 * </Profile.Root>;
 * ```
 */
export interface FormWithoutValidation {
    /** Omit schema to validate through per-field rules. */
    schema?: never;
    /** Omit resolver to validate through per-field rules. */
    resolver?: never;
}
/** Exactly one schema, resolver, or rules-only validation owner.
 * @example
 * ```ts
 * const validation: FormValidationProps<{ name: string }, unknown, { name: string }> = {};
 * // Per-field rules own validation when both schema and resolver are omitted.
 * ```
 */
export type FormValidationProps<TInput extends FieldValues, TContext, TOutput> =
    FormSchemaValidation<TInput, TOutput> | FormResolverValidation<TInput, TContext, TOutput> | FormWithoutValidation;
/** RHF initialization options; a factory-bound validation owner cannot be replaced.
 * defaultValues initializes the baseline; load another record explicitly with reset(newRecord).
 * Async producers own expected loading failures. No whole-model coercion or record merge occurs.
 * @example `Order.useForm({ defaultValues: { reference: '' }, mode: 'onBlur' });`
 */
export type FormUseFormProps<TInput extends FieldValues, TContext = unknown, TOutput = TInput, TBound extends boolean = false> = Omit<UseFormProps<TInput, TContext, TOutput>, 'resolver'> & {
    /** UI dictionary owned by this hook/initializing Root. Omission uses native SPFx resources.
     * Updates affect future failures; stored errors retain their event text. No reset/revalidation occurs.
     * @example `Order.useForm({ strings: { unexpectedFailure: 'Please retry.' } });`
     */
    strings?: FormStringsOverride;
} & (TBound extends true ? {
    /** The factory-bound schema cannot be replaced. */
    schema?: never;
    /** The factory-bound resolver cannot be replaced. */
    resolver?: never;
} : FormValidationProps<TInput, TContext, TOutput>);
/** Native form attributes; content, submission and validation notifications are supplied by FormRootProps. */
type FormNativeProps = Omit<React.FormHTMLAttributes<HTMLFormElement>, 'children' | 'onSubmit' | 'onInvalid' | 'dangerouslySetInnerHTML'>;
/** Native Form Root props: own initialization or provide same-suite methods, never both.
 * noValidate defaults to true. DOM submissions share one gate across validation and save;
 * fields remain editable while Submit/Reset are pending. Direct handleSubmit bypasses that gate.
 * @example `<Order.Root defaultValues={{ reference: '' }} onSubmit={save}>...</Order.Root>`
 */
export type FormRootProps<TInput extends FieldValues = FieldValues, TContext = unknown, TOutput = TInput, TBound extends boolean = false> = FormNativeProps & {
    /** Owned form content. Raw HTML injection is not supported. */
    children?: React.ReactNode;
    /** Called after successful validation with TOutput. Its returned promise is awaited.
     * Records, arrays and Dates are copied before the callback; native browser objects retain identity.
     * Set friendly expected service errors through methods.setError; unexpected failures become root.submit.
     */
    onSubmit: (values: TOutput) => void | Promise<void>;
    /** Called after failed validation with truthful RHF errors and same-suite methods.
     * A returned promise is awaited despite the notification-style void signature; failures are contained.
     */
    onInvalid?: (errors: FormErrors<TInput>, methods: FormMethods<TInput, TContext, TOutput>) => void;
    /** Adds a disabled restriction to every binding in this Root. Omitted/false adds no restriction and cannot override disabled methods or scopes. */
    disabled?: boolean;
} & ((FormUseFormProps<TInput, TContext, TOutput, TBound> & {
    /** Omit methods when Root owns initialization. */
    methods?: never;
}) | ({
    /** Same-suite methods; initialization options must be supplied to useForm instead. */
    methods: FormMethods<TInput, TContext, TOutput>;
} & {
    [TKey in Exclude<keyof FormUseFormProps<TInput, TContext, TOutput, TBound>, 'disabled'>]?: never;
}));
/** Semantic state returned by the custom-control binding hook; values/dirty/touched stay in RHF.
 * @example `const field = Order.useField<string>({ name: 'reference' });`
 */
export interface FormFieldResult<TValue> {
    /** Current RHF live value. */
    value: TValue;
    /** Synchronously writes the complete live-domain value to RHF; validation follows the configured RHF lifecycle. Return value is ignored. */
    onChange: (value: TValue) => void;
    /** Marks the field touched. */
    onBlur: () => void;
    /** Registers the native focusable element. */
    ref: React.RefCallback<HTMLElement>;
    /** First validation message. */
    error?: string;
    /** All messages in deterministic order, without duplicates. */
    errors: readonly string[];
    /** Active validation failure. */
    invalid: boolean;
    /** Effective additive disabled state. */
    disabled: boolean;
    /** Value differs from its RHF baseline. */
    isDirty: boolean;
    /** Field has blurred. */
    isTouched: boolean;
    /** Field-local reset generation for transient draft presentation. */
    resetVersion: number;
}
/** Binding options for a custom control's complete read/write domain.
 * @example `Order.useField<string>({ name: 'reference', rules: { required: 'Reference required' } });`
 */
export interface FormUseFieldProps<TInput extends FieldValues, TValue> {
    /** Field path that can both read and accept TValue. */
    name: CompatibleFieldPath<TInput, TValue>;
    /** Rules-only domain validation; cannot accompany a resolver. */
    rules?: RegisterOptions<TInput>;
    /** Adds to Root/scope disabled state. */
    disabled?: boolean;
}
/** Native alert props; omitted name aggregates global root errors and the reserved blocker.
 * Addressed errors include array-root messages. No arbitrary ref/object traversal occurs.
 * @example `<Order.Error name="reference" />`
 */
export type FormErrorProps<TInput extends FieldValues = FieldValues> = Omit<React.HTMLAttributes<HTMLDivElement>, 'children' | 'dangerouslySetInnerHTML'> & {
    /** Field, object or array path; omit for global errors. */
    name?: FieldPath<TInput>;
};
/** Fluent submit button props; Root submissions are gated while validation/save are active.
 * The native button ref/events are forwarded; anchor rendering and custom type are not supported.
 * @example `<Order.Submit appearance="primary">Save</Order.Submit>`
 */
export type FormSubmitProps = Omit<Extract<ButtonProps, {
    /** Selects Fluent's native button branch for typing; as is omitted from the resulting public props. */
    as?: 'button';
}>, 'type' | 'as'>;
/** Fluent reset button props; native Root reset cancellation is honored.
 * Reset updates RHF and transient drafts. Consumer onReset runs first and may preventDefault.
 * @example `<Order.Reset>Discard changes</Order.Reset>`
 */
export type FormResetProps = Omit<Extract<ButtonProps, {
    /** Selects Fluent's native button branch for typing; as is omitted from the resulting public props. */
    as?: 'button';
}>, 'type' | 'as'>;
/** Field members of a typed Form family, reused with relative names inside composition scopes.
 * @example `<Order.Input name="reference" label="Reference" />`
 */
export interface ScopedFormSuite<TInput extends FieldValues> {
    /** String-valued Fluent Input and Field.
     * @example
     * ```tsx
     * <Order.Input name="reference" label="Reference" />
     * ```
     */
    Input: React.ForwardRefExoticComponent<FormInputProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Multiline string-valued Fluent Textarea with native textarea ref.
     * @example
     * ```tsx
     * <Order.Textarea name="notes" label="Notes" />
     * ```
     */
    Textarea: React.ForwardRefExoticComponent<FormTextareaProps<TInput> & React.RefAttributes<HTMLTextAreaElement>>;
    /** String search binding; clear commits an empty string before consumer notification.
     * @example
     * ```tsx
     * <Order.SearchBox name="query" label="Search" />
     * ```
     */
    SearchBox: React.ForwardRefExoticComponent<FormSearchBoxProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Boolean checkbox; explicit triState permits the mixed control value.
     * @example
     * ```tsx
     * <Order.Checkbox name="approved" label="Approved" />
     * ```
     */
    Checkbox: React.ForwardRefExoticComponent<FormCheckboxProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Boolean switch with additive disabled state and a native input ref.
     * @example
     * ```tsx
     * <Order.Switch name="enabled" label="Enabled" />
     * ```
     */
    Switch: React.ForwardRefExoticComponent<FormSwitchProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Numeric spin binding; empty/nonfinite manual commits use its nullable domain.
     * @example
     * ```tsx
     * <Order.SpinButton name="quantity" label="Quantity" min={0} />
     * ```
     */
    SpinButton: React.ForwardRefExoticComponent<FormSpinButtonProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Numeric slider using caller-supplied defaults and bounds.
     * @example
     * ```tsx
     * <Order.Slider name="progress" label="Progress" min={0} max={100} />
     * ```
     */
    Slider: React.ForwardRefExoticComponent<FormSliderProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Native select; option identity is independent of display text and optional explicit codec.
     * @example
     * ```tsx
     * <Order.Select name="country" options={[{ value: "it", label: "Italy" }]} />
     * ```
     */
    Select: React.ForwardRefExoticComponent<FormSelectProps<TInput> & React.RefAttributes<HTMLSelectElement>>;
    /** Named radio selection group with a truthful group ref and enabled-descendant focus.
     * @example
     * ```tsx
     * <Order.RadioGroup name="priority" label="Priority" options={priorities} />
     * ```
     */
    RadioGroup: React.ForwardRefExoticComponent<FormRadioGroupProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /** Single string-identity dropdown, optionally mapped by an explicit domain codec.
     * @example
     * ```tsx
     * <Order.Dropdown name="country" label="Country" options={countries} />
     * ```
     */
    Dropdown: React.ForwardRefExoticComponent<FormDropdownProps<TInput> & React.RefAttributes<HTMLButtonElement>>;
    /** Ordered multiple dropdown selection with optional explicit domain conversion.
     * @example
     * ```tsx
     * <Order.DropdownMulti name="countries" label="Countries" options={countries} />
     * ```
     */
    DropdownMulti: React.ForwardRefExoticComponent<FormDropdownMultiProps<TInput> & React.RefAttributes<HTMLButtonElement>>;
    /** Ordered checkbox identities; adaptation is limited to documented Fluent Checkbox children.
     * @example
     * ```tsx
     * <Order.CheckboxGroup name="tags" label="Tags" options={tags} />
     * ```
     */
    CheckboxGroup: React.ForwardRefExoticComponent<FormCheckboxGroupProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /** Editable single selection; unmatched query clears its committed identity.
     * @example
     * ```tsx
     * <Order.Combobox name="country" label="Country" options={countries} />
     * ```
     */
    Combobox: React.ForwardRefExoticComponent<FormComboboxProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Editable multiple selection; query edits retain committed selected identities.
     * @example
     * ```tsx
     * <Order.ComboboxMulti name="countries" label="Countries" options={countries} />
     * ```
     */
    ComboboxMulti: React.ForwardRefExoticComponent<FormComboboxMultiProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Selected identity tags with keyboard removal and caller-localized empty text.
     * @example
     * ```tsx
     * <Order.TagPicker name="people" label="People" options={people} />
     * ```
     */
    TagPicker: React.ForwardRefExoticComponent<FormTagPickerProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Numeric Fluent Rating with a real group ref and managed radio error focus.
     * @example
     * ```tsx
     * <Order.Rating name="score" label="Score" max={5} />
     * ```
     */
    Rating: React.ForwardRefExoticComponent<FormRatingProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /** HSV color channels with optional alpha, copied before writing the live field.
     * @example
     * ```tsx
     * <Order.ColorPicker name="accent" label="Accent"><ColorSlider channel="hue" /></Order.ColorPicker>
     * ```
     */
    ColorPicker: React.ForwardRefExoticComponent<FormColorPickerProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /** Stable string swatch identity; the public ref remains on the real picker group.
     * @example
     * ```tsx
     * <Order.SwatchPicker name="theme" label="Theme"><ColorSwatch value="blue" color="#0078d4" /></Order.SwatchPicker>
     * ```
     */
    SwatchPicker: React.ForwardRefExoticComponent<FormSwatchPickerProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /** Nullable local-calendar date with visible invalid drafts and a native input ref.
     * @example
     * ```tsx
     * <Order.DatePicker name="date" label="Date" />
     * ```
     */
    DatePicker: React.ForwardRefExoticComponent<FormDatePickerProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Nullable local wall-clock time using a deterministic anchor and native input ref.
     * @example
     * ```tsx
     * <Order.TimePicker name="time" label="Time" />
     * ```
     */
    TimePicker: React.ForwardRefExoticComponent<FormTimePickerProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Independent nullable date/time parts; merging never creates an implicit timestamp.
     * @example
     * ```tsx
     * <Order.DateTime name="appointment" dateLabel="Date" timeLabel="Time" />
     * ```
     */
    DateTime: React.ForwardRefExoticComponent<FormDateTimeProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /** Actual browser FileList selection; clear/reset never fabricates native file restoration.
     * @example
     * ```tsx
     * <Order.FileInput name="attachments" label="Attachments" multiple />
     * ```
     */
    FileInput: React.ForwardRefExoticComponent<FormFileInputProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Headless nested-record scope with path-specific relative field inference.
     * @example
     * ```tsx
     * <Order.Object name="customer">{({ Fields }) => <Fields.Input name="name" />}</Order.Object>
     * ```
     */
    Object: FormObjectComponent<TInput>;
    /** Headless structured-record array with stable item keys and explicit operations.
     * @example
     * ```tsx
     * <Order.Array name="rows">{({ items }) => items.map(({ key, Fields }) => <Fields.Input key={key} name="sku" />)}</Order.Array>
     * ```
     */
    Array: FormArrayComponent<TInput>;
    /** Named visual nested-record fieldset with direct errors and inherited disabled state.
     * @example
     * ```tsx
     * <Order.ObjectSection name="customer" title="Customer">{({ Fields }) => <Fields.Input name="name" />}</Order.ObjectSection>
     * ```
     */
    ObjectSection: FormObjectSectionComponent<TInput>;
    /** Named ordered editable array with localized actions, bounds and action-owned focus.
     * @example
     * ```tsx
     * <Order.ArraySection name="rows" title="Rows" createItem={() => ({ sku: "" })} renderItem={({ Fields }) => <Fields.Input name="sku" />} />
     * ```
     */
    ArraySection: FormArraySectionComponent<TInput>;
    /** Field or global native alert.
     * @example
     * ```tsx
     * <Order.Error name="reference" />
     * ```
     */
    Error: React.ForwardRefExoticComponent<FormErrorProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /** React hook binding a custom control to live RHF value/state inside its owning Root/scope.
     * Pass a compatible relative name and wire onChange/onBlur/ref; returned state remains owned by RHF.
     * @example
     * ```tsx
     * const field = Order.useField<string>({ name: "reference" });
     * ```
     */
    useField: <TValue>(props: FormUseFieldProps<TInput, TValue>) => FormFieldResult<TValue>;
}
/** Typed form family. Create factories outside render to keep suite ownership stable.
 * @example `const Order = createForm<{ reference: string }>();`
 */
export interface FormSuite<TInput extends FieldValues, TContext = unknown, TOutput = TInput, TBound extends boolean = false> extends ScopedFormSuite<TInput> {
    /** Native form and RHF provider.
     * @example
     * ```tsx
     * <Order.Root defaultValues={{ reference: "" }} onSubmit={save}><Order.Input name="reference" /></Order.Root>
     * ```
     */
    Root: React.ForwardRefExoticComponent<FormRootProps<TInput, TContext, TOutput, TBound> & React.RefAttributes<HTMLFormElement>>;
    /** Fluent submit button, disabled while RHF is submitting.
     * @example
     * ```tsx
     * <Order.Submit appearance="primary">Save</Order.Submit>
     * ```
     */
    Submit: React.ForwardRefExoticComponent<FormSubmitProps & React.RefAttributes<HTMLButtonElement>>;
    /** Fluent reset button, disabled while RHF is submitting.
     * @example
     * ```tsx
     * <Order.Reset>Discard edits</Order.Reset>
     * ```
     */
    Reset: React.ForwardRefExoticComponent<FormResetProps & React.RefAttributes<HTMLButtonElement>>;
    /** React hook creating same-suite RHF methods with optional initialization.
     * Omission uses RHF defaults; factory-bound validation cannot be replaced.
     * Direct handleSubmit bypasses Root's DOM gate; pass these methods to Root instead of initialization props.
     * @example
     * ```ts
     * const methods = Order.useForm({ defaultValues: { reference: "" }, mode: "onBlur" });
     * ```
     */
    useForm: (props?: FormUseFormProps<TInput, TContext, TOutput, TBound>) => FormMethods<TInput, TContext, TOutput>;
}
/** Untyped family for dynamic domain names.
 * @example
 * ```tsx
 * const DynamicForm: UntypedFormSuite = Form;
 * <DynamicForm.Input name="runtime.field" label="Dynamic field" />;
 * ```
 */
export type UntypedFormSuite = FormSuite<FieldValues>;
/** Typed factory overloads preserving distinct live input and submitted output.
 * Create suites outside render. Bound owners cannot be replaced at Root or useForm.
 * @example `const Order = createForm<{ reference: string }>();`
 */
export interface CreateForm {
    /** Creates a rules-only or Root-configured family with this live model. */
    <TInput extends FieldValues>(): FormSuite<TInput>;
    /** Infers the live input and transformed output from a Standard Schema. */
    <TSchema extends StandardSchemaV1>(schema: TSchema & (StandardSchemaV1.InferInput<TSchema> extends FieldValues ? unknown : never)): FormSuite<Extract<StandardSchemaV1.InferInput<TSchema>, FieldValues>, unknown, StandardSchemaV1.InferOutput<TSchema>, true>;
    /** Permanently binds one schema or resolver owner.
     * @param options - Exactly one validation owner; its input stays live and its output reaches onSubmit.
     * @returns A stable family with validation ownership fixed for Root and useForm.
     */
    <TInput extends FieldValues, TOutput = TInput, TContext = unknown>(options: {
        /** Standard Schema v1 owner mapping live input to validated output. */
        schema: StandardSchemaV1<TInput, TOutput>;
        /** Excluded when schema owns validation. */
        resolver?: never;
    } | {
        /** RHF resolver owner receiving live input and the configured context. */
        resolver: Resolver<TInput, TContext, TOutput>;
        /** Excluded when resolver owns validation. */
        schema?: never;
    }): FormSuite<TInput, TContext, TOutput, true>;
}

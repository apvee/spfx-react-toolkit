import type * as React from 'react';
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
/** Rules-only form initialization without a schema/resolver owner. */
export interface FormWithoutValidation {
    /** Omit schema to validate through per-field rules. */
    schema?: never;
    /** Omit resolver to validate through per-field rules. */
    resolver?: never;
}
/** Exactly one schema, resolver, or rules-only validation owner. */
export type FormValidationProps<TInput extends FieldValues, TContext, TOutput> =
    FormSchemaValidation<TInput, TOutput> | FormResolverValidation<TInput, TContext, TOutput> | FormWithoutValidation;
/** RHF initialization options; a factory-bound validation owner cannot be replaced.
 * defaultValues initializes the baseline; load another record explicitly with reset(newRecord).
 * Async producers own expected loading failures. No whole-model coercion or record merge occurs.
 * @example `Order.useForm({ defaultValues: { reference: '' }, mode: 'onBlur' });`
 */
export type FormUseFormProps<TInput extends FieldValues, TContext = unknown, TOutput = TInput, TBound extends boolean = false> = Omit<UseFormProps<TInput, TContext, TOutput>, 'resolver'> & (TBound extends true ? {
    /** The factory-bound schema cannot be replaced. */
    schema?: never;
    /** The factory-bound resolver cannot be replaced. */
    resolver?: never;
} : FormValidationProps<TInput, TContext, TOutput>);
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
    /** Disables all bindings in this Root. */
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
    /** Writes the complete field value. */
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
    as?: 'button';
}>, 'type' | 'as'>;
/** Fluent reset button props; native Root reset cancellation is honored.
 * Reset updates RHF and transient drafts. Consumer onReset runs first and may preventDefault.
 * @example `<Order.Reset>Discard changes</Order.Reset>`
 */
export type FormResetProps = Omit<Extract<ButtonProps, {
    as?: 'button';
}>, 'type' | 'as'>;
/** Field members of a typed Form family, reused with relative names inside composition scopes.
 * @example `<Order.Input name="reference" label="Reference" />`
 */
export interface ScopedFormSuite<TInput extends FieldValues> {
    /** String-valued Fluent Input and Field. */
    Input: React.ForwardRefExoticComponent<FormInputProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Multiline string-valued Fluent Textarea with native textarea ref. */
    Textarea: React.ForwardRefExoticComponent<FormTextareaProps<TInput> & React.RefAttributes<HTMLTextAreaElement>>;
    /** String search binding; clear commits an empty string before consumer notification. */
    SearchBox: React.ForwardRefExoticComponent<FormSearchBoxProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Boolean checkbox; explicit triState permits the mixed control value. */
    Checkbox: React.ForwardRefExoticComponent<FormCheckboxProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Boolean switch with additive disabled state and a native input ref. */
    Switch: React.ForwardRefExoticComponent<FormSwitchProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Numeric spin binding; empty/nonfinite manual commits use its nullable domain. */
    SpinButton: React.ForwardRefExoticComponent<FormSpinButtonProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Numeric slider using caller-supplied defaults and bounds. */
    Slider: React.ForwardRefExoticComponent<FormSliderProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Native select; option identity is independent of display text and optional explicit codec. */
    Select: React.ForwardRefExoticComponent<FormSelectProps<TInput> & React.RefAttributes<HTMLSelectElement>>;
    /** Named radio selection group with a truthful group ref and enabled-descendant focus. */
    RadioGroup: React.ForwardRefExoticComponent<FormRadioGroupProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /** Single string-identity dropdown, optionally mapped by an explicit domain codec. */
    Dropdown: React.ForwardRefExoticComponent<FormDropdownProps<TInput> & React.RefAttributes<HTMLButtonElement>>;
    /** Ordered multiple dropdown selection with optional explicit domain conversion. */
    DropdownMulti: React.ForwardRefExoticComponent<FormDropdownMultiProps<TInput> & React.RefAttributes<HTMLButtonElement>>;
    /** Ordered checkbox identities; adaptation is limited to documented Fluent Checkbox children. */
    CheckboxGroup: React.ForwardRefExoticComponent<FormCheckboxGroupProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /** Editable single selection; unmatched query clears its committed identity. */
    Combobox: React.ForwardRefExoticComponent<FormComboboxProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Editable multiple selection; query edits retain committed selected identities. */
    ComboboxMulti: React.ForwardRefExoticComponent<FormComboboxMultiProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Selected identity tags with keyboard removal and caller-localized empty text. */
    TagPicker: React.ForwardRefExoticComponent<FormTagPickerProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Numeric Fluent Rating with a real group ref and managed radio error focus. */
    Rating: React.ForwardRefExoticComponent<FormRatingProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /** HSV color channels with optional alpha, copied before writing the live field. */
    ColorPicker: React.ForwardRefExoticComponent<FormColorPickerProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /** Stable string swatch identity; the public ref remains on the real picker group. */
    SwatchPicker: React.ForwardRefExoticComponent<FormSwatchPickerProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /** Nullable local-calendar date with visible invalid drafts and a native input ref. */
    DatePicker: React.ForwardRefExoticComponent<FormDatePickerProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Nullable local wall-clock time using a deterministic anchor and native input ref. */
    TimePicker: React.ForwardRefExoticComponent<FormTimePickerProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Independent nullable date/time parts; merging never creates an implicit timestamp. */
    DateTime: React.ForwardRefExoticComponent<FormDateTimeProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /** Actual browser FileList selection; clear/reset never fabricates native file restoration. */
    FileInput: React.ForwardRefExoticComponent<FormFileInputProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /** Headless nested-record scope with path-specific relative field inference. */
    Object: FormObjectComponent<TInput>;
    /** Headless structured-record array with stable item keys and explicit operations. */
    Array: FormArrayComponent<TInput>;
    /** Named visual nested-record fieldset with direct errors and inherited disabled state. */
    ObjectSection: FormObjectSectionComponent<TInput>;
    /** Named ordered editable array with localized actions, bounds and action-owned focus. */
    ArraySection: FormArraySectionComponent<TInput>;
    /** Field or global native alert. */
    Error: React.ForwardRefExoticComponent<FormErrorProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /** Binds custom controls to live RHF value/state. */
    useField: <TValue>(props: FormUseFieldProps<TInput, TValue>) => FormFieldResult<TValue>;
}
/** Typed form family. Create factories outside render to keep suite ownership stable.
 * @example `const Order = createForm<{ reference: string }>();`
 */
export interface FormSuite<TInput extends FieldValues, TContext = unknown, TOutput = TInput, TBound extends boolean = false> extends ScopedFormSuite<TInput> {
    /** Native form and RHF provider. */
    Root: React.ForwardRefExoticComponent<FormRootProps<TInput, TContext, TOutput, TBound> & React.RefAttributes<HTMLFormElement>>;
    /** Fluent submit button, disabled while RHF is submitting. */
    Submit: React.ForwardRefExoticComponent<FormSubmitProps & React.RefAttributes<HTMLButtonElement>>;
    /** Fluent reset button, disabled while RHF is submitting. */
    Reset: React.ForwardRefExoticComponent<FormResetProps & React.RefAttributes<HTMLButtonElement>>;
    /** Creates methods owned by this family. Direct handleSubmit bypasses Root's DOM gate. */
    useForm: (props?: FormUseFormProps<TInput, TContext, TOutput, TBound>) => FormMethods<TInput, TContext, TOutput>;
}
/** Untyped family for dynamic domain names. */
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
    /** Permanently binds one schema or resolver owner. */
    <TInput extends FieldValues, TOutput = TInput, TContext = unknown>(options: {
        schema: StandardSchemaV1<TInput, TOutput>;
        resolver?: never;
    } | {
        resolver: Resolver<TInput, TContext, TOutput>;
        schema?: never;
    }): FormSuite<TInput, TContext, TOutput, true>;
}

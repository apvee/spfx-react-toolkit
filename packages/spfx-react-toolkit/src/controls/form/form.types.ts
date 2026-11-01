/**
 * Typed Form families, RHF value ownership and native Fluent bindings.
 *
 * @packageDocumentation
 * @docPage controls/form
 */
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
/**
 * RHF errors plus the reserved non-root validation failure that blocks submission.
 * @example
 * ```ts
 * import { FormErrors } from '@apvee/spfx-react-toolkit/controls/form';
 * const errors: FormErrors<{ reference: string }> = { __formValidation: { type: 'unexpected', message: 'Please retry.' } };
 * console.log(errors.__formValidation?.message);
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormErrors<TInput extends FieldValues> = FieldErrors<TInput> & {
    /**
     * Internal unexpected validation error; never use this name for domain input.
     * Undefined means no unexpected validation-owner failure is blocking submission.
     */
    __formValidation?: FieldError;
};
/**
 * RHF methods belonging to one Form suite and its private reset coordinator.
 * Use the suite hook to create these methods; raw RHF methods are not interchangeable.
 * @typeParam TInput - Actual live values read by getValues/watch and written by setValue/reset.
 * @typeParam TContext - Context passed through to the validation owner.
 * @typeParam TOutput - Validated submit values; defaults to TInput.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormMethods } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ reference: string }>();
 * export function Editor(): React.ReactElement {
 *   const methods: FormMethods<{ reference: string }> = Order.useForm({ defaultValues: { reference: '' } });
 *   return <Order.Root methods={methods} onSubmit={values => { console.log(values.reference); }}>
 *     <Order.Input name="reference" />
 *   </Order.Root>;
 * }
 * ```
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
/**
 * Standard Schema owner: transforms TInput into submitted TOutput.
 * @example
 * ```ts
 * import type { StandardSchemaV1 } from '@standard-schema/spec';
 * import { FormSchemaValidation } from '@apvee/spfx-react-toolkit/controls/form';
 * const schema: StandardSchemaV1<{ reference: string }> = {
 *   '~standard': { version: 1, vendor: 'reference-example', validate: input =>
 *     typeof input === 'object' && input !== null && 'reference' in input && typeof input.reference === 'string'
 *       ? { value: { reference: input.reference } }
 *       : { issues: [{ message: 'Enter a reference.' }] } },
 * };
 * const validation: FormSchemaValidation<{ reference: string }, { reference: string }> = { schema };
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TOutput - Validated submission model, which can differ from live input; defaults to TInput where optional.
 */
export interface FormSchemaValidation<TInput extends FieldValues, TOutput> {
    /** Standard Schema v1 validator; its output reaches onSubmit. */
    schema: StandardSchemaV1<TInput, TOutput>;
    /**
     * A schema cannot accompany another resolver owner.
     * Omit this member for the other validation-owner branches; it cannot replace a factory-bound owner.
     */
    resolver?: never;
}
/**
 * RHF resolver owner; reads live TInput and submits its validated TOutput.
 * @example
 * ```ts
 * import { FormResolverValidation } from '@apvee/spfx-react-toolkit/controls/form';
 * const validation: FormResolverValidation<{ reference: string }, unknown, { reference: string }> = {
 *   resolver: values => ({ values, errors: {} }),
 * };
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TContext - Context passed to the React Hook Form resolver; defaults to unknown where optional.
 * @typeParam TOutput - Validated submission model, which can differ from live input; defaults to TInput where optional.
 */
export interface FormResolverValidation<TInput extends FieldValues, TContext, TOutput> {
    /** Resolver invoked by RHF at the configured validation lifecycle. */
    resolver: Resolver<TInput, TContext, TOutput>;
    /**
     * A resolver cannot accompany a schema owner.
     * Omit this member for the other validation-owner branches; it cannot replace a factory-bound owner.
     */
    schema?: never;
}
/**
 * Rules-only form initialization without a schema/resolver owner.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormWithoutValidation } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ reference: string }>();
 * const validation: FormWithoutValidation = {};
 * export function Editor(): React.ReactElement {
 *   return <Order.Root {...validation} defaultValues={{ reference: '' }} onSubmit={values => { console.log(values); }}>
 *     <Order.Input name="reference" rules={{ required: "Enter a reference." }} />
 *   </Order.Root>;
 * }
 * ```
 */
export interface FormWithoutValidation {
    /** Omit schema to validate through per-field rules. */
    schema?: never;
    /** Omit resolver to validate through per-field rules. */
    resolver?: never;
}
/**
 * Exactly one schema, resolver, or rules-only validation owner.
 * @example
 * ```ts
 * import { FormValidationProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const validation: FormValidationProps<{ name: string }, unknown, { name: string }> = {};
 * // Field rules own validation when schema and resolver are omitted.
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TContext - Context passed to the React Hook Form resolver; defaults to unknown where optional.
 * @typeParam TOutput - Validated submission model, which can differ from live input; defaults to TInput where optional.
 */
export type FormValidationProps<TInput extends FieldValues, TContext, TOutput> =
    FormSchemaValidation<TInput, TOutput> | FormResolverValidation<TInput, TContext, TOutput> | FormWithoutValidation;
/**
 * RHF initialization options; a factory-bound validation owner cannot be replaced.
 * defaultValues initializes the baseline; load another record explicitly with reset(newRecord).
 * Async producers own expected loading failures. No whole-model coercion or record merge occurs.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormUseFormProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ reference: string }>();
 * export function Editor(): React.ReactElement {
 *   const options: FormUseFormProps<{ reference: string }> = { defaultValues: { reference: '' }, mode: 'onBlur' };
 *   const methods = Order.useForm(options);
 *   return <Order.Root methods={methods} onSubmit={values => { console.log(values.reference); }}>
 *     <Order.Input name="reference" />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TContext - Context passed to the React Hook Form resolver; defaults to unknown where optional.
 * @typeParam TOutput - Validated submission model, which can differ from live input; defaults to TInput where optional.
 * @typeParam TBound - Whether a factory-bound validator forbids replacement at Root/useForm; defaults to false.
 * @remarks
 * RHF owns the inherited initialization options and defaults, including partial/async defaultValues, reactive values and resetOptions. Toolkit adds strings and exclusive schema/resolver selection. Async default producers handle expected loading failures themselves. A methods-owning Root cannot also receive these initialization options.
 */
export type FormUseFormProps<TInput extends FieldValues, TContext = unknown, TOutput = TInput, TBound extends boolean = false> = Omit<UseFormProps<TInput, TContext, TOutput>, 'resolver'> & {
    /**
     * UI dictionary owned by this hook/initializing Root. Omission uses native SPFx resources.
     * Updates affect future failures; stored errors retain their event text. No reset/revalidation occurs.
     * Example: `Order.useForm({ strings: { unexpectedFailure: 'Please retry.' } });`
     * When omitted, native SPFx resources and enclosing per-form dictionaries provide package-owned UI text.
     */
    strings?: FormStringsOverride;
} & (TBound extends true ? {
    /**
     * The factory-bound schema cannot be replaced.
     * Omit this member for the other validation-owner branches; it cannot replace a factory-bound owner.
     */
    schema?: never;
    /**
     * The factory-bound resolver cannot be replaced.
     * Omit this member for the other validation-owner branches; it cannot replace a factory-bound owner.
     */
    resolver?: never;
} : FormValidationProps<TInput, TContext, TOutput>);
/** Native form attributes; content, submission and validation notifications are supplied by FormRootProps. */
type FormNativeProps = Omit<React.FormHTMLAttributes<HTMLFormElement>, 'children' | 'onSubmit' | 'onInvalid' | 'dangerouslySetInnerHTML'>;
/**
 * Native Form Root props: own initialization or provide same-suite methods, never both.
 * noValidate defaults to true. DOM submissions share one gate across validation and save;
 * fields remain editable while Submit/Reset are pending. Direct handleSubmit bypasses that gate.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormRootProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ reference: string }>();
 * const rootProps: FormRootProps<{ reference: string }> = {
 *   defaultValues: { reference: '' }, onSubmit: values => { console.log(values.reference); },
 * };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root {...rootProps}>
 *     <Order.Input name="reference" />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TContext - Context passed to the React Hook Form resolver; defaults to unknown where optional.
 * @typeParam TOutput - Validated submission model, which can differ from live input; defaults to TInput where optional.
 * @typeParam TBound - Whether a factory-bound validator forbids replacement at Root/useForm; defaults to false.
 */
export type FormRootProps<TInput extends FieldValues = FieldValues, TContext = unknown, TOutput = TInput, TBound extends boolean = false> = FormNativeProps & {
    /**
     * Owned form content. Raw HTML injection is not supported.
     * When omitted, the region has no caller content.
     */
    children?: React.ReactNode;
    /**
     * Called after successful validation with TOutput. Its returned promise is awaited.
     * Records, arrays and Dates are copied before the callback; native browser objects retain identity.
     * Set friendly expected service errors through methods.setError; unexpected failures become root.submit.
     * @param values - Validated output snapshot; plain records, arrays and Dates are copied before the callback, while native browser objects retain identity.
     * @returns Void or a promise awaited by Root; unexpected failures become root.submit and pending is released.
     */
    onSubmit: (values: TOutput) => void | Promise<void>;
    /**
     * Called after failed validation with truthful RHF errors and same-suite methods.
     * A returned promise is awaited despite the notification-style void signature; failures are contained.
     * When omitted, failed validation still updates RHF errors and does not call an application notification.
     * @param errors - Truthful RHF validation errors, including the reserved unexpected-validation blocker.
     * @param methods - Same-suite enhanced RHF methods for inspecting values and setting application errors.
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
/**
 * Semantic state returned by the custom-control binding hook; values/dirty/touched stay in RHF.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormFieldResult } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ reference: string }>();
 * function CustomReference(): React.ReactElement {
 *   const field: FormFieldResult<string> = Order.useField<string>({ name: 'reference' });
 *   return <input ref={field.ref} value={field.value} disabled={field.disabled} onBlur={field.onBlur}
 *     onChange={event => field.onChange(event.currentTarget.value)} aria-label="Reference" />;
 * }
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ reference: '' }} onSubmit={values => { console.log(values.reference); }}>
 *     <CustomReference />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TValue - Complete domain read and written by the custom control.
 */
export interface FormFieldResult<TValue> {
    /** Current RHF live value. */
    value: TValue;
    /**
     * Synchronously writes the complete live-domain value to RHF; validation follows the configured RHF lifecycle. Return value is ignored.
     * @param value - Complete live-domain value to write or convert; no implicit domain coercion occurs.
     */
    onChange: (value: TValue) => void;
    /** Marks the field touched. */
    onBlur: () => void;
    /** Registers the native focusable element. */
    ref: React.RefCallback<HTMLElement>;
    /**
     * First validation message.
     * Undefined means no direct validation message is available.
     */
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
/**
 * Binding options for a custom control's complete read/write domain.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormUseFieldProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ reference: string }>();
 * function CustomReference(): React.ReactElement {
 *   const options: FormUseFieldProps<{ reference: string }, string> = { name: 'reference', rules: { required: 'Reference required' } };
 *   const field = Order.useField<string>(options);
 *   return <input ref={field.ref} value={field.value} disabled={field.disabled} onBlur={field.onBlur}
 *     onChange={event => field.onChange(event.currentTarget.value)} aria-label="Reference" />;
 * }
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ reference: '' }} onSubmit={values => { console.log(values.reference); }}>
 *     <CustomReference />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TValue - Complete domain read and written by the custom control.
 */
export interface FormUseFieldProps<TInput extends FieldValues, TValue> {
    /** Field path that can both read and accept TValue. */
    name: CompatibleFieldPath<TInput, TValue>;
    /**
     * Rules-only domain validation; cannot accompany a resolver.
     * When omitted, no field or array domain rules are registered here; schema/resolver ownership remains unchanged.
     */
    rules?: RegisterOptions<TInput>;
    /**
     * Adds to Root/scope disabled state.
     * When omitted or false, adds no disabled restriction; enclosing Root/scope restrictions still apply.
     */
    disabled?: boolean;
}
/**
 * Native alert props; omitted name aggregates global root errors and the reserved blocker.
 * Addressed errors include array-root messages. No arbitrary ref/object traversal occurs.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormErrorProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ reference: string }>();
 * const props: FormErrorProps<{ reference: string }> = { name: 'reference' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ reference: '' }} onSubmit={values => { console.log(values); }}>
 *     <Order.Error {...props} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormErrorProps<TInput extends FieldValues = FieldValues> = Omit<React.HTMLAttributes<HTMLDivElement>, 'children' | 'dangerouslySetInnerHTML'> & {
    /** Field, object or array path; omit for global errors. */
    name?: FieldPath<TInput>;
};
/**
 * Fluent submit button props; Root submissions are gated while validation/save are active.
 * The native button ref/events are forwarded; anchor rendering and custom type are not supported.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormSubmitProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ reference: string }>();
 * const props: FormSubmitProps = { children: 'Save' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ reference: '' }} onSubmit={values => { console.log(values); }}>
 *     <Order.Submit {...props} />
 *   </Order.Root>;
 * }
 * ```
 */
export type FormSubmitProps = Omit<Extract<ButtonProps, {
    /** Selects Fluent's native button branch for typing; as is omitted from the resulting public props. */
    as?: 'button';
}>, 'type' | 'as'>;
/**
 * Fluent reset button props; native Root reset cancellation is honored.
 * Reset updates RHF and transient drafts. Consumer onReset runs first and may preventDefault.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormResetProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ reference: string }>();
 * const props: FormResetProps = { children: 'Discard changes' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ reference: '' }} onSubmit={values => { console.log(values); }}>
 *     <Order.Reset {...props} />
 *   </Order.Root>;
 * }
 * ```
 */
export type FormResetProps = Omit<Extract<ButtonProps, {
    /** Selects Fluent's native button branch for typing; as is omitted from the resulting public props. */
    as?: 'button';
}>, 'type' | 'as'>;
/**
 * Field members of a typed Form family, reused with relative names inside composition scopes.
 * @example
 * ```ts
 * import { createForm, ScopedFormSuite } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order: ScopedFormSuite<{ reference: string }> = createForm<{ reference: string }>();
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @remarks
 * These components and useField require the same suite’s Root. Scopes add relative prefixes while preserving RHF as the value/state owner. Presentation label/hint slots configure fixed inner span/div containers; Fluent owns Label/hint IDs, required/disabled metadata and validation precedence. Slot render callbacks execute during render and must forward their composed props/ref. Raw HTML and replacement native semantics are unsupported.
 * Toolkit-owned field change/blur handlers update RHF before the corresponding consumer notification; notification return values are ignored. Peer slot/child handlers retain their own ordering.
 */
export interface ScopedFormSuite<TInput extends FieldValues> {
    /**
     * String-valued Fluent Input and Field.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ reference: string }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ reference: '' }} onSubmit={values => { console.log(values); }}>
     *     <Order.Input name="reference" label="Reference" />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name, optional presentation and supported Fluent/native control props. Render inside the same suite’s Root; ref targets the HTMLInputElement.
     * @returns The Input control bound to its relative field; object/item scopes supply the path prefix.
     */
    Input: React.ForwardRefExoticComponent<FormInputProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /**
     * Multiline string-valued Fluent Textarea with native textarea ref.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ notes: string }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ notes: '' }} onSubmit={values => { console.log(values); }}>
     *     <Order.Textarea name="notes" label="Notes" resize="vertical" />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name, optional presentation and supported Fluent/native control props. Render inside the same suite’s Root; ref targets the HTMLTextAreaElement.
     * @returns The Textarea control bound to its relative field; object/item scopes supply the path prefix.
     */
    Textarea: React.ForwardRefExoticComponent<FormTextareaProps<TInput> & React.RefAttributes<HTMLTextAreaElement>>;
    /**
     * String search binding; clear commits an empty string before consumer notification.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ query: string }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ query: '' }} onSubmit={values => { console.log(values); }}>
     *     <Order.SearchBox name="query" label="Search" />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name, optional presentation and supported Fluent/native control props. Render inside the same suite’s Root; ref targets the HTMLInputElement.
     * @returns The SearchBox control bound to its relative field; object/item scopes supply the path prefix.
     */
    SearchBox: React.ForwardRefExoticComponent<FormSearchBoxProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /**
     * Boolean checkbox; explicit triState permits the mixed control value.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ approved: boolean }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ approved: false }} onSubmit={values => { console.log(values); }}>
     *     <Order.Checkbox name="approved" label="Approved" />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name, optional presentation and supported Fluent/native control props. Render inside the same suite’s Root; ref targets the HTMLInputElement.
     * @returns The Checkbox control bound to its relative field; object/item scopes supply the path prefix.
     */
    Checkbox: React.ForwardRefExoticComponent<FormCheckboxProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /**
     * Boolean switch with additive disabled state and a native input ref.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ enabled: boolean }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ enabled: false }} onSubmit={values => { console.log(values); }}>
     *     <Order.Switch name="enabled" label="Enabled" />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name, optional presentation and supported Fluent/native control props. Render inside the same suite’s Root; ref targets the HTMLInputElement.
     * @returns The Switch control bound to its relative field; object/item scopes supply the path prefix.
     */
    Switch: React.ForwardRefExoticComponent<FormSwitchProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /**
     * Numeric spin binding; empty/nonfinite manual commits use its nullable domain.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ quantity: number | null }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ quantity: 0 }} onSubmit={values => { console.log(values); }}>
     *     <Order.SpinButton name="quantity" label="Quantity" min={0} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name, optional presentation and supported Fluent/native control props. Render inside the same suite’s Root; ref targets the HTMLInputElement.
     * @returns The SpinButton control bound to its relative field; object/item scopes supply the path prefix.
     */
    SpinButton: React.ForwardRefExoticComponent<FormSpinButtonProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /**
     * Numeric slider using caller-supplied defaults and bounds.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ progress: number }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ progress: 20 }} onSubmit={values => { console.log(values); }}>
     *     <Order.Slider name="progress" label="Progress" min={20} max={40} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name, optional presentation and supported Fluent/native control props. Render inside the same suite’s Root; ref targets the HTMLInputElement.
     * @returns The Slider control bound to its relative field; object/item scopes supply the path prefix.
     */
    Slider: React.ForwardRefExoticComponent<FormSliderProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /**
     * Native select; option identity is independent of display text and optional explicit codec.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ country: string }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ country: '' }} onSubmit={values => { console.log(values); }}>
     *     <Order.Select name="country" options={[{ value: "pt", label: "Portugal" }]} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name and either data options or native option children. A different live domain requires an explicit string codec. Requires the same suite’s Root; ref targets the HTMLSelectElement.
     * @returns The Select control bound to its relative field; object/item scopes supply the path prefix.
     */
    Select: React.ForwardRefExoticComponent<FormSelectProps<TInput> & React.RefAttributes<HTMLSelectElement>>;
    /**
     * Named radio selection group with a truthful group ref and enabled-descendant focus.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ priority: string }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ priority: '' }} onSubmit={values => { console.log(values); }}>
     *     <Order.RadioGroup name="priority" label="Priority" options={[{ value: "high", label: "High" }]} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name and either data options or Fluent Radio children. A different live domain requires an explicit string codec. Requires the same suite’s Root; ref targets the HTMLDivElement group.
     * @returns The RadioGroup control bound to its relative field; object/item scopes supply the path prefix.
     */
    RadioGroup: React.ForwardRefExoticComponent<FormRadioGroupProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /**
     * Single string-identity dropdown, optionally mapped by an explicit domain codec.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ country: string }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ country: '' }} onSubmit={values => { console.log(values); }}>
     *     <Order.Dropdown name="country" label="Country" options={[{ value: "it", label: "Italy" }]} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name and either data options or Fluent Option children. A different live domain requires an explicit string codec. Requires the same suite’s Root; ref targets the HTMLButtonElement trigger.
     * @returns The Dropdown control bound to its relative field; object/item scopes supply the path prefix.
     */
    Dropdown: React.ForwardRefExoticComponent<FormDropdownProps<TInput> & React.RefAttributes<HTMLButtonElement>>;
    /**
     * Ordered multiple dropdown selection with optional explicit domain conversion.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ countries: string[] }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ countries: [] }} onSubmit={values => { console.log(values); }}>
     *     <Order.DropdownMulti name="countries" label="Countries" options={[{ value: "it", label: "Italy" }]} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name and either data options or Fluent Option children. A different live domain requires an explicit string-array codec. Requires the same suite’s Root; ref targets the HTMLButtonElement trigger.
     * @returns The DropdownMulti control bound to its relative field; object/item scopes supply the path prefix.
     */
    DropdownMulti: React.ForwardRefExoticComponent<FormDropdownMultiProps<TInput> & React.RefAttributes<HTMLButtonElement>>;
    /**
     * Ordered checkbox identities; adaptation is limited to documented Fluent Checkbox children.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ tags: string[] }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ tags: [] }} onSubmit={values => { console.log(values); }}>
     *     <Order.CheckboxGroup name="tags" label="Tags" options={[{ value: "work", label: "Work" }]} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name and either data options or direct Fluent Checkbox children. A different live domain requires an explicit string-array codec. Requires the same suite’s Root; ref targets the HTMLDivElement group.
     * @returns The CheckboxGroup control bound to its relative field; object/item scopes supply the path prefix.
     */
    CheckboxGroup: React.ForwardRefExoticComponent<FormCheckboxGroupProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /**
     * Editable single selection; unmatched query clears its committed identity.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ country: string }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ country: '' }} onSubmit={values => { console.log(values); }}>
     *     <Order.Combobox name="country" label="Country" options={[{ value: "it", label: "Italy" }]} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name and either data options or Fluent Option children. A different live domain requires an explicit string codec. Requires the same suite’s Root; ref targets the HTMLInputElement.
     * @returns The Combobox control bound to its relative field; object/item scopes supply the path prefix.
     */
    Combobox: React.ForwardRefExoticComponent<FormComboboxProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /**
     * Editable multiple selection; query edits retain committed selected identities.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ countries: string[] }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ countries: [] }} onSubmit={values => { console.log(values); }}>
     *     <Order.ComboboxMulti name="countries" label="Countries" options={[{ value: "it", label: "Italy" }]} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name and either data options or Fluent Option children. A different live domain requires an explicit string-array codec. Requires the same suite’s Root; ref targets the HTMLInputElement.
     * @returns The ComboboxMulti control bound to its relative field; object/item scopes supply the path prefix.
     */
    ComboboxMulti: React.ForwardRefExoticComponent<FormComboboxMultiProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /**
     * Selected identity tags with keyboard removal and caller-localized empty text.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ people: string[] }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ people: [] }} onSubmit={values => { console.log(values); }}>
     *     <Order.TagPicker name="people" label="People" options={[{ value: "ada", label: "Ada" }]} strings={{ noOptions: "No people" }} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name and either data options or Fluent TagPickerOption children. A different live domain requires an explicit string-array codec. Supply a plain string label or aria-label unless an enclosing SettingsGroup.Item field provides naming. Requires the same suite’s Root; ref targets the HTMLInputElement.
     * @returns The TagPicker control bound to its relative field; object/item scopes supply the path prefix.
     */
    TagPicker: React.ForwardRefExoticComponent<FormTagPickerProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /**
     * Numeric Fluent Rating with a real group ref and managed radio error focus.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ score: number }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ score: 1 }} onSubmit={values => { console.log(values); }}>
     *     <Order.Rating name="score" label="Score" max={5} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name, optional presentation and supported Fluent/native control props. Render inside the same suite’s Root; ref targets the HTMLDivElement group.
     * @returns The Rating control bound to its relative field; object/item scopes supply the path prefix.
     */
    Rating: React.ForwardRefExoticComponent<FormRatingProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /**
     * HSV color channels with optional alpha, copied before writing the live field.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { ColorSlider } from '@fluentui/react-components';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ accent: { h: number; s: number; v: number; a?: number } }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ accent: { h: 210, s: 0.7, v: 0.8 } }} onSubmit={values => { console.log(values); }}>
     *     <Order.ColorPicker name="accent" label="Accent"><ColorSlider channel="hue" /></Order.ColorPicker>
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name, optional presentation and supported Fluent/native control props. Render inside the same suite’s Root; ref targets the HTMLDivElement root.
     * @returns The ColorPicker control bound to its relative field; object/item scopes supply the path prefix.
     */
    ColorPicker: React.ForwardRefExoticComponent<FormColorPickerProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /**
     * Stable string swatch identity; the public ref remains on the real picker group.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { ColorSwatch } from '@fluentui/react-components';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ theme: string }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ theme: 'blue' }} onSubmit={values => { console.log(values); }}>
     *     <Order.SwatchPicker name="theme" label="Theme"><ColorSwatch value="blue" color="#0078d4" /></Order.SwatchPicker>
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name, optional presentation and supported Fluent/native control props. Render inside the same suite’s Root; ref targets the HTMLDivElement root.
     * @returns The SwatchPicker control bound to its relative field; object/item scopes supply the path prefix.
     */
    SwatchPicker: React.ForwardRefExoticComponent<FormSwatchPickerProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /**
     * Nullable local-calendar date with visible invalid drafts and a native input ref.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ date: Date | null }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ date: null }} onSubmit={values => { console.log(values); }}>
     *     <Order.DatePicker name="date" label="Date" allowTextInput />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name, optional presentation and supported Fluent/native control props. Render inside the same suite’s Root; ref targets the HTMLInputElement.
     * @returns The DatePicker control bound to its relative field; object/item scopes supply the path prefix.
     */
    DatePicker: React.ForwardRefExoticComponent<FormDatePickerProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /**
     * Nullable local wall-clock time using a deterministic anchor and native input ref.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ time: Date | null }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ time: null }} onSubmit={values => { console.log(values); }}>
     *     <Order.TimePicker name="time" label="Time" freeform hourCycle="h23" increment={15} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name, optional presentation and supported Fluent/native control props. Render inside the same suite’s Root; ref targets the HTMLInputElement.
     * @returns The TimePicker control bound to its relative field; object/item scopes supply the path prefix.
     */
    TimePicker: React.ForwardRefExoticComponent<FormTimePickerProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /**
     * Independent nullable date/time parts; merging never creates an implicit timestamp.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ appointment: { date: Date | null; time: Date | null } }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ appointment: { date: null, time: null } }} onSubmit={values => { console.log(values); }}>
     *     <Order.DateTime name="appointment" dateLabel="Date" timeLabel="Time" />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name, optional presentation and supported Fluent/native control props. Render inside the same suite’s Root; ref targets the HTMLDivElement group.
     * @returns The DateTime control bound to its relative field; object/item scopes supply the path prefix.
     */
    DateTime: React.ForwardRefExoticComponent<FormDateTimeProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /**
     * Actual browser FileList selection; clear/reset never fabricates native file restoration.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ attachments: FileList | null }>();
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ attachments: null }} onSubmit={values => { console.log(values); }}>
     *     <Order.FileInput name="attachments" label="Attachments" accept="application/pdf" multiple />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required compatible field name, optional presentation and supported Fluent/native control props. Render inside the same suite’s Root; ref targets the HTMLInputElement.
     * @returns The FileInput control bound to its relative field; object/item scopes supply the path prefix.
     */
    FileInput: React.ForwardRefExoticComponent<FormFileInputProps<TInput> & React.RefAttributes<HTMLInputElement>>;
    /**
     * Headless nested-record scope with path-specific relative field inference.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm, FormObjectProps } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ customer: { name: string } }>();
     * const props: FormObjectProps<{ customer: { name: string } }, 'customer'> = {
     *   name: 'customer', children: ({ Fields }) => <Fields.Input name="name" />,
     * };
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ customer: { name: '' } }} onSubmit={values => { console.log(values); }}>
     *     <Order.Object {...props} />
     *   </Order.Root>;
     * }
     * ```
     * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
     * @param props - Required relative record name and children render callback. Render inside the same suite’s Root; this headless scope accepts no native attributes or ref.
     * @returns The children render result under a relative record scope, without adding layout DOM or forwarding a ref.
     */
    Object: FormObjectComponent<TInput>;
    /**
     * Headless structured-record array with stable item keys and explicit operations.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm, FormArrayProps } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ rows: { sku: string }[] }>();
     * const props: FormArrayProps<{ rows: { sku: string }[] }, 'rows'> = {
     *   name: 'rows', children: ({ items }) => items.map(({ key, Fields }) => <Fields.Input key={key} name="sku" />),
     * };
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ rows: [] }} onSubmit={values => { console.log(values); }}>
     *     <Order.Array {...props} />
     *   </Order.Root>;
     * }
     * ```
     * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
     * @param props - Required relative structured-array name and children render callback, plus optional array rules when no schema/resolver owns validation. Requires the same suite’s Root; no native attributes or ref are accepted.
     * @returns The children render result with stable item metadata and array operations, without adding layout DOM or forwarding a ref.
     */
    Array: FormArrayComponent<TInput>;
    /**
     * Named visual nested-record fieldset with direct errors and inherited disabled state.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm, FormObjectSectionProps } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ customer: { name: string } }>();
     * const props = {
     *   name: 'customer', title: 'Customer', children: ({ Fields }) => <Fields.Input name="name" />,
     * } satisfies FormObjectSectionProps<{ customer: { name: string } }, 'customer'>;
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ customer: { name: '' } }} onSubmit={values => { console.log(values); }}>
     *     <Order.ObjectSection {...props} />
     *   </Order.Root>;
     * }
     * ```
     * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
     * @param props - Required relative record name, children renderer and accessible section name through title, aria-label or aria-labelledby. Optional props configure the fieldset/surface; requires the same suite’s Root and does not accept a ref.
     * @returns A named native fieldset containing a Fluent Card, description, direct errors and caller-rendered relative fields; no ref is forwarded.
     */
    ObjectSection: FormObjectSectionComponent<TInput>;
    /**
     * Named ordered editable array with localized actions, bounds and action-owned focus.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm, FormArraySectionProps } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ rows: { sku: string }[] }>();
     * const props = {
     *   name: 'rows', title: 'Rows', createItem: () => ({ sku: '' }), renderItem: ({ Fields }) => <Fields.Input name="sku" />,
     * } satisfies FormArraySectionProps<{ rows: { sku: string }[] }, 'rows'>;
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ rows: [] }} onSubmit={values => { console.log(values); }}>
     *     <Order.ArraySection {...props} />
     *   </Order.Root>;
     * }
     * ```
     * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
     * @param props - Required relative structured-array name, createItem, renderItem and accessible section name through title, aria-label or aria-labelledby. Optional props configure UI bounds/actions and the fieldset/surface; requires the same suite’s Root and does not accept a ref.
     * @returns A named native fieldset containing a Fluent Card, ordered items, localized actions and announcements; no ref is forwarded.
     */
    ArraySection: FormArraySectionComponent<TInput>;
    /**
     * Field or global native alert.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm, FormErrorProps } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ reference: string }>();
     * const props: FormErrorProps<{ reference: string }> = { name: 'reference' };
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ reference: '' }} onSubmit={values => { console.log(values); }}>
     *     <Order.Error {...props} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Optional relative field/object/array name and native alert div attributes; omit name for global errors. Requires the same suite’s Root; ref targets the HTMLDivElement when an alert is rendered.
     * @returns A native alert div containing deduplicated messages, or no DOM element when there are no messages. Object/item scopes supply the prefix for an addressed name.
     */
    Error: React.ForwardRefExoticComponent<FormErrorProps<TInput> & React.RefAttributes<HTMLDivElement>>;
    /**
     * React hook binding a custom control to live RHF value/state inside its owning Root/scope.
     * Pass a compatible relative name and wire onChange/onBlur/ref; returned state remains owned by RHF.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm, FormFieldResult } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ reference: string }>();
     * function CustomReference(): React.ReactElement {
     *   const field: FormFieldResult<string> = Order.useField<string>({ name: 'reference' });
     *   return <input ref={field.ref} value={field.value} disabled={field.disabled} onBlur={field.onBlur}
     *     onChange={event => field.onChange(event.currentTarget.value)} aria-label="Reference" />;
     * }
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ reference: '' }} onSubmit={values => { console.log(values.reference); }}>
     *     <CustomReference />
     *   </Order.Root>;
     * }
     * ```
     * @typeParam TValue - Complete domain read and written by the custom control.
     * @param props - Required compatible relative field path, optional domain rules and additive disabled restriction; requires this suite’s Root.
     * @returns Live RHF value, writers, native focus ref, validation state and transient-reset generation.
     */
    useField: <TValue>(props: FormUseFieldProps<TInput, TValue>) => FormFieldResult<TValue>;
}
/**
 * Typed form family. Create factories outside render to keep suite ownership stable.
 * @example
 * ```ts
 * import { createForm, FormSuite } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order: FormSuite<{ reference: string }> = createForm<{ reference: string }>();
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TContext - Context passed to the React Hook Form resolver; defaults to unknown where optional.
 * @typeParam TOutput - Validated submission model, which can differ from live input; defaults to TInput where optional.
 * @typeParam TBound - Whether a factory-bound validator forbids replacement at Root/useForm; defaults to false.
 */
export interface FormSuite<TInput extends FieldValues, TContext = unknown, TOutput = TInput, TBound extends boolean = false> extends ScopedFormSuite<TInput> {
    /**
     * Native form and RHF provider.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm, FormRootProps } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ reference: string }>();
     * const rootProps: FormRootProps<{ reference: string }> = {
     *   defaultValues: { reference: '' }, onSubmit: values => { console.log(values.reference); },
     * };
     * export function Editor(): React.ReactElement {
     *   return <Order.Root {...rootProps}>
     *     <Order.Input name="reference" />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Required onSubmit plus either RHF initialization options or this suite’s methods; optional children and native form attributes. Ref targets the HTMLFormElement.
     * @returns A native form wrapped in RHF and toolkit providers for this suite’s descendants. Root establishes the form context and does not require an enclosing Form.Root.
     */
    Root: React.ForwardRefExoticComponent<FormRootProps<TInput, TContext, TOutput, TBound> & React.RefAttributes<HTMLFormElement>>;
    /**
     * Fluent submit button, disabled while RHF is submitting.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm, FormSubmitProps } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ reference: string }>();
     * const props: FormSubmitProps = { children: 'Save' };
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ reference: '' }} onSubmit={values => { console.log(values); }}>
     *     <Order.Submit {...props} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Optional Fluent button content, attributes, callbacks and an additive disabled restriction; no field path is required. Ref targets the HTMLButtonElement.
     * @returns A native submit button inside the same suite’s Root. Root/methods disabled state and pending submission keep the button disabled; activating it follows the Root’s native submit gate.
     */
    Submit: React.ForwardRefExoticComponent<FormSubmitProps & React.RefAttributes<HTMLButtonElement>>;
    /**
     * Fluent reset button, disabled while RHF is submitting.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm, FormResetProps } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ reference: string }>();
     * const props: FormResetProps = { children: 'Discard changes' };
     * export function Editor(): React.ReactElement {
     *   return <Order.Root defaultValues={{ reference: '' }} onSubmit={values => { console.log(values); }}>
     *     <Order.Reset {...props} />
     *   </Order.Root>;
     * }
     * ```
     * @param props - Optional Fluent button content, attributes, callbacks and an additive disabled restriction; no field path is required. Ref targets the HTMLButtonElement.
     * @returns A native reset button inside the same suite’s Root. Root/methods disabled state and pending submission keep it disabled; a permitted native reset reaches Root’s cancellable reset handler.
     */
    Reset: React.ForwardRefExoticComponent<FormResetProps & React.RefAttributes<HTMLButtonElement>>;
    /**
     * React hook creating same-suite RHF methods with optional initialization.
     * Omission uses RHF defaults; factory-bound validation cannot be replaced.
     * Direct handleSubmit bypasses Root's DOM gate; pass these methods to Root instead of initialization props.
     * @example
     * ```tsx
     * import * as React from 'react';
     * import { createForm, FormMethods } from '@apvee/spfx-react-toolkit/controls/form';
     * const Order = createForm<{ reference: string }>();
     * export function Editor(): React.ReactElement {
     *   const methods: FormMethods<{ reference: string }> = Order.useForm({ defaultValues: { reference: '' } });
     *   return <Order.Root methods={methods} onSubmit={values => { console.log(values.reference); }}>
     *     <Order.Input name="reference" />
     *   </Order.Root>;
     * }
     * ```
     * @param props - RHF initialization and optional UI dictionary; when omitted, RHF defaults and native SPFx strings apply.
     * @returns Same-suite enhanced RHF methods with reset coordination; pass them to Root instead of initialization props.
     */
    useForm: (props?: FormUseFormProps<TInput, TContext, TOutput, TBound>) => FormMethods<TInput, TContext, TOutput>;
}
/**
 * Untyped family for dynamic domain names.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { Form, UntypedFormSuite } from '@apvee/spfx-react-toolkit/controls/form';
 * const DynamicForm: UntypedFormSuite = Form;
 * export function Editor(): React.ReactElement {
 *   return <DynamicForm.Root defaultValues={{ runtime: { field: '' } }} onSubmit={values => { console.log(values); }}>
 *     <DynamicForm.Input name="runtime.field" label="Dynamic field" />
 *   </DynamicForm.Root>;
 * }
 * ```
 */
export type UntypedFormSuite = FormSuite<FieldValues>;
/**
 * Typed factory overloads preserving distinct live input and submitted output.
 * Create suites outside render. Bound owners cannot be replaced at Root or useForm.
 * @example
 * ```ts
 * import { createForm, CreateForm } from '@apvee/spfx-react-toolkit/controls/form';
 * const factory: CreateForm = createForm;
 * ```
 * @returns A stable typed suite; its validation owner and method identity are retained for Root and field bindings.
 */
export interface CreateForm {
    /**
     * Creates a rules-only or Root-configured family with this live model.
     * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
     * @returns A stable suite whose methods, fields and Root share one private ownership identity.
     */
    <TInput extends FieldValues>(): FormSuite<TInput>;
    /**
     * Infers the live input and transformed output from a Standard Schema.
     * @typeParam TSchema - Standard Schema v1 whose input and output infer the suite models.
     * @param schema - Standard Schema v1 validator used to infer and permanently bind live input and submit output.
     * @returns A stable suite whose methods, fields and Root share one private ownership identity.
     */
    <TSchema extends StandardSchemaV1>(schema: TSchema & (StandardSchemaV1.InferInput<TSchema> extends FieldValues ? unknown : never)): FormSuite<Extract<StandardSchemaV1.InferInput<TSchema>, FieldValues>, unknown, StandardSchemaV1.InferOutput<TSchema>, true>;
    /**
     * Permanently binds one schema or resolver owner.
     * @param options - Exactly one validation owner; its input stays live and its output reaches onSubmit.
     * @returns A stable family with validation ownership fixed for Root and useForm.
     * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
     * @typeParam TOutput - Validated submission model, which can differ from live input; defaults to TInput where optional.
     * @typeParam TContext - Context passed to the React Hook Form resolver; defaults to unknown where optional.
     */
    <TInput extends FieldValues, TOutput = TInput, TContext = unknown>(options: {
        /** Standard Schema v1 owner mapping live input to validated output. */
        schema: StandardSchemaV1<TInput, TOutput>;
        /**
         * Excluded when schema owns validation.
         * Omit this member for the other validation-owner branches; it cannot replace a factory-bound owner.
         */
        resolver?: never;
    } | {
        /** RHF resolver owner receiving live input and the configured context. */
        resolver: Resolver<TInput, TContext, TOutput>;
        /**
         * Excluded when resolver owns validation.
         * Omit this member for the other validation-owner branches; it cannot replace a factory-bound owner.
         */
        schema?: never;
    }): FormSuite<TInput, TContext, TOutput, true>;
}

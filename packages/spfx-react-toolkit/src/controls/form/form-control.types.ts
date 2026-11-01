/**
 * Typed Form families, RHF value ownership and native Fluent bindings.
 *
 * @packageDocumentation
 * @docPage controls/form
 */
import type * as React from 'react';
import type { ContentSlot } from '../content-slot.types.internal';
import type { ColorPickerProps, RatingProps, SwatchPickerProps, CheckboxProps, ComboboxProps, DropdownProps, FieldProps, InputProps, RadioGroupProps, SelectProps, TagPickerProps, TagPickerInputProps, SearchBoxProps, SliderProps, SpinButtonProps, SwitchProps, TextareaProps } from '@fluentui/react-components';
import type { FormDateTimeValue, FormHsvColor, FormOptionsSource, FormValueCodec } from './form-value.types';
import type { FieldPath, FieldPathValue, FieldValues, RegisterOptions } from 'react-hook-form';
/**
 * Paths which can both supply the control's value and accept its complete write domain.
 * Optional fields may initialize absent; controlled inputs require suitable defaults.
 * @example
 * ```ts
 * import { CompatibleFieldPath } from '@apvee/spfx-react-toolkit/controls/form';
 * type StringPath = CompatibleFieldPath<{ title: string; count: number }, string>;
 * const name: StringPath = 'title';
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TValue - Complete domain read and written by the custom control.
 */
export type CompatibleFieldPath<TInput extends FieldValues, TValue> = {
    [TName in FieldPath<TInput>]: [
        Exclude<FieldPathValue<TInput, TName>, undefined>
    ] extends [
        TValue
    ] ? [
        TValue
    ] extends [
        FieldPathValue<TInput, TName>
    ] ? TName : never : never;
}[FieldPath<TInput>];
/**
 * Presentation shared by Fluent Field wrappers; required is metadata, not a validation rule.
 * Slots configure inner containers; native Label/hint IDs and associations remain owned by Field.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormFieldPresentationProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ reference: string }>();
 * const presentation: FormFieldPresentationProps = {
 *   label: { children: 'Reference', className: 'reference-label' },
 *   hint: { children: (Component, props) => <Component {...props}>Help</Component> },
 * };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ reference: '' }} onSubmit={values => { console.log(values); }}>
 *     <Order.Input name="reference" {...presentation} />
 *   </Order.Root>;
 * }
 * ```
 */
export interface FormFieldPresentationProps {
    /** Literal phrasing content inside the native Fluent Label. Objects/callbacks configure an inner span,
     * including its ref; htmlFor, required/disabled and generated Label ID remain native.
     * Only undefined omits the top-level label; legacy null/booleans still request the native Label.
     * Omit inside SettingsGroup.Item field.
     */
    label?: ContentSlot<'span'>;
    /** Supporting content inside the native Field hint. Objects/callbacks configure an inner div;
     * its ref/ID do not replace the associated native hint ID. Undefined retains fieldProps.hint.
     * Omit inside SettingsGroup.Item field.
     */
    hint?: ContentSlot<'div'>;
    /**
     * Marks the control as required visually and accessibly.
     * When omitted, no required presentation metadata is added; supply validation rules separately.
     * @defaultValue false
     */
    required?: boolean;
    /**
     * Fluent Field label layout; omission uses Fluent's vertical layout.
     * When omitted, Fluent Field uses vertical label layout.
     * @defaultValue vertical
     */
    orientation?: FieldProps['orientation'];
    /**
     * Disables the field; cannot override a disabled form or scope.
     * When omitted or false, adds no disabled restriction; enclosing Root/scope restrictions still apply.
     */
    disabled?: boolean;
    /**
     * Class applied to the outer Fluent Field.
     * When omitted, only the component and Fluent classes apply.
     */
    className?: string;
    /**
     * Additional Fluent Field props; active RHF errors take precedence.
     * When omitted, Fluent Field uses the adapter-owned naming, hint and validation props.
     */
    fieldProps?: Omit<FieldProps, 'children' | 'className' | 'disabled' | 'label' | 'orientation' | 'required'>;
}
/**
 * Binding props for a control that reads and writes one complete value domain.
 * @example
 * ```ts
 * import { FormControlProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const props: FormControlProps<{ title: string }, string> = {
 *   name: 'title', label: 'Title', rules: { required: 'Enter a title.' },
 * };
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TValue - Complete domain read and written by the custom control.
 * @typeParam TName - Compatible relative field path; defaults to all paths accepting the complete value domain.
 */
export type FormControlProps<TInput extends FieldValues, TValue, TName extends CompatibleFieldPath<TInput, TValue> = CompatibleFieldPath<TInput, TValue>> = FormFieldPresentationProps & {
    /** Relative field path; object/array scopes supply its prefix. */
    name: TName;
    /**
     * Domain validation; disallowed when a schema/resolver owns validation.
     * When omitted, no field or array domain rules are registered here; schema/resolver ownership remains unchanged.
     */
    rules?: RegisterOptions<TInput, TName>;
};
/**
 * Native text input behaviors supported by Input.
 * @example
 * ```ts
 * import { FormInputType } from '@apvee/spfx-react-toolkit/controls/form';
 * const type: FormInputType = 'email';
 * ```
 */
export type FormInputType = 'text' | 'email' | 'password' | 'search' | 'tel' | 'url';
/**
 * String binding for Fluent Input; it deliberately has no value codec.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormInputProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ reference: string }>();
 * const fieldProps: FormInputProps<{ reference: string }> = { name: 'reference' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ reference: '' }} onSubmit={values => { console.log(values); }}>
 *     <Order.Input {...fieldProps} label="Reference" />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormInputProps<TInput extends FieldValues> = FormControlProps<TInput, string> & Omit<InputProps, keyof FormFieldPresentationProps | 'defaultValue' | 'name' | 'value' | 'type'> & {
    /**
     * Text input behavior; defaults to native text. Numbers, dates and files use dedicated controls.
     * When omitted, Input uses native text behavior.
     * @defaultValue text
     */
    type?: FormInputType;
};

/**
 * Multiline string binding for Fluent Textarea; RHF owns value and defaults.
 * Native change/blur callbacks run after RHF and the ref targets the textarea.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormTextareaProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ notes: string }>();
 * const fieldProps: FormTextareaProps<{ notes: string }> = { name: 'notes' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ notes: '' }} onSubmit={values => { console.log(values); }}>
 *     <Order.Textarea {...fieldProps} label="Notes" resize="vertical" />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormTextareaProps<TInput extends FieldValues> = FormControlProps<TInput, string> & Omit<TextareaProps, keyof FormFieldPresentationProps | 'defaultValue' | 'name' | 'value'>;
/**
 * Search string binding; dismissing writes an empty string before `onChange`.
 * RHF owns value/defaults and the ref targets the native input.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormSearchBoxProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ query: string }>();
 * const fieldProps: FormSearchBoxProps<{ query: string }> = { name: 'query' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ query: '' }} onSubmit={values => { console.log(values); }}>
 *     <Order.SearchBox {...fieldProps} label="Search" />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormSearchBoxProps<TInput extends FieldValues> = FormControlProps<TInput, string> & Omit<SearchBoxProps, keyof FormFieldPresentationProps | 'defaultValue' | 'name' | 'value'>;
/**
 * Checkbox binding with explicit opt-in to Fluent's mixed state.
 * A boolean field accepts the default mode. A `boolean | 'mixed'` field requires
 * `triState: true`. Change/blur callbacks run after RHF; the ref targets the input.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormCheckboxProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ approved: boolean }>();
 * const fieldProps: FormCheckboxProps<{ approved: boolean }> = { name: 'approved' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ approved: false }} onSubmit={values => { console.log(values); }}>
 *     <Order.Checkbox {...fieldProps} label="Approved" />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormCheckboxProps<TInput extends FieldValues> = Omit<CheckboxProps, keyof FormFieldPresentationProps | 'checked' | 'defaultChecked' | 'name' | 'value'> & (
    (FormControlProps<TInput, boolean> & {
        /**
         * Keeps the complete live domain boolean when omitted or false.
         * @defaultValue false
         */
        triState?: false;
    }) | (FormControlProps<TInput, boolean | 'mixed'> & {
        /**
         * Explicit true selects mixed-state mode; the field must accept both booleans and Fluent’s mixed value. This required discriminant has no omission default.
         */
        triState: true;
    })
);
/**
 * Boolean switch binding; RHF owns checked/defaults and the ref targets the input.
 * Change/blur callbacks run after RHF; required is presentation metadata.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormSwitchProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ enabled: boolean }>();
 * const fieldProps: FormSwitchProps<{ enabled: boolean }> = { name: 'enabled' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ enabled: false }} onSubmit={values => { console.log(values); }}>
 *     <Order.Switch {...fieldProps} label="Enabled" />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormSwitchProps<TInput extends FieldValues> = FormControlProps<TInput, boolean> & Omit<SwitchProps, keyof FormFieldPresentationProps | 'checked' | 'defaultChecked' | 'name' | 'value'>;
/**
 * Numeric commit binding for Fluent SpinButton's nullable external domain.
 * Finite display commits become numbers; empty/nonfinite text becomes `null`.
 * Validation belongs to schema/rules. RHF owns value/defaults; the ref targets the
 * input and native change/blur callbacks run after RHF. No text-format codec exists.
 * A non-preserving targeted reset replaces only the Fluent primitive to clear its
 * internal draft; the native ref changes and existing input focus is restored.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormSpinButtonProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ quantity: number | null }>();
 * const fieldProps: FormSpinButtonProps<{ quantity: number | null }> = { name: 'quantity' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ quantity: 0 }} onSubmit={values => { console.log(values); }}>
 *     <Order.SpinButton {...fieldProps} label="Quantity" min={0} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
// Fluent SpinButton explicitly represents an empty committed value with null.
// eslint-disable-next-line @rushstack/no-new-null
export type FormSpinButtonProps<TInput extends FieldValues> = FormControlProps<TInput, number | null> & Omit<SpinButtonProps, keyof FormFieldPresentationProps | 'defaultValue' | 'displayValue' | 'name' | 'value'>;
/**
 * Numeric slider binding; supply a valid RHF default and matching min/max/step.
 * No adapter fallback changes the caller's range or value. Native change/blur
 * callbacks run after RHF; the ref targets the range input.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormSliderProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ progress: number }>();
 * const fieldProps: FormSliderProps<{ progress: number }> = { name: 'progress' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ progress: 20 }} onSubmit={values => { console.log(values); }}>
 *     <Order.Slider {...fieldProps} label="Progress" min={20} max={40} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormSliderProps<TInput extends FieldValues> = FormControlProps<TInput, number> & Omit<SliderProps, keyof FormFieldPresentationProps | 'defaultValue' | 'name' | 'value'>;

/**
 * Selection bindings require an explicit lossless codec for a different live field domain.
 * Each mapped member correlates `name` with that path's exact read/write codec type.
 *
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TValue - Complete domain read and written by the custom control.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormSelectProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Profile = createForm<{ countryCode: number | undefined }>();
 * const countryProps = {
 *   name: 'countryCode', label: 'Country',
 *   codec: {
 *     toControl: (value: number | undefined) => value === undefined ? '' : String(value),
 *     fromControl: (value: string) => value === '' ? undefined : Number(value),
 *   },
 *   options: [{ value: '39', label: 'Italy' }],
 * } satisfies FormSelectProps<{ countryCode: number | undefined }>;
 * export function Editor(): React.ReactElement {
 *   return <Profile.Root defaultValues={{ countryCode: 39 }} onSubmit={values => { console.log(values.countryCode); }}>
 *     <Profile.Select {...countryProps} /><Profile.Submit>Save</Profile.Submit>
 *   </Profile.Root>;
 * }
 * ```
 */
type FormSelectionBinding<TInput extends FieldValues, TValue> =
    | (FormControlProps<TInput, TValue> & {
        /** Omit for the complete native string/string-array domain. */
        codec?: never })
    | { [TName in FieldPath<TInput>]: FormFieldPresentationProps & {
        /** Relative path whose exact live value is converted by this codec. */
        name: TName;
        /**
         * Domain validation when rules own validation.
         * When omitted, no field or array domain rules are registered here; schema/resolver ownership remains unchanged.
         */
        rules?: RegisterOptions<TInput, TName>;
        /** Lossless conversion; every writable control identity must round-trip. */
        codec: FormValueCodec<FieldPathValue<TInput, TName>, TValue>;
    } }[FieldPath<TInput>];
/**
 * Selection adapter props combining field binding, one option source, and unowned Fluent props.
 * @example
 * ```tsx
 * import { FormSelectionProps } from '@apvee/spfx-react-toolkit/controls/form';
 * <Order.Dropdown name="country" label="Country" options={[{ value: 'it', label: 'Italy' }]} />
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TValue - Complete domain read and written by the custom control.
 * @typeParam TPrimitive - Native Fluent primitive props forwarded after adapter-owned props are removed.
 * @typeParam TOwned - Primitive keys owned by the Form adapter and unavailable to consumers.
 */
type FormSelectionProps<TInput extends FieldValues, TValue, TPrimitive, TOwned extends PropertyKey> = FormSelectionBinding<TInput, TValue> & FormOptionsSource & Omit<TPrimitive, TOwned | keyof FormFieldPresentationProps>;
/**
 * Native Select for one stable string identity, or an explicitly converted domain.
 * Supply native `option` children or data options. RHF owns value/defaults.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormSelectProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ country: string }>();
 * const fieldProps = { name: 'country', options: [{ value: 'it', label: 'Italy' }] } satisfies FormSelectProps<{ country: string }>;
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ country: '' }} onSubmit={values => { console.log(values); }}>
 *     <Order.Select {...fieldProps} options={[{ value: "pt", label: "Portugal" }]} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormSelectProps<TInput extends FieldValues> = FormSelectionProps<TInput, string, SelectProps, 'children' | 'defaultValue' | 'name' | 'value' | 'multiple'>;
/**
 * Fluent RadioGroup with a public div ref and descendant error focus.
 * Supply Fluent Radio children or data options. Group callbacks follow RHF.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormRadioGroupProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ priority: string }>();
 * const fieldProps = { name: 'priority', options: [{ value: 'it', label: 'Italy' }] } satisfies FormRadioGroupProps<{ priority: string }>;
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ priority: '' }} onSubmit={values => { console.log(values); }}>
 *     <Order.RadioGroup {...fieldProps} label="Priority" options={[{ value: "high", label: "High" }]} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormRadioGroupProps<TInput extends FieldValues> = FormSelectionProps<TInput, string, RadioGroupProps, 'children' | 'defaultValue' | 'name' | 'value'>;
/**
 * Single Dropdown stores identity and displays independent option text.
 * The ref targets the trigger button; `Option` children or options are exclusive.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormDropdownProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ country: string }>();
 * const fieldProps = { name: 'country', options: [{ value: 'it', label: 'Italy' }] } satisfies FormDropdownProps<{ country: string }>;
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ country: '' }} onSubmit={values => { console.log(values); }}>
 *     <Order.Dropdown {...fieldProps} label="Country" options={[{ value: "it", label: "Italy" }]} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormDropdownProps<TInput extends FieldValues> = FormSelectionProps<TInput, string, DropdownProps, 'children' | 'defaultSelectedOptions' | 'defaultValue' | 'multiselect' | 'name' | 'selectedOptions' | 'value'>;
/**
 * Multiple Dropdown stores an ordered string array; the ref targets the button.
 * Native selection order is retained. Explicit codecs convert the complete array.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormDropdownMultiProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ countries: string[] }>();
 * const fieldProps = { name: 'countries', options: [{ value: 'it', label: 'Italy' }] } satisfies FormDropdownMultiProps<{ countries: string[] }>;
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ countries: [] }} onSubmit={values => { console.log(values); }}>
 *     <Order.DropdownMulti {...fieldProps} label="Countries" options={[{ value: "it", label: "Italy" }]} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormDropdownMultiProps<TInput extends FieldValues> = FormSelectionProps<TInput, string[], DropdownProps, 'children' | 'defaultSelectedOptions' | 'defaultValue' | 'multiselect' | 'name' | 'selectedOptions' | 'value' | 'clearable'>;
/**
 * Single Combobox separates query presentation from selected identity.
 * Unmatched queries write the empty identity; schema/rules decide its validity.
 * Reset clears drafts; explicit `keepDirtyValues` preserves dirty field drafts.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormComboboxProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ country: string }>();
 * const fieldProps = { name: 'country', options: [{ value: 'it', label: 'Italy' }] } satisfies FormComboboxProps<{ country: string }>;
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ country: '' }} onSubmit={values => { console.log(values); }}>
 *     <Order.Combobox {...fieldProps} label="Country" options={[{ value: "it", label: "Italy" }]} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormComboboxProps<TInput extends FieldValues> = FormSelectionProps<TInput, string, ComboboxProps, 'children' | 'defaultSelectedOptions' | 'defaultValue' | 'multiselect' | 'name' | 'selectedOptions' | 'value' | 'freeform'>;
/**
 * Multiple Combobox keeps query edits separate from its ordered selected identities.
 * Selection clears query. Reset clears drafts unless their dirty owner is preserved.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormComboboxMultiProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ countries: string[] }>();
 * const fieldProps = { name: 'countries', options: [{ value: 'it', label: 'Italy' }] } satisfies FormComboboxMultiProps<{ countries: string[] }>;
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ countries: [] }} onSubmit={values => { console.log(values); }}>
 *     <Order.ComboboxMulti {...fieldProps} label="Countries" options={[{ value: "it", label: "Italy" }]} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormComboboxMultiProps<TInput extends FieldValues> = FormSelectionProps<TInput, string[], ComboboxProps, 'children' | 'defaultSelectedOptions' | 'defaultValue' | 'multiselect' | 'name' | 'selectedOptions' | 'value' | 'freeform' | 'clearable'>;
/**
 * Checkbox group for string identities in declared option order.
 * Direct Fluent Checkbox children are the only children adapted; other content is retained.
 * The public div ref is separate from RHF focus on the first eligible checkbox.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormCheckboxGroupProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ tags: string[] }>();
 * const fieldProps = { name: 'tags', options: [{ value: 'it', label: 'Italy' }] } satisfies FormCheckboxGroupProps<{ tags: string[] }>;
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ tags: [] }} onSubmit={values => { console.log(values); }}>
 *     <Order.CheckboxGroup {...fieldProps} label="Tags" options={[{ value: "work", label: "Work" }]} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormCheckboxGroupProps<TInput extends FieldValues> = FormSelectionBinding<TInput, string[]> & FormOptionsSource & Omit<React.HTMLAttributes<HTMLDivElement>, 'children' | 'onChange' | 'onBlur' | 'className'> & {
    /**
     * Runs after RHF writes the group selection, before the child callbacks.
     * When omitted, RHF updates and internal presentation still run without a consumer notification. Returned promises are not awaited.
     */
    onChange?: CheckboxProps['onChange'];
    /**
     * Runs after RHF marks the field touched, before the child's blur callback.
     * When omitted, RHF still marks the field touched without a consumer notification. Returned promises are not awaited.
     */
    onBlur?: CheckboxProps['onBlur'];
};
/**
 * Localized TagPicker query results.
 * @example
 * ```ts
 * import { FormTagPickerStrings } from '@apvee/spfx-react-toolkit/controls/form';
 * const strings: FormTagPickerStrings = { noOptions: 'Nessun risultato' };
 * ```
 */
export interface FormTagPickerStrings {
    /** Literal empty-result content inside the owned status region; objects/callbacks configure an inner div and its ref. */
    noOptions: ContentSlot<'div'>;
}
/**
 * TagPicker naming alternatives; rich labels require explicit text unless a settings field row supplies it.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormTagPickerProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ people: string[] }>();
 * const peopleProps = {
 *   name: 'people', label: <strong>People</strong>, 'aria-label': 'Selected people',
 *   options: [{ value: 'ada', label: 'Ada Lovelace' }],
 * } satisfies FormTagPickerProps<{ people: string[] }>;
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ people: [] }} onSubmit={values => { console.log(values.people); }}>
 *     <Order.TagPicker {...peopleProps} /><Order.Submit>Save</Order.Submit>
 *   </Order.Root>;
 * }
 * ```
 */
type FormTagPickerAccessibleName = {
    /** Plain text visible label also names the selected-tag list. */
    label: string;
    /**
     * Explicit text overriding the label-derived accessible name when supplied.
     * When omitted, the required plain string label names the native input and selected-tag list. An enclosing SettingsGroup.Item field can instead supply their aria-labelledby relationships.
     */
    'aria-label'?: string } | {
    /**
     * Rich/object/callback or absent label needs an explicit accessible name; no text is inferred.
     * When omitted, no top-level Label is requested; inherited settings-row naming can still apply.
     */
    label?: ContentSlot<'span'>;
    /** Required text name when standalone label content is not a string. */
    'aria-label': string } | {
    /** Omit label when SettingsGroup.Item field supplies inherited naming. Standalone usage is checked at runtime. */
    label?: undefined;
    /** Omit with label only when inherited settings-row naming is available. */
    'aria-label'?: undefined };
/**
 * TagPicker with removable tags, string-array identities and a native input ref.
 * Options or Fluent TagPickerOption children provide identities and display text.
 * Omit label/aria-label only when SettingsGroup.Item field supplies inherited naming.
 * Ancestor context cannot be proven by TypeScript; unnamed standalone usage throws at runtime.
 * Query edits retain selections; reset and selection clear query presentation.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormTagPickerProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ people: string[] }>();
 * const fieldProps = { name: 'people', options: [{ value: 'it', label: 'Italy' }], label: 'People' } satisfies FormTagPickerProps<{ people: string[] }>;
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ people: [] }} onSubmit={values => { console.log(values); }}>
 *     <Order.TagPicker {...fieldProps} label="People" options={[{ value: "ada", label: "Ada" }]} strings={{ noOptions: "No people" }} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormTagPickerProps<TInput extends FieldValues> = FormSelectionProps<TInput, string[], TagPickerProps, 'aria-label' | 'children' | 'defaultSelectedOptions' | 'defaultValue' | 'name' | 'noPopover' | 'selectedOptions' | 'value' | 'onChange' | 'onBlur' | 'id'> & FormTagPickerAccessibleName & {
    /**
     * Add a Fluent colorful initials/fallback avatar to each selected tag. Defaults to false, retaining text-only tags.
     * Colors use the string option identity and remain stable when display labels change. Rich consumer labels are preserved.
     * Only presentation changes; form values, query and selection callbacks are unaffected.
     * Example: `<Order.TagPicker name="people" label="People" options={people} colorfulTagAvatars />`
     * When omitted, tags keep the native text-only presentation.
     * @defaultValue false
     */
    colorfulTagAvatars?: boolean;
    /**
     * Additional descriptions for the native query input.
     * When omitted, only adapter-owned hint/error relationships apply.
     */
    'aria-describedby'?: string;
    /**
     * Input id shared with Fluent Field's label.
     * When omitted, the binding uses generated native Field/control identity.
     */
    id?: string;
    /**
     * Query callback after presentation updates; it does not change selected values.
     * When omitted, RHF updates and internal presentation still run without a consumer notification. Returned promises are not awaited.
     */
    onChange?: TagPickerInputProps['onChange'];
    /**
     * Runs after RHF marks the field touched.
     * When omitted, RHF still marks the field touched without a consumer notification. Returned promises are not awaited.
     */
    onBlur?: TagPickerInputProps['onBlur'];
    /**
     * Inner div content for the owned no-options status. Objects support attributes, refs and children callbacks.
     * Nullish values fall back to strings.noOptions; empty/boolean shorthand keeps the status shell.
     * When omitted, strings.noOptions then the per-form dictionary supply the empty-result content.
     */
    noOptionsContent?: ContentSlot<'div'>;
    /**
     * Localized package-owned query-result content.
     * When omitted, native SPFx resources and enclosing per-form dictionaries provide package-owned UI text.
     */
    strings?: Partial<FormTagPickerStrings>;
};

/**
 * Numeric Rating with a public div ref and error focus on an eligible native radio.
 * RHF changes precede consumer callbacks. No codec or default-value override exists.
 * Star color defaults to native `marigold`; explicit `brand` or `neutral` overrides are preserved.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormRatingProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ score: number }>();
 * const fieldProps: FormRatingProps<{ score: number }> = { name: 'score' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ score: 1 }} onSubmit={values => { console.log(values); }}>
 *     <Order.Rating {...fieldProps} label="Score" max={5} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormRatingProps<TInput extends FieldValues> = FormControlProps<TInput, number> & Omit<RatingProps, keyof FormFieldPresentationProps | 'defaultValue' | 'name' | 'value'>;
/**
 * HSV ColorPicker with optional alpha, copied on every read and write.
 * Compose Fluent ColorArea, ColorSlider or AlphaSlider children. The ref targets
 * the root div; RHF error focus targets an eligible native range input.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { ColorSlider } from '@fluentui/react-components';
 * import { createForm, FormColorPickerProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ accent: { h: number; s: number; v: number; a?: number } }>();
 * const fieldProps: FormColorPickerProps<{ accent: { h: number; s: number; v: number; a?: number } }> = { name: 'accent' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ accent: { h: 210, s: 0.7, v: 0.8 } }} onSubmit={values => { console.log(values); }}>
 *     <Order.ColorPicker {...fieldProps} label="Accent"><ColorSlider channel="hue" /></Order.ColorPicker>
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormColorPickerProps<TInput extends FieldValues> = FormControlProps<TInput, FormHsvColor> & Omit<ColorPickerProps, keyof FormFieldPresentationProps | 'color' | 'name'>;
/**
 * SwatchPicker stores the stable string identity, independent of label and color.
 * Compose Fluent ColorSwatch or ImageSwatch children. The public ref targets the
 * root div; RHF error focus targets an enabled swatch button.
 * Native shape defaults to rounded; square and circular remain supported overrides.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { ColorSwatch } from '@fluentui/react-components';
 * import { createForm, FormSwatchPickerProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ theme: string }>();
 * const fieldProps: FormSwatchPickerProps<{ theme: string }> = { name: 'theme' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ theme: 'blue' }} onSubmit={values => { console.log(values); }}>
 *     <Order.SwatchPicker {...fieldProps} label="Theme"><ColorSwatch value="blue" color="#0078d4" /></Order.SwatchPicker>
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormSwatchPickerProps<TInput extends FieldValues> = FormControlProps<TInput, string> & Omit<SwatchPickerProps, keyof FormFieldPresentationProps | 'defaultSelectedValue' | 'name' | 'selectedValue'>;

/**
 * Package-owned DatePicker validation error identifiers.
 * @example
 * ```ts
 * import { FormDatePickerError } from '@apvee/spfx-react-toolkit/controls/form';
 * const error: FormDatePickerError = 'invalid-input';
 * ```
 */
export type FormDatePickerError = 'invalid-input' | 'out-of-bounds' | 'required-input';

/**
 * Package-owned validation result emitted after DatePicker text validation.
 * @example
 * ```ts
 * import { FormDatePickerValidationResult } from '@apvee/spfx-react-toolkit/controls/form';
 * const onValidationResult = ({ error }: FormDatePickerValidationResult): void => {
 *   if (error !== undefined) console.log(error);
 * };
 * ```
 */
export type FormDatePickerValidationResult = {
  /** Validation error, or `undefined` when the current text is valid. */
  error?: FormDatePickerError;
};

/**
 * Localized calendar strings consumed by Form.DatePicker.
 * @example
 * ```ts
 * import { FormDatePickerStrings } from '@apvee/spfx-react-toolkit/controls/form';
 * const strings: FormDatePickerStrings = {
 *   months: ['January','February','March','April','May','June','July','August','September','October','November','December'],
 *   shortMonths: ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],
 *   days: ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'],
 *   shortDays: ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'],
 *   goToToday: 'Today',
 * };
 * ```
 */
export type FormDatePickerStrings = {
  /** Full month names beginning with January. */
  months: readonly string[];
  /** Abbreviated month names beginning with January. */
  shortMonths: readonly string[];
  /** Full weekday names beginning with Sunday. */
  days: readonly string[];
  /** Abbreviated weekday names beginning with Sunday. */
  shortDays: readonly string[];
  /** Label for the action that navigates to today. */
  goToToday: string;
  /**
   * Accessible label for navigating to the previous month.
   * When omitted, the per-form native calendar string remains in use.
   */
  prevMonthAriaLabel?: string;
  /**
   * Accessible label for navigating to the next month.
   * When omitted, the per-form native calendar string remains in use.
   */
  nextMonthAriaLabel?: string;
  /**
   * Accessible label for navigating to the previous year.
   * When omitted, the per-form native calendar string remains in use.
   */
  prevYearAriaLabel?: string;
  /**
   * Accessible label for navigating to the next year.
   * When omitted, the per-form native calendar string remains in use.
   */
  nextYearAriaLabel?: string;
  /**
   * Accessible label for navigating to the previous year range.
   * When omitted, the per-form native calendar string remains in use.
   */
  prevYearRangeAriaLabel?: string;
  /**
   * Accessible label for navigating to the next year range.
   * When omitted, the per-form native calendar string remains in use.
   */
  nextYearRangeAriaLabel?: string;
  /**
   * Accessible format string for the month-picker header.
   * When omitted, the per-form native calendar string remains in use.
   */
  monthPickerHeaderAriaLabel?: string;
  /**
   * Accessible format string for the year-picker header.
   * When omitted, the per-form native calendar string remains in use.
   */
  yearPickerHeaderAriaLabel?: string;
  /**
   * Accessible label for the calendar close action.
   * When omitted, the per-form native calendar string remains in use.
   */
  closeButtonAriaLabel?: string;
  /**
   * Accessible format string for a week number.
   * When omitted, the per-form native calendar string remains in use.
   */
  weekNumberFormatString?: string;
  /**
   * Accessible format string for the selected date.
   * When omitted, the per-form native calendar string remains in use.
   */
  selectedDateFormatString?: string;
  /**
   * Accessible format string for today's date.
   * When omitted, the per-form native calendar string remains in use.
   */
  todayDateFormatString?: string;
  /**
   * Accessible label announced for marked dates.
   * When omitted, the per-form native calendar string remains in use.
   */
  dayMarkedAriaLabel?: string;
};

/**
 * Package-owned localized formatting functions used by Form.DatePicker's calendar.
 * @example
 * ```ts
 * import { FormDatePickerFormatting } from '@apvee/spfx-react-toolkit/controls/form';
 * const formatter: FormDatePickerFormatting = {
 *   formatDay: date => String(date.getDate()),
 *   formatMonth: (date, strings) => strings.months[date.getMonth()],
 *   formatYear: date => String(date.getFullYear()),
 *   formatMonthDayYear: date => date.toLocaleDateString('it-IT'),
 *   formatMonthYear: date => date.toLocaleDateString('it-IT', { month: 'long', year: 'numeric' }),
 * };
 * ```
 */
export type FormDatePickerFormatting = {
  /**
   * Formats a day number.
   * @param date - Required Date whose local calendar parts are formatted; this calendar callback does not accept an absent date.
   * @returns Formatted local date/time text; no UTC or timezone conversion is added.
   */
  formatDay: (date: Date) => string;
  /**
   * Formats a month using the supplied localized strings.
   * @param date - Required Date whose local calendar parts are formatted; this calendar callback does not accept an absent date.
   * @param strings - Resolved localized calendar strings supplied to the synchronous formatter.
   * @returns Formatted local date/time text; no UTC or timezone conversion is added.
   */
  formatMonth: (date: Date, strings: FormDatePickerStrings) => string;
  /**
   * Formats a year.
   * @param date - Required Date whose local calendar parts are formatted; this calendar callback does not accept an absent date.
   * @returns Formatted local date/time text; no UTC or timezone conversion is added.
   */
  formatYear: (date: Date) => string;
  /**
   * Formats a complete month, day, and year.
   * @param date - Required Date whose local calendar parts are formatted; this calendar callback does not accept an absent date.
   * @param strings - Resolved localized calendar strings supplied to the synchronous formatter.
   * @returns Formatted local date/time text; no UTC or timezone conversion is added.
   */
  formatMonthDayYear: (date: Date, strings: FormDatePickerStrings) => string;
  /**
   * Formats a month and year.
   * @param date - Required Date whose local calendar parts are formatted; this calendar callback does not accept an absent date.
   * @param strings - Resolved localized calendar strings supplied to the synchronous formatter.
   * @returns Formatted local date/time text; no UTC or timezone conversion is added.
   */
  formatMonthYear: (date: Date, strings: FormDatePickerStrings) => string;
};

/**
 * Numeric weekday identifiers where Sunday is `0` and Saturday is `6`.
 * @example
 * ```ts
 * import { FormDatePickerDayOfWeek } from '@apvee/spfx-react-toolkit/controls/form';
 * const monday: FormDatePickerDayOfWeek = 1;
 * ```
 */
export type FormDatePickerDayOfWeek = 0 | 1 | 2 | 3 | 4 | 5 | 6;

/**
 * Numeric first-week rules: 0 is first day, 1 is first full week, and 2 is first four-day week.
 * @example
 * ```ts
 * import { FormDatePickerFirstWeekOfYear } from '@apvee/spfx-react-toolkit/controls/form';
 * const firstFourDayWeek: FormDatePickerFirstWeekOfYear = 2;
 * ```
 */
export type FormDatePickerFirstWeekOfYear = 0 | 1 | 2;

/** Native date/time input attributes after removing values, events and presentation owned by the Form binding. */
type FormDateTimeNativeInputProps = Omit<
  React.InputHTMLAttributes<HTMLInputElement>,
  | 'className'
  | 'children'
  | 'defaultValue'
  | 'disabled'
  | 'name'
  | 'onChange'
  | 'required'
  | 'size'
  | 'type'
  | 'value'
>;

/**
 * Package-owned DatePicker configuration without its Form field binding.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormDatePickerControlProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ date: Date | null }>();
 * const datePickerProps: FormDatePickerControlProps = { allowTextInput: true, firstDayOfWeek: 1 };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ date: null }} onSubmit={values => { console.log(values); }}>
 *     <Order.DatePicker name="date" label="Date" allowTextInput />
 *   </Order.Root>;
 * }
 * ```
 */
export type FormDatePickerControlProps = FormDateTimeNativeInputProps & {
  /**
   * Fluent input appearance.
   * When omitted, the native Fluent primitive appearance applies.
   * @defaultValue outline
   */
  appearance?: InputProps['appearance'];
  /**
   * Allows direct date text entry. Defaults to `false`. Uncommitted or rejected text blocks validation and submission until corrected, cleared, or reset.
   * When omitted, direct date text entry is disabled.
   * @defaultValue false
   */
  allowTextInput?: boolean;
  /**
   * Localized error for text that does not match the committed date. Omission inherits the per-form `dateInputValidationMessage` from the native resource or hook dictionary. Domain constraints remain owned by rules, schema, or resolver.
   * When omitted, the per-form dictionary supplies the corresponding date/time draft error.
   */
  inputValidationMessage?: string;
  /**
   * Allows disabled calendar elements to receive focus. Defaults to `false`.
   * When omitted, disabled calendar elements are not focusable.
   * @defaultValue false
   */
  allFocusable?: boolean;
  /**
   * Removes the DatePicker border. Defaults to `false`.
   * When omitted, the native border treatment remains.
   * @defaultValue false
   */
  borderless?: boolean;
  /**
   * Additional class name applied to the DatePicker root.
   * When omitted, only the component and Fluent classes apply.
   */
  className?: string;
  /**
   * Prevents interaction.
   * When omitted or false, adds no disabled restriction; enclosing Root/scope restrictions still apply.
   */
  disabled?: boolean;
  /**
   * Prevents automatic popup opening on focus.
   * When omitted, automatic popup focus is suppressed (compat defaults disableAutoFocus to true).
   * @defaultValue true
   */
  disableAutoFocus?: boolean;
  /**
   * Opens the calendar initially.
   * When omitted, the popup starts closed.
   * @defaultValue false
   */
  defaultOpen?: boolean;
  /**
   * Package-owned localized calendar formatting functions.
   * When omitted, compat calendar formatting functions apply; input formatting remains controlled by formatDate.
   */
  dateTimeFormatter?: FormDatePickerFormatting;
  /**
   * First weekday displayed by the calendar.
   * When omitted, compat uses Sunday (0).
   * @defaultValue 0
   */
  firstDayOfWeek?: FormDatePickerDayOfWeek;
  /**
   * Rule used to determine the first week of a year.
   * When omitted, compat uses FirstDay (0).
   * @defaultValue 0
   */
  firstWeekOfYear?: FormDatePickerFirstWeekOfYear;
  /**
   * Synchronously formats the selected local date as input text; defaults to YYYY-MM-DD, or empty text for undefined. Pair custom formats with parseDateFromString.
   * @param date - Selected local calendar Date; omitted/undefined means no selected date.
   * @returns Formatted local date/time text; no UTC or timezone conversion is added.
   */
  formatDate?: (date?: Date) => string;
  /**
   * Highlights the current month.
   * When omitted, compat does not highlight the current month.
   * @defaultValue false
   */
  highlightCurrentMonth?: boolean;
  /**
   * Highlights the selected month.
   * When omitted, compat does not highlight the selected month.
   * @defaultValue false
   */
  highlightSelectedMonth?: boolean;
  /**
   * Date initially highlighted when the calendar opens.
   * When omitted, compat captures the current client date as its initial picker date; an existing selected date takes priority.
   */
  initialPickerDate?: Date;
  /**
   * Renders the popup in DOM order rather than through a portal.
   * When omitted, Fluent renders the popup through its portal.
   * @defaultValue false
   */
  inlinePopup?: boolean;
  /**
   * Shows the month picker.
   * When omitted, compat shows the month picker.
   * @defaultValue true
   */
  isMonthPickerVisible?: boolean;
  /**
   * Inclusive latest selectable local calendar date; omission supplies no upper bound. Domain constraints still belong to rules/schema/resolver.
   * When omitted, no calendar upper bound applies.
   */
  maxDate?: Date;
  /**
   * Inclusive earliest selectable local calendar date; omission supplies no lower bound. Domain constraints still belong to rules/schema/resolver.
   * When omitted, no calendar lower bound applies.
   */
  minDate?: Date;
  /**
   * Controlled popup open state.
   * When omitted, the popup owns its open state using defaultOpen.
   */
  open?: boolean;
  /**
   * Opens the calendar when the input is clicked. Defaults to `true`.
   * When omitted, clicking the date input opens its popup.
   * @defaultValue true
   */
  openOnClick?: boolean;
  /**
   * Synchronous popup-open notification; receives the proposed boolean state. Update open externally when controlled; the return value is ignored.
   * When omitted, no application popup notification runs; uncontrolled native state still updates. Returned promises are not awaited.
   * @param open - Proposed popup open state; controlled consumers must update their open prop.
   */
  onOpenChange?: (open: boolean) => void;
  /**
   * Synchronous notification after an accepted calendar or text commit writes RHF. Receives the local Date or legacy compat null for clear; the return value is ignored.
   * When omitted, RHF commits still occur without a consumer notification. Returned promises are not awaited.
   * @param date - Committed local calendar Date, or legacy compat null when cleared, after the adapter’s accepted RHF write.
   */
  // Compat date/time selection and parser contracts explicitly use null for empty values.
  // eslint-disable-next-line @rushstack/no-new-null
  onSelectDate?: (date: Date | null) => void;
  /**
   * Synchronous date-validation notification with the package-owned error identifier or undefined for success. Invalid drafts keep the committed value; the return value is ignored.
   * @param data - Date text-validation result: error is undefined for valid text or contains the package-owned error identifier; rejected text retains the previously committed date.
   */
  onValidationResult?: (data: FormDatePickerValidationResult) => void;
  /**
   * Synchronously parses direct input text into a local date; return null for rejected text, following the compat parser contract.
   * When omitted, strict local YYYY-MM-DD parsing accepts only real calendar dates in years 0000–9999; no UTC/timestamp fallback occurs.
   * @param dateText - Direct input text to parse synchronously in local calendar coordinates.
   * @returns Local Date, or legacy compat null for rejected text. No promise is supported.
   */
  // Compat date/time selection and parser contracts explicitly use null for empty values.
  // eslint-disable-next-line @rushstack/no-new-null
  parseDateFromString?: (dateText: string) => Date | null;
  /**
   * Placeholder shown when no date is selected.
   * When omitted, no placeholder text is supplied by the adapter.
   */
  placeholder?: string;
  /**
   * Presents the control as required without replacing validation rules.
   * When omitted, no required presentation metadata is added; supply validation rules separately.
   */
  required?: boolean;
  /**
   * Shows the calendar close action.
   * When omitted, compat hides the calendar close action.
   * @defaultValue false
   */
  showCloseButton?: boolean;
  /**
   * Shows the action that navigates to today.
   * When omitted, compat shows the today action.
   * @defaultValue true
   */
  showGoToToday?: boolean;
  /**
   * Overlays the month picker above the day picker.
   * When omitted, compat keeps the month picker beside the day picker.
   * @defaultValue false
   */
  showMonthPickerAsOverlay?: boolean;
  /**
   * Shows calendar week numbers.
   * When omitted, calendar week numbers are hidden.
   * @defaultValue false
   */
  showWeekNumbers?: boolean;
  /**
   * Package-owned localized calendar strings.
   * When omitted, native SPFx resources and enclosing per-form dictionaries provide package-owned UI text.
   */
  strings?: FormDatePickerStrings;
  /**
   * Date treated as today; defaults to the client's current date.
   * When omitted, compat uses the current client date.
   */
  today?: Date;
  /**
   * Uses the DatePicker underline treatment.
   * When omitted, underline-only treatment is disabled.
   * @defaultValue false
   */
  underlined?: boolean;
};

/**
 * Nullable local date binding; the native ref targets its input.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormDatePickerProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ date: Date | null }>();
 * const fieldProps: FormDatePickerProps<{ date: Date | null }> = { name: 'date' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ date: null }} onSubmit={values => { console.log(values); }}>
 *     <Order.DatePicker {...fieldProps} label="Date" allowTextInput />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormDatePickerProps<TInput extends FieldValues> = FormControlProps<
  TInput,
  // Compat date/time selection and parser contracts explicitly use null for empty values.
  // eslint-disable-next-line @rushstack/no-new-null
  Date | null
> &
  FormDatePickerControlProps;

/**
 * Package-owned TimePicker validation error identifiers.
 * @example
 * ```ts
 * import { FormTimePickerError } from '@apvee/spfx-react-toolkit/controls/form';
 * const error: FormTimePickerError = 'out-of-bounds';
 * ```
 */
export type FormTimePickerError = 'invalid-input' | 'out-of-bounds' | 'required-input';

/**
 * Package-owned TimePicker selection payload.
 * @example
 * ```ts
 * import * as React from 'react';
 * import { FormTimePickerChangeData } from '@apvee/spfx-react-toolkit/controls/form';
 * const onTimeChange = (_event: React.SyntheticEvent<HTMLElement>, data: FormTimePickerChangeData): void => {
 *   if (data.error === undefined) console.log(data.selectedTime);
 * };
 * ```
 */
export type FormTimePickerChangeData = {
  /** Selected or parsed time, or `null` when cleared. */
  // Compat date/time selection and parser contracts explicitly use null for empty values.
  // eslint-disable-next-line @rushstack/no-new-null
  selectedTime: Date | null;
  /** Display text associated with the selection or freeform input. */
  selectedTimeText: string | undefined;
  /**
   * Validation error associated with the submitted time text.
   * Undefined means no direct validation message is available.
   */
  error?: FormTimePickerError;
};

/**
 * Package-owned result returned by a freeform time parser.
 * @example
 * ```ts
 * import { FormTimePickerParseResult } from '@apvee/spfx-react-toolkit/controls/form';
 * const rejected: FormTimePickerParseResult = { date: null, error: 'invalid-input' };
 * ```
 */
export type FormTimePickerParseResult = {
  /** Parsed time, or `null` when no valid time is available. */
  // Compat date/time selection and parser contracts explicitly use null for empty values.
  // eslint-disable-next-line @rushstack/no-new-null
  date: Date | null;
  /**
   * Validation error associated with the parsed text.
   * Undefined means no direct validation message is available.
   */
  error?: FormTimePickerError;
};

/**
 * Inclusive/exclusive hour boundary supported by Form.TimePicker.
 * @example
 * ```ts
 * import { FormTimePickerHour } from '@apvee/spfx-react-toolkit/controls/form';
 * const endHour: FormTimePickerHour = 24;
 * ```
 */
export type FormTimePickerHour =
  | 0
  | 1
  | 2
  | 3
  | 4
  | 5
  | 6
  | 7
  | 8
  | 9
  | 10
  | 11
  | 12
  | 13
  | 14
  | 15
  | 16
  | 17
  | 18
  | 19
  | 20
  | 21
  | 22
  | 23
  | 24;

/**
 * Package-owned TimePicker configuration without its Form field binding.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormTimePickerControlProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ time: Date | null }>();
 * const timePickerProps: FormTimePickerControlProps = { freeform: true, hourCycle: 'h23', increment: 15 };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ time: null }} onSubmit={values => { console.log(values); }}>
 *     <Order.TimePicker name="time" label="Time" {...timePickerProps} />
 *   </Order.Root>;
 * }
 * ```
 */
export type FormTimePickerControlProps = FormDateTimeNativeInputProps & {
  /**
   * Fluent input appearance.
   * When omitted, the native Fluent primitive appearance applies.
   * @defaultValue outline
   */
  appearance?: ComboboxProps['appearance'];
  /**
   * Shows a clear action.
   * When omitted, no TimePicker clear action is requested.
   * @defaultValue false
   */
  clearable?: boolean;
  /**
   * Additional class name applied to the TimePicker root.
   * When omitted, only the component and Fluent classes apply.
   */
  className?: string;
  /**
   * Date used as the fixed basis for generated and parsed times. When omitted, the current field
   * date is preserved, or local 1970-01-01 is used while the field is empty.
   */
  dateAnchor?: Date;
  /**
   * Opens the time list initially.
   * When omitted, the popup starts closed.
   * @defaultValue false
   */
  defaultOpen?: boolean;
  /**
   * Prevents interaction.
   * When omitted or false, adds no disabled restriction; enclosing Root/scope restrictions still apply.
   */
  disabled?: boolean;
  /**
   * Exclusive end hour for generated options.
   * When omitted, generated options end before hour 24.
   * @defaultValue 24
   */
  endHour?: FormTimePickerHour;
  /**
   * Localized error for rejected or uncommitted time text; omission inherits the per-form `timeInputValidationMessage` from the native resource or hook dictionary. Domain validation stays in rules/schema/resolver.
   * When omitted, the per-form dictionary supplies the corresponding date/time draft error.
   */
  inputValidationMessage?: string;
  /**
   * Allows freeform time entry.
   * When omitted, native freeform time commits are disabled.
   * @defaultValue false
   */
  freeform?: boolean;
  /**
   * Synchronously formats generated/selected local time values; omission uses compat formatting with hourCycle and showSeconds.
   * When omitted, compat formatting uses hourCycle and showSeconds.
   * @param date - Required Date containing the local wall-clock time to format for generated or selected options.
   * @returns Formatted local date/time text; no UTC or timezone conversion is added.
   */
  formatDateToTimeString?: (date: Date) => string;
  /**
   * Selects the locale hour cycle.
   * When omitted, compat chooses the locale hour cycle.
   */
  hourCycle?: 'h11' | 'h12' | 'h23' | 'h24';
  /**
   * Minute interval between generated options.
   * When omitted, compat uses 30-minute options.
   * @defaultValue 30
   */
  increment?: number;
  /**
   * Renders the popup in DOM order rather than through a portal.
   * When omitted, Fluent renders the popup through its portal.
   * @defaultValue false
   */
  inlinePopup?: boolean;
  /**
   * Controlled popup open state.
   * When omitted, the popup owns its open state using defaultOpen.
   */
  open?: boolean;
  /**
   * Synchronous popup-open notification; receives the proposed boolean state. Update open externally when controlled; the return value is ignored.
   * When omitted, no application popup notification runs; uncontrolled native state still updates. Returned promises are not awaited.
   * @param open - Proposed popup open state; controlled consumers must update their open prop.
   */
  onOpenChange?: (open: boolean) => void;
  /**
   * Receives the native selection event and package-owned time/text/error payload after any accepted RHF write. Invalid draft notifications retain the committed value. The return value is ignored.
   * When omitted, accepted RHF writes and draft validation still occur without a consumer notification. Returned promises are not awaited.
   * @param event - React event accompanying the time selection, clear or text commit.
   * @param data - The selected or parsed time (legacy compat null for clear), its display text and an optional package-owned validation error; invalid drafts retain the previously committed time.
   */
  onTimeChange?: (event: React.SyntheticEvent<HTMLElement>, data: FormTimePickerChangeData) => void;
  /**
   * Synchronously parses possibly absent freeform text; return the parsed date (or legacy compat null) and optional error. No async parsing is supported.
   * When omitted, compat parses time text with its default parser.
   * @param timeText - Freeform time text, possibly undefined; parse synchronously.
   * @returns Parsed local time and optional validation error; legacy compat null means no parsed time.
   */
  parseTimeStringToDate?: (timeText: string | undefined) => FormTimePickerParseResult;
  /**
   * Placeholder shown when no time is selected.
   * When omitted, no placeholder text is supplied by the adapter.
   */
  placeholder?: string;
  /**
   * Includes seconds in generated and formatted options.
   * When omitted, generated time options omit seconds.
   * @defaultValue false
   */
  showSeconds?: boolean;
  /**
   * Fluent input size.
   * When omitted, native Fluent input size is medium.
   * @defaultValue medium
   */
  size?: InputProps['size'];
  /**
   * Inclusive start hour for generated options.
   * When omitted, generated options start at hour 0.
   * @defaultValue 0
   */
  startHour?: FormTimePickerHour;
};

/**
 * Nullable local time binding; the native ref targets its input.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormTimePickerProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ time: Date | null }>();
 * const fieldProps: FormTimePickerProps<{ time: Date | null }> = { name: 'time' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ time: null }} onSubmit={values => { console.log(values); }}>
 *     <Order.TimePicker {...fieldProps} label="Time" freeform hourCycle="h23" increment={15} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormTimePickerProps<TInput extends FieldValues> = FormControlProps<
  TInput,
  // Compat date/time selection and parser contracts explicitly use null for empty values.
  // eslint-disable-next-line @rushstack/no-new-null
  Date | null
> &
  FormTimePickerControlProps;

/**
 * Separate local date/time parts; the public ref targets the semantic group.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormDateTimeProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ appointment: { date: Date | null; time: Date | null } }>();
 * const fieldProps: FormDateTimeProps<{ appointment: { date: Date | null; time: Date | null } }> = { name: 'appointment' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ appointment: { date: null, time: null } }} onSubmit={values => { console.log(values); }}>
 *     <Order.DateTime {...fieldProps} dateLabel="Date" timeLabel="Time" />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormDateTimeProps<TInput extends FieldValues> = FormControlProps<
  TInput,
  FormDateTimeValue
> & {
  /**
   * Explicit identity of the semantic date/time group.
   * When omitted, the binding uses generated native Field/control identity.
   */
  id?: string;
  /**
   * Additional descriptions appended to group hint and validation relationships.
   * When omitted, only adapter-owned hint/error relationships apply.
   */
  'aria-describedby'?: string;
  /**
   * Literal phrasing slot inside the DatePicker native Label; object/callback refs target its inner span.
   * When omitted, no separate visible date-part Label is requested.
   */
  dateLabel?: ContentSlot<'span'>;
  /**
   * Package-owned props forwarded to the DatePicker part.
   * When omitted, DatePicker part defaults apply; the parent owns its field value and disabled state.
   */
  datePickerProps?: FormDatePickerControlProps;
  /**
   * Literal phrasing slot inside the TimePicker native Label; object/callback refs target its inner span.
   * When omitted, no separate visible time-part Label is requested.
   */
  timeLabel?: ContentSlot<'span'>;
  /**
   * Package-owned props forwarded to the TimePicker part.
   * When omitted, TimePicker part defaults apply; the parent owns its field value and disabled state.
   */
  timePickerProps?: FormTimePickerControlProps;
};


/**
 * Native browser file selection stored as FileList, with null for an empty selection.
 * The ref targets the real HTMLInputElement; change/blur callbacks run after RHF.
 * Browsers cannot restore files from defaults, reset values or external values: package-owned
 * file fields normalize unsupported selections to null. Explicit keepDirtyValues preserves a
 * current native selection only while its input survives. Upload and persisted attachments are
 * consumer responsibilities.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormFileInputProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ attachments: FileList | null }>();
 * const fieldProps: FormFileInputProps<{ attachments: FileList | null }> = { name: 'attachments' };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ attachments: null }} onSubmit={values => { console.log(values); }}>
 *     <Order.FileInput {...fieldProps} label="Attachments" accept="application/pdf" multiple />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormFileInputProps<TInput extends FieldValues> = FormControlProps<
  TInput,
  // HTMLInputElement.files and native file-clear interoperability explicitly use null.
  // eslint-disable-next-line @rushstack/no-new-null
  FileList | null
> & Omit<React.InputHTMLAttributes<HTMLInputElement>, keyof FormFieldPresentationProps | 'children' | 'defaultValue' | 'name' | 'type' | 'value'> & {
  /**
   * Localized visible action that opens the native file chooser; omission inherits per-form `fileInput.browseLabel`.
   * When omitted, the per-form fileInput.browseLabel supplies visible action text.
   */
  browseLabel?: string;
  /**
   * Localized accessible label for clearing the native selection; omission inherits per-form `fileInput.clearLabel`.
   * When omitted, the per-form fileInput.clearLabel supplies the accessible name.
   */
  clearLabel?: string;
  /**
   * Formats the actual live native files; defaults to comma-separated file names.
   * When omitted, comma-separated actual file names are displayed; empty selection displays no names.
   * @param files - Actual live browser FileList; formatting does not upload or restore native files.
   * @returns Caller-owned React presentation of the current native selection.
   */
  formatFiles?: (files: FileList) => React.ReactNode;
};

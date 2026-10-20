import * as React from 'react';
import { FormProvider, useForm as useReactHookForm } from 'react-hook-form';
import type { FieldValues, Resolver } from 'react-hook-form';
import type { FormMethods, FormRootProps, FormSuite, FormUseFieldProps, FormUseFormProps, ScopedFormSuite, UntypedFormSuite } from './form.types';
import { FormInput } from './controls/form-input';
import { FormTextarea } from './controls/form-textarea';
import { FormSearchBox } from './controls/form-search-box';
import { FormCheckbox } from './controls/form-checkbox';
import { FormSwitch } from './controls/form-switch';
import { FormSpinButton } from './controls/form-spin-button';
import { FormSlider } from './controls/form-slider';
import { FormSelect } from './controls/form-select';
import { FormRadioGroup } from './controls/form-radio-group';
import { FormDropdown } from './controls/form-dropdown';
import { FormDropdownMulti } from './controls/form-dropdown-multi';
import { FormCheckboxGroup } from './controls/form-checkbox-group';
import { FormCombobox } from './controls/form-combobox';
import { FormComboboxMulti } from './controls/form-combobox-multi';
import { FormTagPicker } from './controls/form-tag-picker';
import { FormRating } from './controls/form-rating';
import { FormColorPicker } from './controls/form-color-picker';
import { FormSwatchPicker } from './controls/form-swatch-picker';
import { FormDatePicker } from './controls/form-date-picker';
import { FormTimePicker } from './controls/form-time-picker';
import { FormDateTime } from './controls/form-date-time';
import { FormFileInput } from './controls/form-file-input';
import { FormObject } from './composition/form-object';
import { FormArray } from './composition/form-array';
import { FormObjectSection } from './composition/form-object-section';
import { FormArraySection } from './composition/form-array-section';
import type { FormObjectPath, FormObjectProps, FormArrayPath, FormArrayProps, FormObjectComponent, FormArrayComponent, FormObjectSectionProps, FormArraySectionProps, FormObjectSectionComponent, FormArraySectionComponent } from './form-composition.types';
import { FormSubmit } from './form-submit';
import { FormReset } from './form-reset';
import { FormError } from './form-error';
import { FormRuntimeContext, useFormRuntimeContext } from './form-context.internal';
import { FieldPathScope, joinFieldPath } from './field-path.internal';
import { assertFormValues, requireFormMethodsMetadata, useFormMethods, useFormResetSynchronization, useFormResetDefaultsBridge } from './form-methods.internal';
import type { FormInputValidators } from './form-methods.internal';
import { normalizeFormRules, reportFormFailure, resolveFormOptions, snapshotFormOutput, unexpectedFormFailureMessage, withFormInputValidation } from './form-validation.internal';
import { useAbsoluteFormField, useFormField } from './use-form-field';
const initializationKeys = ['mode', 'reValidateMode', 'defaultValues', 'values', 'errors', 'resetOptions', 'resolver', 'schema', 'context', 'shouldFocusError', 'shouldUnregister', 'shouldUseNativeValidation', 'progressive', 'criteriaMode', 'delayError', 'formControl', 'validate'] as const;
type FrameProps<TInput extends FieldValues, TContext, TOutput> = Omit<React.FormHTMLAttributes<HTMLFormElement>, 'onSubmit' | 'onInvalid'> & {
    methods: FormMethods<TInput, TContext, TOutput>;
    disabled?: boolean;
    onSubmit: (values: TOutput) => void | Promise<void>;
    onInvalid?: RuntimeRootProps<TInput, TContext, TOutput>['onInvalid'];
};
type RuntimeRootProps<TInput extends FieldValues, TContext, TOutput> = FormRootProps<TInput, TContext, TOutput, boolean>;
/** Constructs only the real implemented family; integrations extend assembly with real members. */
export function createFormSuiteRuntime<TInput extends FieldValues, TContext, TOutput, TBound extends boolean>(boundResolver?: Resolver<TInput, TContext, TOutput>): FormSuite<TInput, TContext, TOutput, TBound> {
    const suiteIdentity = Symbol('FormSuite');
    function useSuiteForm(props?: FormUseFormProps<TInput, TContext, TOutput, TBound>): FormMethods<TInput, TContext, TOutput> {
        const options = resolveFormOptions(props, boundResolver);
        const [inputValidators] = React.useState<FormInputValidators>(() => new Map());
        const resolver = React.useMemo(() => options.resolver === undefined ? undefined : withFormInputValidation(options.resolver, inputValidators), [options.resolver, inputValidators]);
        const defaultsBridge = useFormResetDefaultsBridge<TInput, TContext, TOutput>(options.defaultValues);
        const methods = useReactHookForm<TInput, TContext, TOutput>({ ...options,
            resetOptions: defaultsBridge.resetOptions,
            ...(defaultsBridge.defaultValues === undefined ? {} : { defaultValues: defaultsBridge.defaultValues }),
            ...(resolver === undefined ? {} : { resolver }) });
        assertFormValues(methods.getValues());
        const enhanced = useFormMethods(methods, suiteIdentity, inputValidators, options.resolver !== undefined, normalizeFormRules);
        useFormResetSynchronization(methods, options, defaultsBridge);
        return enhanced;
    }
    function useSuiteField<TValue>(props: FormUseFieldProps<TInput, TValue>): ReturnType<typeof useFormField<TInput, TValue>> {
        requireFormMethodsMetadata(useFormRuntimeContext().methods, suiteIdentity);
        return useFormField(props);
    }
    function bindMember<TProps, TElement>(Component: React.ForwardRefExoticComponent<React.PropsWithoutRef<TProps> & React.RefAttributes<TElement>>): React.ForwardRefExoticComponent<React.PropsWithoutRef<TProps> & React.RefAttributes<TElement>> {
        return React.forwardRef<TElement, TProps>((props, ref) => {
            requireFormMethodsMetadata(useFormRuntimeContext().methods, suiteIdentity);
            return React.createElement(Component, { ...props, ref } as React.PropsWithoutRef<TProps> & React.RefAttributes<TElement>);
        });
    }
    function ObjectMember<TPath extends FormObjectPath<FieldValues>>(props: FormObjectProps<FieldValues, TPath>): React.ReactElement {
        requireFormMethodsMetadata(useFormRuntimeContext().methods, suiteIdentity);
        return <FormObject<FieldValues, TPath> {...props} />;
    }
    function ArrayMember<TPath extends FormArrayPath<FieldValues>>(props: FormArrayProps<FieldValues, TPath>): React.ReactElement {
        requireFormMethodsMetadata(useFormRuntimeContext().methods, suiteIdentity);
        return <FormArray<FieldValues, TPath> {...props} />;
    }
    function ObjectSectionMember<TPath extends FormObjectPath<FieldValues>>(props: FormObjectSectionProps<FieldValues, TPath>): React.ReactElement {
        requireFormMethodsMetadata(useFormRuntimeContext().methods, suiteIdentity);
        return <FormObjectSection<FieldValues, TPath> {...props} />;
    }
    function ArraySectionMember<TPath extends FormArrayPath<FieldValues>>(props: FormArraySectionProps<FieldValues, TPath>): React.ReactElement {
        requireFormMethodsMetadata(useFormRuntimeContext().methods, suiteIdentity);
        return <FormArraySection<FieldValues, TPath> {...props} />;
    }
    const fields = {
        Input: bindMember(FormInput),
        Textarea: bindMember(FormTextarea),
        SearchBox: bindMember(FormSearchBox),
        Checkbox: bindMember(FormCheckbox),
        Switch: bindMember(FormSwitch),
        SpinButton: bindMember(FormSpinButton),
        Slider: bindMember(FormSlider),
        Select: bindMember(FormSelect),
        RadioGroup: bindMember(FormRadioGroup),
        Dropdown: bindMember(FormDropdown),
        DropdownMulti: bindMember(FormDropdownMulti),
        CheckboxGroup: bindMember(FormCheckboxGroup),
        Combobox: bindMember(FormCombobox),
        ComboboxMulti: bindMember(FormComboboxMulti),
        TagPicker: bindMember(FormTagPicker),
        Rating: bindMember(FormRating),
        ColorPicker: bindMember(FormColorPicker),
        SwatchPicker: bindMember(FormSwatchPicker),
        DatePicker: bindMember(FormDatePicker),
        TimePicker: bindMember(FormTimePicker),
        DateTime: bindMember(FormDateTime),
        FileInput: bindMember(FormFileInput),
        Object: ObjectMember,
        Array: ArrayMember,
        ObjectSection: ObjectSectionMember,
        ArraySection: ArraySectionMember,
        Error: bindMember(FormError),
        useField: useSuiteField,
    } as unknown as ScopedFormSuite<FieldValues>;
    function createScopedFields(usePrefix: () => string): ScopedFormSuite<FieldValues> {
        // Every scoped adapter reads the current owner path at render and retains its native ref.
        function scopeMember<TProps, TElement>(Component: React.ForwardRefExoticComponent<React.PropsWithoutRef<TProps> & React.RefAttributes<TElement>>): React.ForwardRefExoticComponent<React.PropsWithoutRef<TProps> & React.RefAttributes<TElement>> {
            return React.forwardRef<TElement, TProps>((props, ref) => <FieldPathScope prefix={usePrefix()}>
                {React.createElement(Component, { ...props, ref } as React.PropsWithoutRef<TProps> & React.RefAttributes<TElement>)}
            </FieldPathScope>);
        }
        function useScopedField<TValue>(props: FormUseFieldProps<FieldValues, TValue>): ReturnType<typeof useFormField<FieldValues, TValue>> {
            const prefix = usePrefix();
            requireFormMethodsMetadata(useFormRuntimeContext().methods, suiteIdentity);
            return useAbsoluteFormField({ ...props, name: joinFieldPath(prefix, String(props.name)) as typeof props.name });
        }
        const ScopedObject: FormObjectComponent<FieldValues> = props => <FieldPathScope prefix={usePrefix()}><fields.Object {...props} /></FieldPathScope>;
        const ScopedArray: FormArrayComponent<FieldValues> = props => <FieldPathScope prefix={usePrefix()}><fields.Array {...props} /></FieldPathScope>;
        const ScopedObjectSection: FormObjectSectionComponent<FieldValues> = <TPath extends FormObjectPath<FieldValues>>(props: FormObjectSectionProps<FieldValues, TPath>) => <FieldPathScope prefix={usePrefix()}><fields.ObjectSection<TPath> {...props} /></FieldPathScope>;
        const ScopedArraySection: FormArraySectionComponent<FieldValues> = <TPath extends FormArrayPath<FieldValues>>(props: FormArraySectionProps<FieldValues, TPath>) => <FieldPathScope prefix={usePrefix()}><fields.ArraySection<TPath> {...props} /></FieldPathScope>;
        return {
            Input: scopeMember(fields.Input),
            Textarea: scopeMember(fields.Textarea),
            SearchBox: scopeMember(fields.SearchBox),
            Checkbox: scopeMember(fields.Checkbox),
            Switch: scopeMember(fields.Switch),
            SpinButton: scopeMember(fields.SpinButton),
            Slider: scopeMember(fields.Slider),
            Select: scopeMember(fields.Select),
            RadioGroup: scopeMember(fields.RadioGroup),
            Dropdown: scopeMember(fields.Dropdown),
            DropdownMulti: scopeMember(fields.DropdownMulti),
            CheckboxGroup: scopeMember(fields.CheckboxGroup),
            Combobox: scopeMember(fields.Combobox),
            ComboboxMulti: scopeMember(fields.ComboboxMulti),
            TagPicker: scopeMember(fields.TagPicker),
            Rating: scopeMember(fields.Rating),
            ColorPicker: scopeMember(fields.ColorPicker),
            SwatchPicker: scopeMember(fields.SwatchPicker),
            DatePicker: scopeMember(fields.DatePicker),
            TimePicker: scopeMember(fields.TimePicker),
            DateTime: scopeMember(fields.DateTime),
            FileInput: scopeMember(fields.FileInput),
            Object: ScopedObject,
            Array: ScopedArray,
            ObjectSection: ScopedObjectSection,
            ArraySection: ScopedArraySection,
            Error: scopeMember(fields.Error),
            useField: useScopedField,
        };
    }
    const Frame = React.forwardRef<HTMLFormElement, FrameProps<TInput, TContext, TOutput>>(function FormFrame(props, ref) {
        const { methods, children, onSubmit, onInvalid, onReset, noValidate = true, disabled, ...native } = props;
        const runtime = React.useMemo(() => ({ methods: methods as unknown as FormMethods<FieldValues, unknown, unknown>, disabled: disabled === true || methods.formState.disabled, Fields: fields, createScopedFields }), [methods, disabled, methods.formState.disabled]);
        const active = React.useRef(false);
        const submit = methods.handleSubmit(async (value) => {
            try {
                await onSubmit(snapshotFormOutput(value));
            }
            catch (error) {
                reportFormFailure(error);
                methods.setError('root.submit', { type: 'unexpected', message: unexpectedFormFailureMessage });
            }
        }, async (errors) => {
            try {
                await onInvalid?.(errors, methods);
            }
            catch (error) {
                reportFormFailure(error);
                methods.setError('root.submit', { type: 'unexpected', message: unexpectedFormFailureMessage });
            }
        });
        const releaseSubmission = (): void => { active.current = false; };
        const submitFromRoot = async (event: React.FormEvent<HTMLFormElement>): Promise<void> => {
            event.preventDefault();
            if (active.current)
                return;
            active.current = true;
            try {
                await submit(event);
            }
            finally {
                releaseSubmission();
            }
        };
        return <FormProvider {...methods}><FormRuntimeContext.Provider value={runtime}><form {...native} ref={ref} noValidate={noValidate} onSubmit={event => { submitFromRoot(event).catch(reportFormFailure); }} onReset={event => { if (requireFormMethodsMetadata(methods).isResetting) { event.preventDefault(); return; } onReset?.(event); if (!event.defaultPrevented) {
            event.preventDefault();
            methods.reset();
        } }}>{children}</form></FormRuntimeContext.Provider></FormProvider>;
    });
    const OwnedRoot = React.forwardRef<HTMLFormElement, RuntimeRootProps<TInput, TContext, TOutput>>(function OwnedRoot(props, ref) {
        const options: Record<string, unknown> = {};
        const native: Record<string, unknown> = { ...props };
        for (const key of initializationKeys) {
            if (key in native) {
                options[key] = native[key];
                delete native[key];
            }
        }
        if (props.disabled !== undefined)
            options.disabled = props.disabled;
        const methods = useSuiteForm(options as FormUseFormProps<TInput, TContext, TOutput, TBound>);
        return <Frame {...native as RuntimeRootProps<TInput, TContext, TOutput>} methods={methods} ref={ref}/>;
    });
    /** Native Root owns RHF initialization or receives same-suite methods.
     * @example `<Order.Root defaultValues={{reference:''}} onSubmit={save}>...</Order.Root>`
     */
    const Root = React.forwardRef<HTMLFormElement, RuntimeRootProps<TInput, TContext, TOutput>>((props, ref) => {
        if ('dangerouslySetInnerHTML' in props)
            throw new TypeError('Form.Root owns its children and cannot accept raw HTML.');
        if (props.methods !== undefined) {
            requireFormMethodsMetadata(props.methods, suiteIdentity);
            for (const key of initializationKeys)
                if (props[key as keyof typeof props] !== undefined)
                    throw new TypeError('Provided Form.Root methods cannot accompany initialization options.');
            return <Frame {...props} methods={props.methods} ref={ref}/>;
        }
        return <OwnedRoot {...props} ref={ref}/>;
    });
    Root.displayName = 'Form.Root';
    return { ...fields, Root, Submit: bindMember(FormSubmit), Reset: bindMember(FormReset), useForm: useSuiteForm } as unknown as FormSuite<TInput, TContext, TOutput, TBound>;
}
/** Dynamic-name Form family; use createForm for input/output-aware paths.
 * React Hook Form owns values and validation. Render bindings inside this family's Root;
 * Fluent UI and date/time compat peers are required, with a shared FluentProvider.
 * No SPFx provider is required.
 * @example `<Form.Root defaultValues={{title:''}} onSubmit={save}><Form.Input name="title" /></Form.Root>`
 */
export const Form: UntypedFormSuite = createFormSuiteRuntime<FieldValues, unknown, FieldValues, false>();

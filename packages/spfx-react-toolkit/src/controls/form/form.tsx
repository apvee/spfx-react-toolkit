/**
 * Typed Form families, RHF value ownership and native Fluent bindings.
 *
 * @packageDocumentation
 * @docPage controls/form
 */
import * as React from 'react';
import { resolveFormStrings } from './form-strings.internal';
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
import { normalizeFormRules, reportFormFailure, resolveFormOptions, snapshotFormOutput, withFormInputValidation } from './form-validation.internal';
import { useAbsoluteFormField, useFormField } from './use-form-field';
const initializationKeys = ['strings', 'mode', 'reValidateMode', 'defaultValues', 'values', 'errors', 'resetOptions', 'resolver', 'schema', 'context', 'shouldFocusError', 'shouldUnregister', 'shouldUseNativeValidation', 'progressive', 'criteriaMode', 'delayError', 'formControl', 'validate'] as const;
type FrameProps<TInput extends FieldValues, TContext, TOutput> = Omit<React.FormHTMLAttributes<HTMLFormElement>, 'onSubmit' | 'onInvalid'> & {
    methods: FormMethods<TInput, TContext, TOutput>;
    disabled?: boolean;
    onSubmit: (values: TOutput) => void | Promise<void>;
    onInvalid?: RuntimeRootProps<TInput, TContext, TOutput>['onInvalid'];
};
type RuntimeRootProps<TInput extends FieldValues, TContext, TOutput> = FormRootProps<TInput, TContext, TOutput, boolean>;
/**
 * Constructs only the real implemented family; integrations extend assembly with real members.
 */
export function createFormSuiteRuntime<TInput extends FieldValues, TContext, TOutput, TBound extends boolean>(boundResolver?: Resolver<TInput, TContext, TOutput>): FormSuite<TInput, TContext, TOutput, TBound> {
    const suiteIdentity = Symbol('FormSuite');
    function useSuiteForm(props?: FormUseFormProps<TInput, TContext, TOutput, TBound>): FormMethods<TInput, TContext, TOutput> {
        const strings = React.useMemo(() => resolveFormStrings(props?.strings), [props?.strings]);
        const committedStrings = React.useRef(strings);
        const getStrings = React.useCallback(() => committedStrings.current, []);
        const getUnexpectedFailure = React.useCallback(() => getStrings().unexpectedFailure, [getStrings]);
        const options = resolveFormOptions(props, boundResolver);
        const [inputValidators] = React.useState<FormInputValidators>(() => new Map());
        const resolver = React.useMemo(() => options.resolver === undefined ? undefined : withFormInputValidation(options.resolver, inputValidators, getUnexpectedFailure), [options.resolver, inputValidators, getUnexpectedFailure]);
        const defaultsBridge = useFormResetDefaultsBridge<TInput, TContext, TOutput>(options.defaultValues);
        const methods = useReactHookForm<TInput, TContext, TOutput>({ ...options,
            resetOptions: defaultsBridge.resetOptions,
            ...(defaultsBridge.defaultValues === undefined ? {} : { defaultValues: defaultsBridge.defaultValues }),
            ...(resolver === undefined ? {} : { resolver }) });
        assertFormValues(methods.getValues());
        const enhanced = useFormMethods(methods, suiteIdentity, inputValidators, options.resolver !== undefined, normalizeFormRules, getStrings);
        React.useLayoutEffect(() => {
            if (committedStrings.current === strings) return;
            committedStrings.current = strings;
            for (const listener of requireFormMethodsMetadata(methods).stringsListeners) listener();
        }, [methods, strings]);
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
        const metadata = requireFormMethodsMetadata(methods);
        const [, refreshStrings] = React.useReducer((version: number) => version + 1, 0);
        React.useLayoutEffect(() => {
            metadata.stringsListeners.add(refreshStrings);
            return () => { metadata.stringsListeners.delete(refreshStrings); };
        }, [metadata]);
        const strings = metadata.getStrings();
        const runtime = React.useMemo(() => ({ strings, methods: methods as unknown as FormMethods<FieldValues, unknown, unknown>, disabled: disabled === true || methods.formState.disabled, Fields: fields, createScopedFields }), [methods, disabled, methods.formState.disabled, strings]);
        const active = React.useRef(false);
        const submit = methods.handleSubmit(async (value) => {
            try {
                await onSubmit(snapshotFormOutput(value));
            }
            catch (error) {
                reportFormFailure(error);
                methods.setError('root.submit', { type: 'unexpected', message: requireFormMethodsMetadata(methods).getStrings().unexpectedFailure });
            }
        }, async (errors) => {
            try {
                await onInvalid?.(errors, methods);
            }
            catch (error) {
                reportFormFailure(error);
                methods.setError('root.submit', { type: 'unexpected', message: requireFormMethodsMetadata(methods).getStrings().unexpectedFailure });
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
    /**
     * Native Root owns RHF initialization or receives same-suite methods.
     * Example: `<Order.Root defaultValues={{reference:''}} onSubmit={save}>...</Order.Root>`
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
/**
 * Dynamic-name Form family; use createForm for input/output-aware paths.
 * React Hook Form owns values and validation. Render bindings inside this family's Root;
 * Fluent UI and date/time compat peers are required, with a shared FluentProvider.
 * No SPFx provider is required.
 * Example: `<Form.Root defaultValues={{title:''}} onSubmit={save}><Form.Input name="title" /></Form.Root>`
 * @example Dynamic field editor
 * ```tsx
 * import * as React from 'react';
 * import { Form } from '@apvee/spfx-react-toolkit/controls/form';
 *
 * export function Editor(): React.ReactElement {
 *   return (
 *     <Form.Root
 *       defaultValues={{ reference: '' }}
 *       onSubmit={values => {
 *         console.log(values.reference);
 *       }}
 *     >
 *       <Form.Input name="reference" />
 *       <Form.Submit>Save</Form.Submit>
 *     </Form.Root>
 *   );
 * }
 * ```
 * @useWhen Runtime field names prevent a statically typed createForm family.
 * @avoidWhen A known model can use createForm for checked live paths and output types.
 * @scenario Edit a dynamic report with explicit save and reset
 * ```tsx
 * import * as React from 'react';
 * import { Button, Checkbox, Field, Select, FluentProvider, makeStyles } from '@fluentui/react-components';
 * import { Form, flex, gap, useSx, width } from '@apvee/spfx-react-toolkit';
 * interface DemoRegions { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode; }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <section>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</section>;
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
 * const useStyles = makeStyles({ frame: { maxInlineSize: '100%', minInlineSize: 0 }, output: { whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' } });
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx(); const styles = useStyles();
 *   const [size, setSize] = React.useState(640); const [disabled, setDisabled] = React.useState(false);
 *   const [customStrings, setCustomStrings] = React.useState(false); const [rtl, setRtl] = React.useState(false);
 *   const [fail, setFail] = React.useState(false); const [delay, setDelay] = React.useState(false);
 *   const [count, setCount] = React.useState(0); const [status, setStatus] = React.useState('Ready to edit.');
 *   const [saved, setSaved] = React.useState<Record<string, unknown>>();
 *   const formRef = React.useMemo(() => React.createRef<HTMLFormElement>(), []);
 *   const methods = Form.useForm({ defaultValues: { dynamic: 'Quarterly report' }, strings: customStrings ? { unexpectedFailure: 'The report could not be saved. Please retry.' } : undefined });
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
 *     methods.reset({ dynamic: 'Quarterly report' }); setSize(640); setDisabled(false); setCustomStrings(false); setRtl(false);
 *     setFail(false); setDelay(false); setCount(0); setStatus('Ready to edit.'); setSaved(undefined); setHeld(false); setResetPending(false);
 *   }, [resetPending, isSubmitting, methods]);
 *   const busy = isSubmitting || resetPending;
 *   // RHF can defer the callback; capture its source generation during render.
 *   const submissionGeneration = generation.current;
 *   const save = async (value: Record<string, unknown>): Promise<void> => {
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
 *       <Button data-form-demo="new-record" disabled={busy} onClick={() => { methods.reset({ dynamic: 'Annual report' }); setStatus('New model loaded.'); }}>Load new report</Button>
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
 *       <Form.Root ref={formRef} methods={methods} disabled={disabled} onSubmit={save} onInvalid={notifyInvalid} data-form-demo="root">
 *         <Form.Input name="dynamic" label="Title" /><Form.Error /><Form.Error name="dynamic" />
 *         <Form.Submit data-form-demo="submit" disabled={resetPending}>Save report</Form.Submit><Form.Reset data-form-demo="reset" disabled={resetPending}>Reset</Form.Reset>
 *       </Form.Root>
 *     </div>,
 *   })}</FluentProvider></section>;
 * }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'Edit dynamic text, save a snapshot, load a pristine annual report and cancel pending saves on reset.' };
 * ```
 */
export const Form: UntypedFormSuite = createFormSuiteRuntime<FieldValues, unknown, FieldValues, false>();

import * as React from 'react';
import { useId, useMergedRefs, resolveShorthand } from '@fluentui/react-utilities';
import type { FieldProps } from '@fluentui/react-components';
import { useController } from 'react-hook-form';
import type { ControllerRenderProps, FieldValues, RegisterOptions } from 'react-hook-form';
import { useStableCallback } from '../../hooks/useStableCallback';
import { SettingsFieldContext, mergeSettingsFieldIds } from '../settings-field-context.internal';
import type { FormFieldPresentationProps } from './form-control.types';
import { useFormRuntimeContext } from './form-context.internal';
import { useFieldScopeDisabled, useResolvedFieldPath } from './field-path.internal';
import { assertFormFieldName, formResetTargets, registerFormResetOwner, registerFormResetText, requireFormMethodsMetadata, subscribeToFormResets } from './form-methods.internal';
import type { FormResetEvent } from './form-methods.internal';
import { normalizeFormRules } from './form-validation.internal';
import { getFormErrorMessages, renderFormErrorMessages } from './error-messages.internal';
/** Private binding options shared by explicit controls. */
export interface FieldAdapterOptions<TInput extends FieldValues, TValue> extends FormFieldPresentationProps {
    name: string;
    /** Native control id; omitted ids use a stable generated id shared with the Field label. */
    id?: string;
    /** Explicit native slot identity and presentation, used only in field row mode. */
    nativeId?: string;
    /** Native shorthand props override Fluent's propagated descriptions. */
    nativeSlotOwnsDescriptions?: boolean;
    nativeRequired?: boolean;
    nativeDisabled?: boolean;
    absoluteName?: boolean;
    /** Custom public useField bindings do not participate in automatic row naming. */
    inheritSettingsField?: boolean;
    /** Composite roots cannot be activated by a native label. */
    settingsLabelTarget?: 'group';
    /** Additional consumer description relationships on the actual semantic target. */
    describedBy?: string;
    ref?: React.Ref<HTMLElement>;
    rules?: RegisterOptions<TInput>;
    validateInput?: (value: TValue) => true | string;
    /** Transient query only: one requested dirty-reset handoff across RHF array row remounts.
     * read returns undefined when no query needs preservation, e.g. a derived selected label.
     * restore receives the committed query once after the new owner mounts its current path.
     */
    resetText?: { read: () => string | undefined; restore: (text: string) => void };
}
/** One real RHF controller plus presentation-only reset metadata. */
export interface FieldAdapterResult<TValue> {
    controller: Omit<ControllerRenderProps<FieldValues, string>, 'value' | 'ref'> & {
        value: TValue;
        ref: React.RefCallback<HTMLElement>;
    };
    fieldProps: FieldProps;
    controlProps: {
        id: string;
        disabled?: boolean;
        'aria-labelledby'?: string;
        'aria-describedby'?: string;
    };
    name: string;
    resetOwner: symbol;
    focusRef: React.RefCallback<HTMLElement>;
    resetVersion: number;
    resetPreservesValues: boolean;
    resetSource?: FormResetEvent['source'];
    state: {
        error?: string;
        errors: readonly string[];
        invalid: boolean;
        isDirty: boolean;
        isTouched: boolean;
    };
}
/** Binds live values directly to RHF and lets Fluent Field own accessible relationships. */
export function useFieldAdapter<TInput extends FieldValues, TValue>(options: FieldAdapterOptions<TInput, TValue>): FieldAdapterResult<TValue> {
    const { methods, disabled: rootDisabled } = useFormRuntimeContext();
    const scopedName = useResolvedFieldPath(options.name);
    const scopeDisabled = useFieldScopeDisabled();
    const name = options.absoluteName ? options.name : scopedName;
    assertFormFieldName(name);
    const metadata = requireFormMethodsMetadata(methods);
    if (metadata.hasResolver && options.rules !== undefined)
        throw new TypeError('Form field rules cannot accompany a schema/resolver validation owner.');
    const disabled = rootDisabled || scopeDisabled || options.disabled === true;
    const [owner] = React.useState(() => Symbol('FormFieldOwner'));
    const [reset, setReset] = React.useState({ version: 0, preserveValues: false, source: undefined as FormResetEvent['source'] | undefined });
    React.useLayoutEffect(() => () => { metadata.pendingOwnerResets.delete(owner); }, [metadata, owner]);
    React.useLayoutEffect(() => {
        const unregister = registerFormResetOwner(methods, owner, name);
        const unsubscribe = subscribeToFormResets(methods, event => { const currentName = metadata.resetOwners.get(owner); if (currentName !== undefined && (event.owners?.has(owner) ?? formResetTargets(event, currentName)))
            setReset(previous => ({ version: previous.version + 1, preserveValues: event.preserveValues || event.preservedOwners.has(owner), source: event.source })); });
        return () => { unsubscribe(); unregister(); };
    }, [methods, metadata, name, owner]);
    const hasResetText = options.resetText !== undefined;
    const readResetText = useStableCallback(() => options.resetText?.read());
    const restoreResetText = useStableCallback((text: string) => { options.resetText?.restore(text); });
    React.useLayoutEffect(() => {
        if (!hasResetText) return;
        return registerFormResetText(methods, owner, name, readResetText, restoreResetText);
    }, [methods, owner, name, hasResetText, readResetText, restoreResetText]);
    const validateInput = useStableCallback((value: unknown) => disabled ? true : options.validateInput?.(value as TValue) ?? true);
    React.useLayoutEffect(() => {
        if (!options.validateInput)
            return;
        const registration = { name, validate: validateInput };
        metadata.inputValidators.set(owner, registration);
        return () => { if (metadata.inputValidators.get(owner) === registration) metadata.inputValidators.delete(owner); };
    }, [metadata, name, owner, options.validateInput, validateInput]);
    const { field, fieldState } = useController({ control: methods.control, name, disabled, rules: normalizeFormRules(options.rules as RegisterOptions<FieldValues> | undefined, metadata.hasResolver || options.validateInput === undefined ? undefined : validateInput) });
    const inheritedSettingsField = React.useContext(SettingsFieldContext);
    const settingsField = options.inheritSettingsField === false ? undefined : inheritedSettingsField;
    if (settingsField && (options.label !== undefined || options.hint !== undefined || options.fieldProps?.hint !== undefined || (options.fieldProps !== undefined && 'label' in options.fieldProps && options.fieldProps.label !== undefined))) {
        throw new TypeError('SettingsGroup.Item field owns label and hint; omit the Form control label/hint.');
    }
    const generatedId = useId('form-control-');
    const id = options.id ?? (settingsField ? options.nativeId : undefined) ?? generatedId;
    React.useLayoutEffect(() => settingsField?.register({ owner, id, nativeLabel: options.settingsLabelTarget !== 'group', required: options.nativeRequired ?? options.required === true, disabled: disabled || options.nativeDisabled === true }),
        [settingsField, owner, id, options.settingsLabelTarget, options.required, options.nativeRequired, options.nativeDisabled, disabled]);
    const ref = useMergedRefs<HTMLElement>(field.ref, options.ref);
    const errors = getFormErrorMessages(fieldState.error);
    const message = renderFormErrorMessages(errors);
    // Fluent's native shorthand slots may override the Field-propagated description.
    // Give its existing message a stable identity so those slots can retain the same relationship.
    const validationMessage = message ?? options.fieldProps?.validationMessage;
    const validationSlot = settingsField ? resolveShorthand<React.HTMLAttributes<HTMLDivElement>>(validationMessage) : undefined;
    const validationId = validationSlot ? validationSlot.id ?? `${generatedId}-validation` : undefined;
    const fieldProps: FieldProps = {
        ...options.fieldProps,
        className: options.className,
        ...(options.label === undefined ? {} : { label: { children: options.label, htmlFor: id } }),
        ...(options.hint === undefined ? {} : { hint: options.hint }),
        orientation: options.orientation, required: options.required,
        validationMessage: validationSlot ? { ...validationSlot, id: validationId } : validationMessage,
        validationState: errors.length ? 'error' : options.fieldProps?.validationState,
    };
    return { controller: { ...field, value: field.value as TValue, ref }, fieldProps, controlProps: { id, ...(disabled ? { disabled: true } : {}), ...(settingsField ? { 'aria-labelledby': settingsField.labelId, 'aria-describedby': mergeSettingsFieldIds(options.describedBy, settingsField.descriptionId, options.nativeSlotOwnsDescriptions ? validationId : undefined) } : {}) }, name, focusRef: field.ref, resetOwner: owner, resetVersion: reset.version, resetPreservesValues: reset.preserveValues, resetSource: reset.source, state: { error: errors[0], errors, invalid: fieldState.invalid, isDirty: fieldState.isDirty, isTouched: fieldState.isTouched } };
}

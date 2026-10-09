import * as React from 'react';
import { useId, useMergedRefs } from '@fluentui/react-utilities';
import type { FieldProps } from '@fluentui/react-components';
import { useController } from 'react-hook-form';
import type { ControllerRenderProps, FieldValues, RegisterOptions } from 'react-hook-form';
import { useStableCallback } from '../../hooks/useStableCallback';
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
    absoluteName?: boolean;
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
    const generatedId = useId('form-control-');
    const id = options.id ?? generatedId;
    const ref = useMergedRefs<HTMLElement>(field.ref, options.ref);
    const errors = getFormErrorMessages(fieldState.error);
    const message = renderFormErrorMessages(errors);
    const fieldProps: FieldProps = {
        ...options.fieldProps,
        className: options.className,
        ...(options.label === undefined ? {} : { label: { children: options.label, htmlFor: id } }),
        ...(options.hint === undefined ? {} : { hint: options.hint }),
        orientation: options.orientation, required: options.required,
        validationMessage: message ?? options.fieldProps?.validationMessage,
        validationState: errors.length ? 'error' : options.fieldProps?.validationState,
    };
    return { controller: { ...field, value: field.value as TValue, ref }, fieldProps, controlProps: { id, ...(disabled ? { disabled: true } : {}) }, name, focusRef: field.ref, resetOwner: owner, resetVersion: reset.version, resetPreservesValues: reset.preserveValues, resetSource: reset.source, state: { error: errors[0], errors, invalid: fieldState.invalid, isDirty: fieldState.isDirty, isTouched: fieldState.isTouched } };
}

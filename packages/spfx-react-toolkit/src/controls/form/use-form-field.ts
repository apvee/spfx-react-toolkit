import type { FieldValues } from 'react-hook-form';
import type { FormFieldResult, FormUseFieldProps } from './form.types';
import { useFieldAdapter } from './use-field-adapter.internal';
import type { FieldAdapterResult } from './use-field-adapter.internal';
function semanticResult<TValue>(binding: FieldAdapterResult<TValue>): FormFieldResult<TValue> {
    return {
        ...binding.state,
        value: binding.controller.value,
        onChange: binding.controller.onChange,
        onBlur: binding.controller.onBlur,
        ref: binding.controller.ref,
        disabled: binding.controlProps.disabled === true,
        resetVersion: binding.resetVersion,
    };
}
/** Custom-control hook implementation; suite wrappers enforce same-suite ownership. */
export function useFormField<TInput extends FieldValues, TValue>(props: FormUseFieldProps<TInput, TValue>): FormFieldResult<TValue> {
    return semanticResult(useFieldAdapter<TInput, TValue>({ ...props, inheritSettingsField: false }));
}
/** Captured scoped fields already carry an absolute path; avoid applying the rendered scope twice. */
export function useAbsoluteFormField<TInput extends FieldValues, TValue>(props: FormUseFieldProps<TInput, TValue>): FormFieldResult<TValue> {
    return semanticResult(useFieldAdapter<TInput, TValue>({ ...props, absoluteName: true, inheritSettingsField: false }));
}

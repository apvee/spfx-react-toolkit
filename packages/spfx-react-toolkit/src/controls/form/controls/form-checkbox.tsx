import * as React from 'react';
import { Field, Checkbox } from '@fluentui/react-components';
import { resolveShorthand } from '@fluentui/react-utilities';
import type { FieldValues } from 'react-hook-form';
import type { FormCheckboxProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';

/** Checkbox binding through Fluent Field; required is accessibility metadata.
 * Semantic change callbacks run after RHF, top-level first and native-slot second.
 * Blur callbacks run after RHF and the ref targets the native input.
 * @example `<Order.Checkbox name="value" label="Checkbox" />`
 */
export const FormCheckbox = React.forwardRef<HTMLInputElement, FormCheckboxProps<FieldValues>>(function FormCheckbox({ className, disabled, fieldProps, hint, id, label, name, onBlur, onChange, orientation, required, rules, input, triState, ...props }, ref) {
  const nativeInput = resolveShorthand(input);
  const binding = useFieldAdapter<FieldValues, boolean | 'mixed'>({ nativeSlotOwnsDescriptions: true, nativeId: nativeInput?.id, nativeRequired: nativeInput?.required, nativeDisabled: nativeInput?.disabled, describedBy: [props['aria-describedby'], nativeInput?.['aria-describedby']].filter(Boolean).join(' '), className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules });
  const inputProps = { ...nativeInput, onChange: undefined, ...binding.controlProps, name: binding.controller.name, ref: binding.controller.ref, onBlur: (event: React.FocusEvent<HTMLInputElement>) => { binding.controller.onBlur(); onBlur?.(event); nativeInput?.onBlur?.(event); } };
  const checked = triState ? binding.controller.value ?? false : binding.controller.value === true;
  return <Field {...binding.fieldProps}><Checkbox {...props} {...binding.controlProps} name={binding.controller.name} ref={binding.controller.ref} checked={checked}
    input={{ ...inputProps, checked: checked === true }}
    onBlur={event => { binding.controller.onBlur(); onBlur?.(event); }}
    onChange={(event, data) => { binding.controller.onChange(triState ? data.checked : data.checked === true); onChange?.(event, data); nativeInput?.onChange?.(event); }} /></Field>;
});
FormCheckbox.displayName = 'Form.Checkbox';

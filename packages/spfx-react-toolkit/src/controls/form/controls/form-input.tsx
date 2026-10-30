import * as React from 'react';
import { Field, Input } from '@fluentui/react-components';
import { resolveShorthand } from '@fluentui/react-utilities';
import type { FieldValues } from 'react-hook-form';
import type { FormInputProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';

/** String-valued Fluent Input with native input ref and RHF change/blur binding.
 * Required is presentation metadata; validate with rules or the form's schema/resolver.
 * @example `<Order.Input name="reference" label="Reference" hint="Customer reference" />`
 */
export const FormInput: React.ForwardRefExoticComponent<FormInputProps<FieldValues> & React.RefAttributes<HTMLInputElement>> = React.forwardRef<HTMLInputElement, FormInputProps<FieldValues>>(function FormInput({ className, disabled, fieldProps, hint, id, label, name, onBlur, onChange, orientation, required, rules, ...props }, ref) {
  const native = resolveShorthand(props.input);
  const binding = useFieldAdapter<FieldValues, string>({ nativeSlotOwnsDescriptions: true, nativeId: native?.id, nativeRequired: native?.required, nativeDisabled: native?.disabled, describedBy: [props['aria-describedby'], native?.['aria-describedby']].filter(Boolean).join(' '), className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules });
  return <Field {...binding.fieldProps}><Input {...props} {...binding.controlProps} input={{ ...native, ...(binding.controlProps['aria-labelledby'] ? binding.controlProps : {}) }}
    name={binding.controller.name} ref={binding.controller.ref} value={binding.controller.value ?? ''}
    onBlur={event => { binding.controller.onBlur(); onBlur?.(event); }}
    onChange={(event, data) => { binding.controller.onChange(data.value); onChange?.(event, data); }} /></Field>;
});
FormInput.displayName = 'Form.Input';

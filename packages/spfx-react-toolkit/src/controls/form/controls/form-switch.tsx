import * as React from 'react';
import { Field, Switch } from '@fluentui/react-components';
import { resolveShorthand } from '@fluentui/react-utilities';
import type { FieldValues } from 'react-hook-form';
import type { FormSwitchProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';

/** Switch binding through Fluent Field; required is accessibility metadata.
 * Semantic change callbacks run after RHF, top-level first and native-slot second.
 * Blur callbacks run after RHF and the ref targets the native input.
 * @example `<Order.Switch name="value" label="Switch" />`
 */
export const FormSwitch: React.ForwardRefExoticComponent<FormSwitchProps<FieldValues> & React.RefAttributes<HTMLInputElement>> = React.forwardRef<HTMLInputElement, FormSwitchProps<FieldValues>>(function FormSwitch({ className, disabled, fieldProps, hint, id, label, name, onBlur, onChange, orientation, required, rules, input, ...props }, ref) {
  const binding = useFieldAdapter<FieldValues, boolean>({ className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules });
  const nativeInput = resolveShorthand(input);
  // Fluent otherwise merges the native-slot change before the top-level handler.
  // Notify that consumer only after this adapter has updated RHF.
  const inputProps = { ...nativeInput, onChange: undefined, ...binding.controlProps, name: binding.controller.name, ref: binding.controller.ref, onBlur: (event: React.FocusEvent<HTMLInputElement>) => { binding.controller.onBlur(); onBlur?.(event); nativeInput?.onBlur?.(event); } };
  const checked = binding.controller.value ?? false;
  return <Field {...binding.fieldProps}><Switch {...props} {...binding.controlProps} name={binding.controller.name} ref={binding.controller.ref} checked={checked}
    input={{ ...inputProps, checked: checked === true }}
    onBlur={event => { binding.controller.onBlur(); onBlur?.(event); }}
    onChange={(event, data) => { binding.controller.onChange(data.checked); onChange?.(event, data); nativeInput?.onChange?.(event); }} /></Field>;
});
FormSwitch.displayName = 'Form.Switch';

import * as React from 'react';
import { Field, Slider } from '@fluentui/react-components';
import { resolveShorthand } from '@fluentui/react-utilities';
import type { FieldValues } from 'react-hook-form';
import type { FormSliderProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';

/** Slider binding through Fluent Field; required is accessibility metadata.
 * Semantic change callbacks run after RHF, top-level first and native-slot second.
 * Blur callbacks run after RHF and the ref targets the native input.
 * @example `<Order.Slider name="value" label="Slider" />`
 */
export const FormSlider: React.ForwardRefExoticComponent<FormSliderProps<FieldValues> & React.RefAttributes<HTMLInputElement>> = React.forwardRef<HTMLInputElement, FormSliderProps<FieldValues>>(function FormSlider({ className, disabled, fieldProps, hint, id, label, name, onBlur, onChange, orientation, required, rules, input, ...props }, ref) {
  const nativeInput = resolveShorthand(input);
  const binding = useFieldAdapter<FieldValues, number>({ nativeSlotOwnsDescriptions: true, nativeId: nativeInput?.id, nativeRequired: nativeInput?.required, nativeDisabled: nativeInput?.disabled, describedBy: [props['aria-describedby'], nativeInput?.['aria-describedby']].filter(Boolean).join(' '), className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules });
  // Fluent selects a distinct input.onChange instead of its top-level handler.
  // Own the semantic path once, then notify both consumer callback locations.
  const inputProps = { ...nativeInput, onChange: undefined, ...binding.controlProps, name: binding.controller.name, ref: binding.controller.ref, onBlur: (event: React.FocusEvent<HTMLInputElement>) => { binding.controller.onBlur(); onBlur?.(event); nativeInput?.onBlur?.(event); } };
  const value = binding.controller.value;
  return <Field {...binding.fieldProps}><Slider {...props} {...binding.controlProps} name={binding.controller.name} ref={binding.controller.ref} value={value}
    input={{ ...inputProps }}
    onBlur={event => { binding.controller.onBlur(); onBlur?.(event); }}
    onChange={(event, data) => { binding.controller.onChange(data.value); onChange?.(event, data); nativeInput?.onChange?.(event); }} /></Field>;
});
FormSlider.displayName = 'Form.Slider';

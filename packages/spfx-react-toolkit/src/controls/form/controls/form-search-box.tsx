import * as React from 'react';
import { Field, SearchBox } from '@fluentui/react-components';
import { resolveShorthand } from '@fluentui/react-utilities';
import type { SearchBoxChangeEvent } from '@fluentui/react-components';
import type { FieldValues } from 'react-hook-form';
import type { FormSearchBoxProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';

/** Dismiss clicks are span mouse events, not native input changes. */
function isSearchInputChange(event: SearchBoxChangeEvent): event is React.ChangeEvent<HTMLInputElement> {
  return event.type === 'change' && event.currentTarget.tagName === 'INPUT';
}

/** SearchBox binding through Fluent Field; required is accessibility metadata.
 * Semantic change callbacks run after RHF, top-level first and native-slot second.
 * Blur callbacks run after RHF and the ref targets the native input.
 * @example `<Order.SearchBox name="value" label="SearchBox" />`
 */
export const FormSearchBox: React.ForwardRefExoticComponent<FormSearchBoxProps<FieldValues> & React.RefAttributes<HTMLInputElement>> = React.forwardRef<HTMLInputElement, FormSearchBoxProps<FieldValues>>(function FormSearchBox({ className, disabled, fieldProps, hint, id, label, name, onBlur, onChange, orientation, required, rules, input, ...props }, ref) {
  const nativeInput = resolveShorthand(input);
  const binding = useFieldAdapter<FieldValues, string>({ nativeSlotOwnsDescriptions: true, nativeId: nativeInput?.id, nativeRequired: nativeInput?.required, nativeDisabled: nativeInput?.disabled, describedBy: [props['aria-describedby'], nativeInput?.['aria-describedby']].filter(Boolean).join(' '), className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules });
  const inputProps = { ...nativeInput, onChange: undefined, ...binding.controlProps, name: binding.controller.name, ref: binding.controller.ref, onBlur: (event: React.FocusEvent<HTMLInputElement>) => { binding.controller.onBlur(); onBlur?.(event); nativeInput?.onBlur?.(event); } };
  const value = binding.controller.value ?? '';
  return <Field {...binding.fieldProps}><SearchBox {...props} {...binding.controlProps} name={binding.controller.name} ref={binding.controller.ref} value={value}
    input={{ ...inputProps }}
    onBlur={event => { binding.controller.onBlur(); onBlur?.(event); }}
    onChange={(event, data) => { binding.controller.onChange(data.value); onChange?.(event, data); if (isSearchInputChange(event)) nativeInput?.onChange?.(event); }} /></Field>;
});
FormSearchBox.displayName = 'Form.SearchBox';

import * as React from 'react';
import { Field, Textarea } from '@fluentui/react-components';
import { resolveShorthand } from '@fluentui/react-utilities';
import type { FieldValues } from 'react-hook-form';
import type { FormTextareaProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';

/** Textarea binding through Fluent Field; required is accessibility metadata.
 * Semantic change callbacks run after RHF, top-level first and native-slot second.
 * Blur callbacks run after RHF and the ref targets the native textarea.
 * @example `<Order.Textarea name="notes" label="Textarea" />`
 */
export const FormTextarea: React.ForwardRefExoticComponent<FormTextareaProps<FieldValues> & React.RefAttributes<HTMLTextAreaElement>> = React.forwardRef<HTMLTextAreaElement, FormTextareaProps<FieldValues>>(function FormTextarea({ className, disabled, fieldProps, hint, id, label, name, onBlur, onChange, orientation, required, rules, textarea, ...props }, ref) {
  const nativeTextarea = resolveShorthand(textarea);
  const binding = useFieldAdapter<FieldValues, string>({ nativeSlotOwnsDescriptions: true, nativeId: nativeTextarea?.id, nativeRequired: nativeTextarea?.required, nativeDisabled: nativeTextarea?.disabled, describedBy: [props['aria-describedby'], nativeTextarea?.['aria-describedby']].filter(Boolean).join(' '), className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules });
  const textareaProps = { ...nativeTextarea, onChange: undefined, ...binding.controlProps, name: binding.controller.name, ref: binding.controller.ref, onBlur: (event: React.FocusEvent<HTMLTextAreaElement>) => { binding.controller.onBlur(); onBlur?.(event); nativeTextarea?.onBlur?.(event); } };
  const value = binding.controller.value ?? '';
  return <Field {...binding.fieldProps}><Textarea {...props} {...binding.controlProps} name={binding.controller.name} ref={binding.controller.ref} value={value}
    textarea={{ ...textareaProps, value }}
    onBlur={event => { binding.controller.onBlur(); onBlur?.(event); }}
    onChange={(event, data) => { binding.controller.onChange(data.value); onChange?.(event, data); nativeTextarea?.onChange?.(event); }} /></Field>;
});
FormTextarea.displayName = 'Form.Textarea';

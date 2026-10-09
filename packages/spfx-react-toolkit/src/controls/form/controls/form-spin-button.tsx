import * as React from 'react';
import { Field, SpinButton } from '@fluentui/react-components';
import { resolveShorthand, useMergedRefs } from '@fluentui/react-utilities';
import { resolveSpinButtonValue } from '../resolve-spin-button-value.internal';
import type { FieldValues } from 'react-hook-form';
import type { FormSpinButtonProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';

/** SpinButton binding through Fluent Field; required is accessibility metadata.
 * Change/blur callbacks run after RHF and the ref targets the native input.
 * Real resets clear uncommitted text unless RHF preserves this field's values.
 * Fluent has no public draft reset API, so these resets replace only its native
 * input; the forwarded ref updates and focus returns if the input was focused.
 * @example `<Order.SpinButton name="value" label="SpinButton" />`
 */
export const FormSpinButton: React.ForwardRefExoticComponent<FormSpinButtonProps<FieldValues> & React.RefAttributes<HTMLInputElement>> = React.forwardRef<HTMLInputElement, FormSpinButtonProps<FieldValues>>(function FormSpinButton({ className, disabled, fieldProps, hint, id, label, name, onBlur, onChange, orientation, required, rules, input, incrementButton, decrementButton, ...props }, ref) {
  // Fluent's empty committed numeric value is explicitly nullable.
  // eslint-disable-next-line @rushstack/no-new-null
  const binding = useFieldAdapter<FieldValues, number | null>({ className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules });
  const nativeElement = React.useRef<HTMLInputElement>();
  const captureNativeElement = React.useCallback<React.RefCallback<HTMLInputElement>>(node => { nativeElement.current = node ?? undefined; }, []);
  const nativeRef = useMergedRefs<HTMLInputElement>(binding.controller.ref, captureNativeElement);
  const [primitiveVersion, setPrimitiveVersion] = React.useState(0);
  const observedReset = React.useRef(binding.resetVersion);
  const restoreFocus = React.useRef(false);
  React.useLayoutEffect(() => {
    if (observedReset.current === binding.resetVersion) return;
    observedReset.current = binding.resetVersion;
    if (binding.resetPreservesValues) return;
    restoreFocus.current = nativeElement.current !== undefined && nativeElement.current.ownerDocument.activeElement === nativeElement.current;
    // Fluent 9.72.3 always prioritizes internal textValue over controlled displayValue.
    // Its native change/input hooks retain that draft and no public API clears it.
    // Reset only the primitive, preserving this RHF controller and its field owner.
    setPrimitiveVersion(previous => previous + 1);
  }, [binding.resetVersion, binding.resetPreservesValues]);
  React.useLayoutEffect(() => {
    if (!restoreFocus.current) return;
    restoreFocus.current = false;
    nativeElement.current?.focus();
  }, [primitiveVersion]);
  const nativeInput = resolveShorthand(input);
  const inputProps = { ...nativeInput, ...binding.controlProps, name: binding.controller.name, ref: nativeRef, onBlur: (event: React.FocusEvent<HTMLInputElement>) => { binding.controller.onBlur(); onBlur?.(event); nativeInput?.onBlur?.(event); } };
  const { disabled: incrementConsumerDisabled, ...incrementProps } = resolveShorthand(incrementButton) ?? {};
  const { disabled: decrementConsumerDisabled, ...decrementProps } = resolveShorthand(decrementButton) ?? {};
  // Leave Fluent's rounded range-bound defaults intact; slots may only add disabling.
  // The Fluent nullable value keeps empty commits controlled rather than restoring zero.
  // eslint-disable-next-line @rushstack/no-new-null
  const value = binding.controller.value ?? null;
  return <Field {...binding.fieldProps}><SpinButton key={primitiveVersion} {...props} {...binding.controlProps} name={binding.controller.name} ref={nativeRef} value={value}
    input={{ ...inputProps }}
    incrementButton={{ ...incrementProps, ...(binding.controlProps.disabled || incrementConsumerDisabled ? { disabled: true } : {}) }} decrementButton={{ ...decrementProps, ...(binding.controlProps.disabled || decrementConsumerDisabled ? { disabled: true } : {}) }}
    onBlur={event => { binding.controller.onBlur(); onBlur?.(event); }}
    onChange={(event, data) => { binding.controller.onChange(resolveSpinButtonValue(data)); onChange?.(event, data); }} /></Field>;
});
FormSpinButton.displayName = 'Form.SpinButton';

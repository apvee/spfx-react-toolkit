import * as React from 'react';
import { Dropdown, Field } from '@fluentui/react-components';
import { resolveShorthand, useMergedRefs } from '@fluentui/react-utilities';
import type { FieldValues } from 'react-hook-form';
import type { FormDropdownProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';
import { getSelectionSource, getSelectionText, readSelectionValue, renderSelectionOptions, selectionIsDisabled, writeSelectionValue } from '../selection-options.internal';

/** Fluent single selection with independent identity/display text.
 * RHF writes precede option callbacks; the public ref targets the trigger button.
 * @example `<Order.Dropdown name="country" label="Country" options={countries} />`
 */
export const FormDropdown = React.forwardRef<HTMLButtonElement, FormDropdownProps<FieldValues>>(function FormDropdown(
  {
    button,
    children,
    className,
    codec,
    disabled,
    fieldProps,
    hint,
    id,
    label,
    listbox,
    name,
    onBlur,
    onOptionSelect,
    options,
    orientation,
    required,
    rules,
    ...props
  },
  ref
) {
  const source = getSelectionSource(options, children);
  const binding = useFieldAdapter<FieldValues, unknown>({ className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules });
  const value = readSelectionValue<string>(binding.controller.value, codec, '');
  const nativeListbox = resolveShorthand(listbox);
  const optionChildren = renderSelectionOptions(source);
  const native = resolveShorthand(button);
  const nativeRef = useMergedRefs(binding.controller.ref, native?.ref);
  const display = getSelectionText(source, value);
  return <Field {...binding.fieldProps}>
    <Dropdown {...props} {...binding.controlProps}
      name={binding.name}
      ref={nativeRef}
      multiselect={false}
      listbox={{ ...nativeListbox, children: optionChildren, multiselect: false }}
      selectedOptions={value === '' ? [] : [value]}
      value={display}
      button={{
        ...native, ...binding.controlProps, name: binding.name, ref: nativeRef, children: display || props.placeholder, onBlur: event => {
          binding.controller.onBlur();
          onBlur?.(event);
          native?.onBlur?.(event);
        }
      }}
      onOptionSelect={(event, data) => {
        if (binding.controlProps.disabled || (data.optionValue !== undefined && selectionIsDisabled(source, data.optionValue))) return;
        binding.controller.onChange(writeSelectionValue(data.optionValue ?? '', codec));
        onOptionSelect?.(event, data);
      }}>
      {optionChildren}
    </Dropdown></Field>;
});
FormDropdown.displayName = 'Form.Dropdown';

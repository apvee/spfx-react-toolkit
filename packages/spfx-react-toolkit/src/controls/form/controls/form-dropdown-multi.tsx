import * as React from 'react';
import { Dropdown, Field } from '@fluentui/react-components';
import { resolveShorthand, useMergedRefs } from '@fluentui/react-utilities';
import type { FieldValues } from 'react-hook-form';
import type { FormDropdownMultiProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';
import { getSelectionSource, getSelectionText, readSelectionValue, renderSelectionOptions, selectionIsDisabled, writeSelectionValue } from '../selection-options.internal';

/** Fluent ordered multiple selection with independent identity/display text.
 * RHF writes precede option callbacks; the public ref targets the trigger button.
 * @example `<Order.DropdownMulti name="countries" label="Country" options={countries} />`
 */
export const FormDropdownMulti = React.forwardRef<HTMLButtonElement, FormDropdownMultiProps<FieldValues>>(function FormDropdownMulti(
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
  const native = resolveShorthand(button);
  const binding = useFieldAdapter<FieldValues, unknown>({ nativeSlotOwnsDescriptions: true, nativeId: native?.id, nativeDisabled: native?.disabled, describedBy: [props['aria-describedby'], native?.['aria-describedby']].filter(Boolean).join(' '), className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules });
  const value = readSelectionValue<string[]>(binding.controller.value, codec, []);
  const nativeListbox = resolveShorthand(listbox);
  const optionChildren = renderSelectionOptions(source);
  const nativeRef = useMergedRefs(binding.controller.ref, native?.ref);
  const display = value.map(identity => getSelectionText(source, identity)).filter(Boolean).join(', ');
  return <Field {...binding.fieldProps}>
    <Dropdown {...props} {...binding.controlProps}
      name={binding.name}
      ref={nativeRef}
      multiselect={true}
      listbox={{ ...nativeListbox, children: optionChildren, multiselect: true }}
      selectedOptions={value}
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
        binding.controller.onChange(writeSelectionValue(data.selectedOptions, codec));
        onOptionSelect?.(event, data);
      }}>
      {optionChildren}
    </Dropdown></Field>;
});
FormDropdownMulti.displayName = 'Form.DropdownMulti';

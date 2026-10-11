import * as React from 'react';
import { Field, Select } from '@fluentui/react-components';
import { resolveShorthand, useMergedRefs } from '@fluentui/react-utilities';
import type { FieldValues } from 'react-hook-form';
import type { FormSelectProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';
import { getOptionText, getSelectionSource, readSelectionValue, selectionIsDisabled, writeSelectionValue } from '../selection-options.internal';

/** Native single-selection binding; identities remain independent of option labels.
 * RHF writes precede top-level and native select callbacks; the ref targets the select.
 * @example `<Order.Select name="country" label="Country" options={countries} />`
 */
export const FormSelect = React.forwardRef<HTMLSelectElement, FormSelectProps<FieldValues>>(function FormSelect(
  {
    children,
    className,
    codec,
    disabled,
    fieldProps,
    hint,
    id,
    label,
    name,
    onBlur,
    onChange,
    options,
    orientation,
    required,
    rules,
    select,
    ...props
  },
  ref
) {
  const source = getSelectionSource(options, children);
  const native = resolveShorthand(select);
  const binding = useFieldAdapter<FieldValues, unknown>({ nativeSlotOwnsDescriptions: true, nativeId: native?.id, nativeRequired: native?.required, nativeDisabled: native?.disabled, describedBy: [props['aria-describedby'], native?.['aria-describedby']].filter(Boolean).join(' '), className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules });
  const value = readSelectionValue(binding.controller.value, codec, '');
  const nativeRef = useMergedRefs(binding.controller.ref, native?.ref);
  const optionChildren = source.options === undefined ? source.children : source.options.map(option => (
    <option key={option.value} value={option.value} disabled={option.disabled}>
      {typeof option.label === 'string' || typeof option.label === 'number' ? option.label : getOptionText(option)}
    </option>
  ));
  return <Field {...binding.fieldProps}>
    <Select {...props} {...binding.controlProps}
      name={binding.name}
      ref={nativeRef}
      value={value}
      select={{
        ...native, ...binding.controlProps, name: binding.name, ref: nativeRef, value, defaultValue: undefined, multiple: false, children: optionChildren, onChange: undefined, onBlur: event => {
          binding.controller.onBlur();
          onBlur?.(event);
          native?.onBlur?.(event);
        }
      }}
      onChange={(event, data) => {
        if (binding.controlProps.disabled || selectionIsDisabled(source, data.value)) return;
        binding.controller.onChange(writeSelectionValue(data.value, codec));
        onChange?.(event, data);
        native?.onChange?.(event);
      }}>
      {optionChildren}
    </Select></Field>;
});
FormSelect.displayName = 'Form.Select';

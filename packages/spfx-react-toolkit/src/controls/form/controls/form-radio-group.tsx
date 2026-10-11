import * as React from 'react';
import { Field, Radio, RadioGroup, radioClassNames } from '@fluentui/react-components';
import { useSx } from '../../../hooks/useSx';
import * as borderWidth from '../../../helpers/styles/border-width';
import * as padding from '../../../helpers/styles/padding';
import * as margin from '../../../helpers/styles/margin';
import * as minWidth from '../../../helpers/styles/min-width';
import type { FieldValues } from 'react-hook-form';
import type { FormRadioGroupProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';
import { useCompositeFieldFocus } from '../use-composite-field-focus.internal';
import { getSelectionSource, readSelectionValue, selectionIsDisabled, writeSelectionValue } from '../selection-options.internal';

/** Single identity binding with a real group ref and eligible descendant error focus.
 * Fluent Radio children are rendered unchanged; group change/blur callbacks follow RHF.
 * RHF native validation and error focus use the first eligible native descendant.
 * @example `<Order.RadioGroup name="priority" label="Priority" options={priorities} />`
 */
export const FormRadioGroup = React.forwardRef<HTMLDivElement, FormRadioGroupProps<FieldValues>>(function FormRadioGroup(
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
    ...props
  },
  ref
) {
  const sx = useSx();
  const source = getSelectionSource(options, children);
  const binding = useFieldAdapter<FieldValues, unknown>({ settingsLabelTarget: 'group', describedBy: props['aria-describedby'], className, disabled, fieldProps, hint, id, label, name, orientation, required, rules });
  const compositeRef = useCompositeFieldFocus(
    ref,
    binding.focusRef,
    `input.${radioClassNames.input}[type="radio"]`,
    element => element.name === binding.name && source.entries.some(option => option.value === element.value && option.disabled !== true)
  );
  const value = readSelectionValue(binding.controller.value, codec, '');
  // Native fieldset disabling also covers consumer Radio children whose native
  // input slots override Fluent's inherited disabled context; children stay intact.
  return <Field {...binding.fieldProps}>
    <fieldset disabled={binding.controlProps.disabled}
      className={sx(borderWidth.none, padding.none, margin.none, minWidth.zero)} role="presentation">
      <RadioGroup {...props} {...binding.controlProps}
        name={binding.name}
        ref={compositeRef}
        value={value}
        onBlur={event => {
          binding.controller.onBlur();
          onBlur?.(event);
        }}
        onChange={(event, data) => {
          if (binding.controlProps.disabled || selectionIsDisabled(source, data.value)) return;
          binding.controller.onChange(writeSelectionValue(data.value, codec));
          onChange?.(event, data);
        }}>
        {source.options === undefined ? source.children : source.options.map(option => <Radio key={option.value}
          value={option.value} label={{ children: option.label }} disabled={binding.controlProps.disabled || option.disabled} />)}
      </RadioGroup></fieldset></Field>;
});
FormRadioGroup.displayName = 'Form.RadioGroup';

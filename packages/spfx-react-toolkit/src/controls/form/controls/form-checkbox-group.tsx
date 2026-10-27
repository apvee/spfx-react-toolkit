import * as React from 'react';
import { mergeSettingsFieldIds } from '../../settings-field-context.internal';
import { Checkbox, Field, FieldContextProvider, checkboxClassNames } from '@fluentui/react-components';
import { resolveShorthand } from '@fluentui/react-utilities';
import type { CheckboxProps } from '@fluentui/react-components';
import type { FieldValues } from 'react-hook-form';
import type { FormCheckboxGroupProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';
import { useCompositeFieldFocus } from '../use-composite-field-focus.internal';
import { getSelectionSource, readSelectionValue, selectionIsDisabled, writeSelectionValue } from '../selection-options.internal';

/** Visible multiple selections in declared option order, wrapped in an accessible group.
 * Only direct Fluent Checkbox children are adapted. RHF writes precede group,
 * child and native-input callbacks; the public ref targets the real group div.
 * Disabling is additive across root, scope, child, native slot and option identity.
 * RHF native validation and error focus use the first eligible native descendant.
 * @example `<Order.CheckboxGroup name="labels" label="Labels" options={labels} />`
 */
export const FormCheckboxGroup = React.forwardRef<HTMLDivElement, FormCheckboxGroupProps<FieldValues>>(function FormCheckboxGroup(
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
  const source = getSelectionSource(options, children);
  const binding = useFieldAdapter<FieldValues, unknown>({ settingsLabelTarget: 'group', describedBy: props['aria-describedby'], className, disabled, fieldProps, hint, id, label, name, orientation, required, rules });
  const compositeRef = useCompositeFieldFocus(
    ref,
    binding.focusRef,
    `input.${checkboxClassNames.input}[type="checkbox"]`,
    element => element.name === binding.name && source.entries.some(option => option.value === element.value && option.disabled !== true)
  );
  const selected = readSelectionValue<string[]>(binding.controller.value, codec, []);
  const nodes = source.options === undefined ? source.children : source.options.map(option => <Checkbox key={option.value}
    value={option.value} label={{ children: option.label }} disabled={option.disabled} />);
  const declared = source.entries.map(option => option.value);
  const bound = React.Children.map(nodes, child => {
    if (!React.isValidElement<CheckboxProps>(child) || child.type !== Checkbox || typeof child.props.value !== 'string') return child;
    const value = child.props.value;
    const native = resolveShorthand(child.props.input);
    const childDisabled = binding.controlProps.disabled === true || child.props.disabled === true || native?.disabled === true || selectionIsDisabled(source, value);
    const blur = (event: React.FocusEvent<HTMLInputElement>): void => {
      binding.controller.onBlur();
      onBlur?.(event);
      child.props.onBlur?.(event);
      native?.onBlur?.(event);
    };
    return React.cloneElement(child, {
      checked: selected.includes(value), disabled: childDisabled,
      input: { ...native, name: binding.name, value, disabled: childDisabled, checked: selected.includes(value), onChange: undefined, onBlur: blur },
      onBlur: undefined,
      onChange: (event: Parameters<NonNullable<CheckboxProps['onChange']>>[0], data: Parameters<NonNullable<CheckboxProps['onChange']>>[1]) => {
        if (childDisabled) return;
        const nextSet = new Set(selected);
        if (data.checked === true) nextSet.add(value); else nextSet.delete(value);
        binding.controller.onChange(writeSelectionValue(declared.filter(identity => nextSet.has(identity)), codec));
        onChange?.(event, data); child.props.onChange?.(event, data);
        native?.onChange?.(event);
      }
    });
  });
  return <Field {...binding.fieldProps}>{fieldControl => <div {...props} {...fieldControl}
    aria-labelledby={binding.controlProps['aria-labelledby'] ?? fieldControl['aria-labelledby']} id={binding.controlProps.id}
    aria-describedby={mergeSettingsFieldIds(fieldControl['aria-describedby'], binding.controlProps['aria-describedby'], props['aria-describedby'])}
    aria-disabled={binding.controlProps.disabled || undefined}
    ref={compositeRef} role="group">
    <FieldContextProvider
      value={undefined}>{bound}</FieldContextProvider></div>}</Field>;
});
FormCheckboxGroup.displayName = 'Form.CheckboxGroup';

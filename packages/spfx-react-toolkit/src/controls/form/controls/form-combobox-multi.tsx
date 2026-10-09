import * as React from 'react';
import { Combobox, Field } from '@fluentui/react-components';
import { resolveShorthand, useMergedRefs } from '@fluentui/react-utilities';
import type { FieldValues } from 'react-hook-form';
import type { FormComboboxMultiProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';
import { getSelectionSource, getSelectionText, readSelectionValue, renderSelectionOptions, selectionIsDisabled, writeSelectionValue } from '../selection-options.internal';

/** Editable multiple selection separating query from committed identities.
 * At rest, selected labels appear inside the input. Focus shows the search query;
 * leaving the input clears that transient query and restores the labels without changing selection.
 * Targeted reset clears query; explicit dirty-owner preservation retains it.
 * Explicit reset with keepDirtyValues also retains dirty queries through array row remounts.
 * @example `<Order.ComboboxMulti name="countries" label="Country" options={countries} />`
 */
export const FormComboboxMulti = React.forwardRef<HTMLInputElement, FormComboboxMultiProps<FieldValues>>(function FormComboboxMulti(
  {
    children,
    className,
    codec,
    disabled,
    fieldProps,
    hint,
    id,
    input,
    label,
    listbox,
    name,
    onBlur,
    onChange,
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
  const [display, setDisplay] = React.useState('');
  const [editing, setEditing] = React.useState(false);
  const binding = useFieldAdapter<FieldValues, unknown>({ className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules,
    resetText: { read: () => display, restore: text => { setDisplay(text); setEditing(text.length > 0); } }
  });
  const value = readSelectionValue<string[]>(binding.controller.value, codec, []);
  const summary = value.map(identity => getSelectionText(source, identity)).filter(Boolean).join(', ');
  const inputValue = editing ? display : summary;
  const valueKey = JSON.stringify(value);
  const previous = React.useRef({ value: valueKey, reset: binding.resetVersion });
  React.useLayoutEffect(() => {
    const didReset = previous.current.reset !== binding.resetVersion;
    const didChange = previous.current.value !== valueKey;
    previous.current = { value: valueKey, reset: binding.resetVersion };
    if (didReset ? !binding.resetPreservesValues : didChange) setDisplay('');
  }, [valueKey, binding.resetVersion, binding.resetPreservesValues]);
  const nativeListbox = resolveShorthand(listbox);
  const optionChildren = renderSelectionOptions(source);
  const native = resolveShorthand(input);
  const nativeRef = useMergedRefs(binding.controller.ref, native?.ref);
  const queryChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    if (binding.controlProps.disabled) return;
    const query = event.currentTarget.value;
    setEditing(true);
    setDisplay(query);
    onChange?.(event);
    native?.onChange?.(event);
  };
  return <Field {...binding.fieldProps}>{fieldControl =>
    <Combobox {...fieldControl} {...props} {...binding.controlProps}
      name={binding.name}
      ref={nativeRef}
      multiselect={true}
      freeform={false}
      listbox={{ ...nativeListbox, children: optionChildren, multiselect: true }}
      selectedOptions={value}
      value={inputValue}
      input={{
        ...native, ...binding.controlProps, name: binding.name, ref: nativeRef, value: inputValue, defaultValue: undefined, 'aria-describedby': [fieldControl['aria-describedby'], props['aria-describedby'], native?.['aria-describedby']].filter(Boolean).join(' ') || undefined, onChange: queryChange, onFocus: event => {
          setEditing(true);
          native?.onFocus?.(event);
        }, onBlur: event => {
          setEditing(false);
          setDisplay('');
          binding.controller.onBlur();
          onBlur?.(event);
          native?.onBlur?.(event);
        }
      }}
      onOptionSelect={(event, data) => {
        if (binding.controlProps.disabled || (data.optionValue !== undefined && selectionIsDisabled(source, data.optionValue))) return;
        const next = data.selectedOptions;
        binding.controller.onChange(writeSelectionValue(next, codec));
        setDisplay('');
        onOptionSelect?.(event, data);
      }}>
      {optionChildren}
    </Combobox>
  }</Field>;
});
FormComboboxMulti.displayName = 'Form.ComboboxMulti';

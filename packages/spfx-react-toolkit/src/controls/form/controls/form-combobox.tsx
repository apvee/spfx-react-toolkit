import * as React from 'react';
import { Combobox, Field } from '@fluentui/react-components';
import { resolveShorthand, useMergedRefs } from '@fluentui/react-utilities';
import type { FieldValues } from 'react-hook-form';
import type { FormComboboxProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';
import { getSelectionSource, getSelectionText, readSelectionValue, renderSelectionOptions, selectionIsDisabled, writeSelectionValue } from '../selection-options.internal';

/** Editable single selection separating query from committed identity.
 * A query differing from selected text clears identity before query callbacks.
 * Targeted reset clears query; explicit dirty-owner preservation retains it.
 * Explicit reset with keepDirtyValues also retains dirty queries through array row remounts.
 * @example `<Order.Combobox name="country" label="Country" options={countries} />`
 */
export const FormCombobox = React.forwardRef<HTMLInputElement, FormComboboxProps<FieldValues>>(function FormCombobox(
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
  // The layout synchronization below derives selected labels from RHF before paint.
  const [display, setDisplay] = React.useState('');
  const editing = React.useRef(false);
  const native = resolveShorthand(input);
  const binding = useFieldAdapter<FieldValues, unknown>({ nativeSlotOwnsDescriptions: true, nativeId: native?.id, nativeRequired: native?.required, nativeDisabled: native?.disabled, describedBy: [props['aria-describedby'], native?.['aria-describedby']].filter(Boolean).join(' '), className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules,
    resetText: { read: () => editing.current ? display : undefined, restore: text => { editing.current = true; setDisplay(text); } }
  });
  const value = readSelectionValue<string>(binding.controller.value, codec, '');
  const selectedText = getSelectionText(source, value);
  const valueKey = value;
  const previous = React.useRef({ value: valueKey, reset: binding.resetVersion });
  const pendingLocal = React.useRef<string>();
  React.useLayoutEffect(() => {
    const didReset = previous.current.reset !== binding.resetVersion;
    const didChange = previous.current.value !== valueKey;
    previous.current = { value: valueKey, reset: binding.resetVersion };
    const local = pendingLocal.current === valueKey;
    pendingLocal.current = undefined;
    if (didReset) {
      if (binding.resetPreservesValues) return;
      editing.current = false;
      setDisplay(selectedText);
    } else if (didChange && !local) { editing.current = false; setDisplay(selectedText); }
    else if (!editing.current) setDisplay(selectedText);
  }, [valueKey, selectedText, binding.resetVersion, binding.resetPreservesValues]);
  const nativeListbox = resolveShorthand(listbox);
  const optionChildren = renderSelectionOptions(source);
  const nativeRef = useMergedRefs(binding.controller.ref, native?.ref);
  const queryChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    if (binding.controlProps.disabled) return;
    const query = event.currentTarget.value;
    editing.current = true;
    setDisplay(query);
    if (value !== '' && query !== selectedText) {
      pendingLocal.current = '';
      binding.controller.onChange(writeSelectionValue('', codec));
    }
    onChange?.(event);
    native?.onChange?.(event);
  };
  return <Field {...binding.fieldProps}>
    <Combobox {...props} {...binding.controlProps}
      name={binding.name}
      ref={nativeRef}
      multiselect={false}
      freeform={false}
      listbox={{ ...nativeListbox, children: optionChildren, multiselect: false }}
      selectedOptions={value === '' ? [] : [value]}
      value={display}
      input={{
        ...native, ...binding.controlProps, name: binding.name, ref: nativeRef, value: display, defaultValue: undefined, onChange: queryChange, onBlur: event => {
          binding.controller.onBlur();
          onBlur?.(event);
          native?.onBlur?.(event);
        }
      }}
      onOptionSelect={(event, data) => {
        if (binding.controlProps.disabled || (data.optionValue !== undefined && selectionIsDisabled(source, data.optionValue))) return;
        if (data.optionValue === undefined && event.type === 'change') { onOptionSelect?.(event, data); return; }
        if (editing.current && event.type === 'keydown' && 'key' in event && event.key === 'Tab') return;
        const next = data.optionValue ?? '';
        pendingLocal.current = next;
        editing.current = false;
        binding.controller.onChange(writeSelectionValue(next, codec));
        setDisplay(getSelectionText(source, next));
        onOptionSelect?.(event, data);
      }}>
      {optionChildren}
    </Combobox></Field>;
});
FormCombobox.displayName = 'Form.Combobox';

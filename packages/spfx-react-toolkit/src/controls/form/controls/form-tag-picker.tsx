import * as React from 'react';
import { mergeSettingsFieldIds } from '../../settings-field-context.internal';
import { Field, Tag, TagPicker, TagPickerControl, TagPickerGroup, TagPickerInput, TagPickerList, TagPickerOption } from '@fluentui/react-components';
import { mergeClasses } from '@griffel/react';
import { renderContentSlot } from '../../content-slot.internal';
import type { FieldValues } from 'react-hook-form';
import type { FormTagPickerProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';
import { getOptionText, getSelectionSource, getSelectionText, readSelectionValue, selectionIsDisabled, writeSelectionValue } from '../selection-options.internal';

/** Removable selected tags with a separate query and actual native input ref.
 * Naming comes from a plain text label, aria-label or an enclosing SettingsGroup.Item field.
 * Unnamed standalone usage throws; inherited naming also names the selected-tag list.
 * RHF stores string identities or an explicit codec's live domain. Native Fluent
 * keyboard removal writes through RHF before the consumer selection callback.
 * Explicit reset with keepDirtyValues retains dirty queries through array row remounts.
 * @example `<Order.TagPicker name="people" label="People" options={people} strings={{ noOptions: 'No people found' }} />`
 */
export const FormTagPicker = React.forwardRef<HTMLInputElement, FormTagPickerProps<FieldValues>>(function FormTagPicker(
  {
    'aria-label': ariaLabel,
    children,
    className,
    codec,
    disabled,
    fieldProps,
    hint,
    id,
    label,
    name,
    noOptionsContent,
    onBlur,
    onChange,
    onOptionSelect,
    options,
    orientation,
    required,
    rules,
    strings,
    ...props
  },
  ref
) {
  const source = getSelectionSource(options, children);
  const [query, setQuery] = React.useState('');
  const binding = useFieldAdapter<FieldValues, unknown>({ describedBy: props['aria-describedby'], className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules,
    resetText: { read: () => query, restore: text => setQuery(text) }
  });
  const selected = readSelectionValue<string[]>(binding.controller.value, codec, []);
  const valueKey = JSON.stringify(selected);
  const previous = React.useRef({ value: valueKey, reset: binding.resetVersion });
  React.useLayoutEffect(() => {
    const reset = previous.current.reset !== binding.resetVersion;
    const changed = previous.current.value !== valueKey;
    previous.current = { value: valueKey, reset: binding.resetVersion };
    if (reset ? !binding.resetPreservesValues : changed) setQuery('');
  }, [binding.resetVersion, binding.resetPreservesValues, valueKey]);
  const emptyContent = noOptionsContent ?? strings?.noOptions ?? 'No options match the current query.';
  const normalizedQuery = query.trim().toLocaleLowerCase();
  const available = source.entries.filter(option => !selected.includes(option.value) && getOptionText(option).toLocaleLowerCase().includes(normalizedQuery));
  const accessibleName = ariaLabel ?? (typeof label === 'string' ? label : undefined);
  if (!accessibleName?.trim() && !binding.controlProps['aria-labelledby']) throw new TypeError('Form.TagPicker requires a plain text label, aria-label, or SettingsGroup.Item field name.');
  const renderedOptions = source.options === undefined ? source.children : available.length ? available.map(option => <TagPickerOption key={option.value}
    value={option.value} text={getOptionText(option)}
    aria-disabled={option.disabled || undefined}>{option.label}</TagPickerOption>) : <div role="status">{renderContentSlot(emptyContent, 'div', mergeClasses, { shorthand: 'preserve' })}</div>;
  return <Field {...binding.fieldProps}>{fieldControl => <TagPicker {...props} disabled={binding.controlProps.disabled}
    selectedOptions={selected}
    onOptionSelect={(event, data) => {
      if (binding.controlProps.disabled || (data.value !== undefined && !selected.includes(data.value) && selectionIsDisabled(source, data.value))) return;
      binding.controller.onChange(writeSelectionValue(data.selectedOptions, codec));
      setQuery('');
      onOptionSelect?.(event, data);
    }}>
    <TagPickerControl>
      <TagPickerGroup aria-label={accessibleName} aria-labelledby={binding.controlProps['aria-labelledby']}>
        {selected.map(value => <Tag key={value}
          value={value} dismissible>{source.entries.find(option => option.value === value)?.label ?? getSelectionText(source, value)}</Tag>)}
      </TagPickerGroup>
      <TagPickerInput {...fieldControl} {...binding.controlProps}
        name={binding.name} aria-label={accessibleName}
        aria-describedby={mergeSettingsFieldIds(binding.controlProps['aria-describedby'], props['aria-describedby'])}
        ref={binding.controller.ref}
        value={query}
        onBlur={event => {
          binding.controller.onBlur();
          onBlur?.(event);
        }}
        onChange={event => {
          if (binding.controlProps.disabled) return;
          setQuery(event.currentTarget.value);
          onChange?.(event);
        }} />
    </TagPickerControl>
    <TagPickerList>{renderedOptions}</TagPickerList>
  </TagPicker>}</Field>;
});
FormTagPicker.displayName = 'Form.TagPicker';

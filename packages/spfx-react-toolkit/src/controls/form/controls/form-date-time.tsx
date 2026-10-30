import * as React from 'react';
import { mergeSettingsFieldIds } from '../../settings-field-context.internal';
import { SettingsFieldContext } from '../../settings-field-context.internal';
import { Field } from '@fluentui/react-components';
import type { FieldValues } from 'react-hook-form';
import type { FormDateTimeProps } from '../form-control.types';
import type { FormDateTimeValue } from '../form-value.types';
import { useFormRuntimeContext } from '../form-context.internal';
import { useFieldAdapter } from '../use-field-adapter.internal';
import { useCompositeFieldFocus } from '../use-composite-field-focus.internal';
import { FormDatePicker } from './form-date-picker';
import { FormTimePicker } from './form-time-picker';

/** Independent local date/time parts in one semantic field group.
 * Each part commits against the latest current absolute parent value. Clearing
 * one retains the other; no implicit date, midnight, timezone or DST conversion.
 * Public ref targets the group; parent errors focus its first eligible part input.
 * @example `<Order.DateTime name="appointment" label="Appointment" dateLabel="Date" timeLabel="Time" />`
 */
export const FormDateTime = React.forwardRef<HTMLDivElement, FormDateTimeProps<FieldValues>>(function FormDateTime({ id, 'aria-describedby': describedBy, className, disabled, fieldProps, hint, label, name, orientation, required, rules, dateLabel, timeLabel, datePickerProps, timePickerProps }, ref) {
  const { methods } = useFormRuntimeContext();
  const binding = useFieldAdapter<FieldValues, FormDateTimeValue>({ settingsLabelTarget: 'group', id, describedBy, className, disabled, fieldProps, hint, label, name, orientation, required, rules });
  const acceptsInput = React.useCallback((element: HTMLInputElement) => element.name === `${binding.name}.date` || element.name === `${binding.name}.time`, [binding.name]);
  const groupRef = useCompositeFieldFocus(ref, binding.focusRef, 'input', acceptsInput);
  const currentValue = (): FormDateTimeValue => {
    const current = methods.getValues(binding.name) as FormDateTimeValue | undefined;
    // Both compat selection primitives explicitly use null for their empty parts.
    // eslint-disable-next-line @rushstack/no-new-null
    return current ?? { date: null, time: null };
  };
  return <Field {...binding.fieldProps}>{controlProps => <div {...controlProps} aria-labelledby={binding.controlProps['aria-labelledby'] ?? controlProps['aria-labelledby']} aria-disabled={binding.controlProps.disabled || undefined} aria-describedby={mergeSettingsFieldIds(controlProps['aria-describedby'], binding.controlProps['aria-describedby'], describedBy)} id={binding.controlProps.id} ref={groupRef} role="group" tabIndex={-1}>
    <SettingsFieldContext.Provider value={undefined}><FormDatePicker {...datePickerProps} name={`${name}.date`} label={dateLabel} disabled={binding.controlProps.disabled || datePickerProps?.disabled}
      onBlur={event => { binding.controller.onBlur(); datePickerProps?.onBlur?.(event); }}
      onSelectDate={date => { binding.controller.onChange({ ...currentValue(), date }); datePickerProps?.onSelectDate?.(date); }} />
    <FormTimePicker {...timePickerProps} name={`${name}.time`} label={timeLabel} disabled={binding.controlProps.disabled || timePickerProps?.disabled}
      onBlur={event => { binding.controller.onBlur(); timePickerProps?.onBlur?.(event); }}
      onTimeChange={(event, data) => {
        if (data.error === undefined || data.selectedTimeText === '') binding.controller.onChange({ ...currentValue(), time: data.selectedTime });
        timePickerProps?.onTimeChange?.(event, data);
      }} />
    </SettingsFieldContext.Provider>
  </div>}</Field>;
});
FormDateTime.displayName = 'Form.DateTime';

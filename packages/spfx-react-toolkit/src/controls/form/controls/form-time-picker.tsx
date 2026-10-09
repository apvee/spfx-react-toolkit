import * as React from 'react';
import { Field } from '@fluentui/react-components';
import { useMergedRefs } from '@fluentui/react-utilities';
import { TimePicker, formatDateToTimeString as formatCompatTime } from '@fluentui/react-timepicker-compat';
import type { FieldValues } from 'react-hook-form';
import type { FormTimePickerProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';

/** Nullable local wall-clock time, anchored to the current field date or local 1970-01-01.
 * Invalid/uncommitted drafts block stale-value submission without changing the Date.
 * Explicit parsers/formatters retain their local domain; no timestamp/timezone is merged.
 * Ref targets the input; selection/blur callbacks follow RHF writes/touches.
 * @example `<Order.TimePicker name="time" label="Time" freeform hourCycle="h23" increment={15} />`
 */
export const FormTimePicker = React.forwardRef<HTMLInputElement, FormTimePickerProps<FieldValues>>(function FormTimePicker({ className, disabled, fieldProps, hint, id, label, name, orientation, required, rules, dateAnchor, formatDateToTimeString, hourCycle, showSeconds, parseTimeStringToDate, inputValidationMessage = 'Complete or correct the time before submitting.', onBlur, onTimeChange, onOpenChange, ...props }, ref) {
  const [draft, setDraft] = React.useState<string>();
  const draftRef = React.useRef<string>();
  const rejected = React.useRef(false);
  const updateDraft = (text: string | undefined): void => { draftRef.current = text; setDraft(text); };
  const formatTime = (date: Date): string => formatDateToTimeString?.(date) ?? formatCompatTime(date, { hourCycle, showSeconds });
  // TimePicker compat explicitly requires a nullable empty controlled selection.
  // eslint-disable-next-line @rushstack/no-new-null
  const binding = useFieldAdapter<FieldValues, Date | null>({ className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules,
    validateInput: value => (!rejected.current && (draftRef.current === undefined || draftRef.current === (value ? formatTime(value) : ''))) || inputValidationMessage,
    resetText: { read: () => draftRef.current, restore: text => { rejected.current = false; updateDraft(text); } }
  });
  // TimePicker compat uses null for empty controlled values.
  // eslint-disable-next-line @rushstack/no-new-null
  const value = binding.controller.value ?? null;
  const formatted = value ? formatTime(value) : '';
  const valueKey = value?.getTime();
  const nativeElement = React.useRef<HTMLInputElement>();
  const captureNative = React.useCallback<React.RefCallback<HTMLInputElement>>(element => { nativeElement.current = element ?? undefined; }, []);
  const nativeRef = useMergedRefs(binding.controller.ref, captureNative);
  const restoreFocus = React.useRef(false);
  const previous = React.useRef({ value: valueKey, reset: binding.resetVersion });
  const [primitiveVersion, setPrimitiveVersion] = React.useState(0);
  React.useLayoutEffect(() => {
    const didReset = previous.current.reset !== binding.resetVersion;
    const changed = previous.current.value !== valueKey;
    previous.current = { value: valueKey, reset: binding.resetVersion };
    if (didReset ? !binding.resetPreservesValues : changed) { rejected.current = false; updateDraft(undefined); }
    // Compat remembers submittedText internally, so same-value reset must also
    // clear that marker to allow a retry of identical formerly rejected text.
    if (didReset && !binding.resetPreservesValues) {
      restoreFocus.current = nativeElement.current !== undefined && nativeElement.current.ownerDocument.activeElement === nativeElement.current;
      setPrimitiveVersion(version => version + 1);
    }
  }, [valueKey, binding.resetVersion, binding.resetPreservesValues]);
  React.useLayoutEffect(() => {
    if (!restoreFocus.current) return;
    restoreFocus.current = false;
    nativeElement.current?.focus();
  }, [primitiveVersion]);
  return <Field {...binding.fieldProps}><TimePicker key={primitiveVersion} {...props} {...binding.controlProps}
    name={binding.name} ref={nativeRef} selectedTime={value} value={draft ?? formatted}
    dateAnchor={dateAnchor ?? value ?? new Date(1970, 0, 1, 0, 0, 0, 0)}
    formatDateToTimeString={formatDateToTimeString} hourCycle={hourCycle} showSeconds={showSeconds}
    parseTimeStringToDate={parseTimeStringToDate === undefined ? undefined : text => { const result = parseTimeStringToDate(text); return { date: result.date, errorType: result.error }; }}
    onChange={event => { if (binding.controlProps.disabled) return; rejected.current = false; updateDraft(event.currentTarget.value); }}
    onTimeChange={(event, data) => {
      if (binding.controlProps.disabled) return;
      // Non-freeform Combobox may emit an empty selection while typing its
      // query. Keep that actual draft until a selection/clear commits it.
      if (event.type === 'change' && !data.selectedTime && data.selectedTimeText === undefined) return;
      const failed = data.errorType !== undefined && data.selectedTimeText !== '';
      rejected.current = failed;
      if (failed) updateDraft(data.selectedTimeText ?? draftRef.current);
      else { updateDraft(undefined); binding.controller.onChange(data.selectedTime); }
      onTimeChange?.(event, { selectedTime: data.selectedTime, selectedTimeText: data.selectedTimeText, ...(data.errorType === undefined ? {} : { error: data.errorType }) });
    }}
    onBlur={event => { binding.controller.onBlur(); onBlur?.(event); }}
    onOpenChange={onOpenChange === undefined ? undefined : (event, data) => onOpenChange(data.open)}
  /></Field>;
});
FormTimePicker.displayName = 'Form.TimePicker';

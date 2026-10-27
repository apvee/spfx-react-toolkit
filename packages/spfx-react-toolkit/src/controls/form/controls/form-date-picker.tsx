import * as React from 'react';
import { useFormRuntimeContext } from '../form-context.internal';
import { mergeFormStringFields } from '../form-strings.internal';
import { Field } from '@fluentui/react-components';
import { useDatePicker_unstable, useDatePickerStyles_unstable, renderDatePicker_unstable } from '@fluentui/react-datepicker-compat';
import type { DatePickerProps, CalendarStrings } from '@fluentui/react-datepicker-compat';
import type { FieldValues } from 'react-hook-form';
import type { FormDatePickerProps, FormDatePickerError, FormDatePickerStrings } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';

function compareLocalDateParts(first: Date, second: Date): number {
  return first.getFullYear() - second.getFullYear() || first.getMonth() - second.getMonth() || first.getDate() - second.getDate();
}
const defaultFormatDate = (date?: Date): string => date === undefined ? '' : `${String(date.getFullYear()).padStart(4, '0')}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
/** Strict date-only input in local calendar coordinates, without UTC parsing.
 * setFullYear preserves years 0000–0099, which the multi-argument constructor
 * otherwise maps to 1900–1999. Re-read all parts to reject calendar rollover
 * and dates that the local timezone cannot represent.
 */
function parseLocalCalendarDate(text: string): Date | undefined {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (!match) return undefined;
  const year = Number(match[1]);
  const month = Number(match[2]) - 1;
  const day = Number(match[3]);
  const date = new Date(1970, 0, 1);
  date.setFullYear(year, month, day);
  date.setHours(0, 0, 0, 0);
  return date.getFullYear() === year && date.getMonth() === month && date.getDate() === day ? date : undefined;
}
function toCompatStrings(strings: FormDatePickerStrings | undefined): CalendarStrings | undefined {
  return strings === undefined ? undefined : { ...strings, months: [...strings.months], shortMonths: [...strings.shortMonths], days: [...strings.days], shortDays: [...strings.shortDays] };
}
function toCompatFormatting(formatting: FormDatePickerProps<FieldValues>['dateTimeFormatter'], supplied: FormDatePickerStrings | undefined): DatePickerProps['dateTimeFormatter'] {
  if (!formatting) return undefined;
  const publicStrings = (strings: Parameters<NonNullable<DatePickerProps['dateTimeFormatter']>['formatMonth']>[1]): FormDatePickerStrings => ({ ...strings, ...supplied, goToToday: supplied?.goToToday ?? '' });
  return { formatDay: formatting.formatDay, formatYear: formatting.formatYear, formatMonth: (date, strings) => formatting.formatMonth(date, publicStrings(strings)), formatMonthYear: (date, strings) => formatting.formatMonthYear(date, publicStrings(strings)), formatMonthDayYear: (date, strings) => formatting.formatMonthDayYear(date, publicStrings(strings)) };
}
/** Local primitive composition: compat 0.6.16 has no controlled text prop and
 * replaces invalid text with the old selected date on blur. Its published state
 * hooks allow a narrow value/commit override while retaining the real calendar,
 * popup, slots, keyboard navigation and native ref. No DOM mutation is needed.
 */
const LocalDatePicker = React.forwardRef<HTMLInputElement, DatePickerProps & { displayText: string; commitText: () => void }>(function LocalDatePicker({ displayText, commitText, ...props }, ref) {
  const suppressTextCommit = React.useRef(false);
  const rejectedInternalText = React.useRef(false);
  const state = useDatePicker_unstable({
    ...props,
    // Compat's invalid-text branch reports its old selected Date as a selection.
    // Filter that fallback; calendar selections provide a Date directly and do
    // not invoke this parser. The adapter parses the actual visible text below.
    // Compat's parser requires null for a rejected internal text parse.
    // eslint-disable-next-line @rushstack/no-new-null
    parseDateFromString: () => { rejectedInternalText.current = true; return null; },
    onSelectDate: date => { if (!suppressTextCommit.current && !rejectedInternalText.current) props.onSelectDate?.(date); },
    onValidationResult: data => { rejectedInternalText.current = false; if (!suppressTextCommit.current) props.onValidationResult?.(data); }
  }, ref);
  state.root.value = displayText;
  state.root.onBlur = props.onBlur;
  const onKeyDown = state.root.onKeyDown;
  state.root.onKeyDown = event => {
    // Enter still delegates popup/navigation behavior to the real compat primitive.
    // Its internal text validator is bypassed by the inert parser below; this
    // adapter commits the actual controlled draft before that keyboard behavior.
    if (event.key === 'Enter' && props.allowTextInput && !props.disabled) commitText();
    suppressTextCommit.current = event.key === 'Enter' && props.allowTextInput === true;
    try { onKeyDown?.(event); } finally { suppressTextCommit.current = false; }
  };
  useDatePickerStyles_unstable(state);
  return renderDatePicker_unstable(state);
});

/** Nullable local calendar date with explicit parser/formatter and inclusive bounds.
 * Default display/input is strict date-only `YYYY-MM-DD` (years 0000–9999),
 * parsed from local calendar parts. Other formats/timestamps require an explicit
 * consumer parser; its Date and formatter output are retained without conversion.
 * Invalid drafts stay visible and block validation of an older committed Date.
 * Callbacks run after RHF writes/touches; the public ref targets the native input.
 * Reset clears drafts unless this owner's value is explicitly preserved.
 * @example `<Order.DatePicker name="date" label="Date" allowTextInput formatDate={formatDate} parseDateFromString={parseDate} />`
 */
export const FormDatePicker = React.forwardRef<HTMLInputElement, FormDatePickerProps<FieldValues>>(function FormDatePicker({ className, disabled, fieldProps, hint, id, label, name, orientation, required, rules, allowTextInput, formatDate = defaultFormatDate, parseDateFromString, minDate, maxDate, dateTimeFormatter, strings: suppliedStrings, inputValidationMessage: suppliedInputValidationMessage, onBlur, onSelectDate, onValidationResult, ...props }, ref) {
  const familyStrings = useFormRuntimeContext().strings;
  const inputValidationMessage = suppliedInputValidationMessage ?? familyStrings.dateInputValidationMessage;
  const strings = mergeFormStringFields(familyStrings.datePicker, suppliedStrings);
  const [draft, setDraft] = React.useState<string>();
  const draftRef = React.useRef<string>();
  const updateDraft = (text: string | undefined): void => { draftRef.current = text; setDraft(text); };
  // Fluent compat represents empty controlled date selections with null.
  // eslint-disable-next-line @rushstack/no-new-null
  const binding = useFieldAdapter<FieldValues, Date | null>({ describedBy: props['aria-describedby'], className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules,
    validateInput: value => !allowTextInput || draftRef.current === undefined || draftRef.current === (value ? formatDate(value) : '') || inputValidationMessage,
    resetText: { read: () => draftRef.current, restore: text => updateDraft(text) }
  });
  // Compat requires null rather than undefined for controlled empty selections.
  // eslint-disable-next-line @rushstack/no-new-null
  const value = binding.controller.value ?? null;
  const valueKey = value?.getTime();
  const previous = React.useRef({ value: valueKey, reset: binding.resetVersion });
  React.useLayoutEffect(() => {
    const didReset = previous.current.reset !== binding.resetVersion;
    const changed = previous.current.value !== valueKey;
    previous.current = { value: valueKey, reset: binding.resetVersion };
    if (didReset ? !binding.resetPreservesValues : changed) updateDraft(undefined);
  }, [valueKey, binding.resetVersion, binding.resetPreservesValues]);
  const commitText = (): void => {
    if (!allowTextInput || binding.controlProps.disabled || draftRef.current === undefined) return;
    const text = draftRef.current;
    let error: FormDatePickerError | undefined;
    // Compat's explicitly nullable parsed/selected date model includes clears.
    // eslint-disable-next-line @rushstack/no-new-null
    let next: Date | null = null;
    if (text) {
      if (value && text === formatDate(value)) { updateDraft(undefined); return; }
      if (parseDateFromString) next = parseDateFromString(text);
      else {
        // Compat's empty selected/parsed Date model explicitly uses null.
        // eslint-disable-next-line @rushstack/no-new-null
        next = parseLocalCalendarDate(text) ?? null;
      }
      if (!next || !Number.isFinite(next.getTime())) error = 'invalid-input';
      else if ((minDate && compareLocalDateParts(minDate, next) > 0) || (maxDate && compareLocalDateParts(maxDate, next) < 0)) error = 'out-of-bounds';
    } else if (required) error = 'required-input';
    if (error === undefined || error === 'required-input') {
      updateDraft(undefined);
      binding.controller.onChange(next);
      onSelectDate?.(next);
    }
    onValidationResult?.(error === undefined ? {} : { error });
  };
  return <Field {...binding.fieldProps}><LocalDatePicker {...props} {...binding.controlProps}
    children={undefined} name={binding.name} required={required} allowTextInput={allowTextInput} minDate={minDate} maxDate={maxDate}
    value={value} formatDate={formatDate} strings={toCompatStrings(strings)} dateTimeFormatter={toCompatFormatting(dateTimeFormatter, strings)}
    displayText={draft ?? (value ? formatDate(value) : '')} commitText={commitText}
    ref={binding.controller.ref}
    onValidationResult={onValidationResult}
    onChange={(event, data) => { if (!binding.controlProps.disabled && allowTextInput) updateDraft(data.value); }}
    onBlur={event => { commitText(); binding.controller.onBlur(); onBlur?.(event); }}
    onSelectDate={date => {
      if (binding.controlProps.disabled) return;
      // Compat selection uses null for a cleared value.
      // eslint-disable-next-line @rushstack/no-new-null
      const next = date ?? null;
      updateDraft(undefined); binding.controller.onChange(next); onSelectDate?.(next);
    }}
  /></Field>;
});
FormDatePicker.displayName = 'Form.DatePicker';

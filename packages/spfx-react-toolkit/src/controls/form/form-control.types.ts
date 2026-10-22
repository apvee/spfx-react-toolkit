import type * as React from 'react';
import type { ContentSlot } from '../content-slot.types.internal';
import type { ColorPickerProps, RatingProps, SwatchPickerProps, CheckboxProps, ComboboxProps, DropdownProps, FieldProps, InputProps, RadioGroupProps, SelectProps, TagPickerProps, TagPickerInputProps, SearchBoxProps, SliderProps, SpinButtonProps, SwitchProps, TextareaProps } from '@fluentui/react-components';
import type { FormDateTimeValue, FormHsvColor, FormOptionsSource, FormValueCodec } from './form-value.types';
import type { FieldPath, FieldPathValue, FieldValues, RegisterOptions } from 'react-hook-form';
/** Paths which can both supply the control's value and accept its complete write domain.
 * Optional fields may initialize absent; controlled inputs require suitable defaults.
 * @example
 * ```ts
 * type StringPath = CompatibleFieldPath<{ title: string; count: number }, string>;
 * const name: StringPath = 'title';
 * ```
 */
export type CompatibleFieldPath<TInput extends FieldValues, TValue> = {
    [TName in FieldPath<TInput>]: [
        Exclude<FieldPathValue<TInput, TName>, undefined>
    ] extends [
        TValue
    ] ? [
        TValue
    ] extends [
        FieldPathValue<TInput, TName>
    ] ? TName : never : never;
}[FieldPath<TInput>];
/** Presentation shared by Fluent Field wrappers; required is metadata, not a validation rule.
 * Slots configure inner containers; native Label/hint IDs and associations remain owned by Field.
 * @example `<Order.Input name="reference" label={{ children: 'Reference', className: classes.label }} hint={{ children: (Component, props) => <Component {...props}>Help</Component> }} />`
 */
export interface FormFieldPresentationProps {
    /** Literal phrasing content inside the native Fluent Label. Objects/callbacks configure an inner span,
     * including its ref; htmlFor, required/disabled and generated Label ID remain native.
     * Only undefined omits the top-level label; legacy null/booleans still request the native Label.
     * Omit inside SettingsGroup.Item field.
     */
    label?: ContentSlot<'span'>;
    /** Supporting content inside the native Field hint. Objects/callbacks configure an inner div;
     * its ref/ID do not replace the associated native hint ID. Undefined retains fieldProps.hint.
     * Omit inside SettingsGroup.Item field.
     */
    hint?: ContentSlot<'div'>;
    /** Marks the control as required visually and accessibly. */
    required?: boolean;
    /** Fluent Field label layout; omission uses Fluent's vertical layout. */
    orientation?: FieldProps['orientation'];
    /** Disables the field; cannot override a disabled form or scope. */
    disabled?: boolean;
    /** Class applied to the outer Fluent Field. */
    className?: string;
    /** Additional Fluent Field props; active RHF errors take precedence. */
    fieldProps?: Omit<FieldProps, 'children' | 'className' | 'disabled' | 'label' | 'orientation' | 'required'>;
}
/** Binding props for a control that reads and writes one complete value domain.
 * @example
 * ```ts
 * const props: FormControlProps<{ title: string }, string> = {
 *   name: 'title', label: 'Title', rules: { required: 'Enter a title.' },
 * };
 * ```
 */
export type FormControlProps<TInput extends FieldValues, TValue, TName extends CompatibleFieldPath<TInput, TValue> = CompatibleFieldPath<TInput, TValue>> = FormFieldPresentationProps & {
    /** Relative field path; object/array scopes supply its prefix. */
    name: TName;
    /** Domain validation; disallowed when a schema/resolver owns validation. */
    rules?: RegisterOptions<TInput, TName>;
};
/** Native text input behaviors supported by Input.
 * @example
 * ```ts
 * const type: FormInputType = 'email';
 * ```
 */
export type FormInputType = 'text' | 'email' | 'password' | 'search' | 'tel' | 'url';
/** String binding for Fluent Input; it deliberately has no value codec.
 * @example `<Order.Input name="reference" label="Reference" />`
 */
export type FormInputProps<TInput extends FieldValues> = FormControlProps<TInput, string> & Omit<InputProps, keyof FormFieldPresentationProps | 'defaultValue' | 'name' | 'value' | 'type'> & {
    /** Text input behavior; defaults to native text. Numbers, dates and files use dedicated controls. */
    type?: FormInputType;
};

/** Multiline string binding for Fluent Textarea; RHF owns value and defaults.
 * Native change/blur callbacks run after RHF and the ref targets the textarea.
 * @example `<Order.Textarea name="notes" label="Notes" resize="vertical" />`
 */
export type FormTextareaProps<TInput extends FieldValues> = FormControlProps<TInput, string> & Omit<TextareaProps, keyof FormFieldPresentationProps | 'defaultValue' | 'name' | 'value'>;
/** Search string binding; dismissing writes an empty string before `onChange`.
 * RHF owns value/defaults and the ref targets the native input.
 * @example `<Order.SearchBox name="query" label="Search" />`
 */
export type FormSearchBoxProps<TInput extends FieldValues> = FormControlProps<TInput, string> & Omit<SearchBoxProps, keyof FormFieldPresentationProps | 'defaultValue' | 'name' | 'value'>;
/** Checkbox binding with explicit opt-in to Fluent's mixed state.
 * A boolean field accepts the default mode. A `boolean | 'mixed'` field requires
 * `triState: true`. Change/blur callbacks run after RHF; the ref targets the input.
 * @example `<Order.Checkbox name="approved" label="Approved" />`
 */
export type FormCheckboxProps<TInput extends FieldValues> = Omit<CheckboxProps, keyof FormFieldPresentationProps | 'checked' | 'defaultChecked' | 'name' | 'value'> & (
    (FormControlProps<TInput, boolean> & {
        /** Keeps the complete live domain boolean when omitted or false. */
        triState?: false;
    }) | (FormControlProps<TInput, boolean | 'mixed'> & {
        /** Retains Fluent's mixed state in the live model; requires a mixed-capable field. */
        triState: true;
    })
);
/** Boolean switch binding; RHF owns checked/defaults and the ref targets the input.
 * Change/blur callbacks run after RHF; required is presentation metadata.
 * @example `<Order.Switch name="enabled" label="Enabled" />`
 */
export type FormSwitchProps<TInput extends FieldValues> = FormControlProps<TInput, boolean> & Omit<SwitchProps, keyof FormFieldPresentationProps | 'checked' | 'defaultChecked' | 'name' | 'value'>;
/** Numeric commit binding for Fluent SpinButton's nullable external domain.
 * Finite display commits become numbers; empty/nonfinite text becomes `null`.
 * Validation belongs to schema/rules. RHF owns value/defaults; the ref targets the
 * input and native change/blur callbacks run after RHF. No text-format codec exists.
 * A non-preserving targeted reset replaces only the Fluent primitive to clear its
 * internal draft; the native ref changes and existing input focus is restored.
 * @example `<Order.SpinButton name="quantity" label="Quantity" min={0} />`
 */
// Fluent SpinButton explicitly represents an empty committed value with null.
// eslint-disable-next-line @rushstack/no-new-null
export type FormSpinButtonProps<TInput extends FieldValues> = FormControlProps<TInput, number | null> & Omit<SpinButtonProps, keyof FormFieldPresentationProps | 'defaultValue' | 'displayValue' | 'name' | 'value'>;
/** Numeric slider binding; supply a valid RHF default and matching min/max/step.
 * No adapter fallback changes the caller's range or value. Native change/blur
 * callbacks run after RHF; the ref targets the range input.
 * @example `<Order.Slider name="progress" label="Progress" min={20} max={40} />`
 */
export type FormSliderProps<TInput extends FieldValues> = FormControlProps<TInput, number> & Omit<SliderProps, keyof FormFieldPresentationProps | 'defaultValue' | 'name' | 'value'>;

/** Selection bindings require an explicit lossless codec for a different live field domain.
 * Each mapped member correlates `name` with that path's exact read/write codec type.
 * @example
 * ```tsx
 * const Profile = createForm<{ countryCode: number }>();
 * <Profile.Select name="countryCode" codec={{ toControl: String, fromControl: Number }}
 *   options={[{ value: '39', label: 'Italy' }]} />;
 * ```
 */
type FormSelectionBinding<TInput extends FieldValues, TValue> =
    | (FormControlProps<TInput, TValue> & {
        /** Omit for the complete native string/string-array domain. */
        codec?: never })
    | { [TName in FieldPath<TInput>]: FormFieldPresentationProps & {
        /** Relative path whose exact live value is converted by this codec. */
        name: TName;
        /** Domain validation when rules own validation. */
        rules?: RegisterOptions<TInput, TName>;
        /** Lossless conversion; every writable control identity must round-trip. */
        codec: FormValueCodec<FieldPathValue<TInput, TName>, TValue>;
    } }[FieldPath<TInput>];
/** Selection adapter props combining field binding, one option source, and unowned Fluent props.
 * @example
 * ```tsx
 * <Order.Dropdown name="country" label="Country" options={[{ value: 'it', label: 'Italy' }]} />
 * ```
 */
type FormSelectionProps<TInput extends FieldValues, TValue, TPrimitive, TOwned extends PropertyKey> = FormSelectionBinding<TInput, TValue> & FormOptionsSource & Omit<TPrimitive, TOwned | keyof FormFieldPresentationProps>;
/** Native Select for one stable string identity, or an explicitly converted domain.
 * Supply native `option` children or data options. RHF owns value/defaults.
 * @example `<Order.Select name="country" options={[{ value: 'pt', label: 'Portugal' }]} />`
 */
export type FormSelectProps<TInput extends FieldValues> = FormSelectionProps<TInput, string, SelectProps, 'children' | 'defaultValue' | 'name' | 'value' | 'multiple'>;
/** Fluent RadioGroup with a public div ref and descendant error focus.
 * Supply Fluent Radio children or data options. Group callbacks follow RHF.
 * @example `<Order.RadioGroup name="priority" label="Priority" options={priorities} />`
 */
export type FormRadioGroupProps<TInput extends FieldValues> = FormSelectionProps<TInput, string, RadioGroupProps, 'children' | 'defaultValue' | 'name' | 'value'>;
/** Single Dropdown stores identity and displays independent option text.
 * The ref targets the trigger button; `Option` children or options are exclusive.
 * @example `<Order.Dropdown name="country" label="Country" options={countries} />`
 */
export type FormDropdownProps<TInput extends FieldValues> = FormSelectionProps<TInput, string, DropdownProps, 'children' | 'defaultSelectedOptions' | 'defaultValue' | 'multiselect' | 'name' | 'selectedOptions' | 'value'>;
/** Multiple Dropdown stores an ordered string array; the ref targets the button.
 * Native selection order is retained. Explicit codecs convert the complete array.
 * @example `<Order.DropdownMulti name="countries" label="Countries" options={countries} />`
 */
export type FormDropdownMultiProps<TInput extends FieldValues> = FormSelectionProps<TInput, string[], DropdownProps, 'children' | 'defaultSelectedOptions' | 'defaultValue' | 'multiselect' | 'name' | 'selectedOptions' | 'value' | 'clearable'>;
/** Single Combobox separates query presentation from selected identity.
 * Unmatched queries write the empty identity; schema/rules decide its validity.
 * Reset clears drafts; explicit `keepDirtyValues` preserves dirty field drafts.
 * @example `<Order.Combobox name="country" label="Country" options={countries} />`
 */
export type FormComboboxProps<TInput extends FieldValues> = FormSelectionProps<TInput, string, ComboboxProps, 'children' | 'defaultSelectedOptions' | 'defaultValue' | 'multiselect' | 'name' | 'selectedOptions' | 'value' | 'freeform'>;
/** Multiple Combobox keeps query edits separate from its ordered selected identities.
 * Selection clears query. Reset clears drafts unless their dirty owner is preserved.
 * @example `<Order.ComboboxMulti name="countries" label="Countries" options={countries} />`
 */
export type FormComboboxMultiProps<TInput extends FieldValues> = FormSelectionProps<TInput, string[], ComboboxProps, 'children' | 'defaultSelectedOptions' | 'defaultValue' | 'multiselect' | 'name' | 'selectedOptions' | 'value' | 'freeform' | 'clearable'>;
/** Checkbox group for string identities in declared option order.
 * Direct Fluent Checkbox children are the only children adapted; other content is retained.
 * The public div ref is separate from RHF focus on the first eligible checkbox.
 * @example `<Order.CheckboxGroup name="labels" label="Labels" options={labels} />`
 */
export type FormCheckboxGroupProps<TInput extends FieldValues> = FormSelectionBinding<TInput, string[]> & FormOptionsSource & Omit<React.HTMLAttributes<HTMLDivElement>, 'children' | 'onChange' | 'onBlur' | 'className'> & {
    /** Runs after RHF writes the group selection, before the child callbacks. */
    onChange?: CheckboxProps['onChange'];
    /** Runs after RHF marks the field touched, before the child's blur callback. */
    onBlur?: CheckboxProps['onBlur'];
};
/** Localized TagPicker query results.
 * @example `const strings: FormTagPickerStrings = { noOptions: 'Nessun risultato' };`
 */
export interface FormTagPickerStrings {
    /** Literal empty-result content inside the owned status region; objects/callbacks configure an inner div and its ref. */
    noOptions: ContentSlot<'div'>;
}
/** TagPicker naming alternatives; rich labels require explicit text unless a settings field row supplies it.
 * @example
 * ```tsx
 * <Order.TagPicker name="people" aria-label="Selected people" label={<strong>People</strong>} options={people} />
 * ```
 */
type FormTagPickerAccessibleName = {
    /** Plain text visible label also names the selected-tag list. */
    label: string;
    /** Explicit text overriding the label-derived accessible name when supplied. */
    'aria-label'?: string } | {
    /** Rich/object/callback or absent label needs an explicit accessible name; no text is inferred. */
    label?: ContentSlot<'span'>;
    /** Required text name when standalone label content is not a string. */
    'aria-label': string } | {
    /** Omit label when SettingsGroup.Item field supplies inherited naming. Standalone usage is checked at runtime. */
    label?: undefined;
    /** Omit with label only when inherited settings-row naming is available. */
    'aria-label'?: undefined };
/** TagPicker with removable tags, string-array identities and a native input ref.
 * Options or Fluent TagPickerOption children provide identities and display text.
 * Omit label/aria-label only when SettingsGroup.Item field supplies inherited naming.
 * Ancestor context cannot be proven by TypeScript; unnamed standalone usage throws at runtime.
 * Query edits retain selections; reset and selection clear query presentation.
 * @example `<Order.TagPicker name="people" label="People" options={people} strings={{ noOptions: 'No people' }} />`
 */
export type FormTagPickerProps<TInput extends FieldValues> = FormSelectionProps<TInput, string[], TagPickerProps, 'aria-label' | 'children' | 'defaultSelectedOptions' | 'defaultValue' | 'name' | 'noPopover' | 'selectedOptions' | 'value' | 'onChange' | 'onBlur' | 'id'> & FormTagPickerAccessibleName & {
    /** Add a Fluent colorful initials/fallback avatar to each selected tag. Defaults to false, retaining text-only tags.
     * Colors use the string option identity and remain stable when display labels change. Rich consumer labels are preserved.
     * Only presentation changes; form values, query and selection callbacks are unaffected.
     * @example <Order.TagPicker name="people" label="People" options={people} colorfulTagAvatars />
     */
    colorfulTagAvatars?: boolean;
    /** Additional descriptions for the native query input. */
    'aria-describedby'?: string;
    /** Input id shared with Fluent Field's label. */
    id?: string;
    /** Query callback after presentation updates; it does not change selected values. */
    onChange?: TagPickerInputProps['onChange'];
    /** Runs after RHF marks the field touched. */
    onBlur?: TagPickerInputProps['onBlur'];
    /** Inner div content for the owned no-options status. Objects support attributes, refs and children callbacks.
     * Nullish values fall back to strings.noOptions; empty/boolean shorthand keeps the status shell. */
    noOptionsContent?: ContentSlot<'div'>;
    /** Localized package-owned query-result content. */
    strings?: Partial<FormTagPickerStrings>;
};

/** Numeric Rating with a public div ref and error focus on an eligible native radio.
 * RHF changes precede consumer callbacks. No codec or default-value override exists.
 * Star color defaults to native `marigold`; explicit `brand` or `neutral` overrides are preserved.
 * @example `<Order.Rating name="score" label="Score" max={5} />`
 */
export type FormRatingProps<TInput extends FieldValues> = FormControlProps<TInput, number> & Omit<RatingProps, keyof FormFieldPresentationProps | 'defaultValue' | 'name' | 'value'>;
/** HSV ColorPicker with optional alpha, copied on every read and write.
 * Compose Fluent ColorArea, ColorSlider or AlphaSlider children. The ref targets
 * the root div; RHF error focus targets an eligible native range input.
 * @example `<Order.ColorPicker name="accent" label="Accent"><ColorSlider channel="hue" /></Order.ColorPicker>`
 */
export type FormColorPickerProps<TInput extends FieldValues> = FormControlProps<TInput, FormHsvColor> & Omit<ColorPickerProps, keyof FormFieldPresentationProps | 'color' | 'name'>;
/** SwatchPicker stores the stable string identity, independent of label and color.
 * Compose Fluent ColorSwatch or ImageSwatch children. The public ref targets the
 * root div; RHF error focus targets an enabled swatch button.
 * Native shape defaults to rounded; square and circular remain supported overrides.
 * @example `<Order.SwatchPicker name="theme" label="Theme"><ColorSwatch value="blue" color="#0078d4" /></Order.SwatchPicker>`
 */
export type FormSwatchPickerProps<TInput extends FieldValues> = FormControlProps<TInput, string> & Omit<SwatchPickerProps, keyof FormFieldPresentationProps | 'defaultSelectedValue' | 'name' | 'selectedValue'>;

/** Package-owned DatePicker validation error identifiers.
 * @example
 * ```ts
 * const error: FormDatePickerError = 'invalid-input';
 * ```
 */
export type FormDatePickerError = 'invalid-input' | 'out-of-bounds' | 'required-input';

/** Package-owned validation result emitted after DatePicker text validation.
 * @example
 * ```ts
 * const onValidationResult = ({ error }: FormDatePickerValidationResult) => {
 *   if (error !== undefined) reportDateError(error);
 * };
 * ```
 */
export type FormDatePickerValidationResult = {
  /** Validation error, or `undefined` when the current text is valid. */
  error?: FormDatePickerError;
};

/** Localized calendar strings consumed by Form.DatePicker.
 * @example
 * ```ts
 * const strings: FormDatePickerStrings = { ...calendarStrings, goToToday: 'Oggi' };
 * ```
 */
export type FormDatePickerStrings = {
  /** Full month names beginning with January. */
  months: readonly string[];
  /** Abbreviated month names beginning with January. */
  shortMonths: readonly string[];
  /** Full weekday names beginning with Sunday. */
  days: readonly string[];
  /** Abbreviated weekday names beginning with Sunday. */
  shortDays: readonly string[];
  /** Label for the action that navigates to today. */
  goToToday: string;
  /** Accessible label for navigating to the previous month. */
  prevMonthAriaLabel?: string;
  /** Accessible label for navigating to the next month. */
  nextMonthAriaLabel?: string;
  /** Accessible label for navigating to the previous year. */
  prevYearAriaLabel?: string;
  /** Accessible label for navigating to the next year. */
  nextYearAriaLabel?: string;
  /** Accessible label for navigating to the previous year range. */
  prevYearRangeAriaLabel?: string;
  /** Accessible label for navigating to the next year range. */
  nextYearRangeAriaLabel?: string;
  /** Accessible format string for the month-picker header. */
  monthPickerHeaderAriaLabel?: string;
  /** Accessible format string for the year-picker header. */
  yearPickerHeaderAriaLabel?: string;
  /** Accessible label for the calendar close action. */
  closeButtonAriaLabel?: string;
  /** Accessible format string for a week number. */
  weekNumberFormatString?: string;
  /** Accessible format string for the selected date. */
  selectedDateFormatString?: string;
  /** Accessible format string for today's date. */
  todayDateFormatString?: string;
  /** Accessible label announced for marked dates. */
  dayMarkedAriaLabel?: string;
};

/** Package-owned localized formatting functions used by Form.DatePicker's calendar.
 * @example
 * ```ts
 * const formatter: FormDatePickerFormatting = {
 *   formatDay: date => String(date.getDate()),
 *   formatMonth: (date, strings) => strings.months[date.getMonth()],
 *   formatYear: date => String(date.getFullYear()),
 *   formatMonthDayYear: date => date.toLocaleDateString('it-IT'),
 *   formatMonthYear: date => date.toLocaleDateString('it-IT', { month: 'long', year: 'numeric' }),
 * };
 * ```
 */
export type FormDatePickerFormatting = {
  /** Formats a day number. */
  formatDay: (date: Date) => string;
  /** Formats a month using the supplied localized strings. */
  formatMonth: (date: Date, strings: FormDatePickerStrings) => string;
  /** Formats a year. */
  formatYear: (date: Date) => string;
  /** Formats a complete month, day, and year. */
  formatMonthDayYear: (date: Date, strings: FormDatePickerStrings) => string;
  /** Formats a month and year. */
  formatMonthYear: (date: Date, strings: FormDatePickerStrings) => string;
};

/** Numeric weekday identifiers where Sunday is `0` and Saturday is `6`.
 * @example
 * ```ts
 * const monday: FormDatePickerDayOfWeek = 1;
 * ```
 */
export type FormDatePickerDayOfWeek = 0 | 1 | 2 | 3 | 4 | 5 | 6;

/** Numeric first-week rules: 0 is first day, 1 is first full week, and 2 is first four-day week.
 * @example
 * ```ts
 * const firstFourDayWeek: FormDatePickerFirstWeekOfYear = 2;
 * ```
 */
export type FormDatePickerFirstWeekOfYear = 0 | 1 | 2;

/** Native date/time input attributes after removing values, events and presentation owned by the Form binding. */
type FormDateTimeNativeInputProps = Omit<
  React.InputHTMLAttributes<HTMLInputElement>,
  | 'className'
  | 'children'
  | 'defaultValue'
  | 'disabled'
  | 'name'
  | 'onChange'
  | 'required'
  | 'size'
  | 'type'
  | 'value'
>;

/** Package-owned DatePicker configuration without its Form field binding.
 * @example
 * ```tsx
 * const datePickerProps: FormDatePickerControlProps = { allowTextInput: true, firstDayOfWeek: 1 };
 * <Order.DatePicker name="date" label="Date" {...datePickerProps} />;
 * ```
 */
export type FormDatePickerControlProps = FormDateTimeNativeInputProps & {
  /** Fluent input appearance. */
  appearance?: InputProps['appearance'];
  /** Allows direct date text entry. Defaults to `false`. Uncommitted or rejected text blocks validation and submission until corrected, cleared, or reset. */
  allowTextInput?: boolean;
  /** Localized error for text that does not match the committed date. Defaults to `Complete or correct the date before submitting.` Domain constraints remain owned by rules, schema, or resolver. */
  inputValidationMessage?: string;
  /** Allows disabled calendar elements to receive focus. Defaults to `false`. */
  allFocusable?: boolean;
  /** Removes the DatePicker border. Defaults to `false`. */
  borderless?: boolean;
  /** Additional class name applied to the DatePicker root. */
  className?: string;
  /** Prevents interaction. */
  disabled?: boolean;
  /** Prevents automatic popup opening on focus. */
  disableAutoFocus?: boolean;
  /** Opens the calendar initially. */
  defaultOpen?: boolean;
  /** Package-owned localized calendar formatting functions. */
  dateTimeFormatter?: FormDatePickerFormatting;
  /** First weekday displayed by the calendar. */
  firstDayOfWeek?: FormDatePickerDayOfWeek;
  /** Rule used to determine the first week of a year. */
  firstWeekOfYear?: FormDatePickerFirstWeekOfYear;
  /** Synchronously formats the selected local date as input text; defaults to YYYY-MM-DD, or empty text for undefined. Pair custom formats with parseDateFromString. */
  formatDate?: (date?: Date) => string;
  /** Highlights the current month. */
  highlightCurrentMonth?: boolean;
  /** Highlights the selected month. */
  highlightSelectedMonth?: boolean;
  /** Date initially highlighted when the calendar opens. */
  initialPickerDate?: Date;
  /** Renders the popup in DOM order rather than through a portal. */
  inlinePopup?: boolean;
  /** Shows the month picker. */
  isMonthPickerVisible?: boolean;
  /** Inclusive latest selectable local calendar date; omission supplies no upper bound. Domain constraints still belong to rules/schema/resolver. */
  maxDate?: Date;
  /** Inclusive earliest selectable local calendar date; omission supplies no lower bound. Domain constraints still belong to rules/schema/resolver. */
  minDate?: Date;
  /** Controlled popup open state. */
  open?: boolean;
  /** Opens the calendar when the input is clicked. Defaults to `true`. */
  openOnClick?: boolean;
  /** Synchronous popup-open notification; receives the proposed boolean state. Update open externally when controlled; the return value is ignored. */
  onOpenChange?: (open: boolean) => void;
  /** Synchronous notification after an accepted calendar or text commit writes RHF. Receives the local Date or legacy compat null for clear; the return value is ignored. */
  // Compat date/time selection and parser contracts explicitly use null for empty values.
  // eslint-disable-next-line @rushstack/no-new-null
  onSelectDate?: (date: Date | null) => void;
  /** Synchronous date-validation notification with the package-owned error identifier or undefined for success. Invalid drafts keep the committed value; the return value is ignored. */
  onValidationResult?: (data: FormDatePickerValidationResult) => void;
  /** Synchronously parses direct input text into a local date; return null for rejected text, following the compat parser contract. */
  // Compat date/time selection and parser contracts explicitly use null for empty values.
  // eslint-disable-next-line @rushstack/no-new-null
  parseDateFromString?: (dateText: string) => Date | null;
  /** Placeholder shown when no date is selected. */
  placeholder?: string;
  /** Presents the control as required without replacing validation rules. */
  required?: boolean;
  /** Shows the calendar close action. */
  showCloseButton?: boolean;
  /** Shows the action that navigates to today. */
  showGoToToday?: boolean;
  /** Overlays the month picker above the day picker. */
  showMonthPickerAsOverlay?: boolean;
  /** Shows calendar week numbers. */
  showWeekNumbers?: boolean;
  /** Package-owned localized calendar strings. */
  strings?: FormDatePickerStrings;
  /** Date treated as today; defaults to the client's current date. */
  today?: Date;
  /** Uses the DatePicker underline treatment. */
  underlined?: boolean;
};

/** Nullable local date binding; the native ref targets its input.
 * @example `<Order.DatePicker name="date" label="Date" allowTextInput />`
 */
export type FormDatePickerProps<TInput extends FieldValues> = FormControlProps<
  TInput,
  // Compat date/time selection and parser contracts explicitly use null for empty values.
  // eslint-disable-next-line @rushstack/no-new-null
  Date | null
> &
  FormDatePickerControlProps;

/** Package-owned TimePicker validation error identifiers.
 * @example
 * ```ts
 * const error: FormTimePickerError = 'out-of-bounds';
 * ```
 */
export type FormTimePickerError = 'invalid-input' | 'out-of-bounds' | 'required-input';

/** Package-owned TimePicker selection payload.
 * @example
 * ```ts
 * const onTimeChange = (_event: React.SyntheticEvent<HTMLElement>, data: FormTimePickerChangeData) => {
 *   if (data.error === undefined) inspectTime(data.selectedTime);
 * };
 * ```
 */
export type FormTimePickerChangeData = {
  /** Selected or parsed time, or `null` when cleared. */
  // Compat date/time selection and parser contracts explicitly use null for empty values.
  // eslint-disable-next-line @rushstack/no-new-null
  selectedTime: Date | null;
  /** Display text associated with the selection or freeform input. */
  selectedTimeText: string | undefined;
  /** Validation error associated with the submitted time text. */
  error?: FormTimePickerError;
};

/** Package-owned result returned by a freeform time parser.
 * @example
 * ```ts
 * const rejected: FormTimePickerParseResult = { date: null, error: 'invalid-input' };
 * ```
 */
export type FormTimePickerParseResult = {
  /** Parsed time, or `null` when no valid time is available. */
  // Compat date/time selection and parser contracts explicitly use null for empty values.
  // eslint-disable-next-line @rushstack/no-new-null
  date: Date | null;
  /** Validation error associated with the parsed text. */
  error?: FormTimePickerError;
};

/** Inclusive/exclusive hour boundary supported by Form.TimePicker.
 * @example
 * ```ts
 * const endHour: FormTimePickerHour = 24;
 * ```
 */
export type FormTimePickerHour =
  | 0
  | 1
  | 2
  | 3
  | 4
  | 5
  | 6
  | 7
  | 8
  | 9
  | 10
  | 11
  | 12
  | 13
  | 14
  | 15
  | 16
  | 17
  | 18
  | 19
  | 20
  | 21
  | 22
  | 23
  | 24;

/** Package-owned TimePicker configuration without its Form field binding.
 * @example
 * ```tsx
 * const timePickerProps: FormTimePickerControlProps = { freeform: true, hourCycle: 'h23', increment: 15 };
 * <Order.TimePicker name="time" label="Time" {...timePickerProps} />;
 * ```
 */
export type FormTimePickerControlProps = FormDateTimeNativeInputProps & {
  /** Fluent input appearance. */
  appearance?: ComboboxProps['appearance'];
  /** Shows a clear action. */
  clearable?: boolean;
  /** Additional class name applied to the TimePicker root. */
  className?: string;
  /**
   * Date used as the fixed basis for generated and parsed times. When omitted, the current field
   * date is preserved, or local 1970-01-01 is used while the field is empty.
   */
  dateAnchor?: Date;
  /** Opens the time list initially. */
  defaultOpen?: boolean;
  /** Prevents interaction. */
  disabled?: boolean;
  /** Exclusive end hour for generated options. */
  endHour?: FormTimePickerHour;
  /** Localized error for rejected or uncommitted time text; defaults to Complete or correct the time before submitting. Domain validation stays in rules/schema/resolver. */
  inputValidationMessage?: string;
  /** Allows freeform time entry. */
  freeform?: boolean;
  /** Synchronously formats generated/selected local time values; omission uses compat formatting with hourCycle and showSeconds. */
  formatDateToTimeString?: (date: Date) => string;
  /** Selects the locale hour cycle. */
  hourCycle?: 'h11' | 'h12' | 'h23' | 'h24';
  /** Minute interval between generated options. */
  increment?: number;
  /** Renders the popup in DOM order rather than through a portal. */
  inlinePopup?: boolean;
  /** Controlled popup open state. */
  open?: boolean;
  /** Synchronous popup-open notification; receives the proposed boolean state. Update open externally when controlled; the return value is ignored. */
  onOpenChange?: (open: boolean) => void;
  /** Receives the native selection event and package-owned time/text/error payload after any accepted RHF write. Invalid draft notifications retain the committed value. The return value is ignored. */
  onTimeChange?: (event: React.SyntheticEvent<HTMLElement>, data: FormTimePickerChangeData) => void;
  /** Synchronously parses possibly absent freeform text; return the parsed date (or legacy compat null) and optional error. No async parsing is supported. */
  parseTimeStringToDate?: (timeText: string | undefined) => FormTimePickerParseResult;
  /** Placeholder shown when no time is selected. */
  placeholder?: string;
  /** Includes seconds in generated and formatted options. */
  showSeconds?: boolean;
  /** Fluent input size. */
  size?: InputProps['size'];
  /** Inclusive start hour for generated options. */
  startHour?: FormTimePickerHour;
};

/** Nullable local time binding; the native ref targets its input.
 * @example `<Order.TimePicker name="time" label="Time" freeform />`
 */
export type FormTimePickerProps<TInput extends FieldValues> = FormControlProps<
  TInput,
  // Compat date/time selection and parser contracts explicitly use null for empty values.
  // eslint-disable-next-line @rushstack/no-new-null
  Date | null
> &
  FormTimePickerControlProps;

/** Separate local date/time parts; the public ref targets the semantic group.
 * @example `<Order.DateTime name="appointment" dateLabel="Date" timeLabel="Time" />`
 */
export type FormDateTimeProps<TInput extends FieldValues> = FormControlProps<
  TInput,
  FormDateTimeValue
> & {
  /** Explicit identity of the semantic date/time group. */
  id?: string;
  /** Additional descriptions appended to group hint and validation relationships. */
  'aria-describedby'?: string;
  /** Literal phrasing slot inside the DatePicker native Label; object/callback refs target its inner span. */
  dateLabel?: ContentSlot<'span'>;
  /** Package-owned props forwarded to the DatePicker part. */
  datePickerProps?: FormDatePickerControlProps;
  /** Literal phrasing slot inside the TimePicker native Label; object/callback refs target its inner span. */
  timeLabel?: ContentSlot<'span'>;
  /** Package-owned props forwarded to the TimePicker part. */
  timePickerProps?: FormTimePickerControlProps;
};


/** Native browser file selection stored as FileList, with null for an empty selection.
 * The ref targets the real HTMLInputElement; change/blur callbacks run after RHF.
 * Browsers cannot restore files from defaults, reset values or external values: package-owned
 * file fields normalize unsupported selections to null. Explicit keepDirtyValues preserves a
 * current native selection only while its input survives. Upload and persisted attachments are
 * consumer responsibilities.
 * @example `<Order.FileInput name="attachments" label="Attachments" accept="application/pdf" multiple />`
 */
export type FormFileInputProps<TInput extends FieldValues> = FormControlProps<
  TInput,
  // HTMLInputElement.files and native file-clear interoperability explicitly use null.
  // eslint-disable-next-line @rushstack/no-new-null
  FileList | null
> & Omit<React.InputHTMLAttributes<HTMLInputElement>, keyof FormFieldPresentationProps | 'children' | 'defaultValue' | 'name' | 'type' | 'value'> & {
  /** Localized visible action that opens the native file chooser; defaults to Browse. */
  browseLabel?: string;
  /** Localized accessible label for clearing the native selection; defaults to Clear. */
  clearLabel?: string;
  /** Formats the actual live native files; defaults to comma-separated file names. */
  formatFiles?: (files: FileList) => React.ReactNode;
};

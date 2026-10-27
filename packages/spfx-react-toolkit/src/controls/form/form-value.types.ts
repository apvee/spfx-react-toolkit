import type * as React from 'react';
/** Lossless, explicit conversion between a domain field and a control's live value.
 * @example `const codec: FormValueCodec<number, string> = { toControl: String, fromControl: Number };`
 */
export interface FormValueCodec<TField, TControl> {
    /** Converts the stored field value to the control domain. */
    toControl: (value: TField) => TControl;
    /** Converts every writable identity, including supported empty values, back to the field domain. */
    fromControl: (value: TControl) => TField;
}
/** Stable string option identity with independent presentation and typeahead text.
 * @example `const option: FormOption = { value: 'pt', label: 'Portugal' };`
 */
export interface FormOption {
    /** Control identity, stored directly without a codec and independent of the visible label. */
    value: string;
    /** Visible label; strings are rendered literally. */
    label: React.ReactNode;
    /** Prevents user selection of this option. */
    disabled?: boolean;
    /** Plain text used by Fluent faceplates, filtering and typeahead when the label is JSX. */
    textValue?: string;
}

/** Exclusive data options or primitive child composition.
 * Use native `option` children for Select, Fluent Radio/Checkbox for groups,
 * Fluent Option for dropdowns/comboboxes, and TagPickerOption for TagPicker.
 * @example `<Order.Dropdown name="country" options={[{ value: 'pt', label: 'Portugal' }]} />`
 */
export type FormOptionsSource = {
    /** Stable identities and independent labels/typeahead text. */
    options: readonly FormOption[];
    /** Unavailable when data options are supplied. */
    children?: never;
} | {
    /** Unavailable when primitive children are supplied. */
    options?: never;
    /** Native/Fluent option composition, rendered without arbitrary child adaptation. */
    children: React.ReactNode;
};

/** Live HSV color domain shared with Fluent ColorPicker; no CSS-string conversion.
 * @example `const accent: FormHsvColor = { h: 210, s: 0.7, v: 0.8, a: 1 };`
 */
export interface FormHsvColor {
    /** Hue in degrees, following Fluent's color-domain conventions. */
    h: number;
    /** Saturation in Fluent's unit interval. */
    s: number;
    /** Value in Fluent's unit interval. */
    v: number;
    /** Optional alpha in Fluent's unit interval; absence remains absent. */
    a?: number;
}

/** Independent nullable calendar-date and wall-clock parts; no timestamp is created.
 * @example `const value: FormDateTimeValue = { date: new Date(2026, 7, 31), time: new Date(1970, 0, 1, 9) };`
 */
export interface FormDateTimeValue {
    /** Local calendar date. Compat DatePicker explicitly uses null for an empty selection. */
    // eslint-disable-next-line @rushstack/no-new-null
    date: Date | null;
    /** Local wall-clock time. Compat TimePicker explicitly uses null for an empty selection. */
    // eslint-disable-next-line @rushstack/no-new-null
    time: Date | null;
}

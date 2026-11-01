/**
 * Typed Form families, RHF value ownership and native Fluent bindings.
 *
 * @packageDocumentation
 * @docPage controls/form
 */
import type * as React from 'react';
/**
 * Lossless, explicit conversion between a domain field and a control's live value.
 * @example
 * ```ts
 * import { FormValueCodec } from '@apvee/spfx-react-toolkit/controls/form';
 * const codec: FormValueCodec<number, string> = { toControl: String, fromControl: Number };
 * // Supply only option identities for which this mapping round-trips.
 * ```
 * @typeParam TField - Stored live field domain before conversion.
 * @typeParam TControl - Complete control read/write domain, including supported empty values.
 */
export interface FormValueCodec<TField, TControl> {
    /**
     * Converts the stored field value to the control domain.
     * @param value - Complete live-domain value to write or convert; no implicit domain coercion occurs.
     * @returns Control-domain value without changing stored data.
     */
    toControl: (value: TField) => TControl;
    /**
     * Converts every writable identity, including supported empty values, back to the field domain.
     * @param value - Complete live-domain value to write or convert; no implicit domain coercion occurs.
     * @returns Complete field-domain value for every writable control identity.
     */
    fromControl: (value: TControl) => TField;
}
/**
 * Stable string option identity with independent presentation and typeahead text.
 * @example
 * ```ts
 * import { FormOption } from '@apvee/spfx-react-toolkit/controls/form';
 * const option: FormOption = { value: 'pt', label: 'Portugal' };
 * ```
 */
export interface FormOption {
    /** Control identity, stored directly without a codec and independent of the visible label. */
    value: string;
    /** Visible label; strings are rendered literally. */
    label: React.ReactNode;
    /**
     * Prevents user selection of this option.
     * When omitted or false, adds no disabled restriction; enclosing Root/scope restrictions still apply.
     */
    disabled?: boolean;
    /**
     * Plain text used by Fluent faceplates, filtering and typeahead when the label is JSX.
     * When omitted, plain string labels provide text; rich JSX labels need explicit textValue for reliable filtering/typeahead.
     */
    textValue?: string;
}

/**
 * Exclusive data options or primitive child composition.
 * Use native `option` children for Select, Fluent Radio/Checkbox for groups,
 * Fluent Option for dropdowns/comboboxes, and TagPickerOption for TagPicker.
 * @example
 * ```ts
 * import { FormOptionsSource } from '@apvee/spfx-react-toolkit/controls/form';
 * const source: FormOptionsSource = { options: [{ value: 'pt', label: 'Portugal' }] };
 * ```
 */
export type FormOptionsSource = {
    /** Stable identities and independent labels/typeahead text. */
    options: readonly FormOption[];
    /**
     * Unavailable when data options are supplied.
     * When omitted, the region has no caller content.
     */
    children?: never;
} | {
    /**
     * Unavailable when primitive children are supplied.
     * Omit this member in this configuration branch.
     */
    options?: never;
    /** Native/Fluent option composition, rendered without arbitrary child adaptation. */
    children: React.ReactNode;
};

/**
 * Live HSV color domain shared with Fluent ColorPicker; no CSS-string conversion.
 * @example
 * ```ts
 * import { FormHsvColor } from '@apvee/spfx-react-toolkit/controls/form';
 * const accent: FormHsvColor = { h: 210, s: 0.7, v: 0.8, a: 1 };
 * ```
 */
export interface FormHsvColor {
    /** Hue in degrees, following Fluent's color-domain conventions. */
    h: number;
    /** Saturation in Fluent's unit interval. */
    s: number;
    /** Value in Fluent's unit interval. */
    v: number;
    /**
     * Optional alpha in Fluent's unit interval; absence remains absent.
     * When omitted, alpha remains absent on initial reads; Fluent color interactions can normalize it to 1.
     */
    a?: number;
}

/**
 * Independent nullable calendar-date and wall-clock parts; no timestamp is created.
 * @example
 * ```ts
 * import { FormDateTimeValue } from '@apvee/spfx-react-toolkit/controls/form';
 * const value: FormDateTimeValue = { date: new Date(2026, 7, 31), time: new Date(1970, 0, 1, 9) };
 * ```
 */
export interface FormDateTimeValue {
    /** Local calendar date. Compat DatePicker explicitly uses null for an empty selection. */
    // eslint-disable-next-line @rushstack/no-new-null
    date: Date | null;
    /** Local wall-clock time. Compat TimePicker explicitly uses null for an empty selection. */
    // eslint-disable-next-line @rushstack/no-new-null
    time: Date | null;
}

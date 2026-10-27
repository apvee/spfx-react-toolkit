import type { FormDatePickerStrings, FormTagPickerStrings } from './form-control.types';
import type { FormArraySectionStrings } from './form-composition.types';

/** Complete Form UI dictionary. Native SPFx resources supply omitted overrides.
 * Caller validation messages, field labels and values remain caller-owned.
 * @example
 * ```ts
 * function customize(strings: FormStrings): FormStrings {
 *   return { ...strings, unexpectedFailure: 'Please retry.' };
 * }
 * ```
 */
export interface FormStrings {
    /** Fallback created when a resolver, rule, submit or invalid callback throws. Stored errors retain their event's text. */
    unexpectedFailure: string;
    /** Calendar names and accessible messages from FormDatePickerStrings, with every native field resolved. Arrays start with January/Sunday and replace whole arrays. Input parsing is unaffected. */
    datePicker: Required<FormDatePickerStrings>;
    /** Incomplete date draft message, unless the control supplies inputValidationMessage. */
    dateInputValidationMessage: string;
    /** Incomplete time draft message, unless the control supplies inputValidationMessage. */
    timeInputValidationMessage: string;
    /** TagPicker default content; direct content and control strings take precedence. */
    tagPicker: FormTagPickerStrings;
    /** Native file chooser action labels; direct control labels take precedence. */
    fileInput: {
        /** Visible chooser action. Native resources default to Browse in English. */
        browseLabel: string;
        /** Accessible clear-selection action. Native resources default to Clear in English. */
        clearLabel: string;
    };
    /** FormArraySectionStrings action labels and announcements, with every optional leaf field resolved. Callbacks receive caller item names literally. */
    arraySection: Required<FormArraySectionStrings> & {
        /** Default item name from its one-based position. Direct itemLabel receives the full item and takes precedence.
         * @param position - One-based item position.
         * @returns The visible/accessibility item name.
         */
        itemLabel: (position: number) => string;
    };
}

/** Known optional Form dictionary fields, merged per group with native resources.
 * Explicit empty strings are retained; calendar arrays replace complete arrays.
 * Supply to useForm or an initializing Root. A Root with methods inherits that
 * hook's dictionary. Dictionary changes do not reset or revalidate the form.
 * @example
 * ```ts
 * const strings: FormStringsOverride = {
 *   unexpectedFailure: 'Please retry.',
 *   fileInput: { browseLabel: 'Choose files' },
 *   arraySection: { itemLabel: position => `Row ${position}` },
 * };
 * ```
 */
export interface FormStringsOverride {
    /** Optional unexpected exception fallback; omission uses the native resource. */
    unexpectedFailure?: FormStrings['unexpectedFailure'];
    /** Optional calendar fields; omission retains native names and accessibility text. */
    datePicker?: Partial<FormStrings['datePicker']>;
    /** Optional incomplete date draft message; direct control message wins. */
    dateInputValidationMessage?: string;
    /** Optional incomplete time draft message; direct control message wins. */
    timeInputValidationMessage?: string;
    /** Optional tag content using the unchanged FormTagPickerStrings slot contract. */
    tagPicker?: Partial<FormStrings['tagPicker']>;
    /** Optional file action labels; direct control labels win. */
    fileInput?: Partial<FormStrings['fileInput']>;
    /** Optional array callbacks/text/item naming; leaf action strings and itemLabel win. */
    arraySection?: Partial<FormStrings['arraySection']>;
}

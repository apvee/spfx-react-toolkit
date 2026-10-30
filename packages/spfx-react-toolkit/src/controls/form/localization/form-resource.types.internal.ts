/** Complete translator data for the Form family; runtime callbacks remain in the family adapter. */
export interface FormResource {
  readonly unexpectedFailure: string;
  readonly dateInputValidationMessage: string;
  readonly timeInputValidationMessage: string;
  readonly tagPicker: { readonly noOptions: string };
  readonly fileInput: { readonly browseLabel: string; readonly clearLabel: string };
  readonly arraySection: {
    readonly add: string;
    readonly empty: string;
    readonly itemLabelTemplate: string;
    readonly removeTemplate: string;
    readonly moveUpTemplate: string;
    readonly moveDownTemplate: string;
    readonly addedTemplate: string;
    readonly removedTemplate: string;
    readonly movedTemplate: string;
  };
  readonly datePicker: {
    readonly months: readonly string[];
    readonly shortMonths: readonly string[];
    readonly days: readonly string[];
    readonly shortDays: readonly string[];
    readonly goToToday: string;
    readonly prevMonthAriaLabel: string;
    readonly nextMonthAriaLabel: string;
    readonly prevYearAriaLabel: string;
    readonly nextYearAriaLabel: string;
    readonly prevYearRangeAriaLabel: string;
    readonly nextYearRangeAriaLabel: string;
    readonly monthPickerHeaderAriaLabel: string;
    readonly yearPickerHeaderAriaLabel: string;
    readonly closeButtonAriaLabel: string;
    readonly weekNumberFormatString: string;
    readonly selectedDateFormatString: string;
    readonly todayDateFormatString: string;
    readonly dayMarkedAriaLabel: string;
  };
}

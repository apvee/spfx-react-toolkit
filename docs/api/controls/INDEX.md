# Controls

Reusable React controls that do not require an SPFx provider. Theme variables and the shared Fluent/Griffel runtime still apply.

| Control / type | Purpose | Reference |
| --- | --- | --- |
| `Form`, `createForm` | Typed RHF forms, Fluent adapters, scopes and editable sections | [Form](./form.md) |
| `SettingsGroup`, `SettingsGroupProps`, `SettingsGroupItemProps`, `SettingsGroupHeadingLevel` | Caller-owned settings rows in a Fluent Card | [SettingsGroup](./settings-group.md) |
| `AccordionGroup`, `AccordionGroupProps`, `AccordionGroupItemProps`, `AccordionGroupHeadingLevel`, `AccordionGroupValue` | Fluent disclosure with retained panel state | [AccordionGroup](./accordion-group.md) |
| `StateMessage` | Section state with artwork, text and actions | [StateMessage](./state-message.md#statemessage) |
| `StateMessageProps` | Content, preset, visual and native div attributes | [Props](./state-message.md#statemessageprops) |
| `StateMessagePreset` | generic, empty, error, success or welcome | [Presets](./state-message.md#statemessagepreset) |
| `StateMessageHeadingLevel` | Semantic title heading level 1–6 | [Heading](./state-message.md#statemessageheadinglevel-and-statemessagerole) |
| `StateMessageRole` | status or explicit alert for the text region | [Role](./state-message.md#statemessageheadinglevel-and-statemessagerole) |
| `StateMessageAction`, `StateMessageClickAction`, `StateMessageLinkAction` | Typed callback/button or navigation/anchor intentions | [Actions](./state-message.md#statemessageaction-statemessageclickaction-and-statemessagelinkaction) |
| `AutoSkeleton` | Fluent loading items derived from mounted HTML geometry | [AutoSkeleton](./auto-skeleton.md#autoskeleton) |
| `AutoSkeletonProps` | Loading, animation, minimum size, exclusions and native div attributes | [Props](./auto-skeleton.md#autoskeletonprops) |
| `RelativeTime` | Localized relative instant with accessible full date and clock modes | [RelativeTime](./relative-time.md#relativetime) |
| `RelativeTimeProps` | Instant, locale, time zone, clock and native attributes | [Props](./relative-time.md#relativetimeprops) |
| `FormattedText` | Native text element with typography, inline Markdown and overflow | [FormattedText](./formatted-text.md#formattedtext) |
| `FormattedTextProps` | Text, native attributes and presentation options | [Props](./formatted-text.md#formattedtextprops) |
| `FormattedTextVariant` | Keys of the existing typography catalog | [Variants](./formatted-text.md#formattedtextvariant) |
| `FormattedTextSemanticElement` | `span`, `p` or `h1` through `h6` | [Semantic element](./formatted-text.md#formattedtextsemanticelement) |

Import from the package root, `/controls`, or the individual `/controls/auto-skeleton`, `/controls/formatted-text` `/controls/relative-time` `/controls/state-message`, `/controls/settings-group` and `/controls/accordion-group` entries. These entries share the canonical implementation; internal measurement, parser and styling modules are private. See [package imports](../../PACKAGE-IMPORTS.md).

Form public types: `CreateForm`, `FormErrors`, `FormMethods`, `FormRootProps`, `FormUseFormProps`, `FormValidationProps`, `FormSchemaValidation`, `FormResolverValidation`, `FormWithoutValidation`, `FormFieldResult`, `FormUseFieldProps`, `FormErrorProps`, `FormSubmitProps`, `FormResetProps`, `FormSuite`, `ScopedFormSuite`, `UntypedFormSuite`, `CompatibleFieldPath`, `FormFieldPresentationProps`, `FormControlProps`, `FormInputProps`, `FormInputType`, `FormTextareaProps`, `FormSearchBoxProps`, `FormCheckboxProps`, `FormSwitchProps`, `FormSpinButtonProps`, `FormSliderProps`, `FormSelectProps`, `FormRadioGroupProps`, `FormDropdownProps`, `FormDropdownMultiProps`, `FormCheckboxGroupProps`, `FormComboboxProps`, `FormComboboxMultiProps`, `FormTagPickerProps`, `FormTagPickerStrings`, `FormRatingProps`, `FormColorPickerProps`, `FormSwatchPickerProps`, `FormDatePickerError`, `FormDatePickerValidationResult`, `FormDatePickerStrings`, `FormDatePickerFormatting`, `FormDatePickerDayOfWeek`, `FormDatePickerFirstWeekOfYear`, `FormDatePickerControlProps`, `FormDatePickerProps`, `FormTimePickerError`, `FormTimePickerChangeData`, `FormTimePickerParseResult`, `FormTimePickerHour`, `FormTimePickerControlProps`, `FormTimePickerProps`, `FormDateTimeProps`, `FormFileInputProps`, `FormOption`, `FormValueCodec`, `FormOptionsSource`, `FormHsvColor`, `FormDateTimeValue`, `FormObjectPath`, `FormObjectPathValue`, `FormArrayPath`, `FormArrayItemValue`, `FormObjectContext`, `FormObjectProps`, `FormArrayItemContext`, `FormArrayContext`, `FormArrayRules`, `FormArrayProps`, `FormObjectComponent`, `FormArrayComponent`, `FormSectionName`, `FormSectionProps`, `FormObjectSectionProps`, `FormArraySectionProps`, `FormArraySectionStrings`, `FormObjectSectionComponent`, `FormArraySectionComponent`. See the [Form reference](./form.md) for compound members and lifecycle contracts.

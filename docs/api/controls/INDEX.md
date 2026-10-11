# Controls

Most controls are standalone and do not require an SPFx provider. `WebPartShell` requires an actual SPFx instance and supplies its provider internally. Theme variables and the shared Fluent/Griffel runtime still apply.

| Control / type | Purpose | Reference |
| --- | --- | --- |
| `Stage`, `StageProps`, `StageHeaderProps`, `StageBodyProps`, `StageFooterProps` | Explicit Dialog/Drawer/Popover with Header, Body and Footer | [Stage](./stage.md) |
| `DataGrid`, `createDataGrid` | Native table with sixteen cell presentations and typed or dynamic suite ownership | [DataGrid](./data-grid.md) |
| `Form`, `createForm` | Typed RHF forms, Fluent adapters, settings field rows, scopes and editable sections | [Form](./form.md) |
| `SettingsGroup`, `SettingsGroupProps`, `SettingsGroupItemProps`, `SettingsGroupHeadingLevel` | Settings rows in a Fluent Card, with opt-in Form naming | [SettingsGroup](./settings-group.md) |
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

DataGrid public types: `DataGridProps`, `DataGridStrings`, `DataGridSparklineSummaryText`, `DataGridColumnCollection`, `DataGridColumnMarkers`, `DataGridSuite`, `DynamicDataGridColumnMarkers`, `DynamicDataGridSuite`, `DataGridColumnKind`, `DataGridLocale`, `DataGridRowId`, `DataGridValueContext`, `DataGridValueSource`, `DataGridColumnCommonProps`, `DataGridColumnDefinition`, `DataGridColumnBuilder`, `DataGridValueColumnProps`, `DynamicDataGridColumnProps`, `DynamicDataGridColumnBuilder`, `ActionsColumnProps`, `DataGridMappedValues`, `DataGridBadgeValue`, `DataGridIconValue`, `DataGridTagValue`, `DataGridTagValues`, `DataGridLinkNavigateData`, `DataGridSparklinePoint`, `DataGridSparklineMetric`, `DataGridSparklineSummary`, `DataGridActionData`, `DataGridAction`, `DataGridNumberOptions`, `DataGridBooleanOptions`, `DataGridDateOptions`, `DataGridBadgeOptions`, `DataGridIconOptions`, `DataGridLinkOptions`, `DataGridPersonOptions`, `DataGridPeopleOptions`, `DataGridTagsOptions`, `DataGridProgressOptions`, `DataGridRatingOptions`, `DataGridSparklineOptions`, `DataGridCustomOptions`. See the [DataGrid reference](./data-grid.md) for all sixteen column members and controlled state contracts.

## WebPartShell

`WebPartShell`, `IWebPartShellProperties`, `WebPartShellProps`, `WebPartShellSurface`, `WebPartShellPadding` and `getWebPartShellPropertiesMetadata` compose an actual web part with providers, adaptive spacing and a persisted inline title. See the [WebPartShell reference](./web-part-shell.md) for every member, metadata, editor protocol and host limits. Import from the root, `/controls` or `/controls/web-part-shell`.

## Stage

`Stage` is the single value export with `Stage.Header`, `Stage.Body` and `Stage.Footer`. Its thirteen public types are `StagePresentation`, `StageSize`, `StageDismissMode`, `StageDrawerPosition`, `StageHeadingLevel`, `StageOpenChangeReason`, `StageOpenChangeEvent`, `StageOpenChangeData`, `StagePopoverPositioning`, `StageProps`, `StageHeaderProps`, `StageBodyProps` and `StageFooterProps`. See the [Stage reference](./stage.md) for every member, controlled requests, session retention, form composition and host limits. Import from the root, `/controls` or `/controls/stage`; there are no flat part exports.

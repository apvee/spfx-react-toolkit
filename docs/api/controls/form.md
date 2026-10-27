# Form

`Form` and `createForm` compose React Hook Form (RHF) with Fluent UI 9 controls. RHF owns values, validation, dirty/touched state, subscriptions and submission bookkeeping. No SPFx provider is required. Use a shared FluentProvider and the toolkit's `useSx` descriptors for layout.

Import the compound family and its public types from the package root, `/controls` or `/controls/form`. Adapter implementations are private; there are no flat `FormInput` exports or adapter plugins. Create typed suites outside component rendering.

## Dependencies

The tested baseline is React/ReactDOM 17.0.1, TypeScript 5.3.3 and Fluent react-components 9.74.9, with RHF 7.87.0, datepicker-compat 0.6.16, timepicker-compat 0.4.21 and react-icons 2.0.312. RHF, compat controls and icons are explicit host peers. The package supplies `@hookform/resolvers` 5.9.1 and `@standard-schema/spec` 1.1.0 as direct dependencies.

```bash
npm install react-hook-form@7.87.0 @fluentui/react-datepicker-compat@0.6.16 @fluentui/react-timepicker-compat@0.4.21 @fluentui/react-icons@2.0.312
```

The compat subtree is locked with calendar-compat 0.3.11 and the existing shared contexts/utilities/Griffel versions. Preserve a compatible host lockfile and check runtime identity: newer compat candidates require newer shared packages and can otherwise introduce duplicate providers/renderers. The tested baseline preserves context 9.26.4, utilities 9.26.7 and Griffel React 1.7.8; peer ranges are not a tested version matrix. Package-root declaration resolution requires these peers even when another named export is used. Do not upgrade the host toolchain solely to install Form.

## Rules-only forms

```tsx
import * as React from 'react';
import { createForm } from '@apvee/spfx-react-toolkit';

const Profile = createForm<{ name: string; newsletter: boolean }>();

export function ProfileEditor(): React.ReactElement {
  return <Profile.Root
    defaultValues={{ name: '', newsletter: false }}
    onSubmit={async values => { await saveProfile(values); }}
  >
    <Profile.Input name="name" label="Name" required rules={{ required: 'Enter your name.' }} />
    <Profile.Switch name="newsletter" label="Newsletter" />
    <Profile.Error />
    <Profile.Submit>Save</Profile.Submit>
    <Profile.Reset>Reset</Profile.Reset>
  </Profile.Root>;
}
```

`required` supplies visible and accessible metadata. It does **not** create a validation rule. Requirements belong in `rules`, a schema or a resolver. Returned service errors should normally be localized by the consumer with `methods.setError('root.serverError', { message: 'Please try again.' })`.

## Settings rows as Form fields

[SettingsGroup](./settings-group.md#automatic-form-field-rows) can own the visible name and hint of one built-in Form logical control. Use `SettingsGroup.Item field title="..." description="..."` and omit the control's `label` and `hint`. The same integration applies to typed, dynamic and scoped `Fields` adapters, with no manual ARIA naming links or repeated visible text.

```tsx
import { createForm, SettingsGroup } from '@apvee/spfx-react-toolkit';

const Account = createForm<{ profile: { name: string; enabled: boolean } }>();

<Account.Root defaultValues={{ profile: { name: '', enabled: true } }} onSubmit={saveAccount}>
  <Account.Object name="profile">{({ Fields }) =>
    <SettingsGroup title="Profile">
      <SettingsGroup.Item field title="Name" description="The name shown to your colleagues.">
        <Fields.Input name="name" required rules={{ required: 'Enter a name.' }} />
      </SettingsGroup.Item>
      <SettingsGroup.Item field title="Enabled" description="Allow this account to receive updates.">
        <Fields.Switch name="enabled" />
      </SettingsGroup.Item>
    </SettingsGroup>
  }</Account.Object>
  <Account.Error />
  <Account.Submit>Save</Account.Submit>
  <Account.Reset>Reset</Account.Reset>
</Account.Root>
```

The title must supply a meaningful name and may contain noninteractive phrasing JSX. Arbitrary custom components used in the title must themselves render meaningful accessible content; the toolkit does not inspect their rendered DOM. Native labelable controls receive a real label targeting their actual ID; clicking a row title focuses the input or activates the boolean control. Composite controls receive the row's name at their semantic target rather than assigning the same name to every internal control. DateTime is one logical field and still uses `dateLabel`/`timeLabel` for its two parts. TagPicker inherits the row name for its input and tag group; its ordinary standalone naming requirement remains in force.

The row description is merged with caller `aria-describedby` IDs and Field error IDs. TagPicker accepts `id` for its native query input and `aria-describedby` for additional input descriptions. DateTime accepts `id` and `aria-describedby` for its semantic date/time group; its parts retain their own identities and descriptions. Validation stays with the control, and the row label reflects `required` and effective disabled state. `required` does not create validation: choose field rules, a schema or a resolver as usual. Edits update RHF normally; persistence happens through your existing submit handler, with no automatic save.

Exactly one built-in adapter may own a field row. Repeated adapter `label`/`hint` and multiple owners produce diagnostics. Split multiple controls across field rows, or use an ordinary `SettingsGroup.Item` and name each control explicitly. Secondary actions, nested settings rows and DateTime parts do not inherit the parent row's naming. Custom `useField` bindings and standalone Fluent controls need their own labels/hints and should use ordinary rows. This composition uses the [existing dependencies](#dependencies), Fluent theme and renderer, with no new provider or peer.

The **Form** test web part shows one selected form model and, in the variable-field model, one selected adapter at a time. Change the field, edit its value, and use save/reset and the validation controls to inspect the current model's behavior. The versioned [Form regression fixture](../../../tests/fixtures/demo-control-isolation/form-regressions.tsx) retains broader adapter, native-form and isolation scenarios. These local interactions do not establish authenticated SharePoint or screen-reader behavior; see the [host checklist](../../SHAREPOINT-VALIDATION.md#form).

## Presentation slots

Named presentation props follow the [shared content-slot convention](../../CONTROL-SLOTS.md):

| Props | Fixed inner tag | Owned outer region |
| --- | --- | --- |
| `label`, `dateLabel`, `timeLabel` | `span` | Fluent Label/Field |
| `hint` | `div` | Fluent Field hint |
| ObjectSection/ArraySection `title` | `span` | Native legend/fieldset |
| ObjectSection/ArraySection `description` | `div` | Associated native hint |
| TagPicker `noOptionsContent`, `strings.noOptions` | `div` | Owned `role="status"` |

String/number/JSX shorthand retains its existing wrappers and literal text. A props object configures an inner container with attributes, classes, events, a consumer style or ref; its `children` may be a Fluent `(Component, props) => ReactNode` render callback. The ref identifies the inner container. Use `as` only to repeat its fixed tag; raw HTML and direct function shorthand are unsupported. Keep label/legend content noninteractive phrasing. Arbitrary consumer JSX remains the consumer's responsibility.

```tsx
const labelRef = React.createRef<HTMLSpanElement>();
<Profile.Input
  name="name"
  label={{ ref: labelRef, className: classes.label, children: 'Name' }}
  hint={{ children: (Component, props) => <Component {...props}>Contact name</Component> }}
/>;
```

Native Label/hint IDs, `htmlFor`, required/disabled state and validation precedence stay with Fluent Field. An object `id` names the inner container and does not replace those associations. Only `label={undefined}` omits the top-level native Label request; legacy null, false, true and empty labels retain that request. An undefined top-level hint retains native `fieldProps.hint`; native `fieldProps` slots keep their Fluent contracts. Settings field rows still reject every defined control label/hint, including null/false or objects, because the row owns the single name/description bridge. DateTime part labels pass through the same adapter while the parent keeps its separate group name.

Section props retain the required title-or-`aria-label`-or-`aria-labelledby` union. Empty static title objects require an explicit outer name; opaque callbacks/custom components must render meaningful content. Empty description objects render their inner container without adding an empty fieldset description reference. Callbacks are never evaluated to discover semantic presence.

TagPicker uses only a plain string label for automatic input/tag-group naming. Object, JSX and callback labels require explicit `aria-label`; the toolkit does not extract their text. `noOptionsContent ?? strings.noOptions ?? familyStrings.tagPicker.noOptions` retains nullish fallback and the status shell for boolean/empty content. These slots apply only to data-backed options with no available matches; custom option children still own filtering.

The [existing peers](#dependencies) provide slot resolution and native semantics; no provider or dependency is added. Options, value codecs, `formatFiles`, array `renderItem`, validation messages and localized action/accessible-name strings retain their existing domain contracts. The Form sample's Notes field exposes inner label/hint refs, label click events and a static hint. DateTime's time label demonstrates a render callback; Address and empty tag results exercise other slots. The versioned conformance fixture retains the separate stateful callback-hint scenario. Static sample checks do not establish browser or authenticated SharePoint execution.

## Input, output and validation ownership

`createForm<TInput>()` retains the input model at submission. `createForm(schema)` infers distinct input/output models from a Standard Schema. `createForm<TInput, TOutput, TContext>({ schema })` and the equivalent `{ resolver }` bind that owner permanently. Standard Schema is an interface; no Zod or Yup dependency is required.

For example, a schema can accept `{ quantity: string }` and produce `{ quantity: number }`: the Input receives and writes a string, `getValues()`/`watch()` still expose strings, and successful `onSubmit` receives numbers. Validation does not rewrite the live model. Direct Input cannot write a narrower literal type; use a string input domain and transform its output, or a custom `useField` binding.

Choose one validation owner: schema, resolver or field/array rules. Supplying domain rules with a resolver is rejected rather than silently ignored. Internal date/time draft checks compose with the chosen owner so an incomplete draft cannot submit an older committed value.

Root may initialize its own form or receive `methods` from the **same suite's** `useForm`. Supplied methods and initialization options are mutually exclusive; raw RHF methods and methods from another factory are not accepted. The validation owner of a bound suite cannot be replaced. The native HTMLFormElement ref and applicable native attributes/events are forwarded; owned content cannot use `dangerouslySetInnerHTML`. Root defaults to `noValidate`.

```tsx
const Editor = createForm<{ note: string }>();

function ProvidedEditor(): React.ReactElement {
  const methods = Editor.useForm({ defaultValues: { note: '' } });
  return <Editor.Root methods={methods} onSubmit={values => saveNote(values)}>
    <Editor.Textarea name="note" label="Note" />
    <button type="button" onClick={() => methods.reset({ note: 'Another record' })}>Load record</button>
    <Editor.Submit>Save</Editor.Submit>
  </Editor.Root>;
}
```

## Errors, pending and retry

Unexpected resolver failures use the reserved non-root error `__formValidation`. RHF deletes `errors.root` before deciding whether submission is valid; a root-only resolver failure would therefore be unsafe. The reserved field and its paths cannot appear in the input model or control names. `FormErrors` and `methods.formState.errors` expose this blocker truthfully.

Unexpected custom-validator failures become validation errors at the rule boundary. Unexpected save failures are caught inside RHF's valid callback and become `root.submit` before RHF calculates submission success, including promises rejected with `undefined` or other falsy reasons. The original exception is retained in development diagnostics. An invalid notification failure is contained as well. No private RHF state mutation or recovery reset is used; pending clears and later attempts can succeed.

`Error` renders a native alert. Without a name it aggregates the root's own messages, immediate global children such as `root.submit`/`root.serverError`, and `__formValidation`, with deterministic deduplication. Address a field or array root for its local messages. Arbitrary refs and nested object graphs are not recursively interpreted as messages.

Root gates repeated native submissions and `requestSubmit()` through both validation and save. Submit and Reset are disabled while pending; ordinary fields remain editable. Save receives a validated snapshot, so edits after save starts do not mutate it. This does not promise to freeze fields at the initial click. Direct `methods.handleSubmit` preserves RHF behavior and bypasses Root's DOM gate.

Disabled fields follow RHF's omission contract. A disabled Root or section cannot be re-enabled by nested props. `readOnly` retains submitted data where the underlying native control supports it; selection and boolean controls do not acquire a fabricated readOnly implementation. Control change/blur updates RHF before invoking the corresponding consumer callback. Ordinary label, hint and validation relationships belong to Fluent Field. Opt-in settings rows compose row naming and hints with those validation relationships.

Semantic adapter callbacks run after the RHF update, followed by supported native-slot notifications. SpinButton's native input change is an editing draft; its top-level change is the numeric commit. CheckboxGroup owns its documented direct Checkbox child callbacks. ColorPicker and SwatchPicker own their adapter callbacks; arbitrary ColorSlider/ImageSwatch children retain Fluent callback order. Native color interaction can normalize absent alpha to `1`; the initial read preserves an absent alpha. RadioGroup owns its group callback, while raw Radio children retain Fluent's native-slot/child/bubbling order; use the group callback when reading the committed form selection.

## Adapters and stored domains

Controlled fields require appropriate initial values. RHF partial/async defaults remain supported but do not perform whole-model coercion. Consumer-owned async producers must handle expected loading errors themselves.

| Member | Control domain | Empty value and relevant behavior |
| --- | --- | --- |
| `Form.Input` | `string` | `''`; no codec; native input ref. |
| `Form.Textarea` | `string` | `''`; multiline text. |
| `Form.SearchBox` | `string` | Clear commits `''`. |
| `Form.Checkbox` | `boolean`, or mixed when explicitly enabled | `false`; mixed requires `triState`. |
| `Form.Switch` | `boolean` | `false`. |
| `Form.SpinButton` | `number \| null` | Empty/nonfinite manual commits become `null`; finite display values and step callbacks commit numbers. |
| `Form.Slider` | `number` | Supply valid defaults and bounds. |
| `Form.Rating` | `number` | Fluent rating scale and keyboard behavior; native `marigold` stars by default. Explicit `color="brand"` or `color="neutral"` overrides are preserved. |
| `Form.Select` | `string` | `''`; option identity is independent of label. |
| `Form.RadioGroup` | `string` | `''`; named group and enabled-descendant focus. |
| `Form.Dropdown` | `string` | `''`; selected identity and displayed label are separate. |
| `Form.Combobox` | `string` | Unmatched editing clears the single selected identity; visible query is presentation state. |
| `Form.CheckboxGroup` | `string[]` | `[]`; ordered selected identities. |
| `Form.DropdownMulti` | `string[]` | `[]`; ordered selected identities. |
| `Form.ComboboxMulti` | `string[]` | `[]`; editing a query preserves existing selections. |
| `Form.TagPicker` | `string[]` | `[]`; keyboard tag removal and localized filtered-empty text. |
| `Form.DatePicker` | `Date \| null` | `null`; local calendar date and explicit parse/format/bounds. |
| `Form.TimePicker` | `Date \| null` | `null`; local wall-clock time and deterministic anchor. |
| `Form.DateTime` | `FormDateTimeValue` | Separate nullable date/time parts; no implicit timestamp. |
| `Form.ColorPicker` | `FormHsvColor` | HSV channels with optional alpha. |
| `Form.SwatchPicker` | `string` | Stable swatch identity; native `shape` defaults to `rounded`, with `square` and `circular` overrides. |
| `Form.FileInput` | `FileList \| null` | `null`; actual native file selection, no upload. |

Nullable domains above reflect real Fluent compat/native contracts. `watch`, `getValues`, `setValue` and `reset` retain those values instead of disguising them as undefined. Missing optional props and internal absence use undefined. A schema may transform nullable input to a different output at submission.

Typed name compatibility checks both reading the current value and writing the full control domain. Selection and other supported codecs use `FormValueCodec<TField, TControl>` with `toControl` and `fromControl`; consumers own a lossless mapping for numeric/object/literal domains. There is no implicit stringify, parse or option-ID lookup.

`FormOption` carries a stable string `value`, ReactNode `label`, optional `disabled` and optional `textValue` for typeahead. Options and children are alternatives, not simultaneous configurations. Disabled options cannot be selected. CheckboxGroup child adaptation is limited to its documented Fluent Checkbox composition; arbitrary consumer controls are not cloned.

Native Select displays its plain visible label. Fluent faceplate/filter/typeahead text uses `textValue` then a plain label; supply textValue for rich JSX labels. Identities never become a hidden display fallback. DropdownMulti, ComboboxMulti and TagPicker preserve Fluent's emitted selection order; CheckboxGroup writes declared option order. Native select/listbox slots cannot replace the option source or selection mode.

Outside a `SettingsGroup.Item field`, TagPicker requires a string label or explicit `aria-label` at runtime; rich/absent labels require `aria-label` to name its tag group. Its type permits omitted naming for inherited field context, so TypeScript alone cannot establish that an unnamed standalone instance is valid. `strings.noOptions` or `noOptionsContent` localizes data-filtered empty content. Filtering custom TagPickerOption children belongs to the consumer. In the tested Fluent version TagPickerOption exposes `aria-disabled` rather than a typed disabled prop; the adapter honors that identity guard.

## Reset, records and drafts

`defaultValues` initializes the baseline. Load another record with `methods.reset(newRecord)`; this updates the baseline and clears presentation drafts. Refresh the same record with explicit `keepDirtyValues` when preservation is desired. Object identity and domain IDs do not trigger inferred record changes or smart merges.

`resetField(name)` targets that field. Same-value resets still clear combobox/date/time drafts. The private per-form coordinator tracks field owners across scoped arrays; moving an item preserves its identity and draft, removing it cleans up its owners, and unrelated forms remain isolated. Reactive `values` and async defaults follow RHF reset semantics with explicit resetOptions from the last committed render. Options from a suspended render do not change the active reset policy. Explicit reset-call options take precedence. A private, reset-scoped handoff preserves only opted-in dirty presentation text across actual RHF array-key replacement under `keepDirtyValues`; it is consumed once and cleared on targeted reset, removal or the next reset. RHF remains the only owner of semantic values and state.

SpinButton also clears uncommitted text on a real targeted reset. Fluent 9.74.9 prioritizes its internal editing text over the controlled display prop and exposes no draft-clear API, so a reset that does not preserve this field replaces only the SpinButton primitive. Its native input identity changes, its forwarded ref updates, and existing input focus is restored. The Field, RHF controller, array item and form retain identity; a preserved dirty owner keeps its native input and draft. This presentation reset adds no numeric validation registry.

TimePicker likewise replaces only its compat primitive on a targeted reset that does not preserve the field, because compat remembers previously submitted text. The forwarded native input ref changes and existing input focus is restored; preserved dirty presentation keeps the primitive. Freeform times commit on blur or selection using the tested compat popup behavior.

Ordinary Reset invokes the consumer click handler first and respects `preventDefault()`. Native form reset cancellation is honored. Reset does not restore browser file selections or infer focus changes for a new record.

DatePicker represents local calendar dates; TimePicker represents local times. DateTime stores `{ date: Date | null, time: Date | null }` separately and merges edits against the current scoped parent. Invalid drafts remain visible and block Root **and direct handleSubmit**, even if the old committed date is valid. UTC, time zones and DST conversion belong to the consumer/schema.

DatePicker's default display and manual parser use strict local `YYYY-MM-DD`, with four-digit years `0000`–`9999`. The parser constructs and checks local calendar parts, including leap days; invalid dates, rollover, locale/timestamp/offset strings remain visible invalid drafts. Other locale formats or timestamp conversion require explicit consumer parser/formatter functions. There is no UTC parsing fallback. TimePicker retains the current field's local anchor or an explicit anchor; an empty field uses local `1970-01-01`. Its time part is meaningful independently of that anchor.

FileInput owns local browser selection only. Clear/reset synchronizes its native input, display and submitted data. External/default/reset FileList values cannot reconstruct a native selection: keep existing attachments in a separate consumer model. Explicit keepDirtyValues can preserve a still-current dirty native selection. If an array reset actually remounts the native input, its selection cannot survive, including under `keepDirtyValues`: the owned model and display clear truthfully. Custom unowned FileList fields are not normalized by a package FileInput owner.

## Objects, arrays and visual sections

`Form.Object` scopes relative field names without adding layout. `Form.Array` manages structured records with RHF useFieldArray. Its RHF stable key is separate from domain ID/index; render with that key so moving/swapping records preserves mounted identity. Scopes compose for Object-in-Array and Array-in-Object. Captured field components follow the item after reordering.

Array supports append, prepend, insert, remove, update, replace, move and swap. Disabled/invalid operations are no-ops; duplicate removal indices are handled safely; `remove(undefined)` removes every item. Domain array rules conflict with a resolver rather than being ignored. Headless operations do not move focus. Primitive arrays should use an appropriate multi-selection binding or a consumer model of records.

`Form.ObjectSection` adds a named semantic group/fieldset and Fluent surface using `name`, `title`, `description`, `disabled`, `appearance`, children/render content and class/native group attributes. Supply a visible name or aria-label/aria-labelledby.

`Form.ArraySection` adds an ordered editable list with `name`, `createItem`, `renderItem`, `minItems`, `maxItems`, `removable`, `reorderable`, `itemLabel`, localized `strings` and `disabled`. Add focuses the first available item control; Remove chooses the next/previous item or Add fallback; Move keeps the item's action focused. Hidden, inert, disabled, aria-disabled and focus-rejecting targets are skipped. Announcements support repeated identical actions. Record reset does not force a focus change.

Sections inherit disabled state and RHF omission. Titles/descriptions accept the presentation slots above, including literal strings and JSX, with no implicit Markdown. Caller content and `useSx` own custom layout. Source slots/contentLayout/itemLayout configuration is intentionally replaced with these ordinary content/action props and class overrides.

## Custom binding and compatibility evidence

`Form.useField` supplies the live value, onChange/onBlur/ref, errors, dirty/touched/invalid/disabled and resetVersion for a custom control. Use truthful native/RHF refs; do not fabricate DOM nodes. `Form.useForm` supplies enhanced same-suite RHF methods and reset coordination. These suite hooks are not new package-level hook exports.

The lazy Form sample selects one model at a time; its variable-field model selects one of the 22 adapters at a time. Save/reset, validation, draft and pending-save controls operate on that active form. Broader adapter/composition and native/isolation regressions remain in the versioned [Form fixture](../../../tests/fixtures/demo-control-isolation/form-regressions.tsx). Source browser and compiled sample checks are separate from the authenticated [SharePoint checklist](../../SHAREPOINT-VALIDATION.md#form).

See [Development](../../DEVELOPMENT.md) for fresh verification commands and [package imports](../../PACKAGE-IMPORTS.md) for peer resolution and bundle limits. Importing a member of the compound family may retain the complete family; per-slot tree shaking is not promised. Unrelated imports must not retain Form/RHF/schema/date/time dependencies. Build library before the SPFx app and restart a running serve after source changes. Local tests/builds do not establish tenant permissions or screen-reader behavior.

## Public type reference

These declarations are exported from each Form entrypoint. Form-owned members appear below; inherited native Fluent and RHF members retain their upstream contracts. The usage and lifecycle sections above describe their behavior.

| Declaration | Form-owned members |
| --- | --- |
| `Form` | Uses the adapter, binding or lifecycle contract above. |
| `createForm` | Uses the adapter, binding or lifecycle contract above. |
| `CreateForm` | Uses the adapter, binding or lifecycle contract above. |
| `FormErrors` | `FormErrors.__formValidation` |
| `FormMethods` | `FormMethods.formState`, `FormMethods.formState.errors` |
| `FormRootProps` | `FormRootProps.children`, `FormRootProps.onSubmit`, `FormRootProps.onInvalid`, `FormRootProps.disabled`, `FormRootProps.methods` |
| `FormUseFormProps` | `FormUseFormProps.schema`, `FormUseFormProps.resolver`, `FormUseFormProps.strings` |
| `FormValidationProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormSchemaValidation` | `FormSchemaValidation.schema`, `FormSchemaValidation.resolver` |
| `FormResolverValidation` | `FormResolverValidation.resolver`, `FormResolverValidation.schema` |
| `FormWithoutValidation` | `FormWithoutValidation.schema`, `FormWithoutValidation.resolver` |
| `FormFieldResult` | `FormFieldResult.value`, `FormFieldResult.onChange`, `FormFieldResult.onBlur`, `FormFieldResult.ref`, `FormFieldResult.error`, `FormFieldResult.errors`, `FormFieldResult.invalid`, `FormFieldResult.disabled`, `FormFieldResult.isDirty`, `FormFieldResult.isTouched`, `FormFieldResult.resetVersion` |
| `FormUseFieldProps` | `FormUseFieldProps.name`, `FormUseFieldProps.rules`, `FormUseFieldProps.disabled` |
| `FormErrorProps` | `FormErrorProps.name` |
| `FormSubmitProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormResetProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormSuite` | `FormSuite.Root`, `FormSuite.Submit`, `FormSuite.Reset`, `FormSuite.useForm` |
| `ScopedFormSuite` | `ScopedFormSuite.Input`, `ScopedFormSuite.Textarea`, `ScopedFormSuite.SearchBox`, `ScopedFormSuite.Checkbox`, `ScopedFormSuite.Switch`, `ScopedFormSuite.SpinButton`, `ScopedFormSuite.Slider`, `ScopedFormSuite.Select`, `ScopedFormSuite.RadioGroup`, `ScopedFormSuite.Dropdown`, `ScopedFormSuite.DropdownMulti`, `ScopedFormSuite.CheckboxGroup`, `ScopedFormSuite.Combobox`, `ScopedFormSuite.ComboboxMulti`, `ScopedFormSuite.TagPicker`, `ScopedFormSuite.Rating`, `ScopedFormSuite.ColorPicker`, `ScopedFormSuite.SwatchPicker`, `ScopedFormSuite.DatePicker`, `ScopedFormSuite.TimePicker`, `ScopedFormSuite.DateTime`, `ScopedFormSuite.FileInput`, `ScopedFormSuite.Object`, `ScopedFormSuite.Array`, `ScopedFormSuite.ObjectSection`, `ScopedFormSuite.ArraySection`, `ScopedFormSuite.Error`, `ScopedFormSuite.useField` |
| `UntypedFormSuite` | Uses the adapter, binding or lifecycle contract above. |
| `CompatibleFieldPath` | Uses the adapter, binding or lifecycle contract above. |
| `FormFieldPresentationProps` | `FormFieldPresentationProps.label`, `FormFieldPresentationProps.hint`, `FormFieldPresentationProps.required`, `FormFieldPresentationProps.orientation`, `FormFieldPresentationProps.disabled`, `FormFieldPresentationProps.className`, `FormFieldPresentationProps.fieldProps` |
| `FormControlProps` | `FormControlProps.name`, `FormControlProps.rules` |
| `FormInputProps` | `FormInputProps.type` |
| `FormInputType` | Uses the adapter, binding or lifecycle contract above. |
| `FormTextareaProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormSearchBoxProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormCheckboxProps` | `FormCheckboxProps.triState` |
| `FormSwitchProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormSpinButtonProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormSliderProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormSelectProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormRadioGroupProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormDropdownProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormDropdownMultiProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormCheckboxGroupProps` | `FormCheckboxGroupProps.onChange`, `FormCheckboxGroupProps.onBlur` |
| `FormComboboxProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormComboboxMultiProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormTagPickerProps` | `FormTagPickerProps.colorfulTagAvatars`, `FormTagPickerProps.'aria-describedby'`, `FormTagPickerProps.id`, `FormTagPickerProps.onChange`, `FormTagPickerProps.onBlur`, `FormTagPickerProps.noOptionsContent`, `FormTagPickerProps.strings` |
| `FormTagPickerStrings` | `FormTagPickerStrings.noOptions` |
| `FormRatingProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormColorPickerProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormSwatchPickerProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormDatePickerError` | Uses the adapter, binding or lifecycle contract above. |
| `FormDatePickerValidationResult` | `FormDatePickerValidationResult.error` |
| `FormDatePickerStrings` | `FormDatePickerStrings.months`, `FormDatePickerStrings.shortMonths`, `FormDatePickerStrings.days`, `FormDatePickerStrings.shortDays`, `FormDatePickerStrings.goToToday`, `FormDatePickerStrings.prevMonthAriaLabel`, `FormDatePickerStrings.nextMonthAriaLabel`, `FormDatePickerStrings.prevYearAriaLabel`, `FormDatePickerStrings.nextYearAriaLabel`, `FormDatePickerStrings.prevYearRangeAriaLabel`, `FormDatePickerStrings.nextYearRangeAriaLabel`, `FormDatePickerStrings.monthPickerHeaderAriaLabel`, `FormDatePickerStrings.yearPickerHeaderAriaLabel`, `FormDatePickerStrings.closeButtonAriaLabel`, `FormDatePickerStrings.weekNumberFormatString`, `FormDatePickerStrings.selectedDateFormatString`, `FormDatePickerStrings.todayDateFormatString`, `FormDatePickerStrings.dayMarkedAriaLabel` |
| `FormDatePickerFormatting` | `FormDatePickerFormatting.formatDay`, `FormDatePickerFormatting.formatMonth`, `FormDatePickerFormatting.formatYear`, `FormDatePickerFormatting.formatMonthDayYear`, `FormDatePickerFormatting.formatMonthYear` |
| `FormDatePickerDayOfWeek` | Uses the adapter, binding or lifecycle contract above. |
| `FormDatePickerFirstWeekOfYear` | Uses the adapter, binding or lifecycle contract above. |
| `FormDatePickerControlProps` | `FormDatePickerControlProps.appearance`, `FormDatePickerControlProps.allowTextInput`, `FormDatePickerControlProps.inputValidationMessage`, `FormDatePickerControlProps.allFocusable`, `FormDatePickerControlProps.borderless`, `FormDatePickerControlProps.className`, `FormDatePickerControlProps.disabled`, `FormDatePickerControlProps.disableAutoFocus`, `FormDatePickerControlProps.defaultOpen`, `FormDatePickerControlProps.dateTimeFormatter`, `FormDatePickerControlProps.firstDayOfWeek`, `FormDatePickerControlProps.firstWeekOfYear`, `FormDatePickerControlProps.formatDate`, `FormDatePickerControlProps.highlightCurrentMonth`, `FormDatePickerControlProps.highlightSelectedMonth`, `FormDatePickerControlProps.initialPickerDate`, `FormDatePickerControlProps.inlinePopup`, `FormDatePickerControlProps.isMonthPickerVisible`, `FormDatePickerControlProps.maxDate`, `FormDatePickerControlProps.minDate`, `FormDatePickerControlProps.open`, `FormDatePickerControlProps.openOnClick`, `FormDatePickerControlProps.onOpenChange`, `FormDatePickerControlProps.onSelectDate`, `FormDatePickerControlProps.onValidationResult`, `FormDatePickerControlProps.parseDateFromString`, `FormDatePickerControlProps.placeholder`, `FormDatePickerControlProps.required`, `FormDatePickerControlProps.showCloseButton`, `FormDatePickerControlProps.showGoToToday`, `FormDatePickerControlProps.showMonthPickerAsOverlay`, `FormDatePickerControlProps.showWeekNumbers`, `FormDatePickerControlProps.strings`, `FormDatePickerControlProps.today`, `FormDatePickerControlProps.underlined` |
| `FormDatePickerProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormTimePickerError` | Uses the adapter, binding or lifecycle contract above. |
| `FormTimePickerChangeData` | `FormTimePickerChangeData.selectedTime`, `FormTimePickerChangeData.selectedTimeText`, `FormTimePickerChangeData.error` |
| `FormTimePickerParseResult` | `FormTimePickerParseResult.date`, `FormTimePickerParseResult.error` |
| `FormTimePickerHour` | Uses the adapter, binding or lifecycle contract above. |
| `FormTimePickerControlProps` | `FormTimePickerControlProps.appearance`, `FormTimePickerControlProps.clearable`, `FormTimePickerControlProps.className`, `FormTimePickerControlProps.dateAnchor`, `FormTimePickerControlProps.defaultOpen`, `FormTimePickerControlProps.disabled`, `FormTimePickerControlProps.endHour`, `FormTimePickerControlProps.inputValidationMessage`, `FormTimePickerControlProps.freeform`, `FormTimePickerControlProps.formatDateToTimeString`, `FormTimePickerControlProps.hourCycle`, `FormTimePickerControlProps.increment`, `FormTimePickerControlProps.inlinePopup`, `FormTimePickerControlProps.open`, `FormTimePickerControlProps.onOpenChange`, `FormTimePickerControlProps.onTimeChange`, `FormTimePickerControlProps.parseTimeStringToDate`, `FormTimePickerControlProps.placeholder`, `FormTimePickerControlProps.showSeconds`, `FormTimePickerControlProps.size`, `FormTimePickerControlProps.startHour` |
| `FormTimePickerProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormDateTimeProps` | `FormDateTimeProps.id`, `FormDateTimeProps.'aria-describedby'`, `FormDateTimeProps.dateLabel`, `FormDateTimeProps.datePickerProps`, `FormDateTimeProps.timeLabel`, `FormDateTimeProps.timePickerProps` |
| `FormFileInputProps` | `FormFileInputProps.browseLabel`, `FormFileInputProps.clearLabel`, `FormFileInputProps.formatFiles` |
| `FormOption` | `FormOption.value`, `FormOption.label`, `FormOption.disabled`, `FormOption.textValue` |
| `FormValueCodec` | `FormValueCodec.toControl`, `FormValueCodec.fromControl` |
| `FormOptionsSource` | `FormOptionsSource.options`, `FormOptionsSource.children` |
| `FormHsvColor` | `FormHsvColor.h`, `FormHsvColor.s`, `FormHsvColor.v`, `FormHsvColor.a` |
| `FormDateTimeValue` | `FormDateTimeValue.date`, `FormDateTimeValue.time` |
| `FormObjectPath` | Uses the adapter, binding or lifecycle contract above. |
| `FormObjectPathValue` | Uses the adapter, binding or lifecycle contract above. |
| `FormArrayPath` | Uses the adapter, binding or lifecycle contract above. |
| `FormArrayItemValue` | Uses the adapter, binding or lifecycle contract above. |
| `FormObjectContext` | `FormObjectContext.Fields`, `FormObjectContext.error`, `FormObjectContext.errors` |
| `FormObjectProps` | `FormObjectProps.name`, `FormObjectProps.children` |
| `FormArrayItemContext` | `FormArrayItemContext.key`, `FormArrayItemContext.index`, `FormArrayItemContext.total`, `FormArrayItemContext.isFirst`, `FormArrayItemContext.isLast`, `FormArrayItemContext.Fields` |
| `FormArrayContext` | `FormArrayContext.items`, `FormArrayContext.append`, `FormArrayContext.prepend`, `FormArrayContext.insert`, `FormArrayContext.remove`, `FormArrayContext.update`, `FormArrayContext.replace`, `FormArrayContext.move`, `FormArrayContext.swap`, `FormArrayContext.error`, `FormArrayContext.errors` |
| `FormArrayRules` | Uses the adapter, binding or lifecycle contract above. |
| `FormArrayProps` | `FormArrayProps.name`, `FormArrayProps.rules`, `FormArrayProps.children` |
| `FormObjectComponent` | Uses the adapter, binding or lifecycle contract above. |
| `FormArrayComponent` | Uses the adapter, binding or lifecycle contract above. |
| `FormSectionName` | `FormSectionName.title`, `FormSectionName.'aria-label'`, `FormSectionName.'aria-labelledby'` |
| `FormSectionProps` | `FormSectionProps.description`, `FormSectionProps.disabled`, `FormSectionProps.appearance` |
| `FormObjectSectionProps` | Uses the adapter, binding or lifecycle contract above. |
| `FormArraySectionProps` | `FormArraySectionProps.name`, `FormArraySectionProps.createItem`, `FormArraySectionProps.renderItem`, `FormArraySectionProps.minItems`, `FormArraySectionProps.maxItems`, `FormArraySectionProps.removable`, `FormArraySectionProps.reorderable`, `FormArraySectionProps.itemLabel`, `FormArraySectionProps.strings` |
| `FormArraySectionStrings` | `FormArraySectionStrings.add`, `FormArraySectionStrings.empty`, `FormArraySectionStrings.remove`, `FormArraySectionStrings.moveUp`, `FormArraySectionStrings.moveDown`, `FormArraySectionStrings.added`, `FormArraySectionStrings.removed`, `FormArraySectionStrings.moved` |
| `FormObjectSectionComponent` | Uses the adapter, binding or lifecycle contract above. |
| `FormArraySectionComponent` | Uses the adapter, binding or lifecycle contract above. |

ComboboxMulti displays selected labels inside the input at rest. Focus switches to the search query; blur clears that transient query and restores selected labels without changing the stored selections. No selected-value caption is rendered beneath the field. ColorPicker defaults to a responsive square color area capped at 300px, with sliders fitting the same available width. ArraySection uses compact arrow/delete actions with localized accessible labels and tooltips, plus an Add icon alongside its action text.

ObjectSection content and ArraySection item content stack direct fields with `gap.medium` (12px). Their description/error output uses `gap.small` (8px) from the content. Custom layout inside a caller-supplied wrapper remains caller-owned.

The coordinated Fluent/Charts baseline uses Charts 9.3.27 and locks Tabster 8.8.0 with shared Keyborg 2.14.1. Preserve this compatible navigation subtree: Tabster 8.8.1 resolves Keyborg 3 alongside Fluent’s Keyborg 2 and the actual React 17 cleanup probe detects a registry collision. This lock selection satisfies Fluent’s published range; no overrides are used. Native providers may have distinct package versions, but must resolve the same shared contexts, theme and Griffel renderer.

## Colorful TagPicker avatars

`FormTagPickerProps.colorfulTagAvatars?: boolean` is an optional presentation flag, default false. Omission/false retains today's text-only selected tags. True adds a Fluent colorful avatar to each selected tag, using its string option identity as `idForColor` and its plain display text for initials. Missing/plain-text-free labels use Fluent's fallback icon; rich consumer labels are preserved alongside the avatar. Labels can change without changing the identity's color; the finite palette can repeat colors.

```tsx
<Order.TagPicker name="people" label="People" options={people} colorfulTagAvatars />
```

This uses the existing Form/Fluent providers and theme. It preserves the current input, typed query, RHF values, codec and selection events. The option does not recolor suggestions or the tag container. The sample's **Form → Variable field → TagPicker** exposes **Colorful tag avatars**, initially off, for direct comparison.

## Localization

Form uses its native SPFx `ApveeSpfxReactToolkitFormStrings` resource by default. Configure a partial `FormStringsOverride` through `FormUseFormProps.strings` on the owning `useForm`, or on a Root that initializes its own methods. A Root supplied with `methods` inherits that hook's dictionary and rejects a separate `strings` option. Localization options are removed before RHF initialization.

```tsx
const Order = createForm<{ attachments: FileList | null }>();
function OrderForm() {
  const methods = Order.useForm({
    defaultValues: { attachments: null },
    strings: {
      unexpectedFailure: 'Please retry saving your order.',
      fileInput: { browseLabel: 'Choose attachments', clearLabel: 'Remove attachments' },
      datePicker: { goToToday: 'Today' },
      arraySection: { itemLabel: position => `Row ${position}` },
    },
  });
  return <Order.Root methods={methods} onSubmit={save}>
    <Order.FileInput name="attachments" label="Attachments" />
  </Order.Root>;
}
```

`FormStrings` is the resolved dictionary. `FormStringsOverride` accepts only known optional fields. Nested groups merge field by field; calendar arrays replace whole arrays (months start with January, weekdays with Sunday). Explicit empty strings remain empty.

| Group | Purpose and precedence |
| --- | --- |
| `FormStrings.unexpectedFailure`, `FormStringsOverride.unexpectedFailure` | Unexpected resolver, rule, submit and invalid-callback failures. Expected rule/resolver messages and errors installed with `setError` remain caller-owned. |
| `FormStrings.datePicker`, `FormStringsOverride.datePicker` | Full calendar names, navigation labels and accessible format strings, using `FormDatePickerStrings`. Direct DatePicker `strings` fields win. |
| `FormStrings.dateInputValidationMessage`, `FormStrings.timeInputValidationMessage`, `FormStringsOverride.dateInputValidationMessage`, `FormStringsOverride.timeInputValidationMessage` | Incomplete manual draft messages. Direct `inputValidationMessage` wins. |
| `FormStrings.tagPicker`, `FormStringsOverride.tagPicker` (`noOptions`) | Existing `FormTagPickerStrings` content slot. Direct `noOptionsContent`, then leaf `strings.noOptions`, then family content applies. Empty/boolean content retains the owned status region. |
| `FormStrings.fileInput`, `FormStringsOverride.fileInput` (`browseLabel`, `clearLabel`) | File chooser action labels. Direct control labels win. File names and `formatFiles` remain caller content. |
| `FormStrings.arraySection`, `FormStringsOverride.arraySection` | Existing `FormArraySectionStrings` text/callbacks plus `itemLabel(position)` for one-based default item names. Leaf `strings` fields and direct full-item `itemLabel` win. Action callbacks receive literal labels; native templates preserve dollar/braced caller text. |

Hook metadata isolates each form and publishes the dictionary at React commit. Asynchronous resolver/rule/submit failures read the latest committed dictionary when the failure is created, including hook methods used outside a mounted Root. Stored errors and announcements are event snapshots: changing `strings` does not rewrite them, reset values, rerun validation, or alter dirty/touched state. The next validation/submission uses current text.

UI localization does not change strict default date parsing, caller parsers/formatters, time anchors/hour cycles, browser-native file selection, RHF error contracts, array focus behavior, or domain values. The existing RHF and Fluent compat peers remain required. Authenticated SharePoint rendering requires native resource registration in the host.

# SettingsGroup

`SettingsGroup` places related preferences in a Fluent Card with an optional header. It composes native semantic text, `useSx` typography and Fluent UI 9; no SPFx provider is required. Application controls own values, validation and persistence. Ordinary rows keep caller-owned labels; opt-in `field` rows inherit naming from the row for one built-in [Form](./form.md#settings-rows-as-form-fields) control. Choose [AccordionGroup](./accordion-group.md) for optional details users can reveal.

```tsx
import { SettingsGroup } from '@apvee/spfx-react-toolkit';
import { Input, Switch } from '@fluentui/react-components';

<SettingsGroup title="Notification preferences" description="Choose how we contact you.">
  <SettingsGroup.Item title="Delivery address" description="Your work email.">
    <Input aria-label="Notification email" value={email} onChange={(_, data) => setEmail(data.value)} />
  </SettingsGroup.Item>
  <SettingsGroup.Item title={<span>Account <em>alerts</em></span>}>
    <Switch aria-label="Receive account alerts" />
  </SettingsGroup.Item>
</SettingsGroup>
```

Import from the package root, `/controls`, or `/controls/settings-group`; these share the canonical implementation. `SettingsGroup.Item` is the row component; there is no separate public flat Item export.

## SettingsGroupProps

Native div attributes/events, `className`, consumer `style`, `dir`, `lang`, `hidden` and the root HTMLDivElement ref are forwarded. The group adds no implicit accessible name, group role, form behavior or generated label relationships. `dangerouslySetInnerHTML` is unsupported because the component owns its content.

| Member | Type | Behavior |
| --- | --- | --- |
| `SettingsGroupProps.title` | `React.ReactNode` | Optional phrasing content inside the semantic title heading; strings remain literal. |
| `SettingsGroupProps.description` | `React.ReactNode` | Optional supporting phrasing content inside a span. |
| `SettingsGroupProps.icon` | `React.ReactNode` | Optional decorative icon in a fixed 24px header slot. |
| `SettingsGroupProps.action` | `React.ReactNode` | Optional caller-owned trailing header action. |
| `SettingsGroupProps.headingLevel` | `SettingsGroupHeadingLevel` | Default 2; selects the title heading wrapper for strings and JSX. |
| `SettingsGroupProps.appearance` | `'filled' \| 'filled-alternative'` | Fluent Card appearance, default filled. |
| `SettingsGroupProps.children` | `React.ReactNode` | Items and arbitrary content in declaration order. |
| `SettingsGroupProps.dangerouslySetInnerHTML` | `never` | Unsupported HTML injection pathway. |

## SettingsGroupItemProps

The row is a div with its own native attributes/events/ref. With `field` omitted or false, titles are visible text, **not implicit labels** for controls. Supply a Fluent control label, Field, or explicit application-owned label IDs. Rows do not clone children, intercept control events or own application state. Keyed rows retain child identity when reordered by the application.

| Member | Type | Behavior |
| --- | --- | --- |
| `SettingsGroupItemProps.title` | `React.ReactNode` | Optional phrasing content; a visible span normally, or the accessible name of one Form control in field mode. |
| `SettingsGroupItemProps.field` | `boolean` | Default false; opt in to automatic label/hint composition with exactly one built-in Form logical control. |
| `SettingsGroupItemProps.description` | `React.ReactNode` | Optional supporting phrasing content inside a span. |
| `SettingsGroupItemProps.icon` | `React.ReactNode` | Optional decorative icon in a fixed 20px Item slot. |
| `SettingsGroupItemProps.action` | `React.ReactNode` | Optional secondary trailing action. |
| `SettingsGroupItemProps.children` | `React.ReactNode` | Caller-owned primary controls or custom content. |
| `SettingsGroupItemProps.dangerouslySetInnerHTML` | `never` | Unsupported HTML injection pathway. |

## Automatic Form field rows

Set `SettingsGroup.Item field` inside a Form root and supply a meaningful `title`. The row renders its title and optional `description` once; the built-in Form adapter consumes them as its name and hint. Omit the adapter's `label` and `hint`: repeating either is rejected with a diagnostic. No manually authored `aria-label` or `aria-labelledby` is needed. Ordinary rows continue to support independent Fluent controls and custom content as above.

```tsx
import { createForm, SettingsGroup } from '@apvee/spfx-react-toolkit';

const Preferences = createForm<{ email: string; enabled: boolean }>();

<Preferences.Root defaultValues={{ email: '', enabled: true }} onSubmit={savePreferences}>
  <SettingsGroup title="Notifications">
    <SettingsGroup.Item field title="Delivery address" description="Your work email.">
      <Preferences.Input name="email" required rules={{ required: 'Enter an email.' }} />
    </SettingsGroup.Item>
    <SettingsGroup.Item field title={<span>Account <em>alerts</em></span>} description="Receive account updates.">
      <Preferences.Switch name="enabled" />
    </SettingsGroup.Item>
  </SettingsGroup>
  <Preferences.Error />
  <Preferences.Submit>Save</Preferences.Submit>
  <Preferences.Reset>Reset</Preferences.Reset>
</Preferences.Root>
```

A native control uses a real label associated with its actual input, preserving explicit control IDs. Clicking the title focuses the input or activates a checkbox/switch through native label behavior. Composite adapters name their semantic group or control automatically; a group label does not promise to activate each inner control. DateTime retains distinct `dateLabel`/`timeLabel` names for its parts; TagPicker receives the row name for its input and tag group, including rich JSX titles. Names do not spread to secondary actions, nested settings rows or DateTime parts.

The row description, caller-supplied `aria-describedby` IDs and Fluent Field validation messages coexist. Validation output stays beside the control. The visible label reflects the adapter's `required` and effective disabled state, including disabled Form roots/scopes. `required` remains presentation metadata; add `rules`, a schema or a resolver to enforce it. Form continues to own values, errors, omission of disabled values, save and reset. Editing a settings row does not save automatically.

Use exactly one built-in logical adapter in each field row; DateTime counts as one. Multiple owning adapters are rejected. For several independently named controls, use separate field rows or omit `field` and label each control explicitly. `Form.useField` custom bindings and standalone Fluent controls do not consume this integration: give them application-owned labels/hints and use an ordinary row. This composition requires the existing Form peers and Fluent theme described in [Form dependencies](./form.md#dependencies); it adds no dependency or SPFx provider requirement. See [scoped Fields examples](./form.md#settings-rows-as-form-fields).

## SettingsGroupHeadingLevel

Accepts 1 through 6. Choose the group title wrapper for its document hierarchy; it applies to both strings and JSX. Ordinary Item title/description wrappers are spans; field mode uses a label for a native labelable target. All title/description slots receive inherited defaults: group title `subtitle2`/primary, Item title `body1Strong`/primary and description `caption1`/subtle, with zero block margin and zero title-description gap.

Strings are rendered literally: `**bold**` stays `**bold**`; the group does not parse Markdown or depend on FormattedText. Use noninteractive phrasing JSX such as `<strong>Bold</strong>` or `<em>Emphasis</em>` for formatting. JSX children retain their props, identity and explicit styling; they inherit the wrapper defaults when they do not provide their own. Do not supply nested headings or block markup inside these text wrappers. Supply `undefined` for absent optional content.

## Layout, dependencies and limitations

Fluent Card medium owns the 12px exterior padding, border and radius. The body sets its row gap to zero; the inner content row has a 44px minimum matching native Accordion triggers, independent of the outer divider padding. Adjacent direct Item DOM siblings receive a thin separator with one Card-padding interval of additional spacing, split equally before and after it (6px + 6px at medium). Both halves derive from Fluent’s public `cardCSSVars.cardSizeVar` (`--fui-Card--size`), the same variable as Card padding; arbitrary intervening content breaks adjacency. Consumer wrappers can therefore affect separators. The control preserves declaration order rather than collecting Items into a registry.

Each Settings Item uses a fluid wrapping flex row: text and controls remain side by side while they fit. Leading text grows and wraps around a 160px preferred/minimum width, capped at the available width for narrower groups. A compact Switch can remain beside text at 500px; a wider control wraps only when the combined regions and their 12px horizontal gap do not fit. Wrapped lines use an 8px vertical gap. Title and description remain adjacent with no extra gap. Custom leading content must support wrapping or shrinking within that bounded region.

The group establishes the toolkit `apvee-sx` inline-size container. Row placement follows actual available space and content widths; there is no blanket column layout below a fixed breakpoint. Existing layout/spacing descriptors use the host Griffel renderer and Fluent direction. The optional header sits 12px from the Card (Fluent medium spacing). Header icons have fixed 24px boxes and Item icons fixed 20px boxes; direct font icons inherit slot font size and SVG icons fill the slot. Consumer children keep their props and identity; complex custom icon content retains its own structure. Theme variables require a Fluent theme in scope; the package uses shared Fluent/Griffel peers without a Markdown parser dependency. Native `hidden` hides the group/row without unmounting controls; it is presentation, not authorization.

The lazy **SettingsGroup and AccordionGroup** sample edits caller-labeled values, actions, theme, direction, width and visibility and observes root refs. The **Form** sample exercises automatic field rows with required Reference and Enabled, shared save/reset and observable RHF state. Automated marker checks establish sample coverage only. Local browser execution and authenticated SharePoint validation are separate; see the [host checklist](../../SHAREPOINT-VALIDATION.md#settingsgroup-and-accordiongroup).

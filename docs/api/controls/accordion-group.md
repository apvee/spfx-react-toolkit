# AccordionGroup

[Documentation](../../INDEX.md) > [API reference](../INDEX.md#react-controls) > AccordionGroup

`AccordionGroup` groups optional details inside a Fluent Card. Fluent Accordion owns expansion and native trigger behavior; native semantic wrappers and `useSx` provide toolkit typography/layout. Use [SettingsGroup](settings-group.md) for preferences that should remain visible. No SPFx provider is required.

```tsx
import { AccordionGroup } from '@apvee/spfx-react-toolkit';
import { Field, Input } from '@fluentui/react-components';

<AccordionGroup title="Account details" defaultOpenItems={['profile']}>
  <AccordionGroup.Item value="profile" title="Profile preferences" description="Optional account details.">
    <Field label="Profile note"><Input defaultValue="My note" /></Field>
  </AccordionGroup.Item>
  <AccordionGroup.Item value="advanced" title={<span>Advanced <em>preferences</em></span>}>
    Additional settings
  </AccordionGroup.Item>
</AccordionGroup>
```

Import from the root, `/controls`, or `/controls/accordion-group`; these share one implementation. Compose `AccordionGroup.Item`; no flat Item alias is public.

## AccordionGroupProps

The root forwards native div attributes/events, `className`, caller `style`, `dir`, `lang`, `hidden` and its HTMLDivElement ref. It introduces no implicit root name/group role and does not own form state or persistence. Fluent Card uses medium native exterior padding; internal row gap is zero.

| Member | Type | Behavior |
| --- | --- | --- |
| `AccordionGroupProps.title` | `span` content slot | Optional phrasing content inside the semantic title heading; strings remain literal. |
| `AccordionGroupProps.description` | `span` content slot | Optional header supporting text. |
| `AccordionGroupProps.icon` | `span` content slot | Optional decorative icon in a fixed 24px header slot. |
| `AccordionGroupProps.action` | `div` content slot | Caller-owned header action outside triggers. |
| `AccordionGroupProps.headingLevel` | `AccordionGroupHeadingLevel` | Default 2; selects the title heading wrapper for strings and JSX. |
| `AccordionGroupProps.appearance` | `'filled' \| 'alternate' \| 'outline'` | Default `filled`; `alternate` maps to Fluent `filled-alternative`; `outline` uses the native outlined Card. |
| `AccordionGroupProps.children` | `React.ReactNode` | Declared Items/content without cloning or a custom registry. |
| `AccordionGroupProps.openItems` | Fluent Accordion openItems | String/number or array of values; controlled current values; use with onToggle. |
| `AccordionGroupProps.defaultOpenItems` | Fluent Accordion defaultOpenItems | String/number or array of initial uncontrolled values. |
| `AccordionGroupProps.onToggle` | Fluent Accordion onToggle | Exact Fluent event/data signature, including data.openItems. |
| `AccordionGroupProps.multiple` | `boolean` | Default false; permits multiple expanded Items when true. |
| `AccordionGroupProps.collapsible` | `boolean` | Default true; allows the last expanded Item to close. |
| `AccordionGroupProps.dangerouslySetInnerHTML` | `never` | Unsupported owned-content injection pathway. |

```tsx
const [openItems, setOpenItems] = React.useState<Array<string | number>>(['profile']);
<AccordionGroup openItems={openItems} onToggle={(_, data) => setOpenItems(data.openItems)} multiple>
  <AccordionGroup.Item value="profile" title="Profile">Profile settings</AccordionGroup.Item>
  <AccordionGroup.Item value="advanced" title="Advanced">Advanced settings</AccordionGroup.Item>
</AccordionGroup>
```

Controlled values remain application-authoritative; the wrapper does not mirror them in separate state. Assign stable distinct Item values and React keys when rendering collections.

## AccordionGroupItemProps

The native div ref/attributes target the Item root. Content belongs in its panel; the trigger is supplied by Fluent. Supply a nonempty title or description to name that button: an aria-label on the Item div does not label the trigger. Before programmatically closing a panel that contains focus, the application should move focus to an appropriate visible control.

| Member | Type | Behavior |
| --- | --- | --- |
| `AccordionGroupItemProps.value` | `AccordionGroupValue` | Required stable expansion identifier. |
| `AccordionGroupItemProps.title` | `span` content slot | Optional trigger phrasing content inside a span. |
| `AccordionGroupItemProps.description` | `span` content slot | Optional supporting trigger phrasing content inside a span. |
| `AccordionGroupItemProps.icon` | `span` content slot | Optional decorative icon in a fixed 20px trigger slot. |
| `AccordionGroupItemProps.headingLevel` | `AccordionGroupHeadingLevel` | Default 3; semantic trigger heading wrapper. |
| `AccordionGroupItemProps.disabled` | `boolean` | Disables the Fluent trigger and preserves inherited disabled text color. |
| `AccordionGroupItemProps.unmountOnClose` | `boolean` | Default false; true unmounts closed panel children. |
| `AccordionGroupItemProps.children` | `React.ReactNode` | Panel content, including caller-owned labeled controls. |
| `AccordionGroupItemProps.dangerouslySetInnerHTML` | `never` | Unsupported owned-content injection pathway. |

## AccordionGroupHeadingLevel and AccordionGroupValue

`AccordionGroupHeadingLevel` accepts 1 through 6; root default 2 and Item default 3 should fit the surrounding document hierarchy. The root title heading wrapper applies to strings and JSX; the Item uses its own heading wrapper around the native Fluent trigger. `AccordionGroupValue` accepts string or number.

All text slot wrappers provide inherited defaults: group titles use subtitle2/primary, Item titles body1Strong/primary and descriptions caption1/subtle, with zero block margin and zero title-description gap. Disabled trigger text inherits Fluent disabled color. JSX children retain their props and explicit styling; they inherit the defaults when they do not specify their own.

Strings remain literal: `**bold**` stays `**bold**`; the group does not parse Markdown or depend on FormattedText. Use noninteractive phrasing JSX such as `<strong>Bold</strong>` or `<em>Emphasis</em>` for formatting. **Never place interactive JSX links, buttons, inputs or other interactive descendants in Item title, description or icon:** these regions render inside a button. Put actions in the group header's action or inside the panel. Nested headings and block markup inside text wrappers are unsupported. The component does not sanitize or rewrite arbitrary consumer JSX.

## Closed panels, layout and compatibility

The body Card supports only `filled`, `alternate` and `outline`, with `filled` as the default. `alternate` maps to Fluent Card's `filled-alternative` surface. `outline` uses Fluent's outlined surface and native border, without a filled background or shadow. Native Card owns padding, radius and child ancestry in every appearance; colors inherit the host theme. Changing appearance retains mounted panel controls and their state. The former `transparent` value is removed; use `outline` for an unfilled, visibly outlined Card. Replace former `filled-alternative` values with `alternate`.

```tsx
<AccordionGroup title="Details" appearance="filled">
  <AccordionGroup.Item value="profile" title="Profile">{profileContent}</AccordionGroup.Item>
</AccordionGroup>
```

The test web part exposes all three appearance values, requested versus rendered width, host theme inheritance and genuine Fluent SVG icons. Its direction-only provider is transparent and fills the bounded example container. Provider classes are generated for its effective direction, including RTL overrides.

By default closed panel children remain mounted, preserving uncontrolled inputs and component state. Closed content is hidden and excluded from keyboard focus through native inert compatibility for React 17. This requires a browser with native inert support; build success alone cannot prove focus behavior. Closing does not provide application persistence or a focus-restoration policy. `unmountOnClose` unmounts after Fluent closing motion and resets uncontrolled state after reopening; initially closed resettable content does not mount. Changing this policy while closed applies to the next opening/closing cycle. Retained child effects continue while closed; controlled values still belong to the application. Native hidden on the root/Item hides it without changing controlled expansion values.

The optional header sits 12px from the Card (Fluent medium spacing). Header icons have fixed 24px boxes and Item icons fixed 20px boxes; direct font icons inherit slot font size and SVG icons fill the slot without cloned child props. Complex custom icon content retains its own structure. Fluent Card medium provides the single 12px lateral inset, border and radius. The trigger has zero extra horizontal padding and the panel zero horizontal margin, so panel content and trigger text share the Card inset when no leading icon is present. Native trigger minimum height remains 44px; adjacent Items receive a thin separator with additional spacing equal to one Card-padding interval, split equally before and after it (6px + 6px at medium). Both halves derive from Fluent’s public `cardCSSVars.cardSizeVar` (`--fui-Card--size`), matching native Card padding; body gap stays zero. Divider spacing is measured between the native trigger/content row boxes; the 44px trigger minimum remains independent of the outer separator padding, so short and multiline titles retain native vertical centering. Fluent owns interaction and motion; the wrapper adds no animation timers or custom expansion state. Theme and direction use existing toolkit/Fluent/Griffel primitives and shared peer contracts; no Markdown parser or new dependency is required.

The lazy **SettingsGroup and AccordionGroup** sample selects one active group. Its AccordionGroup contains Profile details with a retained note and Advanced preferences with a notification switch; both triggers have Fluent icons. The sample and public control default to `filled`; `alternate` and `outline` remain selectable. Controlled expansion, multiple mode, disabled triggers, retained note state and unmount reset, inherited host theme/direction, width and root refs remain observable. Verify actual Tab behavior and state in a browser. Local and authenticated SharePoint results are distinct; see the [host checklist](../../development/sharepoint-validation/checklists/groups.md#settingsgroup-and-accordiongroup).

## Named content slots

Titles, descriptions and icons accept React content or a native `span` props object; actions accept a native `div` props object. The action prop remains `action`. Props objects support attributes, classes, consumer styles, refs and `children: (Component, props) => JSX` callbacks. The callback receives the fixed inner tag and composed props during rendering; inspection never executes it. Forward those props to preserve its attributes and ref.

```tsx
<AccordionGroup
  title={{ as: 'span', className: 'custom-title', children: 'Preferences' }}
  description={{ children: 'Literal **supporting text**' }}
/>
```

The headings, generated IDs, native trigger and panel association remain owned by AccordionGroup. A slot ref points to its inner container. Strings/JSX shorthand retain their existing DOM and literal formatting; props objects explicitly request an inner container even with empty children. Booleans and absent values suppress content; zero remains visible. Tags cannot replace owned headings or triggers; raw HTML is unsupported. Phrasing slots must contain appropriate noninteractive content (especially inside accordion triggers). The existing Fluent theme/host renderer is still required.

Description text inside an Item contributes to the native trigger content; the group adds no separate description association. Shorthand, including fragments, preserves the existing presence policy: empty text/fragments are absent, while whitespace text is present. Object/callback icon containers fill the owned shell; their direct SVG children retain 24px header/20px Item sizing, even with larger authored SVG dimensions. Callback output and custom components remain caller-owned and opaque.

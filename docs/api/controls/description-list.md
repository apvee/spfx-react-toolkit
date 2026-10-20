# DescriptionList

`DescriptionList` presents read-only labeled details of one record using native `dl`, `dt` and `dd` elements. Import the single compound value from the package root, `/controls` or `/controls/description-list`. Its only part is `DescriptionList.Item`; there is no flat Item runtime export.

```tsx
import { DescriptionList } from '@apvee/spfx-react-toolkit/controls/description-list';

<DescriptionList items={[
  { id: 'owner', label: 'Owner', value: 'Ada Lovelace' },
  { id: 'count', label: 'Count', value: 0 },
]} />;
```

Use stable unique IDs for data arrays. For conditional pairs, native pair attributes or rich JSX, compose Items instead:

```tsx
<DescriptionList aria-label="Record details" layout="horizontal">
  <DescriptionList.Item label="Owner">Ada Lovelace</DescriptionList.Item>
  <DescriptionList.Item label="Status"><Badge appearance="filled">Active</Badge></DescriptionList.Item>
</DescriptionList>
```

The Badge in this example comes from `@fluentui/react-components`. Use a form for editable fields and DataGrid to compare multiple records. Formatting, missing-value labels, loading, errors and empty-state messages belong to the consumer. Strings remain literal; use an explicit FormattedText value when markdown rendering is desired.

## DescriptionListProps

The native root is `dl`; its forwarded ref targets `HTMLDListElement`. Native attributes, events, `className`, `style`, language, direction, accessible naming and `hidden` are forwarded. `dangerouslySetInnerHTML` is excluded. Supply either `items` or compound `children`, never both.

| Member | Type | Behavior/default |
| --- | --- | --- |
| `DescriptionListProps.items` | `readonly DescriptionListEntry[]` | Ordered immutable pairs. Mutually exclusive with children. Empty arrays render nothing. |
| `DescriptionListProps.children` | `React.ReactNode` | Compound Items, fragments, conditions and mapped children. Mutually exclusive with items. Absent or boolean children render nothing; an empty fragment can leave an empty dl. |
| `DescriptionListProps.dangerouslySetInnerHTML` | `never` | Unsupported raw-HTML pathway; use items or children. |
| `DescriptionListProps.layout` | `DescriptionListLayout` | auto; stacks below 480px and aligns columns from 480px in the nearest named `apvee-sx` container. |

## DescriptionListLayout

`DescriptionListLayout` is `'auto' | 'horizontal' | 'vertical'`. Explicit horizontal or vertical overrides container responsiveness. Auto uses CSS container queries rather than viewport width. Without a named container, auto remains vertical, even on a wide screen. Establish the ancestor with the existing `container.inlineSize` descriptor:

```tsx
import { DescriptionList, container, useSx, width } from '@apvee/spfx-react-toolkit';

function RecordDetails() {
  const sx = useSx();
  return <section className={sx(container.inlineSize, width.full)}>
    <DescriptionList items={[{ id: 'owner', label: 'Owner', value: 'Ada' }]} />
  </section>;
}
```

The 480px transition matches toolkit `responsive.small` and the source Blocks medium threshold. Nested lists use their nearest named container. Layout changes retain child DOM, state and focus. Horizontal layout shares two columns across pairs, with the label column limited to 40% of the list width. Long labels and values wrap without truncation. Logical spacing supports RTL.

## DescriptionListEntry

| Member | Type | Behavior |
| --- | --- | --- |
| `DescriptionListEntry.id` | `string \| number` | Required stable unique React identity within the list; not a DOM identifier. |
| `DescriptionListEntry.label` | Non-nullable Fluent div content slot | Required meaningful visible term; text, JSX or Fluent div-slot customization. |
| `DescriptionListEntry.value` | Fluent div content slot with conditions | Associated definition; text, JSX or Fluent div-slot customization. Undefined leaves an empty definition. |

Zero is visible; empty strings, boolean conditions and absent values retain their definitions. Legacy Fluent/React slot values may include null, which also yields an empty definition; use undefined for new absence defaults. IDs, data formatting and meaningful labels are consumer responsibilities.

## DescriptionListItemProps

`DescriptionList.Item` requires a parent DescriptionList and inherits its layout. Rendering it outside a parent throws an actionable error. Native attributes/events and the forwarded `HTMLDivElement` ref target the grouping div. This div has no independent box in horizontal layout; customize the content slots for visible styling.

| Member | Type | Behavior |
| --- | --- | --- |
| `DescriptionListItemProps.label` | Non-nullable Fluent div content slot | Required visible term; customization stays inside the preserved dt. |
| `DescriptionListItemProps.dangerouslySetInnerHTML` | `never` | Unsupported raw-HTML pathway; use content slots. |
| `DescriptionListItemProps.children` | Fluent div content slot with conditions | Required associated value; customization stays inside the preserved dd. |

Native children and `dangerouslySetInnerHTML` are replaced/excluded by the slot contract. Consumer `style` escape hatches are forwarded. Repository-authored styling uses classes with the shared host renderer.

## Slots and accessibility

Scalar labels use semibold Fluent typography and scalar values regular typography. Rich JSX keeps its markup; paragraphs are not wrapped inside text spans. Fluent slot objects preserve allowed native content attributes, refs, consumer style, classes and render callbacks. Strings and JSX retain the existing inner div wrapper; undefined, booleans and legacy null leave the native cell empty. Explicit props objects request a div even when their static children are empty. `as` may only repeat `div`; it cannot select another tag. Slot refs identify the inner div, while root and Item refs continue to identify their native containers:

```tsx
<DescriptionList>
  <DescriptionList.Item label={{ title: 'Account owner', children: 'Owner' }}>
    {{ children: (Component, props) => <Component {...props}><p>Ada Lovelace</p></Component> }}
  </DescriptionList.Item>
</DescriptionList>
```

Slot types exclude `dangerouslySetInnerHTML`, and runtime resolution removes raw HTML and unsupported native keys and guards the fixed div tag, including for JavaScript callers. Callback children receive `(Component, props)` during normal rendering; they are never invoked to inspect presence. Keep the provided props on the rendered inner component to retain refs/classes/events. Arbitrary consumer JSX stays caller-owned and is not sanitized or cloned. React 17 can erase object restrictions through a `ReactNode` variable, so static checking alone cannot promise protection for erased values. Slots customize content inside invariant native term/definition cells. The list adds no keyboard behavior or focus stops; links and buttons retain their own native behavior. Supply an accessible name on the root when needed. Native hidden works on root and individual compound pairs in either layout. Fluent theme variables and normal forced-color adjustment preserve host colors; actual screen-reader and host behavior require separate validation.

## Dependencies and working sample

No SPFx provider or new dependency is required. The control uses the existing shared React 17, Fluent UI 9 and Griffel peers, inherits the host theme/direction/renderer and registers no PnP features. Auto layout requires browser support for CSS container queries; explicit layouts work without a named container.

The lazy Description List tab shows one document detail with five entries. Layout, width, data/compound API, order and content variations modify that same list; diagnostics inspect its refs and callbacks without adding a nested list. The heading and one outline configuration precede the undecorated subject. Nested container, pathological values and independent/RTL comparisons remain in the [versioned regression fixture](../../../tests/fixtures/demo-control-isolation/description-list-regressions.tsx). Browser runners exercise the product route and then all preserved conformance assertions.

See [development checks](../../DEVELOPMENT.md#description-list-checks) and [authenticated acceptance](../../SHAREPOINT-VALIDATION.md#description-list). Authenticated SharePoint, Teams/iframe and screen-reader checks remain **NOT EXECUTED** until separately recorded. Static demo checks do not execute the rendered scenario or establish tenant behavior.

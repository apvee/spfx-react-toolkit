# Control presentation slots

Named presentation props such as `title`, `description`, `icon`, `action` and `actions` support Fluent shorthand and props objects. This convention adds customization of an inner content container while the control owns its semantic elements, state and layout. Each public control documents the tags, content constraints and defaults of its own regions; the shared resolver and type aliases are private implementation details, with no public slot helper or `/controls/slots` entry.

## Supported values

| Value | Rendering contract |
| --- | --- |
| String, number, JSX, fragment, array, iterable or portal | React shorthand. Zero is content; JSX keeps its props and identity. |
| `false` or `true` | Empty content, allowing `condition && content`. Existing control-specific shell/default policies still apply. |
| `undefined` | Ordinary absence; a control may supply its documented default. Use this for new absence/default state. |
| `null` | Legacy React/Fluent interoperability. Whether it suppresses content or selects a default depends on the existing control contract. |
| Props object | Requests an inner container, including when its static children are empty. Supports permitted native attributes, `data-*`, classes, consumer style, ref and events. |
| Props object with `children: (Component, props) => ...` | Fluent render callback for the inner container, called only when that region renders. |
| Direct function value | Unsupported; put a render callback in the object's `children`. |

```tsx
import * as React from 'react';
import { Button } from '@fluentui/react-components';
import { ContentCard, useSx, typography } from '@apvee/spfx-react-toolkit';

function Summary({ showDescription }: { showDescription: boolean }): React.ReactElement {
  const sx = useSx();
  const titleRef = React.useRef<HTMLSpanElement>(null); // React 17 ref initialization.
  return (
    <ContentCard
      headingLevel={2}
      title={{ children: 'Workspace', className: sx(typography.body1Strong), ref: titleRef }}
      description={showDescription && 'Three members'}
      actions={{
        'data-region': 'workspace-actions',
        children: (Component, props) => (
          <Component {...props}><Button appearance="primary">Open workspace</Button></Component>
        ),
      }}
    />
  );
}
```

The title ref here identifies the inner span. `headingLevel` still selects the outer h2. Callback props include composed classes, attributes and refs: forward them to the supplied `Component`. Callbacks may run multiple times under ordinary React rendering. Returning `undefined` is supported; it does not remove an outer semantic shell owned by the control. Never execute callbacks or render consumer components to discover whether a region is present.

Concrete slot objects reject unsupported tags, raw HTML, feature-owned keys and direct functions in the type contract. React 17 `ReactNode` variables can erase those restrictions because its legacy fragment type admits broadly shaped objects. Do not treat acceptance of an erased value as complete static validation; the runtime container guards still apply. Arbitrary JSX/callback output remains opaque.

## Tags and semantic ownership

Ordinary slot objects have a fixed native tag. `as` may repeat that tag, but cannot replace it with a heading, link, button or arbitrary component. A documented bounded alternative is allowed only where the feature needs it: `PreviewCard.logo` permits `div` or `img`. This does not make every slot polymorphic.

| Region | Owned outer structure | Customizable content |
| --- | --- | --- |
| Group, Stage, StateMessage and card titles | Heading chosen by the control's `headingLevel` | Phrasing content; props/callback container is a span. |
| Settings field title, Form label and DateTime part labels | Native Label/Field, IDs, htmlFor, required/disabled state | Inner phrasing content; refs identify the span. |
| Accordion item title, description and icon | Heading, native trigger, aria-controls and associated panel | Phrasing, non-interactive content inside the trigger. |
| Stage header action | Surface naming and native Close button | Application content in the `action` div region; action and Close share one row, vertically centered with the text block at all widths. Close stays separately owned at the logical end. |
| StateMessage | Live text region separate from artwork/actions; native action Button/link | Text, visual div and phrasing action label; the action icon retains its native Button contract. |
| DescriptionList | dl, dt and dd | Existing inner div slots, including shorthand wrappers. |
| DataGrid messages and column content | MessageBar/status, table headers/cells and retry behavior | Inner div messages or span header/placeholder objects; sorting and naming remain owned. |
| Form sections | Fieldset, legend and accessible name/description associations | Title span and description div; title content does not replace the fieldset naming contract. |
| Cards | Native Card/Header/Preview/Footer, root div/ref and heading | Title/icon spans, description/headerActions/footer/actions divs; PreviewCard adds preview/fallback divs and bounded logo. |

Phrasing regions cannot contain nested headings or block markup. Content inside a native trigger or label must also meet that region's interaction/naming constraints. The consumer owns the accessibility and semantics of arbitrary JSX and callback output. The resolver filters container props; it does not sanitize the consumer subtree, clone its elements, disable its controls or infer a reliable accessible name from opaque components.

`dangerouslySetInnerHTML`, unsupported native keys and feature-owned keys are excluded from concrete slot objects and filtered at runtime. Fixed-tag guards also apply to JavaScript callers. Consumer attributes cannot override owned outer IDs, associations, roles or trigger behavior. Each control defines the specific invariant keys it reserves.

## Presence, defaults and preserved DOM

Ordinary shorthand retains the existing DOM; adding object support does not introduce a wrapper around every string or JSX element. DescriptionList deliberately preserves its existing div wrappers. Props objects introduce the fixed inner container; their refs do not replace the control root, heading, Label or trigger ref.

Container request presence and meaningful content for an accessibility association are separate decisions. An empty object requests a region, while an empty static description need not add aria-describedby. Static inspection can consider children and explicit accessible naming. Opaque callbacks request their region without being invoked; their output remains the consumer's responsibility. Use repeatable iterable collections if a control must inspect presence before rendering.

Control defaults remain explicit and feature-specific:

- Groups preserve their previous shorthand presence policy, including whitespace handling. Empty description objects request the inner container without registering an empty Settings/Stage description association.
- StateMessage retains its preset/default policies and separate live-region boundary. RelativeTime resolves `fallback` only for invalid input; valid time/Tooltip output is unchanged. WebPartShell keeps its persisted title/editor contract; `action` objects request the existing action region only while showTitle is true. showTitle=false suppresses the entire header and hidden action callbacks. titlePosition moves title and actions together inside/outside an optional Card.
- DataGrid retains nullish defaults for status/placeholder content, defined empty/boolean suppression and the legacy `error={true}` empty error bar. Its actions-column header default applies only when header is `undefined`.
- Form labels are omitted only for `undefined`; legacy null, boolean or empty labels still request the native Label. Hint null omits content, and undefined preserves native `fieldProps.hint` fallback. Section associations and TagPicker accessible-name validation remain feature-owned. No-options callbacks do not run when options or custom option children render.
- PreviewCard derives media from `image` only when `preview` is `undefined`. Any defined preview overrides the image; empty shorthand suppresses the region, while object/callback slots request it even with empty output. Logo requires a requested preview region. Suppressed preview/fallback/logo callbacks are not evaluated. Image failure uses `previewFallback` when defined, otherwise meaningful alt text; failure resets only for a new src.

See the individual control references for exact defaults and branch precedence. Do not use one generic presence rule to silently change existing behavior.

## Formatting, styling and refs

Slot resolution does not interpret Markdown or impose global truncation. Stage, SettingsGroup, AccordionGroup, DescriptionList, DataGrid presentation content and Form presentation content retain literal strings. Stage header title strings have a one-line ellipsis limit and description strings a two-line ellipsis limit across Dialog, Drawer and Popover, including props object string children. The full text remains in the DOM with the owned IDs and accessible associations; JSX/callbacks own their presentation. StateMessage retains its literal title and existing inline Markdown description formatting. `FormattedText` keeps its own children and presentation API. Card strings are literal; title/description strings, including object children strings, have a two-line limit. JSX/callbacks own their presentation. Compose `FormattedText` explicitly when Markdown is wanted.

Compose container props in this order: toolkit defaults, allowed consumer attributes, control-owned invariants. Classes compose in default/toolkit/consumer/owned order. Use the existing `useSx` and Griffel host renderer for repository-authored styling. Where an audited dependency boundary forbids the SX engine, use the same private renderer with the existing Griffel `mergeClasses` string composer, as RelativeTime and the Form field adapter do. This is class composition, not permission to generate inline styles or a new renderer.

Consumer `style` remains a forwarded escape hatch. Default, consumer and owned refs are composed and deduplicated; required toolkit refs must survive customization. Preserve the control's existing event ordering and cancellation behavior. Native action buttons keep their own appearance, size, ref and events; Card disabled state does not set disabled on caller-owned child controls; the native Fluent disabled overlay can block pointer interaction. Preserve that peer behavior and explicitly disable child controls when the application requires it.

## Boundaries for control authors

Use this standard for new or extended named presentation regions. Import the private slot types and renderer directly from their implementation modules. Public controls declare their own props and semantic types; do not depend on Stage types merely to share a heading convention. Document each slot's tag, ref target, allowed content, formatting, omission/default policy and owned semantic shell in JSDoc and its public guide. Add meaningful behavior/type coverage and an observable sample scenario for new behavior.

Keep these contracts feature-specific:

- Structural compound children, Stage triggers, DataGrid column markers, FormattedText/AutoSkeleton children and caller-owned card bodies.
- Data/domain callbacks and descriptors: renderCell, Form Object/Array/renderItem, formatFiles, option/value models and action records.
- Pagination labels, aria/tooltip strings and WebPartShell persisted properties.
- Native Fluent slots inherited from peer props, such as Button icons and Card checkbox/floatingAction; retain their native peer contract.

Use the documented `action` and `actions` prop names for each control. Stage.Header uses singular `action`; its former `actions` prop has no compatibility alias. Other controls preserve their existing names: cards use `headerActions` and `actions` for distinct regions. Do not add competing aliases, a generic action descriptor, state handling or navigation to obtain visual uniformity.

References: [controls index](./api/controls/INDEX.md), [ContentCard](./api/controls/content-card.md), [PreviewCard](./api/controls/preview-card.md), [development](./DEVELOPMENT.md), [package imports](./PACKAGE-IMPORTS.md).

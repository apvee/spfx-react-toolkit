# StateMessage

[Documentation](../../INDEX.md) > [API reference](../INDEX.md#react-controls) > StateMessage

`StateMessage` replaces the main content of a bounded section with a localized explanation, optional decorative artwork and actions. Use Fluent `MessageBar` for compact feedback beside ordinary content. It composes the toolkit's `FormattedText` and `useSx` with Fluent Button; no SPFx provider or additional dependency is required.

```tsx
import { StateMessage } from '@apvee/spfx-react-toolkit';

<StateMessage preset="empty" title="No results"
  description="Try changing the **active filters**."
  primaryAction={{ label: 'Clear filters', onClick: clearFilters }}
  tertiaryAction={{ label: 'Search help', href: '/help' }} />
```

Import from the root, `/controls`, or `/controls/state-message`. The paths share the same implementation. Artwork, theme classification and styles are private.

## StateMessageProps

Native div attributes/events, `className`, consumer `style`, `dir` and `lang` are forwarded. The ref targets the root HTMLDivElement. Content is owned by the control: `children` and `dangerouslySetInnerHTML` are unsupported, including raw HTML on actions.

| Member | Type | Behavior |
| --- | --- | --- |
| `StateMessageProps.title` | span content slot | Required literal title; strings are not parsed as Markdown. |
| `StateMessageProps.description` | span content slot | Optional inline Markdown or caller-owned phrasing JSX inside a paragraph. |
| `StateMessageProps.preset` | `StateMessagePreset` | Optional artwork and announcement defaults; never supplies application copy. |
| `StateMessageProps.headingLevel` | `StateMessageHeadingLevel` | Default 3, independent of subtitle2Stronger typography. |
| `StateMessageProps.visual` | div content slot | Undefined uses preset artwork; false hides it; custom content overrides the preset. |
| `StateMessageProps.primaryAction` | `StateMessageAction` | Highest emphasis, primary appearance. |
| `StateMessageProps.secondaryAction` | `StateMessageAction` | Supporting action, secondary appearance. |
| `StateMessageProps.tertiaryAction` | `StateMessageAction` | Lowest emphasis, subtle appearance. |
| `StateMessageProps.role` | `StateMessageRole` | Overrides the role of the text region. Alert is always explicit. |
| `StateMessageProps.dangerouslySetInnerHTML` | `never` | Unsupported; removed from untyped input too. |

### StateMessagePreset

| Preset | Artwork | Default text-region role |
| --- | --- | --- |
| `generic` | Default | None |
| `empty` | Empty | status |
| `error` | Error | status |
| `success` | Thanks | status |
| `welcome` | Hello | None |

Omitting the preset renders no automatic artwork or role. `role={undefined}` keeps the preset default; it does not disable its announcement semantics.

### StateMessageHeadingLevel and StateMessageRole

`StateMessageHeadingLevel` accepts 1 through 6; choose it for the surrounding heading hierarchy. `StateMessageRole` accepts `status` or `alert`. Title and description occupy that region; artwork and actions stay outside it. The control does not autofocus or restore focus. Role presence alone does not prove that a screen reader announces initial or updated content.

### StateMessageAction, StateMessageClickAction and StateMessageLinkAction

`StateMessageAction` is a union of a callback action and a navigation action. Both have a visible span content-slot `label`, optional Fluent `icon`, and compatible native/Fluent attributes such as disabled and className. Appearance, size, shape, root element and content belong to the control. Actions render in primary, secondary, tertiary order, with medium size and rounded shape.

`StateMessageClickAction` requires `onClick: React.MouseEventHandler<HTMLButtonElement>` and forbids href. Its native button has type=button, so it does not submit a containing form. `StateMessageLinkAction` requires `href: string` and accepts optional target, rel and `onClick: React.MouseEventHandler<HTMLAnchorElement>`; it uses Fluent Button as an anchor. Destinations are assigned as supplied by the application.

| Action member | Contract |
| --- | --- |
| `StateMessageClickAction.onClick` | Required button callback; asynchronous operations and their pending/error state belong to the application. |
| `StateMessageClickAction.href` | `never`; navigation belongs to the link branch. |
| `StateMessageLinkAction.href` | Required consumer-owned destination string; the action does not apply FormattedText's Markdown URL policy. |
| `StateMessageLinkAction.target` | Optional native anchor target. |
| `StateMessageLinkAction.rel` | Optional native anchor relationship; choose the relationship required by the destination/target. |
| `StateMessageLinkAction.onClick` | Optional native anchor callback; normal browser cancellation applies. |

```tsx
<StateMessage preset="error" title="Could not save" role="alert" headingLevel={2}
  primaryAction={{ label: 'Retry', onClick: retry }}
  secondaryAction={{ label: 'Copy details', onClick: copyDetails }}
  tertiaryAction={{ label: 'Open help', href: '/help', target: '_blank', rel: 'noopener noreferrer' }} />
```

Title/description JSX must fit the heading/paragraph semantics; labels must not nest interactive controls. Caller-owned JSX is not rewritten.

## Artwork and theme

The five subjects keep their original light and dark geometry, viewBox and palette. SVGs are decorative (`aria-hidden=true`, `focusable=false`) and gradient IDs are unique per mounted instance. Their OfficeDev/fluent-blocks MIT attribution is included in the package's THIRD-PARTY-NOTICES.md.

The nearest FluentProvider supplies ThemeContext_unstable. A private pure resolver compares relative luminance of colorNeutralBackground1 and colorNeutralForeground1: darker background than text selects dark; otherwise light. It accepts hex and rgb/rgba token colors. Missing or uninterpretable tokens fall back to light. This is an **heuristic**, not a universal Fluent isDark flag; CSS variables, unusual custom palettes and transparency may not describe the actual rendered surface. No DOM reads, SPFx hook, subscription or public theme option is added.

```tsx
<StateMessage preset="empty" title="No results" visual={false} />
<StateMessage title="Welcome" visual={<MyIllustration />} />
```

Custom visual content keeps its own accessibility semantics; it is not automatically hidden from assistive technology. Built-in SVGs stay private, and are loaded synchronously with the control. Bundle size includes all preset variants and FormattedText's Markdown support; importing generic hooks does not pull this control in.

## Layout and responsive behavior

The root establishes the toolkit's named inline-size container. All authored layout uses useSx/descriptors and local Griffel classes, with no generated inline style or imported Layout control.

| Part | Base | At `responsive.medium` |
| --- | --- | --- |
| Visual | Fluid width up to 280px; natural height | Same |
| Text maximum width | 240px | 600px |
| Action maximum width | 240px | 280px |

Medium means **640px of root content width**, excluding padding and border. It follows the toolkit's 480/640/1024px breakpoint vocabulary, rather than the source repository's names. The viewport does not decide presentation. A host with outer width 664px and default 12px padding per side has 640px available. Below the threshold, or without container-query support, base sizes remain active. The visual follows the BigMessage sizing approach: width fills the available space up to 280px, and SVG height follows its viewBox ratio. There is no square box or equal-area normalization. Wide subjects are naturally shorter; tall subjects take more vertical space, and the text/actions follow them without a reserved empty square. Custom SVGs should provide an intrinsic aspect ratio, usually through viewBox. Widths can shrink inside a narrow host; content has no fixed height. Custom CSS may change containment and sizing.

The root uses min-height semantics; vertical centering needs an available containing height. Long content can grow naturally. Responsive changes keep the same text/action DOM elements. See [styles and useSx](../helpers/styles.md) for container semantics and [SharePoint validation](../../development/sharepoint-validation/checklists/state-message.md#statemessage) for authenticated checks.

## Content slots

`title`, `description` and action descriptor `label` accept phrasing shorthand or fixed `span` props objects. `visual` accepts shorthand or a fixed `div` props object. These use the existing Fluent/React peers; no additional dependency is required. Shorthand retains its existing wrappers. StateMessage owns the heading, paragraph, live text region and Button/link; artwork and actions remain outside the live region.

```tsx
<StateMessage title={{ className: classes.title, children: '**Literal title**' }}
  description={{ children: 'Try **other filters**.' }}
  visual={{ children: <span role="img" aria-label="Folder">📁</span> }}
  primaryAction={{ label: { children: 'Retry' }, onClick: retry }} />
```

Static description strings/fragments keep existing inline Markdown, including object children. Title/action labels stay literal. `children: (Component, props) => <Component {...props}>…</Component>` inside an object supplies a Fluent render callback: output is opaque and must compose `FormattedText` explicitly for Markdown. Callbacks execute only during rendering, never for presence inspection, and may run more than once. Slot refs target their inner span/div; the public ref still targets the root div. Empty objects request containers. Booleans/legacy null suppress content; undefined visual selects preset artwork.

`as` can only repeat the fixed tag. Native attributes, classes, refs and consumer style are supported; raw HTML and unsupported native keys are discarded. Callers own valid phrasing/noninteractive title and label content. Arbitrary JSX/callback subtrees are not sanitized. React 17 erased `ReactNode` types can bypass literal-object type restrictions; runtime guards still protect toolkit containers.

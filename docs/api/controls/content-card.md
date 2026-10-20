# ContentCard

`ContentCard` groups a titled summary, arbitrary body content and explicit caller controls using Fluent `Card`, `CardHeader` and `CardFooter`. Import from the package root, `/controls` or `/controls/content-card`.

```tsx
import { Button } from '@fluentui/react-components';
import { ContentCard, FormattedText } from '@apvee/spfx-react-toolkit';

<ContentCard
  title={{ children: 'Workspace', className: classes.title }}
  headingLevel={2}
  description="Three people have access."
  headerActions={<Button appearance="subtle">Share</Button>}
  footer="Updated today"
  actions={<Button as="a" href="/workspace" appearance="primary">Open workspace</Button>}
>
  <FormattedText>Explicit **Markdown** composition.</FormattedText>
</ContentCard>
```

| Prop | Contract |
| --- | --- |
| `ContentCardProps.title` | Required meaningful phrasing content; shorthand, inner `span` props or children render callback. Do not nest headings. |
| `ContentCardProps.headingLevel` | `1 \| 2 \| 3 \| 4 \| 5 \| 6`; default `3`, owns the outer heading for all content forms. |
| `ContentCardProps.description` | Optional inner `div` content slot. |
| `ContentCardProps.icon` | Optional decorative inner `span` slot inside an aria-hidden header region. |
| `ContentCardProps.headerActions` | Optional inner `div` slot inside the header action region. |
| `ContentCardProps.children` | Caller-owned React body, including literal strings and arbitrary controls. |
| `ContentCardProps.footer` | Optional inner `div` slot for supplemental information. |
| `ContentCardProps.actions` | Optional inner `div` slot for footer controls, independent of `footer`. |
| `ContentCardProps.appearance` | Native `filled`, `filled-alternative`, `outline` or `subtle`; default `filled`. |
| `size` (native) | Native `small`, `medium` or `large`; default `medium`. |
| `ContentCardProps.dangerouslySetInnerHTML` | `never`; raw HTML is unsupported. |

Public types: `ContentCardProps`, `ContentCardAppearance`, `ContentCardHeadingLevel`.

All strings stay literal, including Markdown-looking content in props, object `children` and the body. Title/description strings retain the reference's two-line limit. JSX and render callbacks are opaque and own their presentation; wrap in `FormattedText` when parsing is desired. Existing `useSx`, typography and foreground descriptors supply text styles without importing the Markdown parser. Shared private frame semantics serve both card families; Griffel only supplies card geometry, clamping and media selectors under the host renderer.

Slot objects request inner containers, including empty objects. Their refs identify that inner element; the root ref identifies the native `HTMLDivElement`. Callbacks receive the fixed native tag and composed props during rendering and must forward those props. Shorthand JSX is rendered unchanged. Generic owned region shells survive an empty callback result. Raw HTML and alternative tags are unsupported. React 17 `ReactNode` variables may erase object type restrictions; runtime guards still omit raw HTML and unsupported tags. Explicit boolean, null and undefined titles are rejected by the title type. Runtime rejects meaningless static titles (including empty strings/objects/fragments) with TypeError without executing callbacks ahead of rendering. Callers remain responsible for meaningful callback content and accessibility. See the [slot convention](../../CONTROL-SLOTS.md).

Native Card props, events, focus modes, selection, checkbox and floating action pass through. Buttons keep their own appearance, size, ref and handlers. Root `onClick` can receive bubbled child events and does not promise Enter/Space activation. `disabled` affects the native Card without cloning children or changing their disabled flags. The installed Fluent Card adds an overlay that blocks pointer input to descendants even when their native disabled flag remains false; keyboard behavior also follows the native peer. Selectable Card's Enter behavior on child controls follows the installed Fluent peer; compare with native Card before assuming a separate activation contract. The toolkit does not add a route, implicit href, action descriptor, loading state or timestamp. Compose `AutoSkeleton` externally.

The card requires the existing shared React 17, Fluent and Griffel peers and host theme/renderer; no SPFx provider or new dependencies are required. Stable detail ancestry preserves caller controls across theme, direction and preview orientation changes. See [PreviewCard](./preview-card.md) for media.

The lazy CardsPanel selects one ContentCard or PreviewCard report in the same position, with shared title, description and retained notes. One outline configuration contains the relevant properties, simulation commands, compact status and collapsed diagnostics. Control-owned actions remain inside the card. Native comparison, literal/callback slots and independent/RTL regression cases remain in the versioned [card conformance fixture](../../../tests/fixtures/demo-control-isolation/cards-regressions.tsx), exercised alongside the product by source/compiled browser checks. Report local and authenticated host results separately.

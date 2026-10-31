# FormattedText

[Documentation](../../INDEX.md) > [API reference](../INDEX.md#react-controls) > FormattedText

## FormattedText

`FormattedText` renders a native HTML text element with the toolkit's typography and style helpers. It interprets a restricted inline Markdown subset and accepts mixed React content. It needs no SPFx provider and embeds no theme provider.

```tsx
import * as React from 'react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { FormattedText } from '@apvee/spfx-react-toolkit';

export const Example: React.FC = () => {
  const ref = React.useRef<HTMLElement>(null);
  return <FluentProvider theme={webLightTheme}>
    <FormattedText>Inline **strong** text and [a link](https://example.com).</FormattedText>
    <FormattedText as="h2" variant="title2">A semantic heading</FormattedText>
    <FormattedText as="p" subtle>A secondary paragraph</FormattedText>
    <FormattedText ref={ref} maxLines={2} style={{ width: 240 }}>
      A long inline source whose span receives block clamp layout without changing its HTML tag.
    </FormattedText>
  </FluentProvider>;
};
```

`as` defaults to `span` for every variant. `variant` chooses typography only; it never chooses a heading or changes the HTML tag. Without clamp, the element uses its natural inline or block display. Select appropriate heading levels and ensure authored JSX is valid inside the chosen element. The root remains present for empty content.

## FormattedTextProps

The component uses `React.forwardRef<HTMLElement, FormattedTextProps>`. Native HTML attributes, event handlers, `className`, `style`, ARIA and data attributes reach the root. Component-specific props do not reach the DOM. `children`, `color` and `dangerouslySetInnerHTML` are omitted from the native attribute base; `children` is redeclared as `React.ReactNode`, and `dangerouslySetInnerHTML` is forbidden and also removed at runtime. There are no `inline`, `size`, `weight` or `block` props.

| Prop | Type / accepted values | Default / behavior |
| --- | --- | --- |
| `FormattedTextProps.children` | `React.ReactNode` | Text and authored JSX; each string is interpreted independently |
| `FormattedTextProps.as` | `FormattedTextSemanticElement` | `span`, independent of variant |
| `FormattedTextProps.variant` | `FormattedTextVariant` | `body1` |
| `FormattedTextProps.font` | `base`, `monospace`, `numeric` | `base`; a font family override, not number formatting |
| `FormattedTextProps.align` | `start`, `center`, `end`, `justify` | `start`; start/end follow direction |
| `FormattedTextProps.markdown` | `boolean` | `true`; false returns the original children directly |
| `FormattedTextProps.wrap` | `boolean` | `true`; false enables nowrap without clipping |
| `FormattedTextProps.truncate` | `boolean` | `false`; single-line ellipsis, nowrap and hidden overflow |
| `FormattedTextProps.maxLines` | `number` | Unset; finite positive integers enable multiline clamp |
| `FormattedTextProps.subtle` | `boolean` | `false`; uses `foreground.subtle` (`colorNeutralForeground2`) |
| `FormattedTextProps.highlight` | `boolean` | `false`; wraps resolved content in a `mark` |
| `FormattedTextProps.italic` | `boolean` | `false` |
| `FormattedTextProps.underline` | `boolean` | `false`; composes with strikethrough |
| `FormattedTextProps.strikethrough` | `boolean` | `false`; composes with underline |
| `FormattedTextProps.dangerouslySetInnerHTML` | `never` | Forbidden; removed at runtime to keep children as the sole content pathway |

## FormattedTextVariant

This type is `keyof typeof typography`, reusing the existing [typography catalog](../helpers/styles.md). Its current 17 presets are `display`, `largeTitle`, `title1`, `title2`, `title3`, `subtitle1`, `subtitle2`, `subtitle2Stronger`, `body2`, `body1`, `body1Strong`, `body1Stronger`, `caption1`, `caption1Strong`, `caption1Stronger`, `caption2`, and `caption2Strong`.

## FormattedTextSemanticElement

`'span' | 'p' | 'h1' | 'h2' | 'h3' | 'h4' | 'h5' | 'h6'`. All accept `maxLines`. The consumer owns document semantics and authored JSX validity.

## Overflow and styling

Precedence is valid `maxLines` > `truncate` > `wrap`. A finite positive integer applies `display: -webkit-box`, vertical orientation, normal wrapping and hidden overflow while preserving the chosen HTML tag. Zero, negative, fractional, `NaN` and infinite values are ignored. Removing clamp restores the natural element layout unless consumer styles override it.

Ellipsis and clamp need constrained width and an appropriate layout. An unclamped span remains naturally inline; for visible single-line ellipsis, use `as="p"` with a bounded width or provide an appropriate display style. The full text stays in the DOM; the component adds no tooltip or duplicate accessible text. The line count is emitted through a dynamic Griffel class using the host renderer; the control generates no inline style. Consumer `style` is forwarded unchanged and can override `WebkitLineClamp`. Consumer `className` participates in the normal CSS cascade; argument order alone does not guarantee that arbitrary classes override generated styles.

```tsx
<FormattedText as="p" truncate style={{ width: 240 }}>Long single-line text...</FormattedText>
<FormattedText maxLines={2} style={{ width: 240, WebkitLineClamp: 3 }}>Long clamped text...</FormattedText>
```

Normal foreground inherits. Subtle is the existing `foreground.subtle` descriptor, with no extra color override. Inline code reuses `background.alternative` (`colorNeutralBackground2`), `paddingInline.extraSmall` and `borderRadius.small`, plus monospace and word breaking. Highlight reuses `foreground.primary`, the same XS padding and small radius, plus a yellow Fluent token and cloned box decoration. Highlighted Markdown links inherit the mark color; authored JSX remains untouched. Forced colors gives code Canvas/CanvasText with an outline and mark Highlight/HighlightText with forced color adjustment disabled.

## Markdown and mixed JSX

Each string supports one paragraph containing text, emphasis, strong, GFM strikethrough, inline code, hard line breaks, links and GFM autolinks. Soft line breaks follow normal HTML whitespace. Leading/trailing whitespace and whitespace-only strings are preserved as React text; visual whitespace follows CSS. Inline HTML is literal React text, never injected markup.

Unsupported syntax returns the entire original string literally: headings, multiple paragraphs, lists/task lists, block quotes, thematic breaks, fenced/indented code blocks, tables, images, reference links, and input mixing supported and excluded constructs.

Links use Fluent `Link` from `@fluentui/react-components`, preserving `href` and optional title without forcing a target. Destinations are checked after Remark decoding and trimming. HTTP, HTTPS, mailto, tel, relative/absolute paths, query strings and hashes are accepted. Empty destinations, protocol-relative `//`, backslashes, ASCII controls/whitespace inside a destination, invalid URLs and other schemes are rejected. Rejected links retain their descendant text/emphasis without an anchor. Consumer JSX is opaque and is not sanitized.

Arrays and `React.Fragment` children are visited; other React elements keep their identity, props, refs and children. Markdown delimiters never combine across separate string children or JSX elements.

When grouping Markdown strings with JSX actions, use `React.Fragment` (or `<>...</>`). A wrapper such as `<span>` is preserved together with all its descendants, so `markdown={true}` does not interpret strings inside it. This rule concerns child wrappers; the control's own `as="span"` root still interprets its direct string children normally.

```tsx
<FormattedText as="p">
  <React.Fragment>
    {'**Project update** '}
    <button type="button" onClick={() => console.log('clicked')}>Open project</button>
    {' *review available*'}
  </React.Fragment>
</FormattedText>
// The child span preserves its contents: the asterisks remain visible.
<FormattedText markdown={true}><span>{'**Literal inside JSX**'}</span></FormattedText>
<FormattedText markdown={false}>**Literal delimiters**</FormattedText>
```

The fragment example renders “Project update” in bold, keeps the authored button, and renders “review available” in italics. The sample's **Text and JSX action** mode uses this composition; turning Markdown off shows the delimiters while retaining the button.

Strings above 64 KiB of UTF-8 bytes fall back before parsing. More than 2,000 original AST nodes, depth above 32 (root depth 1), unexpected node structure or parser errors also trigger complete literal fallback. Node/depth limits are checked after parsing and do not bound the parser's own work. No global text cache is maintained.

## Dependencies, imports and verification

The host supplies compatible shared React 17, Griffel and Fluent packages, including the mandatory `@fluentui/react-components` peer `^9.74.9`. Fluent theme variables must be in CSS scope; `useSx` uses the existing renderer and direction. The library supplies runtime dependencies `unified@11.0.5`, `remark-parse@11.0.0` and `remark-gfm@4.0.1`. Use an ESM-aware bundler. `markdown={false}` skips parsing at runtime but does not promise to remove the static parser dependency from a bundle that imports the control.

The lazy **FormattedText** sample includes editable text, all presentation props, width and RTL controls, inherited host theme, mixed JSX, a ref inspection button and click count. Build the library before the app and restart serve after library edits. The marker test checks sample coverage; it does not execute browser layout. Local DOM/browser/package checks and authenticated SharePoint results are separate. The [authenticated checklist](../../development/sharepoint-validation/checklists/formatted-text.md#formattedtext) is **NOT EXECUTED** until a tenant run records evidence.

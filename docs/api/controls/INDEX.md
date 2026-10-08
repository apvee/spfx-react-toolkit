# Controls

Reusable React controls that do not require an SPFx provider. Theme variables and the shared Fluent/Griffel runtime still apply.

| Control / type | Purpose | Reference |
| --- | --- | --- |
| `FormattedText` | Native text element with typography, inline Markdown and overflow | [FormattedText](./formatted-text.md#formattedtext) |
| `FormattedTextProps` | Text, native attributes and presentation options | [Props](./formatted-text.md#formattedtextprops) |
| `FormattedTextVariant` | Keys of the existing typography catalog | [Variants](./formatted-text.md#formattedtextvariant) |
| `FormattedTextSemanticElement` | `span`, `p` or `h1` through `h6` | [Semantic element](./formatted-text.md#formattedtextsemanticelement) |

Import from the package root, `/controls`, or `/controls/formatted-text`. These entries share the canonical implementation; internal parser and styling modules are private. See [package imports](../../PACKAGE-IMPORTS.md).

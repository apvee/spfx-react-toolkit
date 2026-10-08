# Controls

Reusable React controls that do not require an SPFx provider. Theme variables and the shared Fluent/Griffel runtime still apply.

| Control / type | Purpose | Reference |
| --- | --- | --- |
| `AutoSkeleton` | Fluent loading items derived from mounted HTML geometry | [AutoSkeleton](./auto-skeleton.md#autoskeleton) |
| `AutoSkeletonProps` | Loading, animation, minimum size, exclusions and native div attributes | [Props](./auto-skeleton.md#autoskeletonprops) |
| `RelativeTime` | Localized relative instant with accessible full date and clock modes | [RelativeTime](./relative-time.md#relativetime) |
| `RelativeTimeProps` | Instant, locale, time zone, clock and native attributes | [Props](./relative-time.md#relativetimeprops) |
| `FormattedText` | Native text element with typography, inline Markdown and overflow | [FormattedText](./formatted-text.md#formattedtext) |
| `FormattedTextProps` | Text, native attributes and presentation options | [Props](./formatted-text.md#formattedtextprops) |
| `FormattedTextVariant` | Keys of the existing typography catalog | [Variants](./formatted-text.md#formattedtextvariant) |
| `FormattedTextSemanticElement` | `span`, `p` or `h1` through `h6` | [Semantic element](./formatted-text.md#formattedtextsemanticelement) |

Import from the package root, `/controls`, or the individual `/controls/auto-skeleton`, `/controls/formatted-text` and `/controls/relative-time` entries. These entries share the canonical implementation; internal measurement, parser and styling modules are private. See [package imports](../../PACKAGE-IMPORTS.md).

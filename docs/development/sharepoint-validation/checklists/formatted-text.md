# FormattedText

[Validation index](../INDEX.md) · [Setup](../setup.md)

Authenticated checks below are **NOT EXECUTED**. Local marker/DOM/browser/package checks do not establish SharePoint-host behavior. Build library before app and restart serve, then open the lazy **FormattedText** panel. See the [control reference](../../../api/controls/formatted-text.md).

| Scenario | Expected observation | Status |
| --- | --- | --- |
| Semantic elements and typography | Try all variants with span, p and headings. Changing variant leaves the chosen tag unchanged; inspect-ref reports the actual root. | NOT EXECUTED |
| Markdown and JSX | Edit emphasis, strong, strike, code, breaks, links and literal HTML. Unsupported blocks stay fully literal; unsafe links retain text without anchors. Authored JSX button works and the preview click count updates. | NOT EXECUTED |
| Overflow | Constrain width; compare wrapping, nowrap without clipping, truncate, and maxLines 1/2/4. Clamp wins and adds block layout on span without changing the tag. Remove clamp and verify natural display; use p for visible single-line ellipsis. | NOT EXECUTED |
| Themes and subtle | Compare subtle preview with the foreground.subtle reference in light/dark and site themes. Check highlighted links, code Background2, XS inline padding and small radii. | NOT EXECUTED |
| Accessibility and direction | Test actual keyboard link/button focus, RTL narrow width/long words, real browser 200% zoom and forced colors. Code uses Canvas/CanvasText; mark uses Highlight/HighlightText. | NOT EXECUTED |
| Lifecycle and isolation | Open/close/reopen the lazy panel and use two web parts. Preview controls, clicks and ref observations stay instance-scoped; record warnings/errors and regressions in existing panels. | NOT EXECUTED |

Record browser/version, actual theme/direction, computed styles, keyboard observations, screenshots, console output and outcomes. The control requires shared Fluent/Griffel peers and scoped theme variables; it introduces no SPFx provider or tenant permissions.

# Styles and useSx

[Validation index](../INDEX.md). These procedures require an authenticated host; their presence is not an executed result.


These authenticated host checks are **NOT EXECUTED**. The local React/Griffel DOM tests and standalone browser checkpoint do not constitute tenant evidence. Open the lazy **Styles** panel after rebuilding library/app and restarting serve. See [Styles and useSx](../../../api/helpers/styles.md) for exact mappings, query/state priority, theme and scrollbar limits.

| Check | Procedure and expected observation | Status |
| --- | --- | --- |
| Catalog and layout | Select every namespace/member, including overflow axes, spacing sides, typography, borders and decoration. Change width, gap and columns; the preview visibly follows the selected value. | NOT EXECUTED |
| Theme and roles | Switch sample light/dark/high contrast themes and site theme. Foreground subtle and background subtle/alternative remain distinct mapped roles; scoped variables change computed styles without replacing classes with resolved RGB values. Inspect presets on canvas and alternative surfaces. | NOT EXECUTED |
| Contrast limitations | Record inverted only on supported light/dark palettes. The floor high contrast inverted pair is black on black (1:1) and unsupported for active content; use canvas/alternative. Disabled preset is disabled-only. Evaluate any custom host palette/surface independently. | NOT EXECUTED |
| Independent query regions | Resize one of the two named ancestor regions across 480/640/1024px while holding viewport and the other region unchanged. Query content follows its nearest apvee-sx ancestor; its own container does not measure itself. | NOT EXECUTED |
| Viewport and direction | Resize window separately from the query regions. Verify explicit viewport scopes and logical layout in LTR/RTL with consistent host direction. If vertical writing mode is exercised, container inline-size measures height. | NOT EXECUTED |
| Pointer, focus and disabled | Focus with keyboard, hover and press. Overlapping properties follow focus-visible > active > hover. Change selected/preset and disable with pointer stationary; enabled state recipes are removed and native disabled semantics apply. Retain an observable focus indicator. | NOT EXECUTED |
| Scroll region | Scroll constrained content on both axes with overflow.auto and scrollbar.fluent. Check 6px thickness and 3px rounded thumb corners on both axes in Edge/Chrome, plus the native thin fallback elsewhere. Hover each actual thumb, press/drag it, release and leave; verify scoped Fluent accessible stroke base/hover/pressed colors, with pressed winning while hovered. Check the thumb against the actual underlying surface; overlay visibility varies by platform. Test browser forced-colors separately from the Fluent high contrast theme; native thin sizing/system colors and interactions remain available. | NOT EXECUTED |
| Two instances and lifecycle | Exercise theme/query/state controls independently in two WebParts. Remove/reinsert one and observe isolation and console. Renderer-owned CSS can remain after unmount; no CSS removal or bounded cache guarantee is implied. | NOT EXECUTED |

Record browser/version, theme, direction, actual computed styles, screenshots, console output and statuses. In development, a host using a nonempty Griffel salt may expose the documented upstream diagnostic; do not silently discard it. Separate unavailable cases from passing cases.

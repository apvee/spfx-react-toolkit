# AutoSkeleton

[Validation index](../INDEX.md) · [Setup](../setup.md)

Authenticated checks below are **NOT EXECUTED**. Local marker, DOM, browser and package checks do not establish SharePoint-host behavior. Build the library before the app and restart serve before opening the lazy **AutoSkeleton** panel. See the [control guide](../../../api/controls/auto-skeleton.md).

| Check | Expected observation | Status |
| --- | --- | --- |
| Geometry and shape | Toggle loading for the 160×24 rectangle, 40-pixel circle and wrapped/mixed text. Generated items use only circle/rectangle and align with the authored content. | NOT EXECUTED |
| Static placeholders | Select `none` while loading: uniform themed shapes have no animation or shimmer, including with reduced motion, forced colors and RTL. Resize while static, then select wave/pulse to restore animation. Note state, mount count and root busy state remain unchanged. | NOT EXECUTED |
| Resize and direction | Resize while loading and switch LTR/RTL; physical geometry remains aligned. Test root border, padding and scrolling in a host scenario. | NOT EXECUTED |
| Persistent content and native forwarding | Type a note, load, then finish loading. Value and mount count persist; inspect-ref observes div/busy/item count and loaded clicks update observations. | NOT EXECUTED |
| Focus and native inert | Focus note then load. Focus leaves owned content, tab skips it, and descendants declaring visibility:visible remain hidden. Outside focus is preserved; completion does not restore focus. Without native inert, content remains available with aria-busy. | NOT EXECUTED |
| Exclusions, minimum and empty content | Marked/selector-ignored subtrees produce no items; changing minimum size changes eligible geometry. Empty/cold content clears prior items and retains busy state without invented placeholders. | NOT EXECUTED |
| Theme and accessibility media | Compare light/dark themes, reduced motion and Windows forced colors with native Fluent SkeletonItem behavior. Overlay remains decorative without an extra progressbar/live announcement. | NOT EXECUTED |
| Lifecycle and isolation | Use two web parts, remove/reinsert one and inspect console. Controls/geometry stay instance-scoped and pending subscriptions are cleaned up. Renderer CSS can remain after unmount. | NOT EXECUTED |

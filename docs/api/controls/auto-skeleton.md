# AutoSkeleton

[Documentation](../../INDEX.md) > [API reference](../INDEX.md#react-controls) > AutoSkeleton

`AutoSkeleton` derives Fluent skeleton items from real geometry of mounted, bounded HTML content. Children stay mounted, preserving their state through loading. Text uses actual browser text fragments; visual leaves produce rectangles or circles. Every generated `SkeletonItem` explicitly uses `circle` or `rectangle`, including a rectangle whose width equals its height.

```tsx
import { AutoSkeleton } from '@apvee/spfx-react-toolkit';

<AutoSkeleton loading={isLoading} animation="wave">
  <section data-skeleton-container="">
    <img data-skeleton-shape="circle" className={classes.avatar} src={photoUrl} alt="" />
    <p>{profileName}</p>
    <div data-skeleton-ignore=""><ProfileActions /></div>
  </section>
</AutoSkeleton>
```

Import `AutoSkeleton` and `AutoSkeletonProps` from the root, `/controls`, or `/controls/auto-skeleton`. All paths share one implementation. The control needs no SPFx provider. Existing React 17, Fluent UI 9 and Griffel peers supply the runtime, theme variables, direction and host renderer; no scanner, storage, animation or accessibility dependency is added.

## AutoSkeletonProps

The props extend native `React.ComponentPropsWithoutRef<'div'>`, excluding owned children and direct HTML injection. The forwarded ref targets the root `HTMLDivElement`. Native events, `id`, `className`, consumer `style`, ARIA attributes and other native div attributes are forwarded. `id` is optional and has no cache meaning.

| Member | Type | Default and behavior |
| --- | --- | --- |
| `AutoSkeletonProps.children` | `React.ReactNode` | Required; rendered content to measure. It stays mounted while loading. |
| `AutoSkeletonProps.loading` | `boolean` | Required; controls root `aria-busy` and requests the decorative overlay when geometry and native inert are available. |
| `AutoSkeletonProps.animation` | `'wave' \| 'pulse' \| 'none'` | `'wave'`; animated modes use Fluent's animation and reduced-motion rules. `'none'` shows a uniform themed fill without animation or shimmer; loading semantics are unchanged. |
| `AutoSkeletonProps.minSize` | `number` | `8`; both normalized width and height must meet this minimum. Values must be finite and greater than zero; invalid values throw `RangeError`. |
| `AutoSkeletonProps.ignoreSelectors` | `readonly string[]` | Empty; each matching descendant and its entire subtree are excluded. Invalid CSS selectors throw `SyntaxError`, even when nothing matches. |
| `AutoSkeletonProps.dangerouslySetInnerHTML` | `never` | Unsupported; supply rendered children. Removed from untyped runtime input as well. |

There is no public measurement hook, shape prop, options object, cache, watch or debug API. Do not rely on the private scanner or styling modules.

## Static placeholders

Use `animation="none"` to show geometry-derived placeholders without the visual motion of loading. AutoSkeleton continues to render Fluent `Skeleton` and `SkeletonItem` components with a uniform themed fill, preserving their rounded shapes and forced-colors treatment. Omit `animation` to keep the default `wave`, or select `pulse` for Fluent's pulsing effect.

```tsx
import { AutoSkeleton } from '@apvee/spfx-react-toolkit';

<AutoSkeleton loading={showPlaceholder} animation="none">
  <ProfileCard />
</AutoSkeleton>
```

`animation` only changes presentation. `loading` still controls the overlay and root `aria-busy`; measurable content remains mounted, hidden and inert while the overlay is active. Setting `animation="none"` does not create a permanent empty-state mode or change focus handling. The same native inert and measurable-content requirements apply. Switching between `none`, `wave` and `pulse` preserves child state and does not affect other AutoSkeleton instances.

In the test web part, enable **Loading** and choose **none** under **Animation**. Resize the preview or change direction while static placeholders are visible, then choose `wave` or `pulse` to restore animation. Finish loading to check that the edited note and content mount count are preserved.

## Content markers

| Attribute on authored content | Effect |
| --- | --- |
| `data-skeleton-ignore=""` | Excludes the marked element and its complete subtree. |
| `data-skeleton-container=""` | Treats the element as a layout container: scans its children, but never creates a visual leaf for the empty container itself. |
| `data-skeleton-shape="circle"` | Forces one circular item from this element's box and stops scanning its descendants. |
| `data-skeleton-shape="rectangle"` | Forces one rectangular item from this element's box and stops scanning its descendants. |

The documented shape values are `circle` and `rectangle`. Neither `rect` nor `square` is an alias. A forced circle uses measured width and height; use a circular box in your authored layout. There is no artificial line-count marker: text follows actual `Range.getClientRects()` geometry. Author-hidden and excluded subtrees produce no items.

### One skeleton for an entire Card

Apply `data-skeleton-shape="rectangle"` to the Card to generate one item covering its measured box, including its approximated border radius. Its children stay mounted, but are not scanned for separate skeletons.

```tsx
import { Card, CardHeader, Text } from '@fluentui/react-components';
import { AutoSkeleton } from '@apvee/spfx-react-toolkit';

<AutoSkeleton loading={isLoading}>
  <Card data-skeleton-shape="rectangle">
    <CardHeader header={<Text>Profile</Text>} />
    <Text>Profile details</Text>
  </Card>
</AutoSkeleton>
```

### Containers and excluded content

Without a forced shape, elements containing child elements or nonblank text are scanned recursively by default. `data-skeleton-container` is usually unnecessary for these populated elements. It is useful for layout wrappers that may become empty: an empty marked wrapper produces no skeleton, even if CSS gives it dimensions or a background. When children are present, they are scanned normally.

```tsx
<AutoSkeleton loading={isLoading}>
  <div data-skeleton-container="" className={classes.layout}>
    {showSummary && <ProfileSummary />}
  </div>
</AutoSkeleton>
```

Use `data-skeleton-ignore` to exclude both an element and all its children from measurement. This differs from a container, whose children remain eligible. Ignored content is still hidden with the rest of the owned content wrapper during loading when an overlay is active.

Ignore markers and `ignoreSelectors` take precedence over forced shapes. Otherwise, a valid forced shape takes precedence over `data-skeleton-container` and stops descendant scanning. Put markers on HTML elements, or on components such as Fluent Card that forward them to their rendered HTML root; custom components must forward these attributes too.

### Marker syntax

`data-skeleton-ignore` and `data-skeleton-container` are presence markers: the scanner checks whether the attribute exists, not its value. In plain HTML, both `data-skeleton-ignore` and `data-skeleton-ignore=""` work. In React/TSX, use the explicit empty string shown in the examples; a bare JSX attribute passes `true` rather than an empty string. An attribute rendered with `"false"` still activates the marker because it is present. For a conditional marker, use `data-skeleton-ignore={shouldIgnore ? '' : undefined}` to omit it when inactive.

`data-skeleton-shape` requires the explicit string value `"circle"` or `"rectangle"`; presence alone does not force a shape.

## Wrapper and geometry

Text fragments are inset by 2px per side to separate adjacent skeleton bars, with a smaller inset for tiny fragments so they remain visible. `minSize` filters the measured fragment before this inset. Visual leaves and explicit shapes keep their measured dimensions; content layout is unchanged.

The root always contains an owned **block div wrapper**, with the decorative overlay as a sibling. This changes selector ancestry and means children no longer participate directly in root flex/grid layout. Put that layout on an element inside `children`:

```tsx
<AutoSkeleton loading={isLoading} className={classes.frame}>
  <div className={classes.profileGrid} data-skeleton-container="">
    <ProfileAvatar />
    <ProfileSummary />
  </div>
</AutoSkeleton>
```

Keep the root's positioning compatible with the overlay. Normalized pixel coordinates account for root borders and scroll position. Physical coordinates remain physical under RTL, while Fluent animation and presentation retain the host direction. Positive content radii are approximated as one conservative radius rather than independent corner geometry. When content has no radius, including text fragments, rectangular items retain Fluent's native rounded corners (4px in the installed Fluent version); a measured zero does not override them. Circular items use a 50% radius. No `square` items are generated.

The first measurement runs after mount in a layout effect; loading, children and configuration changes trigger measurement. Root resizing, content mutations, owner-window resizing, captured media load and supported font completion events coalesce into one animation frame. Observers stay active while loading, so hidden content can reflow and update the overlay. Owned loading opacity is ignored during measurement without revealing content temporarily. An empty scan clears previous geometry.

If observers are unavailable, initial and React-driven measurement still work; automatic detection is reduced. Geometry uses the node's own document and window. Importing on the server is safe, but server rendering has no measured overlay.

## Loading, focus and empty content

When loading, measurable content is hidden with opacity, marked `aria-hidden`, made natively `inert`, and given no pointer interactions. Opacity applies to the whole wrapper, so a descendant declaring `visibility: visible` cannot reveal itself. Focus already inside content is released; outside focus is preserved and completion does not restore focus automatically.

Native `HTMLElement.inert` support is required for this replacement behavior. When inert is unavailable, content stays visible and available with root `aria-busy=true`; no partial hiding or tabindex polyfill is applied. The decorative overlay has `aria-hidden=true`, `role="presentation"` and no independent busy state, progressbar or live announcement. Provide consumer status text where your flow requires an announcement.

Cold loading with absent, empty, author-hidden or otherwise unmeasurable content does not invent a placeholder or dimensions. It keeps `aria-busy` and leaves content available. Render representative content if you need a geometry-derived placeholder before data arrives. Exclusion markers exclude measurement; they do not exempt that content from hiding when the surrounding wrapper is loading.

## Performance and limits

Use bounded, nonvirtualized HTML. Portals, shadow DOM, pseudo-elements, detailed SVG shapes, scaled/rotated geometry, nested clipping/scrolling and continuously animated layouts are unsupported. A layout change outside the observed sources that only moves descendants may require a new render. The control does not observe every descendant's dimensions or poll for movement.

Geometry is local to the mounted instance; there is no shared cache or session storage. Equal normalized geometry avoids state updates, duplicate bones are removed, and Griffel deduplicates repeated CSS declarations. New geometry values can still insert rules that persist for the host renderer's lifetime; memory is not guaranteed constant during arbitrary resizing. Generated nodes use classes, without authored inline geometry styles; forwarded consumer `style` remains an escape hatch.

Fluent provides theme colors, forced-colors treatment and reduced-motion behavior. The control does not duplicate those colors or media rules; static mode disables motion and the shimmer gradient on the existing Fluent items. The lazy test-web-part panel shows one complete profile with a named Fluent Avatar, visible name and role, project description, editable note and metadata icons. A native Fluent Card surrounds AutoSkeleton and retains its themed surface and padding during loading. Its 40-pixel avatar and 160×24 title retain explicit circle/rectangle geometry hints. The same mounted profile exercises loading, wave/pulse/static presentation, minimum size, exclusions, persistent child state, empty content, resize, theme/RTL and native ref/focus behavior. Local checks and authenticated host validation are separate; see the [SharePoint checklist](../../development/sharepoint-validation/checklists/auto-skeleton.md#autoskeleton).

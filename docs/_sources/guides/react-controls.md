# Choosing and composing React controls

Compose controls around explicit state, native form ownership and host lifetimes. Choose presentation independently of persistence and verify real-host behavior separately.

[Documentation](../INDEX.md) > React controls

The toolkit's controls compose the host's React 17, Fluent UI 9 and Griffel runtime. Start with the control that owns the interaction you need, then keep application data, requests and persistence in the consumer. The [API reference](../../api/INDEX.md#references) is the complete export catalog; the references below explain behavior and configuration.

## Choose a control

Start with the interaction and its state owner. Use [forms](forms.md) for validation and persistence, [data grids](data-grids.md) for supplied rows and navigation, [editing](editing.md) for one saved HTML region, and [pickers](pickers.md) for discoverable references. For a specific symbol or its supported imports, use the [generated API catalog](../../api/INDEX.md).

## Host dependencies and imports

Controls are available from the package root, `/controls` and the documented family entry, such as `@apvee/spfx-react-toolkit/controls/content-card`. These aliases share one implementation. Compound parts belong to their family: use `Stage.Header`, `SettingsGroup.Item`, `AccordionGroup.Item` and `DescriptionList.Item`. Create `createForm` and `createDataGrid` suites outside rendering; each suite owns its methods/column markers and descriptors. See [package imports](package-imports.md) for published paths, declaration dependencies and bundle boundaries.

Ordinary presentation, Form, DataGrid, Stage and Pagination controls do not require an SPFx context provider. They still require their declared shared peers and the host Fluent theme/direction/Griffel renderer. Form additionally uses its RHF/compat peers; DataGrid includes its Charts peer. An individual family import does not guarantee that only one member of a compound family is retained by the bundler.

The seven SharePoint pickers, FileFolderPicker and all three access trims require the matching [SPFx provider](../../api/INDEX.md#references). `WebPartShell` installs its own web-part and Fluent providers for an actual web-part instance. Do not wrap the shell in a second toolkit provider merely to make its children use toolkit hooks. Native [Property Pane factories](../../api/INDEX.md#references) own separate provider/render roots and must be imported through their documented Property Pane paths.

Native UI localization is independent of the SPFx provider requirement. Form, DataGrid, Stage, Pagination, WebPartShell, EditableText, SharePoint pickers and FileFolderPicker statically load their own native resource modules. Register them in an SPFx solution with the [localization setup](localization.md). A non-SPFx host needs an explicit resource-module adapter even when passing a complete `strings` dictionary; overrides alone do not remove the static resource dependency. RelativeTime uses browser Intl locale/time-zone data instead.

## Ownership and composition

| Interaction | The control owns | The consumer owns |
| --- | --- | --- |
| Form submission | RHF validation, pending gate, field state, reset/draft coordination and error presentation | Initial values, validation rules/schema/resolver, service calls and record changes |
| DataGrid sorting and selection | Native navigation, request callbacks and cell presentation | Item order, selected-ID Set, hidden/sort state, filtering, fetching and pagination |
| Pagination | Available-action UI and a discriminated request | Page/cursor state, request cancellation, pending locks, rows and result announcements |
| Stage | Surface semantics, compound associations, native focus integration and dismissal requests | Explicit presentation, acceptance of controlled requests, save/discard policy and persistence |
| EditableText | Safe one-region Read/Edit UI and effective user-change notifications | Saved HTML, controlled echoes and external persistence |
| Picker selection | Identity/source verification, UI selection, hydration and documented caches | Controlled keys/JSON arrays, display seeds and eventual storage |
| Access trim | Current-user reads and mounting only the active structural branch | Protected operation authorization on the server and caller-owned branch content |

Place Form.Root inside a portaled Stage surface when its footer submits the form: React context across a portal does not make an outer native form own the portaled inputs. Await application persistence inside Root's `onSubmit` before resetting values or closing the surface. A controlled Stage rejects a close request by keeping `open`; preventing the callback event is not a Stage veto. See the complete [Stage/Form example](forms.md#form-composition).

DataGrid preserves supplied item order. Pair it with Pagination only after the consumer applies sorting/filtering and supplies the corresponding rows. Page changes are intent, so set `loading` while an asynchronous request is pending and commit the new result/state deliberately. Selection can retain off-page IDs; do not silently reconstruct it from the currently visible rows.

## Presentation and absence

Named presentation regions use the [control slot standard](control-slots.md): shorthand retains its DOM, while an object requests the documented inner tag and may supply native attributes, classes, refs or a Fluent children render callback. Semantic headings, triggers, IDs and accessible associations stay with the control. Keep noninteractive phrasing inside labels/headings/triggers and put explicit controls in the owning action region. Strings are literal unless the owning reference explicitly documents Markdown; compose FormattedText for parsing elsewhere.

Use undefined for new absence/default state. Each reference defines its own presence/default policy: an empty object can request a region, boolean/empty shorthand can suppress it, and some legacy nullish inputs retain defaults. A defined PreviewCard preview overrides image even when empty. Form's built-in date/time/file/SpinButton domains retain the null required by their native/compat APIs. Do not infer one universal empty-value rule across controls.

Use `useSx` descriptors or the host's Griffel classes for authored layout and dynamic styles. Existing consumer `style` props remain escape hatches. Slot refs target inner containers; a control's public root ref continues to target its documented root. Forward the callback-supplied props so composed refs, classes and attributes survive.

## Local examples and real-host validation

The private sample's generated source scenarios and remaining transitional control entries exercise actual components: ContentCard, PreviewCard, DescriptionList, Pagination, DataGrid, Form, Stage, SettingsGroup/AccordionGroup, StateMessage, AutoSkeleton, RelativeTime, EditableText, FormattedText, pickers and all three trims. Native Property Pane examples exercise array persistence and cascading entity sources. Versioned conformance fixtures preserve broader state, native-comparison and isolation cases separately from those product examples.

Build the library before the app because the sample imports compiled `lib`; restart a running serve process after library source changes. [Development verification](../development/verification.md) and [browser checks](../development/verification.md#local-browser-checks) distinguish static/Node evidence from rendered checks. Local tests, mock transports and builds do not establish authenticated SharePoint, Graph consent, native Apply/save/reopen, keyboard/screen-reader behavior or Microsoft's hosted picker traffic. Record those outcomes through the [SharePoint validation guides](../development/sharepoint-validation/INDEX.md).

## Stage portals and host validation

### Dependencies, portals and validation

Stage uses the existing shared Fluent components/icons and Griffel peers. The repository baseline is React/ReactDOM 17.0.1, TypeScript 5.3.3, Fluent 9.74.9 and SPFx 1.21.1, with Node `>=22.14.0 <23.0.0`; it introduces no dependency or peer upgrade. Inherit theme, direction, targetDocument and renderer from the enclosing Fluent provider, including an iframe's owner document. Stage creates no global renderer/provider and adds no arbitrary z-index.

Omitting mountNode retains Fluent's themed portal/virtual-parent behavior. An explicit mountNode is a caller-owned escape hatch: use the correct owner document and verify theme variables, direction, renderer styles, focus and nested portal relationships at that destination. Do not automatically move content to document.body.

The lazy Stage tab shows one configurable two-field editor in Dialog, Drawer or Popover. Presentation, size, title and description drive the same active surface; header action and Close stay aligned, while save/dismiss/retention guards remain observable. Configuration, open trigger, compact status and advanced diagnostics share one outline card. Modal-only manual completion remains inside the surface to preserve focus ownership. Nested popup, long-content and independent/RTL/native-ancestor regressions remain in the [versioned Stage fixture](../../../tests/fixtures/demo-control-isolation/stage-regressions.tsx). The product and conformance browser results, including any console gate failure, must be reported separately from authenticated host observations.

## Report cards

The generated PreviewCard scenario configures one report with editable title/description, orientation, standard/compact widths, loading/disabled and explicit image failure/recovery commands. ContentCard has its own working interaction/slot scenario. A compact PreviewCard-owned integration scenario switches between the two card types while retaining caller-owned title, description and notes; changing the component type replaces the card subtree. Each selected scenario has one outline configuration, compact status and collapsed diagnostics above its actual subject. Full slot/empty-media/native/isolation regressions remain in [versioned conformance fixtures](../../../tests/fixtures/demo-control-isolation/cards-regressions.tsx). Browser and authenticated host execution must be reported separately.

## Accordion and settings groups

`AccordionGroup` groups optional details inside a Fluent Card. Fluent Accordion owns expansion and native trigger behavior; native semantic wrappers and `useSx` provide toolkit typography/layout. Use [SettingsGroup](../../api/controls/settings-group.md) for preferences that should remain visible. No SPFx provider is required.

Import from the root, `/controls`, or `/controls/accordion-group`; these share one implementation. Compose `AccordionGroup.Item`; no flat Item alias is public.

The root forwards native div attributes/events, `className`, caller `style`, `dir`, `lang`, `hidden` and its HTMLDivElement ref. It introduces no implicit root name/group role and does not own form state or persistence. Fluent Card uses medium native exterior padding; internal row gap is zero.

Controlled values remain application-authoritative; the wrapper does not mirror them in separate state. Assign stable distinct Item values and React keys when rendering collections.

The native div ref/attributes target the Item root. Content belongs in its panel; the trigger is supplied by Fluent. Supply a nonempty title or description to name that button: an aria-label on the Item div does not label the trigger. Before programmatically closing a panel that contains focus, the application should move focus to an appropriate visible control.

### AccordionGroupHeadingLevel and AccordionGroupValue

`AccordionGroupHeadingLevel` accepts 1 through 6; root default 2 and Item default 3 should fit the surrounding document hierarchy. The root title heading wrapper applies to strings and JSX; the Item uses its own heading wrapper around the native Fluent trigger. `AccordionGroupValue` accepts string or number.

All text slot wrappers provide inherited defaults: group titles use subtitle2/primary, Item titles body1Strong/primary and descriptions caption1/subtle, with zero block margin and zero title-description gap. Disabled trigger text inherits Fluent disabled color. JSX children retain their props and explicit styling; they inherit the defaults when they do not specify their own.

Strings remain literal: `**bold**` stays `**bold**`; the group does not parse Markdown or depend on FormattedText. Use noninteractive phrasing JSX such as `<strong>Bold</strong>` or `<em>Emphasis</em>` for formatting. **Never place interactive JSX links, buttons, inputs or other interactive descendants in Item title, description or icon:** these regions render inside a button. Put actions in the group header's action or inside the panel. Nested headings and block markup inside text wrappers are unsupported. The component does not sanitize or rewrite arbitrary consumer JSX.

### Closed panels, layout and compatibility

The body Card supports only `filled`, `alternate` and `outline`, with `filled` as the default. `alternate` maps to Fluent Card's `filled-alternative` surface. `outline` uses Fluent's outlined surface and native border, without a filled background or shadow. Native Card owns padding, radius and child ancestry in every appearance; colors inherit the host theme. Changing appearance retains mounted panel controls and their state. The former `transparent` value is removed; use `outline` for an unfilled, visibly outlined Card. Replace former `filled-alternative` values with `alternate`.

The test web part exposes all three appearance values, requested versus rendered width, host theme inheritance and genuine Fluent SVG icons. Its direction-only provider is transparent and fills the bounded example container. Provider classes are generated for its effective direction, including RTL overrides.

By default closed panel children remain mounted, preserving uncontrolled inputs and component state. Closed content is hidden and excluded from keyboard focus through native inert compatibility for React 17. This requires a browser with native inert support; build success alone cannot prove focus behavior. Closing does not provide application persistence or a focus-restoration policy. `unmountOnClose` unmounts after Fluent closing motion and resets uncontrolled state after reopening; initially closed resettable content does not mount. Changing this policy while closed applies to the next opening/closing cycle. Retained child effects continue while closed; controlled values still belong to the application. Native hidden on the root/Item hides it without changing controlled expansion values.

The optional header sits 12px from the Card (Fluent medium spacing). Header icons have fixed 24px boxes and Item icons fixed 20px boxes; direct font icons inherit slot font size and SVG icons fill the slot without cloned child props. Complex custom icon content retains its own structure. Fluent Card medium provides the single 12px lateral inset, border and radius. The trigger has zero extra horizontal padding and the panel zero horizontal margin, so panel content and trigger text share the Card inset when no leading icon is present. Native trigger minimum height remains 44px; adjacent Items receive a thin separator with additional spacing equal to one Card-padding interval, split equally before and after it (6px + 6px at medium). Both halves derive from Fluent’s public `cardCSSVars.cardSizeVar` (`--fui-Card--size`), matching native Card padding; body gap stays zero. Divider spacing is measured between the native trigger/content row boxes; the 44px trigger minimum remains independent of the outer separator padding, so short and multiline titles retain native vertical centering. Fluent owns interaction and motion; the wrapper adds no animation timers or custom expansion state. Theme and direction use existing toolkit/Fluent/Griffel primitives and shared peer contracts; no Markdown parser or new dependency is required.

The lazy **SettingsGroup and AccordionGroup** sample selects one active group. Its AccordionGroup contains Profile details with a retained note and Advanced preferences with a notification switch; both triggers have Fluent icons. The sample and public control default to `filled`; `alternate` and `outline` remain selectable. Controlled expansion, multiple mode, disabled triggers, retained note state and unmount reset, inherited host theme/direction, width and root refs remain observable. Verify actual Tab behavior and state in a browser. Local and authenticated SharePoint results are distinct; see the [host checklist](../development/sharepoint-validation/INDEX.md#focus-input-and-accessibility).

### Named content slots

Titles, descriptions and icons accept React content or a native `span` props object; actions accept a native `div` props object. The action prop remains `action`. Props objects support attributes, classes, consumer styles, refs and `children: (Component, props) => JSX` callbacks. The callback receives the fixed inner tag and composed props during rendering; inspection never executes it. Forward those props to preserve its attributes and ref.

The headings, generated IDs, native trigger and panel association remain owned by AccordionGroup. A slot ref points to its inner container. Strings/JSX shorthand retain their existing DOM and literal formatting; props objects explicitly request an inner container even with empty children. Booleans and absent values suppress content; zero remains visible. Tags cannot replace owned headings or triggers; raw HTML is unsupported. Phrasing slots must contain appropriate noninteractive content (especially inside accordion triggers). The existing Fluent theme/host renderer is still required.

Description text inside an Item contributes to the native trigger content; the group adds no separate description association. Shorthand, including fragments, preserves the existing presence policy: empty text/fragments are absent, while whitespace text is present. Object/callback icon containers fill the owned shell; their direct SVG children retain 24px header/20px Item sizing, even with larger authored SVG dimensions. Callback output and custom components remain caller-owned and opaque.

## Geometry-derived loading

`AutoSkeleton` derives Fluent skeleton items from real geometry of mounted, bounded HTML content. Children stay mounted, preserving their state through loading. Text uses actual browser text fragments; visual leaves produce rectangles or circles. Every generated `SkeletonItem` explicitly uses `circle` or `rectangle`, including a rectangle whose width equals its height.

Import `AutoSkeleton` and `AutoSkeletonProps` from the root, `/controls`, or `/controls/auto-skeleton`. All paths share one implementation. The control needs no SPFx provider. Existing React 17, Fluent UI 9 and Griffel peers supply the runtime, theme variables, direction and host renderer; no scanner, storage, animation or accessibility dependency is added.

The props extend native `React.ComponentPropsWithoutRef<'div'>`, excluding owned children and direct HTML injection. The forwarded ref targets the root `HTMLDivElement`. Native events, `id`, `className`, consumer `style`, ARIA attributes and other native div attributes are forwarded. `id` is optional and has no cache meaning.

There is no public measurement hook, shape prop, options object, cache, watch or debug API. Do not rely on the private scanner or styling modules.

### Static placeholders

Use `animation="none"` to show geometry-derived placeholders without the visual motion of loading. AutoSkeleton continues to render Fluent `Skeleton` and `SkeletonItem` components with a uniform themed fill, preserving their rounded shapes and forced-colors treatment. Omit `animation` to keep the default `wave`, or select `pulse` for Fluent's pulsing effect.

`animation` only changes presentation. `loading` still controls the overlay and root `aria-busy`; measurable content remains mounted, hidden and inert while the overlay is active. Setting `animation="none"` does not create a permanent empty-state mode or change focus handling. The same native inert and measurable-content requirements apply. Switching between `none`, `wave` and `pulse` preserves child state and does not affect other AutoSkeleton instances.

In the test web part, enable **Loading** and choose **none** under **Animation**. Resize the preview or change direction while static placeholders are visible, then choose `wave` or `pulse` to restore animation. Finish loading to check that the edited note and content mount count are preserved.

### Content markers

The documented shape values are `circle` and `rectangle`. Neither `rect` nor `square` is an alias. A forced circle uses measured width and height; use a circular box in your authored layout. There is no artificial line-count marker: text follows actual `Range.getClientRects()` geometry. Author-hidden and excluded subtrees produce no items.

### One skeleton for an entire Card

Apply `data-skeleton-shape="rectangle"` to the Card to generate one item covering its measured box, including its approximated border radius. Its children stay mounted, but are not scanned for separate skeletons.

### Containers and excluded content

Without a forced shape, elements containing child elements or nonblank text are scanned recursively by default. `data-skeleton-container` is usually unnecessary for these populated elements. It is useful for layout wrappers that may become empty: an empty marked wrapper produces no skeleton, even if CSS gives it dimensions or a background. When children are present, they are scanned normally.

Use `data-skeleton-ignore` to exclude both an element and all its children from measurement. This differs from a container, whose children remain eligible. Ignored content is still hidden with the rest of the owned content wrapper during loading when an overlay is active.

Ignore markers and `ignoreSelectors` take precedence over forced shapes. Otherwise, a valid forced shape takes precedence over `data-skeleton-container` and stops descendant scanning. Put markers on HTML elements, or on components such as Fluent Card that forward them to their rendered HTML root; custom components must forward these attributes too.

### Marker syntax

`data-skeleton-ignore` and `data-skeleton-container` are presence markers: the scanner checks whether the attribute exists, not its value. In plain HTML, both `data-skeleton-ignore` and `data-skeleton-ignore=""` work. In React/TSX, use the explicit empty string shown in the examples; a bare JSX attribute passes `true` rather than an empty string. An attribute rendered with `"false"` still activates the marker because it is present. For a conditional marker, use `data-skeleton-ignore={shouldIgnore ? '' : undefined}` to omit it when inactive.

`data-skeleton-shape` requires the explicit string value `"circle"` or `"rectangle"`; presence alone does not force a shape.

### Wrapper and geometry

Text fragments are inset by 2px per side to separate adjacent skeleton bars, with a smaller inset for tiny fragments so they remain visible. `minSize` filters the measured fragment before this inset. Visual leaves and explicit shapes keep their measured dimensions; content layout is unchanged.

The root always contains an owned **block div wrapper**, with the decorative overlay as a sibling. This changes selector ancestry and means children no longer participate directly in root flex/grid layout. Put that layout on an element inside `children`:

### Compose markers around measured content

<!-- toolkit:include symbol="AutoSkeleton" section="example" title="Compose shape container and conditional ignore markers" -->

Keep the root's positioning compatible with the overlay. Normalized pixel coordinates account for root borders and scroll position. Physical coordinates remain physical under RTL, while Fluent animation and presentation retain the host direction. Positive content radii are approximated as one conservative radius rather than independent corner geometry. When content has no radius, including text fragments, rectangular items retain Fluent's native rounded corners (4px in the installed Fluent version); a measured zero does not override them. Circular items use a 50% radius. No `square` items are generated.

The first measurement runs after mount in a layout effect; loading, children and configuration changes trigger measurement. Root resizing, content mutations, owner-window resizing, captured media load and supported font completion events coalesce into one animation frame. Observers stay active while loading, so hidden content can reflow and update the overlay. Owned loading opacity is ignored during measurement without revealing content temporarily. An empty scan clears previous geometry.

If observers are unavailable, initial and React-driven measurement still work; automatic detection is reduced. Geometry uses the node's own document and window. Importing on the server is safe, but server rendering has no measured overlay.

### Loading, focus and empty content

When loading, measurable content is hidden with opacity, marked `aria-hidden`, made natively `inert`, and given no pointer interactions. Opacity applies to the whole wrapper, so a descendant declaring `visibility: visible` cannot reveal itself. Focus already inside content is released; outside focus is preserved and completion does not restore focus automatically.

Native `HTMLElement.inert` support is required for this replacement behavior. When inert is unavailable, content stays visible and available with root `aria-busy=true`; no partial hiding or tabindex polyfill is applied. The decorative overlay has `aria-hidden=true`, `role="presentation"` and no independent busy state, progressbar or live announcement. Provide consumer status text where your flow requires an announcement.

Cold loading with absent, empty, author-hidden or otherwise unmeasurable content does not invent a placeholder or dimensions. It keeps `aria-busy` and leaves content available. Render representative content if you need a geometry-derived placeholder before data arrives. Exclusion markers exclude measurement; they do not exempt that content from hiding when the surrounding wrapper is loading.

### Performance and limits

Use bounded, nonvirtualized HTML. Portals, shadow DOM, pseudo-elements, detailed SVG shapes, scaled/rotated geometry, nested clipping/scrolling and continuously animated layouts are unsupported. A layout change outside the observed sources that only moves descendants may require a new render. The control does not observe every descendant's dimensions or poll for movement.

Geometry is local to the mounted instance; there is no shared cache or session storage. Equal normalized geometry avoids state updates, duplicate bones are removed, and Griffel deduplicates repeated CSS declarations. New geometry values can still insert rules that persist for the host renderer's lifetime; memory is not guaranteed constant during arbitrary resizing. Generated nodes use classes, without authored inline geometry styles; forwarded consumer `style` remains an escape hatch.

Fluent provides theme colors, forced-colors treatment and reduced-motion behavior. The control does not duplicate those colors or media rules; static mode disables motion and the shimmer gradient on the existing Fluent items. The lazy test-web-part panel shows one complete profile with a named Fluent Avatar, visible name and role, project description, editable note and metadata icons. A native Fluent Card surrounds AutoSkeleton and retains its themed surface and padding during loading. Its 40-pixel avatar and 160×24 title retain explicit circle/rectangle geometry hints. The same mounted profile exercises loading, wave/pulse/static presentation, minimum size, exclusions, persistent child state, empty content, resize, theme/RTL and native ref/focus behavior. Local checks and authenticated host validation are separate; see the [SharePoint checklist](../development/sharepoint-validation/INDEX.md#focus-input-and-accessibility).

## Card actions and native interactions

`ContentCard` groups a titled summary, arbitrary body content and explicit caller controls using Fluent `Card`, `CardHeader` and `CardFooter`. Import from the package root, `/controls` or `/controls/content-card`.

All strings stay literal, including Markdown-looking content in props, object `children` and the body. Title/description strings retain the reference's two-line limit. JSX and render callbacks are opaque and own their presentation; wrap in `FormattedText` when parsing is desired. Existing `useSx`, typography and foreground descriptors supply text styles without importing the Markdown parser. Shared private frame semantics serve both card families; Griffel only supplies card geometry, clamping and media selectors under the host renderer.

Slot objects request inner containers, including empty objects. Their refs identify that inner element; the root ref identifies the native `HTMLDivElement`. Callbacks receive the fixed native tag and composed props during rendering and must forward those props. Shorthand JSX is rendered unchanged. Generic owned region shells survive an empty callback result. Raw HTML and alternative tags are unsupported. React 17 `ReactNode` variables may erase object type restrictions; runtime guards still omit raw HTML and unsupported tags. Explicit boolean, null and undefined titles are rejected by the title type. Runtime rejects meaningless static titles (including empty strings/objects/fragments) with TypeError without executing callbacks ahead of rendering. Callers remain responsible for meaningful callback content and accessibility. See the [slot convention](control-slots.md).

Native Card events, focus modes, selection, checkbox and floating action pass through. The toolkit fixes the root to a div and ContentCard to vertical orientation; native `as` and `orientation` props are excluded. Use PreviewCard's explicit `orientation` for a media layout. Buttons keep their own appearance, size, ref and handlers. Root `onClick` can receive bubbled child events and does not promise Enter/Space activation. `disabled` affects the native Card without cloning children or changing their disabled flags. The installed Fluent Card adds an overlay that blocks pointer input to descendants even when their native disabled flag remains false; keyboard behavior also follows the native peer. Selectable Card's Enter behavior on child controls follows the installed Fluent peer; compare with native Card before assuming a separate activation contract. The toolkit does not add a route, implicit href, action descriptor, loading state or timestamp. Compose `AutoSkeleton` externally.

The card requires the existing shared React 17, Fluent and Griffel peers and host theme/renderer; no SPFx provider or new dependencies are required. Stable detail ancestry preserves caller controls across theme, direction and preview orientation changes. See [PreviewCard](../../api/controls/preview-card.md) for media.

The source-owned ContentCard and PreviewCard scenarios each mount an independent real card with caller controls. A separate type-switch scenario keeps title, description and notes in its caller while replacing one card type with the other. Ordinary layout, loading, theme and direction changes retain native children within a card. Native comparison, literal/callback slots and independent/RTL regression cases remain in the versioned [card conformance fixture](../../../tests/fixtures/demo-control-isolation/cards-regressions.tsx). Report local and authenticated host results separately.

## Read-only metadata

`DescriptionList` presents read-only labeled details of one record using native `dl`, `dt` and `dd` elements. Import the single compound value from the package root, `/controls` or `/controls/description-list`. Its only part is `DescriptionList.Item`; there is no flat Item runtime export.

Use stable unique IDs for data arrays. For conditional pairs, native pair attributes or rich JSX, compose Items instead:

<!-- toolkit:include symbol="DescriptionList" section="example" title="Compose conditional metadata pairs with inner slot refs" -->

The Badge in this example comes from `@fluentui/react-components`. Use a form for editable fields and DataGrid to compare multiple records. Formatting, missing-value labels, loading, errors and empty-state messages belong to the consumer. Strings remain literal; use an explicit FormattedText value when markdown rendering is desired.

The native root is `dl`; its forwarded ref targets `HTMLDListElement`. Native attributes, events, `className`, `style`, language, direction, accessible naming and `hidden` are forwarded. `dangerouslySetInnerHTML` is excluded. Supply either `items` or compound `children`, never both.

### DescriptionListLayout

`DescriptionListLayout` is `'auto' | 'horizontal' | 'vertical'`. Explicit horizontal or vertical overrides container responsiveness. Auto uses CSS container queries rather than viewport width. Without a named container, auto remains vertical, even on a wide screen. Establish the ancestor with the existing `container.inlineSize` descriptor:

### Establish the metadata container

<!-- toolkit:include symbol="DescriptionList" section="example" title="Use the nearest named size container" -->

The 480px transition matches toolkit `responsive.small` and the source Blocks medium threshold. Nested lists use their nearest named container. Layout changes retain child DOM, state and focus. Horizontal layout shares two columns across pairs, with the label column limited to 40% of the list width. Long labels and values wrap without truncation. Logical spacing supports RTL.

### DescriptionListEntry

Zero is visible; empty strings, boolean conditions and absent values retain their definitions. Legacy Fluent/React slot values may include null, which also yields an empty definition; use undefined for new absence defaults. IDs, data formatting and meaningful labels are consumer responsibilities.

`DescriptionList.Item` requires a parent DescriptionList and inherits its layout. Rendering it outside a parent throws an actionable error. Native attributes/events and the forwarded `HTMLDivElement` ref target the grouping div. This div has no independent box in horizontal layout; customize the content slots for visible styling.

Native children and `dangerouslySetInnerHTML` are replaced/excluded by the slot contract. Consumer `style` escape hatches are forwarded. Repository-authored styling uses classes with the shared host renderer.

### Slots and accessibility

Scalar labels use semibold Fluent typography and scalar values regular typography. Rich JSX keeps its markup; paragraphs are not wrapped inside text spans. Fluent slot objects preserve allowed native content attributes, refs, consumer style, classes and render callbacks. Strings and JSX retain the existing inner div wrapper; undefined, booleans and legacy null leave the native cell empty. Explicit props objects request a div even when their static children are empty. `as` may only repeat `div`; it cannot select another tag. Slot refs identify the inner div, while root and Item refs continue to identify their native containers:

Use the complete [conditional-pair and inner-ref recipe](#compose-conditional-metadata-pairs-with-inner-slot-refs) above for the imported Badge, native ref and callback prop forwarding.

Slot types exclude `dangerouslySetInnerHTML`, and runtime resolution removes raw HTML and unsupported native keys and guards the fixed div tag, including for JavaScript callers. Callback children receive `(Component, props)` during normal rendering; they are never invoked to inspect presence. Keep the provided props on the rendered inner component to retain refs/classes/events. Arbitrary consumer JSX stays caller-owned and is not sanitized or cloned. React 17 can erase object restrictions through a `ReactNode` variable, so static checking alone cannot promise protection for erased values. Slots customize content inside invariant native term/definition cells. The list adds no keyboard behavior or focus stops; links and buttons retain their own native behavior. Supply an accessible name on the root when needed. Native hidden works on root and individual compound pairs in either layout. Fluent theme variables and normal forced-color adjustment preserve host colors; actual screen-reader and host behavior require separate validation.

### Dependencies and working sample

No SPFx provider or new dependency is required. The control uses the existing shared React 17, Fluent UI 9 and Griffel peers, inherits the host theme/direction/renderer and registers no PnP features. Auto layout requires browser support for CSS container queries; explicit layouts work without a named container.

The source-owned DescriptionList scenario shows one document detail with five entries. Layout, six named-container widths, order and hidden states modify the actual list; diagnostics inspect its native refs, computed columns/term width and root/item events. Stable IDs preserve counter state through reordering within either API. Switching data/compound composition or removing data can remount children; explicit reset restores the local baseline. Nested-container, pathological-value and independent/RTL comparisons remain in the [versioned regression fixture](../../../tests/fixtures/demo-control-isolation/description-list-regressions.tsx). Browser runners exercise the emitted source product route and then all preserved conformance assertions.

See [development checks](../development/verification.md#sample-and-component-checks) and [authenticated acceptance](../development/sharepoint-validation/INDEX.md#focus-input-and-accessibility). Authenticated SharePoint, Teams/iframe and screen-reader checks remain **NOT EXECUTED** until separately recorded. Static demo checks do not execute the rendered scenario or establish tenant behavior.

## Formatted text

### FormattedText

`FormattedText` renders a native HTML text element with the toolkit's typography and style helpers. It interprets a restricted inline Markdown subset and accepts mixed React content. It needs no SPFx provider and embeds no theme provider.

`as` defaults to `span` for every variant. `variant` chooses typography only; it never chooses a heading or changes the HTML tag. Without clamp, the element uses its natural inline or block display. Select appropriate heading levels and ensure authored JSX is valid inside the chosen element. The root remains present for empty content.

The component uses `React.forwardRef<HTMLElement, FormattedTextProps>`. Native HTML attributes, event handlers, `className`, `style`, ARIA and data attributes reach the root. Component-specific props do not reach the DOM. `children`, `color` and `dangerouslySetInnerHTML` are omitted from the native attribute base; `children` is redeclared as `React.ReactNode`, and `dangerouslySetInnerHTML` is forbidden and also removed at runtime. There are no `inline`, `size`, `weight` or `block` props.

### FormattedTextVariant

This type is `keyof typeof typography`, reusing the existing [typography catalog](../../api/helpers/styles.md). Its current 17 presets are `display`, `largeTitle`, `title1`, `title2`, `title3`, `subtitle1`, `subtitle2`, `subtitle2Stronger`, `body2`, `body1`, `body1Strong`, `body1Stronger`, `caption1`, `caption1Strong`, `caption1Stronger`, `caption2`, and `caption2Strong`.

### FormattedTextSemanticElement

`'span' | 'p' | 'h1' | 'h2' | 'h3' | 'h4' | 'h5' | 'h6'`. All accept `maxLines`. The consumer owns document semantics and authored JSX validity.

### Overflow and styling

Precedence is valid `maxLines` > `truncate` > `wrap`. A finite positive integer applies `display: -webkit-box`, vertical orientation, normal wrapping and hidden overflow while preserving the chosen HTML tag. Zero, negative, fractional, `NaN` and infinite values are ignored. Removing clamp restores the natural element layout unless consumer styles override it.

Ellipsis and clamp need constrained width and an appropriate layout. An unclamped span remains naturally inline; for visible single-line ellipsis, use `as="p"` with a bounded width or provide an appropriate display style. The full text stays in the DOM; the component adds no tooltip or duplicate accessible text. The line count is emitted through a dynamic Griffel class using the host renderer; the control generates no inline style. Consumer `style` is forwarded unchanged and can override `WebkitLineClamp`. Consumer `className` participates in the normal CSS cascade; argument order alone does not guarantee that arbitrary classes override generated styles.

Normal foreground inherits. Subtle is the existing `foreground.subtle` descriptor, with no extra color override. Inline code reuses `background.alternative` (`colorNeutralBackground2`), `paddingInline.extraSmall` and `borderRadius.small`, plus monospace and word breaking. Highlight reuses `foreground.primary`, the same XS padding and small radius, plus a yellow Fluent token and cloned box decoration. Highlighted Markdown links inherit the mark color; authored JSX remains untouched. Forced colors gives code Canvas/CanvasText with an outline and mark Highlight/HighlightText with forced color adjustment disabled.

### Markdown and mixed JSX

Each string supports one paragraph containing text, emphasis, strong, GFM strikethrough, inline code, hard line breaks, links and GFM autolinks. Soft line breaks follow normal HTML whitespace. Leading/trailing whitespace and whitespace-only strings are preserved as React text; visual whitespace follows CSS. Inline HTML is literal React text, never injected markup.

Unsupported syntax returns the entire original string literally: headings, multiple paragraphs, lists/task lists, block quotes, thematic breaks, fenced/indented code blocks, tables, images, reference links, and input mixing supported and excluded constructs.

Links use Fluent `Link` from `@fluentui/react-components`, preserving `href` and optional title without forcing a target. Destinations are checked after Remark decoding and trimming. HTTP, HTTPS, mailto, tel, relative/absolute paths, query strings and hashes are accepted. Empty destinations, protocol-relative `//`, backslashes, ASCII controls/whitespace inside a destination, invalid URLs and other schemes are rejected. Rejected links retain their descendant text/emphasis without an anchor. Consumer JSX is opaque and is not sanitized.

Arrays and `React.Fragment` children are visited; other React elements keep their identity, props, refs and children. Markdown delimiters never combine across separate string children or JSX elements.

When grouping Markdown strings with JSX actions, use `React.Fragment` (or `<>...</>`). A wrapper such as `<span>` is preserved together with all its descendants, so `markdown={true}` does not interpret strings inside it. This rule concerns child wrappers; the control's own `as="span"` root still interprets its direct string children normally.

The fragment example renders “Project update” in bold, keeps the authored button, and renders “review available” in italics. The sample's **Text and JSX action** mode uses this composition; turning Markdown off shows the delimiters while retaining the button.

Strings above 64 KiB of UTF-8 bytes fall back before parsing. More than 2,000 original AST nodes, depth above 32 (root depth 1), unexpected node structure or parser errors also trigger complete literal fallback. Node/depth limits are checked after parsing and do not bound the parser's own work. No global text cache is maintained.

### Dependencies, imports and verification

The host supplies compatible shared React 17, Griffel and Fluent packages, including the mandatory `@fluentui/react-components` peer `^9.74.9`. Fluent theme variables must be in CSS scope; `useSx` uses the existing renderer and direction. The library supplies runtime dependencies `unified@11.0.5`, `remark-parse@11.0.0` and `remark-gfm@4.0.1`. Use an ESM-aware bundler. `markdown={false}` skips parsing at runtime but does not promise to remove the static parser dependency from a bundle that imports the control.

The lazy **FormattedText** sample includes editable text, all presentation props, width and RTL controls, inherited host theme, mixed JSX, a ref inspection button and click count. Build the library before the app and restart serve after library edits. The marker test checks sample coverage; it does not execute browser layout. Local DOM/browser/package checks and authenticated SharePoint results are separate. The [authenticated checklist](../development/sharepoint-validation/INDEX.md#focus-input-and-accessibility) is **NOT EXECUTED** until a tenant run records evidence.

## Controlled pagination

`Pagination` renders accessible controlled result navigation. The application owns the supplied rows, page/cursor availability, requests and loading state. It does not fetch, slice results or retain pagination state. Import the same implementation from the package root, `/controls`, or `/controls/pagination`.

Supply the corresponding result rows separately, for example to a typed `createDataGrid` suite created outside rendering. For local arrays, slice using `pageIndex * pageSize`. For server data, resolve the request and replace controlled state only when the results arrive. Reset the application page/cursor and pending request when filters, page size or the dataset change. Announce the loaded page/item interval in a consumer-owned live region; the control does not know result contents.

The native root is `nav`, with an `HTMLElement` ref. Native attributes, events, `hidden`, classes and consumer `style` are forwarded; `children` and `dangerouslySetInnerHTML` are excluded. Root layout uses classes and the host renderer. Native `hidden` remains effective. A nonempty `aria-label` overrides the localized navigation name; native `aria-labelledby` retains normal browser precedence. Loading sets `aria-busy` on the navigation.

`appearance` supports `'primary' | 'secondary' | 'outline' | 'subtle' | 'transparent'`. Omit it to retain subtle navigation, or use `appearance="outline"` for bordered actions. Fluent owns theme, interaction and disabled colors; appearance does not change button size, request semantics or spinner placement.

### PaginationSize

`PaginationSize` follows native Fluent Button: `'small' | 'medium' | 'large'`, with medium as the default. Progress follows button density: 16px (`extra-tiny`) for small, 20px (`tiny`) for medium and 24px (`extra-small`) for large. Buttons use the requested native Fluent appearance, defaulting to subtle. The centered layout wraps within its container. Loading replaces the icon of the last activated action with a spinner in the same icon slot. There is no reserved gap between navigation actions. Load-more has no icon at rest; during loading, only the proportional spinner appears beside its visible text. Each instance remembers only its last clicked action for presentation; page state, results and loading remain controlled. Loading supplied before any action uses Next for pages/cursor and the Load more action for incremental mode. Every button remains disabled during loading, and the button containing the spinner retains its action accessible name. Logical navigation icons mirror with host RTL direction.

### PaginationState

`PaginationState` is `PaginationPagesKnownTotal | PaginationPagesUnknownTotal | PaginationCursor | PaginationLoadMore`. All modes are controlled. There are no numbered-page links, ellipsis, page-size selector or uncontrolled defaults.

### PaginationPagesKnownTotal

Page count is `Math.ceil(totalItems / pageSize)`. First requests page 0; last requests `Math.max(0, pageCount - 1)`. Previous requests `pageIndex - 1`; next requests `pageIndex + 1`. First/previous disable at page 0, next/last disable at or beyond the final page. An empty result at page 0 disables all four actions. A stale index beyond the total is deliberately retained: first can recover to 0, previous requests the received index minus one, and next/last remain disabled. No implicit clamping, state mutation or correction callback occurs; reset stale state in the application.

### PaginationPagesUnknownTotal

Only previous/next render. Previous disables at page 0; next disables when `hasNextPage` is false or the index is `Number.MAX_SAFE_INTEGER`, so no unsafe index is emitted. The application decides whether more results exist.

### PaginationCursor

Cursor tokens/history remain private to the application. Resolve the requested direction and then supply new capability flags and result rows.

### PaginationLoadMore

Append results in the application. At rest, the button shows text only. During loading, a proportional spinner appears in the icon slot, retaining the text; exhaustion disables it without hiding the existing results.

### PaginationChangeData

Load-more data contains only `{ type: 'loadMore' }`. Each available click emits one request. Navigation controls have `type="button"`, so an enclosing form is not submitted.

### PaginationStrings

SPFx selects the native Pagination resource from the host UI culture. Labels are trimmed; missing, undefined or whitespace-only overrides use that native baseline. The table shows the en-US resource values. The navigation contains one labeled group and native Fluent buttons. It does not move focus or announce fetched results automatically.

### Validation and dependencies

Malformed runtime states render an empty fragment. Numeric fields reject negatives, fractions, NaN, Infinity and unsafe integers; page size also rejects zero. Known total and next-page capability are mutually exclusive. Cursor/load-more availability must be boolean. Use undefined for absence.

No SPFx provider or new dependency is required. The control uses existing shared React 17, Fluent UI 9 Button/Spinner/tokens, Fluent Icons and Griffel peers. It inherits host theme, direction and renderer through the established styling integration. It does not create a renderer or register PnP features.

### Working sample and validation limits

The lazy Pagination tab shows one small result list and its navigation. Mode, page size, filter, size/appearance and explicit pending completion drive the current result. Configuration/status/diagnostics share one outline card. The separate enclosing-form/native/isolation comparisons remain in the [versioned regression fixture](../../../tests/fixtures/demo-control-isolation/pagination-regressions.tsx). Source/compiled runners exercise the product route before the preserved regression scenarios; local execution does not establish authenticated host behavior.

Authenticated SharePoint, Teams/iframe and screen-reader validation is **NOT EXECUTED** until separately recorded. Static demo checks do not establish rendered or tenant behavior.

### Native UI localization

Pagination requires the host to register its toolkit resource ID through the [localization configuration CLI](localization.md). SPFx selects the UI resource; no toolkit provider is required. Outside the SPFx loader, supply an explicit resource-module adapter: the native resource is a static dependency even when passing a complete dictionary.

`strings?: Partial<PaginationStrings>` overrides labels per instance. Overrides are trimmed, and omitted, undefined or blank labels use the selected resource. A complete nonblank `PaginationStrings` dictionary replaces every owned label. This does not translate result summaries or change controlled navigation callbacks.

## Relative instants

`RelativeTime` formats an explicit instant as localized elapsed time and exposes its full date through an accessible description. It renders a native `time` element and requires no SPFx provider. Its typography, color and direction inherit from the surrounding UI. Tooltip uses the existing Fluent UI 9 peer dependency and shared Fluent/Griffel host runtime; no extra date library is required.

Import `RelativeTime` and `RelativeTimeProps` from the package root, `/controls`, or `/controls/relative-time`. These paths share one implementation. Parsing and clock helpers are private.

The props extend native `React.ComponentPropsWithoutRef<'time'>`, excluding formatter-owned content and replacing the native string dateTime type with the required instant input. The forwarded ref targets `HTMLTimeElement`. Native events, `className`, consumer `style`, ARIA attributes and explicit `tabIndex` are forwarded.

Unsupported content props are also removed from untyped runtime input. The old `date` prop is not an alias. The native time `dateTime` attribute is normalized UTC ISO derived from the input. Invalid event/reference input renders fallback without a time root, ref target, full-date description or timer. Instant validation occurs before Intl construction, so invalid input still renders fallback if locale or timeZone is invalid. With valid instants, invalid Intl configuration propagates `RangeError`.

### Instant and unit rules

Date-only and timezone-less strings, impossible calendar dates, invalid offset/hour/minute/second fields, nonfinite epochs and values outside the native Date range are invalid. Fractional ISO seconds truncate to JavaScript millisecond precision. Valid signed six-digit extended ISO years are supported within that range.

Less than one minute in either direction renders the localized equivalent of “now” using zero seconds and `numeric: 'auto'`. Other labels use `numeric: 'always'` and completed minutes, hours, 24-hour days and seven-day weeks. Months and years use UTC anniversaries anchored to the original event, clamping to the end of the target month. January 31 to February 28 is one completed month; March 31 is two months. Leap-day anniversaries similarly clamp. Months are not approximated by thirty days; days are not rendered as yesterday/tomorrow. Daylight-saving changes affect the parsed instant but do not alter absolute elapsed hours.

### Clock modes and cleanup

A supplied valid `referenceTime` is controlled: it creates no timer or visibility listener. With `autoUpdate={false}`, the local sample stays frozen until the parsed event instant or clock mode changes, then resamples after commit. Equivalent instant inputs and presentation changes do not reset that sample. This mode does not promise synchronous sampling during render.

Live mode uses one timeout per instance at the next label transition, with a minimum delay of 1ms and maximum of one hour for clock resynchronization. It pauses while the Fluent host document is hidden and resynchronizes when visible. Input, document or mode changes and unmount clean up timers/listeners. When no document exists, it uses timers without visibility listeners. At Date extrema, unrepresentable future boundaries use the hourly cap safely rather than an overflowing tight loop.

### Full date and accessibility

The Fluent tooltip uses a description relationship and stable full-date ID. The root merges its ID with consumer `aria-describedby` tokens. Tooltip-enabled valid times receive `tabIndex={0}` unless the consumer overrides it; hover/focus, Escape/blur and ref/handlers compose through Fluent. With `showTooltip={false}`, a hidden description remains and no tab stop is added. `lang` defaults to the resolved locale unless overridden.

There is no automatic `aria-live` announcement. If access to exact dates is essential on touch devices, provide an explicit consumer full-date surface. Use `FormattedText` or Fluent `Badge` for presentation; the control adds no typography/appearance API or generated inline styling.

Requires browser `Intl.RelativeTimeFormat`, `Intl.DateTimeFormat` and appropriate locale/time-zone data. No polyfill is installed automatically. Full-date spelling, punctuation and relative abbreviations vary by runtime.

### SPFx locale composition

The [locale hook](../../api/hooks/environment.md#usespfxlocaleinfo) requires an SPFx provider; RelativeTime itself does not. `ianaTimeZone` identifies a representative standard SharePoint numeric zone group. When undefined, this example uses the runtime zone: it does not establish the site's zone. A locale such as `it-IT` does not imply `Europe/Rome`; verified group 4 maps to `Europe/Berlin`. Consumers should choose their unavailable-mapping UI or runtime fallback explicitly.

The sample's lazy RelativeTime panel exercises all props, three clock modes, invalid input, composition and native ref/events. ContextPanel separately compares fixed winter/summer instants only for a mapped site zone. Local tests and browser checks do not establish authenticated SharePoint-host behavior; see [SharePoint validation](../development/sharepoint-validation/INDEX.md#localization-and-instance-isolation).

### Invalid-input fallback slot

`fallback` accepts literal string/JSX shorthand or a fixed `span` props object using existing Fluent/Griffel peers. Shorthand gains no wrapper; objects supply native attributes, classes, a ref and static or Fluent callback children:

<!-- toolkit:include symbol="RelativeTime" section="example" title="Customize an invalid-date fallback without replacing the time ref" -->

Callbacks run only for invalid event/reference input during rendering, without presence preflight. The fallback ref targets span; the component ref remains reserved for the valid time root. Booleans, undefined and legacy null render no fallback; empty props still request a span. Strings stay literal. `as` can only repeat span, raw HTML and unsupported attributes are discarded, and JSX/callback subtrees stay caller-owned. React 17 erased `ReactNode` types can bypass literal-object restrictions. Valid time, clock, Tooltip and localization remain independent of fallback customization. No new dependency or SPFx provider is required.

## Settings presentation

`SettingsGroup` places related preferences in a Fluent Card with an optional header. It composes native semantic text, `useSx` typography and Fluent UI 9; no SPFx provider is required. Application controls own values, validation and persistence. Ordinary rows keep caller-owned labels; opt-in `field` rows inherit naming from the row for one built-in [Form](forms.md#settings-rows-as-form-fields) control. Choose [AccordionGroup](../../api/controls/accordion-group.md) for optional details users can reveal.

Import from the package root, `/controls`, or `/controls/settings-group`; these share the canonical implementation. `SettingsGroup.Item` is the row component; there is no separate public flat Item export.

Native div attributes/events, `className`, consumer `style`, `dir`, `lang`, `hidden` and the root HTMLDivElement ref are forwarded. The group adds no implicit accessible name, group role, form behavior or generated label relationships. `dangerouslySetInnerHTML` is unsupported because the component owns its content.

The row is a div with its own native attributes/events/ref. With `field` omitted or false, titles are visible text, **not implicit labels** for controls. Supply a Fluent control label, Field, or explicit application-owned label IDs. Rows do not clone children, intercept control events or own application state. Keyed rows retain child identity when reordered by the application.

### Automatic Form field rows

Set `SettingsGroup.Item field` inside a Form root and supply a meaningful `title`. The row renders its title and optional `description` once; the built-in Form adapter consumes them as its name and hint. Omit the adapter's `label` and `hint`: repeating either is rejected with a diagnostic. No manually authored `aria-label` or `aria-labelledby` is needed. Ordinary rows continue to support independent Fluent controls and custom content as above.

A native control uses a real label associated with its actual input, preserving explicit control IDs. Clicking the title focuses the input or activates a checkbox/switch through native label behavior. Composite adapters name their semantic group or control automatically; a group label does not promise to activate each inner control. DateTime retains distinct `dateLabel`/`timeLabel` names for its parts; TagPicker receives the row name for its input and tag group, including rich JSX titles. Names do not spread to secondary actions, nested settings rows or DateTime parts.

The row description, caller-supplied `aria-describedby` IDs and Fluent Field validation messages coexist. Validation output stays beside the control. The visible label reflects the adapter's `required` and effective disabled state, including disabled Form roots/scopes. `required` remains presentation metadata; add `rules`, a schema or a resolver to enforce it. Form continues to own values, errors, omission of disabled values, save and reset. Editing a settings row does not save automatically.

Use exactly one built-in logical adapter in each field row; DateTime counts as one. Multiple owning adapters are rejected. For several independently named controls, use separate field rows or omit `field` and label each control explicitly. `Form.useField` custom bindings and standalone Fluent controls do not consume this integration: give them application-owned labels/hints and use an ordinary row. This composition requires the existing Form peers and Fluent theme described in [Form dependencies](../getting-started/installation.md#form-adapter-baseline); it adds no dependency or SPFx provider requirement. See [scoped Fields examples](forms.md#settings-rows-as-form-fields).

### SettingsGroupHeadingLevel

Accepts 1 through 6. Choose the group title wrapper for its document hierarchy; it applies to both strings and JSX. Ordinary Item title/description wrappers are spans; field mode uses a label for a native labelable target. All title/description slots receive inherited defaults: group title `subtitle2`/primary, Item title `body1Strong`/primary and description `caption1`/subtle, with zero block margin and zero title-description gap.

Strings are rendered literally: `**bold**` stays `**bold**`; the group does not parse Markdown or depend on FormattedText. Use noninteractive phrasing JSX such as `<strong>Bold</strong>` or `<em>Emphasis</em>` for formatting. JSX children retain their props, identity and explicit styling; they inherit the wrapper defaults when they do not provide their own. Do not supply nested headings or block markup inside these text wrappers. Supply `undefined` for absent optional content.

### Layout, dependencies and limitations

The body Card supports only `filled`, `alternate` and `outline`, with `filled` as the default. `alternate` maps to Fluent Card's `filled-alternative` surface. `outline` uses Fluent's outlined surface and native border, without a filled background or shadow. Native Card owns padding, radius and child ancestry in every appearance; colors inherit the host theme. Changing appearance retains caller controls and their state. The former `transparent` value is removed; use `outline` for an unfilled, visibly outlined Card. Replace former `filled-alternative` values with `alternate`.

Fluent Card medium owns the 12px exterior padding, border and radius. The body sets its row gap to zero; the inner content row has a 44px minimum matching native Accordion triggers, independent of the outer divider padding. Adjacent direct Item DOM siblings receive a thin separator with one Card-padding interval of additional spacing, split equally before and after it (6px + 6px at medium). Both halves derive from Fluent’s public `cardCSSVars.cardSizeVar` (`--fui-Card--size`), the same variable as Card padding; arbitrary intervening content breaks adjacency. Consumer wrappers can therefore affect separators. The control preserves declaration order rather than collecting Items into a registry.

Each Settings Item uses a fluid wrapping flex row: text and controls remain side by side while they fit. Leading text grows and wraps around a 160px preferred/minimum width, capped at the available width for narrower groups. A compact Switch can remain beside text at 500px; a wider control wraps only when the combined regions and their 12px horizontal gap do not fit. Wrapped lines use an 8px vertical gap. Title and description remain adjacent with no extra gap. Custom leading content must support wrapping or shrinking within that bounded region.

The group establishes the toolkit `apvee-sx` inline-size container. Row placement follows actual available space and content widths; there is no blanket column layout below a fixed breakpoint. Existing layout/spacing descriptors use the host Griffel renderer and Fluent direction. The optional header sits 12px from the Card (Fluent medium spacing). Header icons have fixed 24px boxes and Item icons fixed 20px boxes; direct font icons inherit slot font size and SVG icons fill the slot. Consumer children keep their props and identity; complex custom icon content retains its own structure. Theme variables require a Fluent theme in scope; the package uses shared Fluent/Griffel peers without a Markdown parser dependency. Native `hidden` hides the group/row without unmounting controls; it is presentation, not authorization.

The lazy **SettingsGroup and AccordionGroup** sample selects one active group. Its SettingsGroup has separate Delivery address and Account alerts rows, each with a Fluent icon and one caller-labeled control, plus a Manage notifications header action. The sample and public control default to `filled`; `alternate` and `outline` remain selectable. Shared values, appearance, width and host direction remain configurable. The **Form** sample exercises automatic field rows with required Reference and Enabled, shared save/reset and observable RHF state. Automated marker checks establish sample coverage only. Local browser execution and authenticated SharePoint validation are separate; see the [host checklist](../development/sharepoint-validation/INDEX.md#focus-input-and-accessibility).

The group heading and field-mode Label, generated IDs and label/description associations remain owned by SettingsGroup. A slot ref points to its inner container. Strings/JSX shorthand retain their existing DOM and literal formatting; props objects explicitly request an inner container even with empty children. Booleans and absent values suppress content; zero remains visible. Tags cannot replace owned headings or Labels; raw HTML is unsupported. Phrasing slots must contain valid content for their heading, span or field-mode Label. The existing Fluent theme/host renderer is still required.

In field mode, a description props object with empty or whitespace-only static children adds no accessible description association unless the object supplies aria-label/aria-labelledby. Render callbacks request that association without probing their output. Shorthand, including fragments, preserves the existing presence policy: empty text/fragments are absent, while whitespace text is present. Object/callback icon containers fill the owned shell; their direct SVG children retain 24px header/20px Item sizing, even with larger authored SVG dimensions. Callback output and custom components remain caller-owned and opaque.

In `SettingsGroup.Item field`, title validation inspects static object children recursively, rejecting whitespace/empty intrinsic markup. Render callbacks and custom components must produce meaningful accessible label content. The outer Label and its `htmlFor`, required/disabled state and ID remain intact; the action slot stays outside the field bridge.

## Section states

`StateMessage` replaces the main content of a bounded section with a localized explanation, optional decorative artwork and actions. Use Fluent `MessageBar` for compact feedback beside ordinary content. It composes the toolkit's `FormattedText` and `useSx` with Fluent Button; no SPFx provider or additional dependency is required.

Import from the root, `/controls`, or `/controls/state-message`. The paths share the same implementation. Artwork, theme classification and styles are private.

Native div attributes/events, `className`, consumer `style`, `dir` and `lang` are forwarded. The ref targets the root HTMLDivElement. Content is owned by the control: `children` and `dangerouslySetInnerHTML` are unsupported, including raw HTML on actions.

### StateMessagePreset

Omitting the preset renders no automatic artwork or role. `role={undefined}` keeps the preset default; it does not disable its announcement semantics.

### StateMessageHeadingLevel and StateMessageRole

`StateMessageHeadingLevel` accepts 1 through 6; choose it for the surrounding heading hierarchy. `StateMessageRole` accepts `status` or `alert`. Title and description occupy that region; artwork and actions stay outside it. The control does not autofocus or restore focus. Role presence alone does not prove that a screen reader announces initial or updated content.

### StateMessageAction, StateMessageClickAction and StateMessageLinkAction

`StateMessageAction` is a union of a callback action and a navigation action. Both have a visible span content-slot `label`, optional Fluent `icon`, and compatible native/Fluent attributes such as disabled and className. Appearance, size, shape, root element and content belong to the control. Actions render in primary, secondary, tertiary order, with medium size and rounded shape.

`StateMessageClickAction` requires `onClick: React.MouseEventHandler<HTMLButtonElement>` and forbids href. Its native button has type=button, so it does not submit a containing form. `StateMessageLinkAction` requires `href: string` and accepts optional target, rel and `onClick: React.MouseEventHandler<HTMLAnchorElement>`; it uses Fluent Button as an anchor. Destinations are assigned as supplied by the application.

Title/description JSX must fit the heading/paragraph semantics; labels must not nest interactive controls. Caller-owned JSX is not rewritten.

### Artwork and theme

The five subjects keep their original light and dark geometry, viewBox and palette. SVGs are decorative (`aria-hidden=true`, `focusable=false`) and gradient IDs are unique per mounted instance. Their OfficeDev/fluent-blocks MIT attribution is included in the package's THIRD-PARTY-NOTICES.md.

The nearest FluentProvider supplies ThemeContext_unstable. A private pure resolver compares relative luminance of colorNeutralBackground1 and colorNeutralForeground1: darker background than text selects dark; otherwise light. It accepts hex and rgb/rgba token colors. Missing or uninterpretable tokens fall back to light. This is an **heuristic**, not a universal Fluent isDark flag; CSS variables, unusual custom palettes and transparency may not describe the actual rendered surface. No DOM reads, SPFx hook, subscription or public theme option is added.

Custom visual content keeps its own accessibility semantics; it is not automatically hidden from assistive technology. Built-in SVGs stay private, and are loaded synchronously with the control. Bundle size includes all preset variants and FormattedText's Markdown support; importing generic hooks does not pull this control in.

### Layout and responsive behavior

The root establishes the toolkit's named inline-size container. All authored layout uses useSx/descriptors and local Griffel classes, with no generated inline style or imported Layout control.

Medium means **640px of root content width**, excluding padding and border. It follows the toolkit's 480/640/1024px breakpoint vocabulary, rather than the source repository's names. The viewport does not decide presentation. A host with outer width 664px and default 12px padding per side has 640px available. Below the threshold, or without container-query support, base sizes remain active. The visual follows the BigMessage sizing approach: width fills the available space up to 280px, and SVG height follows its viewBox ratio. There is no square box or equal-area normalization. Wide subjects are naturally shorter; tall subjects take more vertical space, and the text/actions follow them without a reserved empty square. Custom SVGs should provide an intrinsic aspect ratio, usually through viewBox. Widths can shrink inside a narrow host; content has no fixed height. Custom CSS may change containment and sizing.

The root uses min-height semantics; vertical centering needs an available containing height. Long content can grow naturally. Responsive changes keep the same text/action DOM elements. See [styles and useSx](../../api/helpers/styles.md) for container semantics and [SharePoint validation](../development/sharepoint-validation/INDEX.md#focus-input-and-accessibility) for authenticated checks.

### Content slots

`title`, `description` and action descriptor `label` accept phrasing shorthand or fixed `span` props objects. `visual` accepts shorthand or a fixed `div` props object. These use the existing Fluent/React peers; no additional dependency is required. Shorthand retains its existing wrappers. StateMessage owns the heading, paragraph, live text region and Button/link; artwork and actions remain outside the live region.

Static description strings/fragments keep existing inline Markdown, including object children. Title/action labels stay literal. `children: (Component, props) => <Component {...props}>…</Component>` inside an object supplies a Fluent render callback: output is opaque and must compose `FormattedText` explicitly for Markdown. Callbacks execute only during rendering, never for presence inspection, and may run more than once. Slot refs target their inner span/div; the public ref still targets the root div. Empty objects request containers. Booleans/legacy null suppress content; undefined visual selects preset artwork.

`as` can only repeat the fixed tag. Native attributes, classes, refs and consumer style are supported; raw HTML and unsupported native keys are discarded. Callers own valid phrasing/noninteractive title and label content. Arbitrary JSX/callback subtrees are not sanitized. React 17 erased `ReactNode` types can bypass literal-object type restrictions; runtime guards still protect toolkit containers.

## Explore source-owned state and disclosure scenarios

The sample explorer exposes [SettingsGroup](../../api/controls/settings-group.md), [AccordionGroup](../../api/controls/accordion-group.md), [StateMessage](../../api/controls/state-message.md) and [Stage](../../api/controls/stage.md) as separate generated features. Their complete runnable examples live in the owning API JSDoc and use the existing Fluent theme and renderer; they do not require an SPFx provider or perform external persistence.

“Preserve notification settings across group presentations” keeps email, alert choices and action counts while its local Type control switches the actual group. Changing explorer features instead mounts the selected scenario with its own baseline. “Control disclosure selection and retained content” independently exercises AccordionGroup's native callback values, controlled selection, disabled trigger and retained versus unmounted note.

“Configure state artwork semantics and actions” demonstrates all artwork presets, custom/hidden visuals, Markdown shorthand versus opaque slot callbacks, semantic headings and explicit alert roles. “Edit a report with guarded Stage dismissal” keeps one real Form.Root inside each native portal. Dirty or pending saves refuse ordinary close requests. Its in-surface **Reset editor** command deliberately cancels/discards the local draft, closes the editor and waits for the real submission gate before restoring defaults; the next submission remains usable. The explorer's Reset command uses the same owned reset request, but native modal surfaces can make outside navigation unavailable while open. No simulated save writes to SharePoint.

## Stage presentation and migration

### StagePresentation, StageSize and StageDrawerPosition

`StagePresentation` is `'dialog' | 'drawer' | 'popover'`. `StageSize` is `'small' | 'medium' | 'large' | 'full'`, narrowed by StageProps for Popover. `StageDrawerPosition` is `'start' | 'end'`; end is the default logical edge and follows provider RTL direction.

Popover uses preferred widths of 320/480/640px, reduced by native Fluent auto-size when viewport space is smaller. The default remains `small`, now 320px. This replaces the previous Popover scale of 240/320/480px; existing small and medium Popovers become wider, and large increases to 640px. Dialog and Drawer presets retain their existing widths. Padding remains 12px per side for Popover. Select each size in the test web part and resize the viewport to observe the available-space constraint in LTR and RTL.

Dialog's equal outer gaps are `clamp(8px, 1vw, 16px)` on all sides. Measure full/narrow geometry against the usable CSS viewport: Fluent's native modal body-scroll lock applies `scrollbar-gutter: stable`, which can reserve a classic scrollbar gutter. Consequently CSS `100vw` can differ from `window.innerWidth`; do not treat that difference as a Stage sizing failure. The responsive gap still follows CSS `1vw`, bounded from 8px to 16px. Full fills both dimensions inside that gap; other Dialog sizes follow content height up to the available maximum. Drawer fills available height and reaches the selected viewport edge, including full size. Popover retains native anchoring, flip/shift and auto-size within viewport limits. Dynamic viewport height is used when supported with vh fallback. Modal surface padding is 24px; Popover padding is 12px. Header/Body/Footer add no duplicate padding.

### StageHeadingLevel and StagePopoverPositioning

Use `withArrow` with Popover to show Fluent's native arrow. Toggle the boolean while open; content and form state stay mounted. Native positioning keeps the arrow pointed toward the trigger, including flip/shift and RTL. The native surface keeps the arrow outside its border while an inner wrapper scrolls content on constrained viewports.

`StageHeadingLevel` accepts 1 through 6, default 2. The wrapper heading supplies Stage's accessible title association; a description is associated only while content is actually present.

Dialog/Drawer use native modal dialog semantics and focus management. Popover uses the native nonmodal group surface without a focus trap. Supply a named native `<button>` or Fluent `Button` trigger; arbitrary elements/custom wrappers are unsupported. Stage preserves trigger ref, disabled state, name and event handlers. It owns `aria-expanded` and `aria-controls`; Popover omits `aria-haspopup` because its surface is a group. Consumer activation cancellation prevents opening. Native focus initial placement/restoration is Fluent-owned; removed triggers are not a usable restore destination.

### StageDismissMode, StageOpenChangeReason, StageOpenChangeEvent and StageOpenChangeData

`StageDismissMode` is `'standard' | 'explicit'`. Standard allows Escape and applicable backdrop/outside/focus-outside interactions. Explicit filters those requests before notification, permitting only trigger and close-button requests; the Popover trigger still toggles closed. Header and Footer actions never dismiss automatically.

`StageOpenChangeReason` is `'triggerClick' | 'closeButton' | 'escapeKeyDown' | 'backdropClick' | 'outsideClick' | 'focusOutside'`. `StageOpenChangeEvent` is a React MouseEvent, KeyboardEvent or FocusEvent of HTMLElement, or a native MouseEvent, TouchEvent, FocusEvent or KeyboardEvent.

Uncontrolled requests update state and notify once; prop changes do not notify. Controlled consumers accept by updating `open`, or refuse dirty/pending dismissal by retaining it. Calling `event.preventDefault()` in onOpenChange is not a veto. Fluent can already prevent native Popover Escape before delivering its callback; Stage still accepts that request. An Escape consumed first by a nested Menu/Combobox/DatePicker must dismiss the child first. The installed legacy DatePicker compat input stops propagation and prevents Escape even when its calendar is already closed. Repeated Escape while focus remains on that closed input therefore does not close Stage. Use X or an application footer close action, or move focus to another Stage control before pressing Escape. Stage respects child consumption and does not rewrite Form or third-party keyboard handlers. Nested Fluent portals preserve native virtual-parent relationships for outside interactions.

Presentation is latched for each open session. A requested mode change applies next session, retaining the active adapter's compatible trigger and placement meanwhile. Size can change live; a pending full Dialog/Drawer leaves an active Popover at its last compatible size. Refused controlled closes keep the session intact. Do not depend on child-state retention across adapter changes.

With `unmountOnClose=true`, native presence/motion removes content on close. With false, content first mounts when opened, then remains hidden, aria-hidden and inert while closed; mounted effects remain active. Compatible size, theme and action rerenders preserve child state. Adapter changes can reset retained content.

<!-- toolkit:include symbol="PreviewCard" section="example" title="Report card with loading and notes" -->

<!-- toolkit:include symbol="AutoSkeleton" section="example" title="Loading a mounted profile" -->

<!-- toolkit:include symbol="DescriptionList" section="example" title="Basic usage" -->

<!-- toolkit:include symbol="RelativeTime" section="example" title="Basic usage" -->

<!-- toolkit:include symbol="Pagination" section="example" title="Request the next load-more batch" -->

<!-- toolkit:include symbol="Pagination" section="example" title="Browse controlled rows" -->

<!-- toolkit:include symbol="Pagination" section="example" title="Browse controlled rows with pending completion" -->

<!-- toolkit:include symbol="Pagination" section="example" title="Show an empty controlled result" -->

### Choose the site-zone fallback

<!-- toolkit:include symbol="RelativeTime" section="example" title="Separate SPFx UI culture from the site time zone" -->

### Own page state in the consumer

<!-- toolkit:include symbol="PaginationProps" section="example" -->

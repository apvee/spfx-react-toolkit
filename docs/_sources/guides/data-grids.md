# Own grid data and navigation

Create a typed column suite or render dynamic records, then own sorting, filtering, off-page selection and request state in the application.

[Documentation home](../INDEX.md)

## Data grids

### Typed columns and descriptors

Typed fields are compatible top-level property names. Accessors support computed or nested values and retain their value domains. Mapped values and person/tag/custom callbacks retain the field or accessor type. There is exactly one value source, `field` or synchronous `accessor`; Actions has neither.

`defineColumns` freezes a suite-owned descriptor collection. Use descriptors from the same suite. Hand-written descriptor objects, foreign-suite descriptors, arbitrary JSX children, duplicate or empty column IDs and malformed options are rejected with actionable configuration errors. Inputs and callback results must be immutable; accessors and renderers must be pure. Callback exceptions remain observable.

### JSX and dynamic data

Dynamic fields are strings and dynamic source values are `unknown`. Known-domain cells validate before presentation; generic callbacks must narrow unknown values explicitly. Dynamic mode is selected by `DataGrid`; a typed suite remains typed even when its model has an index signature.

Supply exactly one declared collection: JSX `children` or `columns`. Omitting both or supplying both is invalid. Nested arrays and fragments are supported; an explicitly empty collection renders the no-visible-columns status. JSX erases component type identity, so suite ownership is also validated at runtime.

### Named presentation slots

`error`, `emptyContent`, `noResultsContent`, column `header`, `placeholder` and `invalidPlaceholder` accept literal text, numbers, JSX, boolean conditions or a Fluent-style props object. Objects support native attributes, `className`, consumer `style`, events, `ref` and `children: (Component, nativeProps) => JSX`. This uses the existing Fluent/Griffel peers; no additional dependency is required.

Error and status object slots create an inner **div**. Header and placeholder object slots create an inner **span**: supply phrasing content, with noninteractive header content. `as` may only repeat that fixed tag. Refs and events belong to the inner container; the native MessageBarBody, status region, header/sort button and existing placeholder span remain owned by the grid. `headerAriaLabel` and `headerTooltip` continue to name and describe the native header. A callback runs during ordinary rendering, never to classify content or decide whether a region exists.

Shorthand keeps its existing wrappers and JSX identity; strings remain literal. Nullish status/placeholders use their existing localized defaults (or the em dash). Boolean and empty content suppress default text while retaining the owned wrapper; an empty object requests an explicit inner container. Actions headers default only when undefined. Error is omitted for undefined, legacy null and false; **true**, empty strings, empty fragments and zero still request the error bar. No visible columns retains precedence over loading/empty status, and errors can coexist with rows and loading. Retry, sorting, sizing, selection and row activation are unchanged.

Raw HTML and unsupported object props are rejected/filtered; React 17's erased `ReactNode` can defeat static object restrictions. Arbitrary JSX/callback subtrees remain the consumer's responsibility. Domain labels/icons/actions, value readers and `renderCell` callbacks retain their existing contracts and rendering paths. The source-owned scenarios expose inner refs, header callback clicks, object error/status and missing/invalid slots, and conditional suppression controls.

### Column inventory

Badge colors and appearance are configured per mapped value, for example `values={{ ready: { label: 'Ready', color: 'success', appearance: 'tint' } }}`. Appearance defaults to `tint`; set `filled`, `outline` or `ghost` explicitly to use another native Fluent presentation. Person text defaults to vertically centered next to the avatar; `personaProps.textAlignment` can override it. People overflow renders one native avatar per person with name, optional secondary text and presence. Its text-only Persona label lets AvatarGroupItem own the avatar without duplicating it.

People uses native `AvatarGroup layout="stack"`. The Tags `+N` button sizes to its text and native padding, without a fixed minimum width; native Overflow still measures and recovers hidden tags. Authored Boolean and action-menu icons use native 20px Regular variants, as do the sample's supplied icons. Consumer-provided icon elements keep their own presentation.

Rating displays only filled marigold stars. Its numeric value and optional vote count are included in the accessible description, without visible value/count text. Fluent's numeric slots require the documented external `null` shorthand to disable their automatic rendering. The default icon is `Star20Filled` with native `size="large"` (20px); `ratingDisplayProps` retains explicit consumer options, including icon, color and compact mode. An explicitly supplied `count` is independent of the maximum: zero is described, while `undefined` omits it. Localize its accessible vote label through the source-owned grid string contract.

`formattedCount` follows the column or grid locale; `count` remains the original number for pluralization. SPFx selects native UI labels independently of the numeric locale.

Common options include `id`, `header`, `headerAriaLabel`, `headerTooltip`, `tooltip`, `placeholder`, `invalidPlaceholder`, `sortable`, `width`, `minWidth` and `locale`. Numeric map keys require a `ReadonlyMap`, preserving numeric versus string identity. Tags require a mapping or paired label/identity callbacks. The callback pair supports scalar and object tags; object tags require the pair so identities are explicit. A matching mapping presentation takes precedence over the callback label. Scalar tags with a declared mapping retain their original scalar IDs. People and tag IDs must be unique within each cell; duplicates render invalid data because native overflow requires unique identities. Objects are never implicitly stringified. Custom content can compose other toolkit controls explicitly.

Progress accepts finite values from zero through `max` (default one). With `format={{ style: 'percent' }}`, the visible value uses `value / max`: `75` with `max={100}` displays `75%`. Rating is readonly and accepts zero through its maximum (default five).

### Order, identity and controlled selection

`getRowId` is required and returns a stable string or finite number. IDs are not derived from row indices. Duplicate IDs and React-key collisions such as `1` and `"1"` are rejected. Original items and ID types reach callbacks unchanged. Source reads are cached once per prepared pass, including undefined values and repeated accessor identities; the cache does not survive changed item/column inputs.

`sortState` and `onSortChange` report controlled sort intent. A sortable header changes the requested column/direction; rows stay in consumer-supplied order until the consumer provides reordered items. There is no default sort or local comparator API.

Use `selectionMode`, `selectedItems: ReadonlySet<string | number>` and `onSelectionChange` for controlled selection. Omitted selection is controlled empty. Multiselect/select-all changes the supplied rows while retaining off-page IDs. Callbacks receive a fresh Set; caller state is never mutated or pruned on refresh/removal. Single selection follows native Fluent semantics. `hiddenColumnIds` is controlled; stale hidden/sort IDs are ignored.

`onRowActivate` emits click/Enter intent with the original item and ID. It is mutually exclusive with selection. Prevented events and interactions in nested links, buttons, editors and widgets do not activate the row. Fluent owns cell/composite/group focus, keyboard navigation and resizing.

### Values, locale and dates

Missing external null/undefined values use `placeholder`; present values that fail the domain use `invalidPlaceholder` or the localized invalid-value label. Zero, false and empty strings remain valid. Numeric strings, NaN and Infinity are invalid numeric data. Bad cell data is distinct from invalid locale/time-zone/format or column configuration.

Function tooltips receive the original valid source, or `undefined` for missing data, including a legacy null source. They are skipped for invalid values so data outside the declared domain never reaches a typed callback; a static tooltip or the localized invalid-value description is used instead. Exceptions from invoked callbacks remain observable. Link `linkProps` describes an anchor; `href`, content and navigation are owned by the grid, while native consumer classes and styles are forwarded.

Column `locale` overrides the grid locale. Boolean text, informative icons, tooltips and accessible names use one resolved label. Calendar strings are validated, including leap days and years `0000`–`0099`, and displayed as their UTC calendar day without a browser-zone day shift, even when a display zone is configured. `Date` instances preserve their original instant and use the configured display zone. DateTime strings require four-digit years, uppercase `Z` or an explicit offset and at most three fractional digits; naive timestamps, signed years and epoch numbers are not accepted. Use `timeZone` for supported explicit display zones.

### Sparkline summaries and updates

The chart preserves source order and original values for summaries and color callbacks. Only plot coordinates are normalized. Empty, constant, negative, very large and extreme finite series are supported without mutating arrays or Date objects. First/last/minimum/maximum/average stay finite. Each derived metric is evaluated independently: an unrepresentable `change` or `changePercent` is undefined, and a zero first-value denominator has no percentage. For `[Number.MAX_VALUE, -Number.MAX_VALUE]`, change is undefined while changePercent is `-200`.

Immutable point updates change the actual chart; color, theme and RTL changes reach the mounted chart. A meaningful accessible summary remains when `showValue={false}`. The graphic is decorative and adds no competing tab stop or duplicate announcement. Chart failures are isolated locally, keeping the summary available and allowing recovery.

### Sizing, states and native ref

Use explicit `width`/`minWidth` where fixed layout constrains long action labels. Actions use intrinsic content; no label-length measurement heuristic or generated inline width variables are used. Native resizing owns header/body sizes. The scroll region contains narrow-width overflow. Toolkit-authored presentation uses `useSx` and host-renderer Griffel classes with Fluent tokens; consumer `style` remains a native escape hatch.

`loading` reuses one [AutoSkeleton](../../api/controls/auto-skeleton.md) around the native grid body. Each visible column receives a template appropriate to its kind: text/number/date value blocks, icon squares, people/tag groups, progress bars, rating items, chart and action shapes. These simulate value geometry inside cells rather than covering the whole cell. Their sizes vary deterministically by row and column identity, so rerenders do not cause jitter. Boolean display, visible-group limits, optional values and compact rating options influence the templates. Custom columns use a neutral value block because their renderer requires a real consumer item.

All template leaves request `shape="rectangle"` through AutoSkeleton's supported `data-skeleton-shape` marker and use the Fluent medium corner radius, clamped to their measured dimensions. AutoSkeleton continues to own measurement, overlay, native inert and accessibility behavior; no separate skeleton engine is added. The private live-content wrapper keeps original cell controls mounted while templates are measured, preserving their state across loading. There is no generated consumer data and no access to absent items or their callbacks.

Padding rows bring the loading body to at least four rows, with IDs checked against supplied IDs. Native selection and row activation are disabled while loading; controlled IDs are retained. No visible `Loading` message or loading progress bar is rendered; `aria-busy` signals the state. AutoSkeleton's documented browser requirements and fallback apply. Errors remain outside the overlay, and `onRetry` is consumer intent. Empty/no-results/all-hidden statuses are compact and localized; `emptyContent`/`noResultsContent` can reuse caller JSX such as StateMessage.

The `HTMLDivElement` ref, native attributes/events, `className` and consumer `style` target the actual Fluent grid. When a status replaces the grid, that native ref is unavailable until the grid mounts again. Owned content does not accept `dangerouslySetInnerHTML`.

### Empty and no-results messages

Default empty/no-results text is horizontally centered with Fluent medium vertical padding and natural height. `strings.empty` and `strings.noResults` localize it. These compact messages have no artwork or fixed height. The no-visible-columns status remains separate.

`emptyContent` and `noResultsContent` preserve caller-owned presentation instead of applying the default text styling. Compose [StateMessage](../../api/controls/state-message.md) when richer content is useful; default DataGrid imports do not include StateMessage or its artwork. Its actions remain consumer-owned.

### Scope and verification

There is no View, toolbar, search, filter, pagination, fetching or virtualization API. Supply data and orchestrate those behaviors in the consumer.

The explorer offers two createDataGrid scenarios and one DataGrid scenario. Choose typed document columns to switch five compact presets across all sixteen column families while preserving sort intent, selected IDs and supplied rows. Choose descriptors for one stable defineColumns collection with explicit sort application, or dynamic records to distinguish empty text, zero and false from missing and invalid values. Each scenario owns one grid, its consumer actions and its status/ref observations; switching scenarios starts that scenario's baseline. Broader column, chart and native-behavior regression cases remain in the versioned [DataGrid fixture](../../../tests/fixtures/demo-control-isolation/data-grid-regressions.tsx). Automated source/compiled browser runs and standalone tarball checks are local evidence. Authenticated SharePoint and screen-reader validation remain separate; see the [host checklist](../development/sharepoint-validation/INDEX.md).

Reset restores the current scenario without remounting Demo. The typed baseline returns to Documents: leaving Custom cells removes that child, and returning creates a new count. Width, theme and loading changes retain its mounted count while the custom preset stays present. Native grid refs detach when empty/no-results/all-hidden output replaces the grid; restoration reattaches a DIV. Header and missing-value refs apply to the typed and dynamic cases respectively; every case exposes its error DIV ref.

The examples require the host FluentProvider/theme, shared Griffel renderer and Fluent Charts peer. Their standalone fallback renders controls, commands, status, grid and diagnostics in that order; optional render maps those same regions to a host presentation, and optional resetRevision defaults to zero. Its first observation is not a reset.

The canonical {@link DataGrid} and {@link createDataGrid} reference owns every column, callback, type and default. The complete typed recipe below creates its suite once outside rendering.

<!-- toolkit:include symbol="createDataGrid" section="example" title="Create a typed descriptor grid" -->

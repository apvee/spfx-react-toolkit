# DataGrid

`DataGrid` and `createDataGrid<TItem>()` compose a table with native Fluent UI 9 navigation, focus, selection cells and column resizing. The consumer owns the items, their order and controlled state. No SPFx provider is required.

Import from the package root, `/controls` or `/controls/data-grid`. Create typed suites outside rendering. The implementation modules and column markers are private to each suite; columns are used through the suite rather than flat exports.

## Dependencies

The verified baseline uses React/ReactDOM 17.0.1, TypeScript 5.3.3 and SPFx 1.21.1, with components 9.74.9, Charts 9.3.27, shared contexts 9.26.4, utilities 9.26.7, theme 9.2.2, Griffel React 1.7.8/core 1.21.4 and migration 9.10.20. The shared packages, including `@fluentui/react-charts`, are mandatory host peers. A shared FluentProvider and renderer supply theme, direction and classes. See [installation](../../../README.md) and [development](../../DEVELOPMENT.md) for the full peer contract and compatible lockfile.

The complete grid family includes its cell renderers, reused AutoSkeleton and Charts/D3 even when a particular declaration uses only text. Root, domain and control aliases share the same implementation. There is no promise of per-column tree shaking; unrelated imports are verified to exclude the grid and charts. The optimized complete-family reference consumer measured 572,545 raw / 164,102 gzip / 128,306 Brotli bytes; these are fixture measurements with React/ReactDOM supplied by the host, rather than a size promise for every application. Root/domain/control aliases had zero size difference in that comparison.

## Typed columns and descriptors

```tsx
import * as React from 'react';
import { createDataGrid } from '@apvee/spfx-react-toolkit';

interface Order { id: string; title: string; total: number; paid: boolean; }
const Orders = createDataGrid<Order>();
const columns = Orders.defineColumns(column => [
  column.text({ id: 'title', header: 'Order', field: 'title', sortable: true }),
  column.number({ id: 'total', header: 'Total', field: 'total', format: { style: 'currency', currency: 'EUR' } }),
  column.boolean({ id: 'paid', header: 'Paid', field: 'paid', display: 'icon', trueLabel: 'Paid', falseLabel: 'Pending' }),
]);

export function OrderTable({ items }: { items: readonly Order[] }): React.ReactElement {
  return <Orders items={items} getRowId={item => item.id} columns={columns} locale="it-IT" aria-label="Orders" />;
}
```

Typed fields are compatible top-level property names. Accessors support computed or nested values and retain their value domains. Mapped values and person/tag/custom callbacks retain the field or accessor type. There is exactly one value source, `field` or synchronous `accessor`; Actions has neither.

`defineColumns` freezes a suite-owned descriptor collection. Use descriptors from the same suite. Hand-written descriptor objects, foreign-suite descriptors, arbitrary JSX children, duplicate or empty column IDs and malformed options are rejected with actionable configuration errors. Inputs and callback results must be immutable; accessors and renderers must be pure. Callback exceptions remain observable.

## JSX and dynamic data

```tsx
import { DataGrid } from '@apvee/spfx-react-toolkit';

const records: readonly Record<string, unknown>[] = [
  { id: 'one', title: 'First', count: 0 },
  { id: 'two', title: '', count: 'invalid numeric text' },
];

const table = <DataGrid items={records} getRowId={item => String(item.id)} aria-label="Dynamic records">
  <DataGrid.TextColumn id="title" header="Title" field="title" />
  <DataGrid.NumberColumn id="count" header="Count" field="count" invalidPlaceholder="Invalid count" />
</DataGrid>;
```

Dynamic fields are strings and dynamic source values are `unknown`. Known-domain cells validate before presentation; generic callbacks must narrow unknown values explicitly. Dynamic mode is selected by `DataGrid`; a typed suite remains typed even when its model has an index signature.

Supply exactly one declared collection: JSX `children` or `columns`. Omitting both or supplying both is invalid. Nested arrays and fragments are supported; an explicitly empty collection renders the no-visible-columns status. JSX erases component type identity, so suite ownership is also validated at runtime.

## Named presentation slots

`error`, `emptyContent`, `noResultsContent`, column `header`, `placeholder` and
`invalidPlaceholder` accept literal text, numbers, JSX, boolean conditions or a
Fluent-style props object. Objects support native attributes, `className`, consumer
`style`, events, `ref` and `children: (Component, nativeProps) => JSX`. This uses the
existing Fluent/Griffel peers; no additional dependency is required.

```tsx
const headerRef = React.createRef<HTMLSpanElement>();
const presentationColumns = Orders.defineColumns(c => [
  c.text({ id: 'title', field: 'title', headerAriaLabel: 'Order',
    header: { ref: headerRef, children: (Component, props) =>
      <Component {...props}>Order</Component> },
    placeholder: { children: 'Missing order', title: 'No source value' },
    invalidPlaceholder: { children: 'Invalid order' },
  }),
]);
const presented = <Orders items={[]} getRowId={item => item.id}
  columns={presentationColumns}
  error={{ children: <p>Request failed</p> }}
  emptyContent={{ children: 'No orders yet' }}
  noResultsContent={false} />;
```

Error and status object slots create an inner **div**. Header and placeholder object
slots create an inner **span**: supply phrasing content, with noninteractive header
content. `as` may only repeat that fixed tag. Refs and events belong to the inner
container; the native MessageBarBody, status region, header/sort button and existing
placeholder span remain owned by the grid. `headerAriaLabel` and `headerTooltip`
continue to name and describe the native header. A callback runs during ordinary
rendering, never to classify content or decide whether a region exists.

Shorthand keeps its existing wrappers and JSX identity; strings remain literal.
Nullish status/placeholders use their existing localized defaults (or the em dash).
Boolean and empty content suppress default text while retaining the owned wrapper;
an empty object requests an explicit inner container. Actions headers default only
when undefined. Error is omitted for undefined, legacy null and false; **true**,
empty strings, empty fragments and zero still request the error bar. No visible
columns retains precedence over loading/empty status, and errors can coexist with
rows and loading. Retry, sorting, sizing, selection and row activation are unchanged.

Raw HTML and unsupported object props are rejected/filtered; React 17's erased
`ReactNode` can defeat static object restrictions. Arbitrary JSX/callback subtrees
remain the consumer's responsibility. Domain labels/icons/actions, value readers and
`renderCell` callbacks retain their existing contracts and rendering paths.
The lazy sample exposes inner refs, header callback clicks, object error/status and
missing/invalid slots, and conditional suppression controls.

## Column inventory

| Suite member / builder | Input and presentation |
| --- | --- |
| `TextColumn` / `text` | String, including empty text; native text and truncation |
| `NumberColumn` / `number` | Finite number with `Intl.NumberFormat` options |
| `BooleanColumn` / `boolean` | Boolean, rendered as text or a centered native 20px icon with the same Yes/No label |
| `DateColumn` / `date` | `Date` or calendar `YYYY-MM-DD`; validated calendar day |
| `DateTimeColumn` / `dateTime` | `Date` or explicit zoned instant string |
| `BadgeColumn` / `badge` | Exact string/number value mapped to `DataGridBadgeValue` |
| `IconColumn` / `icon` | Exact string/number value mapped to `DataGridIconValue` and its accessible label |
| `LinkColumn` / `link` | Valid string with configured destination and original row identity in navigation callbacks |
| `PersonColumn` / `person` | Caller-provided name/avatar/presence presentation using native Persona |
| `PeopleColumn` / `people` | Immutable people array with stable person IDs and recoverable avatar overflow |
| `TagsColumn` / `tags` | Immutable scalar/object tags, configured and responsive recoverable overflow |
| `ProgressColumn` / `progress` | Finite value in the configured positive maximum; native ProgressBar |
| `RatingColumn` / `rating` | Read-only native RatingDisplay with configured maximum and optional count |
| `SparklineColumn` / `sparkline` | Immutable number series or `{ x: number | Date, y: number }` points; official Charts Sparkline |
| `ActionsColumn` / `actions` | Source-free inline/menu actions with labels, row IDs and disabled/busy states |
| `Column` / `custom` | Caller JSX with typed item, value and row ID; content identity is preserved |

Badge colors and appearance are configured per mapped value, for example `values={{ ready: { label: 'Ready', color: 'success', appearance: 'tint' } }}`. Appearance defaults to `tint`; set `filled`, `outline` or `ghost` explicitly to use another native Fluent presentation. Person text defaults to vertically centered next to the avatar; `personaProps.textAlignment` can override it. People overflow renders one native avatar per person with name, optional secondary text and presence. Its text-only Persona label lets AvatarGroupItem own the avatar without duplicating it.

People uses native `AvatarGroup layout="stack"`. The Tags `+N` button sizes to its text and native padding, without a fixed minimum width; native Overflow still measures and recovers hidden tags. Authored Boolean and action-menu icons use native 20px Regular variants, as do the sample's supplied icons. Consumer-provided icon elements keep their own presentation.

Rating displays only filled marigold stars. Its numeric value and optional vote count are included in the accessible description, without visible value/count text. Fluent's numeric slots require the documented external `null` shorthand to disable their automatic rendering. The default icon is `Star20Filled` with native `size="large"` (20px); `ratingDisplayProps` retains explicit consumer options, including icon, color and compact mode. An explicitly supplied `count` is independent of the maximum: zero is described, while `undefined` omits it. Localize its accessible vote label through grid strings:

```tsx
<Grid items={rows} getRowId={row => row.id} columns={columns}
  strings={{ ratingCount: (count, formattedCount) => `${formattedCount} ${count === 1 ? 'voto' : 'voti'}` }} />
```

`formattedCount` follows the column or grid locale; `count` remains the original number for pluralization. SPFx selects native UI labels independently of the numeric locale.

Common options include `id`, `header`, `headerAriaLabel`, `headerTooltip`, `tooltip`, `placeholder`, `invalidPlaceholder`, `sortable`, `width`, `minWidth` and `locale`. Numeric map keys require a `ReadonlyMap`, preserving numeric versus string identity. Tags require a mapping or paired label/identity callbacks. The callback pair supports scalar and object tags; object tags require the pair so identities are explicit. A matching mapping presentation takes precedence over the callback label. Scalar tags with a declared mapping retain their original scalar IDs. People and tag IDs must be unique within each cell; duplicates render invalid data because native overflow requires unique identities. Objects are never implicitly stringified. Custom content can compose other toolkit controls explicitly.

Progress accepts finite values from zero through `max` (default one). With `format={{ style: 'percent' }}`, the visible value uses `value / max`: `75` with `max={100}` displays `75%`. Rating is readonly and accepts zero through its maximum (default five).

## Order, identity and controlled selection

`getRowId` is required and returns a stable string or finite number. IDs are not derived from row indices. Duplicate IDs and React-key collisions such as `1` and `"1"` are rejected. Original items and ID types reach callbacks unchanged. Source reads are cached once per prepared pass, including undefined values and repeated accessor identities; the cache does not survive changed item/column inputs.

`sortState` and `onSortChange` report controlled sort intent. A sortable header changes the requested column/direction; rows stay in consumer-supplied order until the consumer provides reordered items. There is no default sort or local comparator API.

Use `selectionMode`, `selectedItems: ReadonlySet<string | number>` and `onSelectionChange` for controlled selection. Omitted selection is controlled empty. Multiselect/select-all changes the supplied rows while retaining off-page IDs. Callbacks receive a fresh Set; caller state is never mutated or pruned on refresh/removal. Single selection follows native Fluent semantics. `hiddenColumnIds` is controlled; stale hidden/sort IDs are ignored.

`onRowActivate` emits click/Enter intent with the original item and ID. It is mutually exclusive with selection. Prevented events and interactions in nested links, buttons, editors and widgets do not activate the row. Fluent owns cell/composite/group focus, keyboard navigation and resizing.

## Values, locale and dates

Missing external null/undefined values use `placeholder`; present values that fail the domain use `invalidPlaceholder` or the localized invalid-value label. Zero, false and empty strings remain valid. Numeric strings, NaN and Infinity are invalid numeric data. Bad cell data is distinct from invalid locale/time-zone/format or column configuration.

Function tooltips receive the original valid source, or `undefined` for missing data, including a legacy null source. They are skipped for invalid values so data outside the declared domain never reaches a typed callback; a static tooltip or the localized invalid-value description is used instead. Exceptions from invoked callbacks remain observable. Link `linkProps` describes an anchor; `href`, content and navigation are owned by the grid, while native consumer classes and styles are forwarded.

Column `locale` overrides the grid locale. Boolean text, informative icons, tooltips and accessible names use one resolved label. Calendar strings are validated, including leap days and years `0000`–`0099`, and displayed as their UTC calendar day without a browser-zone day shift, even when a display zone is configured. `Date` instances preserve their original instant and use the configured display zone. DateTime strings require four-digit years, uppercase `Z` or an explicit offset and at most three fractional digits; naive timestamps, signed years and epoch numbers are not accepted. Use `timeZone` for supported explicit display zones.

## Sparkline summaries and updates

The chart preserves source order and original values for summaries and color callbacks. Only plot coordinates are normalized. Empty, constant, negative, very large and extreme finite series are supported without mutating arrays or Date objects. First/last/minimum/maximum/average stay finite. Each derived metric is evaluated independently: an unrepresentable `change` or `changePercent` is undefined, and a zero first-value denominator has no percentage. For `[Number.MAX_VALUE, -Number.MAX_VALUE]`, change is undefined while changePercent is `-200`.

Immutable point updates change the actual chart; color, theme and RTL changes reach the mounted chart. A meaningful accessible summary remains when `showValue={false}`. The graphic is decorative and adds no competing tab stop or duplicate announcement. Chart failures are isolated locally, keeping the summary available and allowing recovery.

## Sizing, states and native ref

Use explicit `width`/`minWidth` where fixed layout constrains long action labels. Actions use intrinsic content; no label-length measurement heuristic or generated inline width variables are used. Native resizing owns header/body sizes. The scroll region contains narrow-width overflow. Toolkit-authored presentation uses `useSx` and host-renderer Griffel classes with Fluent tokens; consumer `style` remains a native escape hatch.

`loading` reuses one [AutoSkeleton](auto-skeleton.md) around the native grid body. Each visible column receives a template appropriate to its kind: text/number/date value blocks, icon squares, people/tag groups, progress bars, rating items, chart and action shapes. These simulate value geometry inside cells rather than covering the whole cell. Their sizes vary deterministically by row and column identity, so rerenders do not cause jitter. Boolean display, visible-group limits, optional values and compact rating options influence the templates. Custom columns use a neutral value block because their renderer requires a real consumer item.

All template leaves request `shape="rectangle"` through AutoSkeleton's supported `data-skeleton-shape` marker and use the Fluent medium corner radius, clamped to their measured dimensions. AutoSkeleton continues to own measurement, overlay, native inert and accessibility behavior; no separate skeleton engine is added. The private live-content wrapper keeps original cell controls mounted while templates are measured, preserving their state across loading. There is no generated consumer data and no access to absent items or their callbacks.

Padding rows bring the loading body to at least four rows, with IDs checked against supplied IDs. Native selection and row activation are disabled while loading; controlled IDs are retained. No visible `Loading` message or loading progress bar is rendered; `aria-busy` signals the state. AutoSkeleton's documented browser requirements and fallback apply. Errors remain outside the overlay, and `onRetry` is consumer intent. Empty/no-results/all-hidden statuses are compact and localized; `emptyContent`/`noResultsContent` can reuse caller JSX such as StateMessage.

The `HTMLDivElement` ref, native attributes/events, `className` and consumer `style` target the actual Fluent grid. When a status replaces the grid, that native ref is unavailable until the grid mounts again. Owned content does not accept `dangerouslySetInnerHTML`.

### Empty and no-results messages

Default empty/no-results text is horizontally centered with Fluent medium vertical padding and natural height. `strings.empty` and `strings.noResults` localize it. These compact messages have no artwork or fixed height. The no-visible-columns status remains separate.

`emptyContent` and `noResultsContent` preserve caller-owned presentation instead of applying the default text styling. Compose [StateMessage](state-message.md) when richer content is useful; default DataGrid imports do not include StateMessage or its artwork. Its actions remain consumer-owned:

```tsx
<Grid items={rows} getRowId={row => row.id} columns={columns}
  emptyContent={<StateMessage title="Nessun elemento disponibile" />}
  noResultsContent={<StateMessage title="Nessun risultato"
    description="Prova a modificare i criteri impostati." />} />
```

## Scope and verification

There is no View, toolbar, search, filter, pagination, fetching or virtualization API. Supply data and orchestrate those behaviors in the consumer.

The lazy DataGrid test-web-part scenario shows one compact grid with selectable typed, descriptor or dynamic APIs and column presets. Its controls update the active grid's rows, sizing, selection and status. Broader column, chart and native-behavior regression cases remain in the versioned [DataGrid fixture](../../../tests/fixtures/demo-control-isolation/data-grid-regressions.tsx). Automated source/compiled browser runs and standalone tarball checks are local evidence. Authenticated SharePoint and screen-reader validation remain separate; see the [host checklist](../../SHAREPOINT-VALIDATION.md).

## Public API catalog

The qualified names below correspond to source IntelliSense documentation. Native Fluent props retain their native contracts; toolkit-owned members are listed explicitly.

### `DataGrid`

Native controlled grid for dictionary rows. Values remain unknown until classified. Requires immutable items, stable getRowId and one same-suite column collection. Native div ref is unavailable while a status replaces the grid. Fluent supplies navigation, resize, selection and cell widgets; no SPFx provider is required.

### `createDataGrid`

Create an isolated typed grid and sixteen inferred JSX/builder column kinds. Call once outside rendering. Descriptor and JSX columns must belong to this suite; inputs and callbacks are immutable and pure. The native div ref follows Fluent's grid and detaches while an empty/no-visible-column status is shown. Sorting reports controlled intent without changing supplied item order.

### `DataGridProps`

Controlled grid configuration. Input order is preserved.

| Member | Contract |
| --- | --- |
| `DataGridProps.items` | Immutable source rows; objects are passed to callbacks unchanged. |
| `DataGridProps.onClick` | Native div click event. |
| `DataGridProps.getRowId` | Stable finite number or string ID. Never inferred from position. |
| `DataGridProps.sortState` | Controlled sort intent; rows are not sorted internally. |
| `DataGridProps.onSortChange` | Native controlled sort intent callback. |
| `DataGridProps.selectionMode` | Controlled selection mode. Mutually exclusive with row activation. |
| `DataGridProps.selectedItems` | Controlled IDs, retaining off-supplied selections. |
| `DataGridProps.onSelectionChange` | Receives a fresh native selection Set. |
| `DataGridProps.hiddenColumnIds` | Controlled hidden IDs; stale IDs are ignored. |
| `DataGridProps.columnSizingOptions` | Native resizing options keyed by declared column IDs. Descriptor width and minWidth override matching native entries; other native options are preserved. Active when resizableColumns is enabled. |
| `DataGridProps.containerWidthOffset` | Native offset applied to measured resize container width. |
| `DataGridProps.locale` | Root Intl locale fallback. |
| `DataGridProps.strings` | Overrides selected localized strings. |
| `DataGridProps.loading` | AutoSkeleton over type-aware rounded templates for every visible column; at least four rows, stable varied geometry, real cell state preserved, no visible label/progress indicator. Selection/activation disabled until completion. |
| `DataGridProps.error` | Consumer-owned error content. |
| `DataGridProps.onRetry` | Retry intent without fetching. |
| `DataGridProps.emptyState` | Defaults to `empty`; choose `noResults` for the alternate message. Both use compact, horizontally centered default text with native Fluent padding and natural height. |
| `DataGridProps.emptyContent` | Custom empty input content retaining its own presentation; can compose StateMessage. |
| `DataGridProps.noResultsContent` | Custom no-results content retaining its own presentation; can compose StateMessage. |
| `DataGridProps.onRowActivate` | Original row activation from unhandled click or Enter. |

### `DataGridStrings`

Localized status and cell strings.

| Member | Contract |
| --- | --- |
| `DataGridStrings.empty` | Empty input status. |
| `DataGridStrings.noResults` | No matches status. |
| `DataGridStrings.noVisibleColumns` | No visible columns status. |
| `DataGridStrings.loading` | Retained loading label for source compatibility; loading is represented by skeletons and aria-busy. |
| `DataGridStrings.actions` | Default actions heading. |
| `DataGridStrings.moreActions` | Overflow actions trigger. |
| `DataGridStrings.yes` | True label shared by text, icon and tooltip. |
| `DataGridStrings.no` | False label shared by text, icon and tooltip. |
| `DataGridStrings.noTags` | Empty tag list label. |
| `DataGridStrings.tagOverflow` | Accessible trigger label for hidden tags. |
| `DataGridStrings.ratingCount` | Accessible localized vote count when count is explicitly supplied, receiving its original number and locale-formatted text. Defaults to the native resource vote-count forms. |
| `DataGridStrings.noObservations` | Empty chart series label. |
| `DataGridStrings.invalidValue` | Invalid present cell data label. |
| `DataGridStrings.selectAllRows` | Select all supplied rows. |
| `DataGridStrings.selectRow` | Accessible selection label for an original row ID. |
| `DataGridStrings.sparklineSummary` | Accessible chart summary from locale-formatted original observations. Defaults to host-selected native metric labels. Derived metrics independently use an em dash when unavailable; showValue=false does not suppress this summary. |
| `DataGridStrings.retry` | Retry label. |

### `DataGridSparklineSummaryText`

Locale-formatted chart metrics supplied to the summary string callback.

| Member | Contract |
| --- | --- |
| `DataGridSparklineSummaryText.first` | First observation. |
| `DataGridSparklineSummaryText.last` | Latest observation. |
| `DataGridSparklineSummaryText.minimum` | Minimum observation. |
| `DataGridSparklineSummaryText.maximum` | Maximum observation. |
| `DataGridSparklineSummaryText.average` | Overflow-safe mean. |
| `DataGridSparklineSummaryText.change` | Formatted change or placeholder when independently unrepresentable. |
| `DataGridSparklineSummaryText.changePercent` | Formatted percent change or placeholder for zero denominator/unrepresentable result. |

### `DataGridColumnCollection`

Exactly one declared column collection, including an explicitly empty collection.

| Member | Contract |
| --- | --- |
| `DataGridColumnCollection.columns` | Excluded when JSX is declared. |
| `DataGridColumnCollection.children` | Same-suite JSX markers, including an explicitly empty collection. |

### `DataGridColumnMarkers`

Same-suite JSX column markers; runtime additionally validates erased JSX ownership.

| Member | Contract |
| --- | --- |
| `DataGridColumnMarkers.TextColumn` | Declares a text column with callbacks inferred from its source. |
| `DataGridColumnMarkers.NumberColumn` | Declares a number column with callbacks inferred from its source. |
| `DataGridColumnMarkers.BooleanColumn` | Declares a boolean column with callbacks inferred from its source. |
| `DataGridColumnMarkers.DateColumn` | Declares a date column with callbacks inferred from its source. |
| `DataGridColumnMarkers.DateTimeColumn` | Declares a dateTime column with callbacks inferred from its source. |
| `DataGridColumnMarkers.BadgeColumn` | Declares a badge column with callbacks inferred from its source. |
| `DataGridColumnMarkers.IconColumn` | Declares a icon column with callbacks inferred from its source. |
| `DataGridColumnMarkers.LinkColumn` | Declares a link column with callbacks inferred from its source. |
| `DataGridColumnMarkers.PersonColumn` | Declares a person column with callbacks inferred from its source. |
| `DataGridColumnMarkers.PeopleColumn` | Declares a people column with callbacks inferred from its source. |
| `DataGridColumnMarkers.TagsColumn` | Declares a tags column with callbacks inferred from its source. |
| `DataGridColumnMarkers.ProgressColumn` | Declares a progress column with callbacks inferred from its source. |
| `DataGridColumnMarkers.RatingColumn` | Declares a rating column with callbacks inferred from its source. |
| `DataGridColumnMarkers.SparklineColumn` | Declares a sparkline column with callbacks inferred from its source. |
| `DataGridColumnMarkers.Column` | Declares a custom column with callbacks inferred from its source. |
| `DataGridColumnMarkers.ActionsColumn` | Declares source-free row actions. |

### `DataGridSuite`

Typed grid and its suite-owned marker/builder family. Construct outside rendering.

| Member | Contract |
| --- | --- |
| `DataGridSuite.defineColumns` | Freezes descriptors from this suite's builder. |

### `DynamicDataGridColumnMarkers`

Explicit dynamic JSX markers; generic source callbacks receive unknown.

| Member | Contract |
| --- | --- |
| `DynamicDataGridColumnMarkers.TextColumn` | Declares a guarded dynamic text column. |
| `DynamicDataGridColumnMarkers.NumberColumn` | Declares a guarded dynamic number column. |
| `DynamicDataGridColumnMarkers.BooleanColumn` | Declares a guarded dynamic boolean column. |
| `DynamicDataGridColumnMarkers.DateColumn` | Declares a guarded dynamic date column. |
| `DynamicDataGridColumnMarkers.DateTimeColumn` | Declares a guarded dynamic dateTime column. |
| `DynamicDataGridColumnMarkers.BadgeColumn` | Declares a guarded dynamic badge column. |
| `DynamicDataGridColumnMarkers.IconColumn` | Declares a guarded dynamic icon column. |
| `DynamicDataGridColumnMarkers.LinkColumn` | Declares a guarded dynamic link column. |
| `DynamicDataGridColumnMarkers.PersonColumn` | Declares a guarded dynamic person column. |
| `DynamicDataGridColumnMarkers.PeopleColumn` | Declares a guarded dynamic people column. |
| `DynamicDataGridColumnMarkers.TagsColumn` | Declares a guarded dynamic tags column. |
| `DynamicDataGridColumnMarkers.ProgressColumn` | Declares a guarded dynamic progress column. |
| `DynamicDataGridColumnMarkers.RatingColumn` | Declares a guarded dynamic rating column. |
| `DynamicDataGridColumnMarkers.SparklineColumn` | Declares a guarded dynamic sparkline column. |
| `DynamicDataGridColumnMarkers.Column` | Declares a guarded dynamic custom column. |
| `DynamicDataGridColumnMarkers.ActionsColumn` | Declares source-free dynamic row actions. |

### `DynamicDataGridSuite`

Dynamic grid for dictionary rows. Typed factories remain typed even for index signatures.

| Member | Contract |
| --- | --- |
| `DynamicDataGridSuite.defineColumns` | Freezes dynamic descriptors belonging to this family. |

### `DataGridColumnKind`

One supported cell presentation.

### `DataGridLocale`

Intl locale or ordered fallback locales.

### `DataGridRowId`

Stable original row identity. Numbers must be finite.

### `DataGridValueContext`

Cell callback context. External nullable data is supported at the source boundary.

| Member | Contract |
| --- | --- |
| `DataGridValueContext.item` | Original row object. |
| `DataGridValueContext.value` | Original source value; legacy nullable source fields remain nullable. |

### `DataGridValueSource`

Exactly one synchronous source. Nullable fields/accessors interoperate with legacy row models.

| Member | Contract |
| --- | --- |
| `DataGridValueSource.field` | Excluded for computed sources. |
| `DataGridValueSource.accessor` | Pure synchronous source returning original values. |

### `DataGridColumnCommonProps`

Shared column configuration. Sources and callbacks must be pure.

| Member | Contract |
| --- | --- |
| `DataGridColumnCommonProps.id` | Unique non-empty column ID. |
| `DataGridColumnCommonProps.header` | Visible header. |
| `DataGridColumnCommonProps.headerAriaLabel` | Accessible name for a non-text header. |
| `DataGridColumnCommonProps.headerTooltip` | Plain-text header description. |
| `DataGridColumnCommonProps.tooltip` | Plain-text description of valid original data. Missing values are normalized to undefined for this callback, including legacy null; invalid values skip the callback and use a static tooltip or the localized invalid label. |
| `DataGridColumnCommonProps.placeholder` | Missing source presentation; defaults to an em dash. |
| `DataGridColumnCommonProps.invalidPlaceholder` | Present but invalid source presentation. |
| `DataGridColumnCommonProps.sortable` | Enables controlled sorting intent, without changing item order. |
| `DataGridColumnCommonProps.width` | Preferred width in pixels. |
| `DataGridColumnCommonProps.minWidth` | Minimum width in pixels. |
| `DataGridColumnCommonProps.locale` | Overrides the root locale. |

### `DataGridColumnDefinition`

Immutable opaque descriptor owned by one suite. Create it with defineColumns.

| Member | Contract |
| --- | --- |
| `DataGridColumnDefinition.kind` | Display kind, without exposing captured callbacks. |
| `DataGridColumnDefinition.id` | Stable unique column ID. |

### `DataGridColumnBuilder`

Builder for all sixteen presentation kinds.

| Member | Contract |
| --- | --- |
| `DataGridColumnBuilder.text` | Captures a text source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.number` | Captures a number source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.boolean` | Captures a boolean source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.date` | Captures a date source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.dateTime` | Captures a dateTime source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.badge` | Captures a badge source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.icon` | Captures a icon source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.link` | Captures a link source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.person` | Captures a person source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.people` | Captures a people source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.tags` | Captures a tags source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.progress` | Captures a progress source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.rating` | Captures a rating source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.sparkline` | Captures a sparkline source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.custom` | Captures a custom source and creates a same-suite immutable descriptor. |
| `DataGridColumnBuilder.actions` | Creates a source-free action descriptor. |

### `DataGridValueColumnProps`

Props for one strongly typed value column.

### `DynamicDataGridColumnProps`

Explicit dynamic source props. Generic callbacks receive unknown.

| Member | Contract |
| --- | --- |
| `DynamicDataGridColumnProps.field` | Excluded for computed sources. |
| `DynamicDataGridColumnProps.accessor` | Pure synchronous source; its return remains unknown. |

### `DynamicDataGridColumnBuilder`

Explicit dynamic builder; generic source callbacks receive unknown.

| Member | Contract |
| --- | --- |
| `DynamicDataGridColumnBuilder.text` | Creates a guarded dynamic text descriptor. |
| `DynamicDataGridColumnBuilder.number` | Creates a guarded dynamic number descriptor. |
| `DynamicDataGridColumnBuilder.boolean` | Creates a guarded dynamic boolean descriptor. |
| `DynamicDataGridColumnBuilder.date` | Creates a guarded dynamic date descriptor. |
| `DynamicDataGridColumnBuilder.dateTime` | Creates a guarded dynamic dateTime descriptor. |
| `DynamicDataGridColumnBuilder.badge` | Creates a guarded dynamic badge descriptor. |
| `DynamicDataGridColumnBuilder.icon` | Creates a guarded dynamic icon descriptor. |
| `DynamicDataGridColumnBuilder.link` | Creates a guarded dynamic link descriptor. |
| `DynamicDataGridColumnBuilder.person` | Creates a guarded dynamic person descriptor. |
| `DynamicDataGridColumnBuilder.people` | Creates a guarded dynamic people descriptor. |
| `DynamicDataGridColumnBuilder.tags` | Creates a guarded dynamic tags descriptor. |
| `DynamicDataGridColumnBuilder.progress` | Creates a guarded dynamic progress descriptor. |
| `DynamicDataGridColumnBuilder.rating` | Creates a guarded dynamic rating descriptor. |
| `DynamicDataGridColumnBuilder.sparkline` | Creates a guarded dynamic sparkline descriptor. |
| `DynamicDataGridColumnBuilder.custom` | Creates a guarded dynamic custom descriptor. |
| `DynamicDataGridColumnBuilder.actions` | Creates a source-free dynamic action descriptor. |

### `ActionsColumnProps`

Actions deliberately have no source or sorting option.

| Member | Contract |
| --- | --- |
| `ActionsColumnProps.header` | Optional actions heading; defaults to the localized actions label. |
| `ActionsColumnProps.actions` | Ordered synchronous actions evaluated for the original row. |
| `ActionsColumnProps.field` | Actions cannot read a field. |
| `ActionsColumnProps.accessor` | Actions cannot read an accessor. |

### `DataGridMappedValues`

Exact dictionaries for string domains; numeric keys require a Map.

### `DataGridBadgeValue`

Badge presentation mapped from an exact scalar value.

| Member | Contract |
| --- | --- |
| `DataGridBadgeValue.label` | Display and accessible label. |
| `DataGridBadgeValue.color` | Fluent color. |
| `DataGridBadgeValue.appearance` | Fluent appearance; defaults to tint. |
| `DataGridBadgeValue.shape` | Fluent shape. |
| `DataGridBadgeValue.size` | Fluent size. |
| `DataGridBadgeValue.icon` | Optional icon. |

### `DataGridIconValue`

Informative icon and its text alternative.

| Member | Contract |
| --- | --- |
| `DataGridIconValue.icon` | Visible icon. |
| `DataGridIconValue.label` | Accessible label. |

### `DataGridTagValue`

Tag presentation mapped from an exact scalar ID.

| Member | Contract |
| --- | --- |
| `DataGridTagValue.label` | Display and accessible label. |
| `DataGridTagValue.appearance` | Fluent appearance. |
| `DataGridTagValue.shape` | Fluent shape. |
| `DataGridTagValue.size` | Fluent size. |
| `DataGridTagValue.icon` | Optional icon. |

### `DataGridTagValues`

Exact scalar tag maps; object-tag callback IDs use a guarded string/number map.

### `DataGridLinkNavigateData`

Original data supplied to link navigation.

| Member | Contract |
| --- | --- |
| `DataGridLinkNavigateData.item` | Original row. |
| `DataGridLinkNavigateData.rowId` | Original stable ID. |
| `DataGridLinkNavigateData.value` | Valid original text. |
| `DataGridLinkNavigateData.href` | Resolved destination. |

### `DataGridSparklinePoint`

A source-order chart observation.

| Member | Contract |
| --- | --- |
| `DataGridSparklinePoint.x` | Numeric or date X coordinate. |
| `DataGridSparklinePoint.y` | Finite original observation. |

### `DataGridSparklineMetric`

Metric available to consumers reading a chart summary.

### `DataGridSparklineSummary`

Immutable chart summary; each derived change is independently optional.

| Member | Contract |
| --- | --- |
| `DataGridSparklineSummary.first` | First observation. |
| `DataGridSparklineSummary.last` | Last observation. |
| `DataGridSparklineSummary.minimum` | Minimum observation. |
| `DataGridSparklineSummary.maximum` | Maximum observation. |
| `DataGridSparklineSummary.average` | Overflow-safe arithmetic mean. |
| `DataGridSparklineSummary.change` | Last minus first, absent when not representable. |
| `DataGridSparklineSummary.changePercent` | Percent change, absent for zero denominator or unrepresentable result. |

### `DataGridActionData`

Original row and action identities.

| Member | Contract |
| --- | --- |
| `DataGridActionData.item` | Original row. |
| `DataGridActionData.rowId` | Stable row ID. |
| `DataGridActionData.actionId` | Stable action ID. |

### `DataGridAction`

One synchronous row action.

| Member | Contract |
| --- | --- |
| `DataGridAction.id` | Non-empty unique action ID. |
| `DataGridAction.label` | Visible and accessible label. |
| `DataGridAction.icon` | Optional icon. |
| `DataGridAction.placement` | Placement in the row. |
| `DataGridAction.visible` | Static or row-dependent visibility. |
| `DataGridAction.disabled` | Static or row-dependent disabled state. |
| `DataGridAction.loading` | Static or row-dependent busy state. |
| `DataGridAction.onClick` | Receives the original item and IDs. |

### `DataGridNumberOptions`

Intl number presentation.

| Member | Contract |
| --- | --- |
| `DataGridNumberOptions.format` | Formatting changes presentation only. |

### `DataGridBooleanOptions`

Readonly boolean presentation.

| Member | Contract |
| --- | --- |
| `DataGridBooleanOptions.display` | Text or a vertically centered native 20px informative icon; defaults to text. |
| `DataGridBooleanOptions.trueLabel` | Visible and accessible true label. |
| `DataGridBooleanOptions.falseLabel` | Visible and accessible false label. |

### `DataGridDateOptions`

Absolute calendar or instant presentation.

| Member | Contract |
| --- | --- |
| `DataGridDateOptions.timeZone` | Explicit IANA zone for Date objects and dateTime instants; defaults to UTC. Calendar-only YYYY-MM-DD strings always display their UTC calendar day. |
| `DataGridDateOptions.format` | Date formatting; configure the zone with timeZone. |

### `DataGridBadgeOptions`

Badge mapping keeps the exact source scalar domain.

| Member | Contract |
| --- | --- |
| `DataGridBadgeOptions.values` | Exact scalar values mapped to Fluent badge presentations. |

### `DataGridIconOptions`

Informative icon mapping keeps exact source keys.

| Member | Contract |
| --- | --- |
| `DataGridIconOptions.values` | Exact scalar values mapped to icons and labels. |

### `DataGridLinkOptions`

Link destination and navigation intent.

| Member | Contract |
| --- | --- |
| `DataGridLinkOptions.href` | Static or row-aware destination for valid text. |
| `DataGridLinkOptions.linkProps` | Native anchor presentation; the grid owns its element, content and navigation. |
| `DataGridLinkOptions.onNavigate` | Receives original item, row ID, value and destination. |

### `DataGridPersonOptions`

Person display callbacks retain their inferred person type.

| Member | Contract |
| --- | --- |
| `DataGridPersonOptions.getName` | Required person name. |
| `DataGridPersonOptions.getSecondaryText` | Optional secondary text. |
| `DataGridPersonOptions.getAvatarProps` | Avatar options except its owned name. |
| `DataGridPersonOptions.getPresenceProps` | Native presence configuration. |
| `DataGridPersonOptions.personaProps` | Native Persona presentation options; textAlignment defaults to center. |

### `DataGridPeopleOptions`

Multiple-person identity and overflow options.

| Member | Contract |
| --- | --- |
| `DataGridPeopleOptions.getPersonId` | Returns a stable ID unique within this cell. |
| `DataGridPeopleOptions.maxVisible` | Maximum visible avatars before the recoverable overflow affordance. |

### `DataGridTagsOptions`

Exact tag mappings or paired scalar/object tag callbacks. Missing configuration and incomplete callback pairs are rejected by types and runtime.

| Member | Contract |
| --- | --- |
| `DataGridTagsOptions.maxVisible` | Maximum visible tags before recoverable overflow. |
| `DataGridTagsOptions.values` | Optional mapped presentation for IDs returned by getTagId. |
| `DataGridTagsOptions.getTagId` | Stable scalar ID for each tag. Supply together with getLabel. |
| `DataGridTagsOptions.getLabel` | Explicit display label for each tag. Supply together with getTagId. |

### `DataGridProgressOptions`

Determinate readonly progress presentation.

| Member | Contract |
| --- | --- |
| `DataGridProgressOptions.max` | Positive finite upper bound; defaults to one. |
| `DataGridProgressOptions.showValue` | Display the localized value next to the indicator. |
| `DataGridProgressOptions.progressBarProps` | Native visual options, excluding owned value and accessibility. |

### `DataGridRatingOptions`

Readonly rating presentation.

| Member | Contract |
| --- | --- |
| `DataGridRatingOptions.max` | Whole number greater than one; defaults to five. |
| `DataGridRatingOptions.count` | Optional non-negative whole number of votes for the accessible description only. Zero is included; undefined omits it. Localize with strings.ratingCount. |
| `DataGridRatingOptions.ratingDisplayProps` | Native visual options, excluding owned value and range. Defaults to marigold, Star20Filled and native large (20px); explicit consumer overrides remain supported. Numeric slots are hidden. |

### `DataGridSparklineOptions`

Official Charts Sparkline presentation.

| Member | Contract |
| --- | --- |
| `DataGridSparklineOptions.showValue` | Show the latest observation; accessible summary remains when false. |
| `DataGridSparklineOptions.color` | Static or summary-aware Fluent color string. |

### `DataGridCustomOptions`

Custom content preserves original source and JSX identity.

| Member | Contract |
| --- | --- |
| `DataGridCustomOptions.renderCell` | Renders one valid value with original row identity. |

## Native UI localization

SPFx selects the DataGrid native resource from the host UI culture after the host registers its toolkit resource ID through the localization configuration CLI. `strings?: Partial<DataGridStrings>` replaces static labels and callbacks per grid instance. Omitted or undefined entries and blank static strings retain that baseline; nonblank strings preserve surrounding whitespace. Invalid override object/member types throw. A complete dictionary replaces every owned label.

Root and column `locale`, `format` and `timeZone` still govern data formatting independently of the UI resource. The source resource culture selects default plural grammar. `ratingCount` receives the original numeric count and its locale-formatted text; `tagOverflow` receives the original hidden count; `sparklineSummary` receives the existing formatted metrics. Source rows, mapped labels, dates, relative time text and callback arguments are unchanged.

```tsx
<Grid items={rows} getRowId={row => row.id} columns={columns} locale="it-IT"
  strings={{ empty: 'No records', ratingCount: (count, formattedCount) =>
    `${formattedCount} ${count === 1 ? 'vote' : 'votes'}` }} />
```

This example keeps Italian data formatting with an English vote label; omitted labels use the host resource. No additional provider is required.

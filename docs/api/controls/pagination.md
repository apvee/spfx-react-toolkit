# Pagination

[Documentation](../../INDEX.md) > [API reference](../INDEX.md#react-controls) > Pagination

`Pagination` renders accessible controlled result navigation. The application owns the supplied rows, page/cursor availability, requests and loading state. It does not fetch, slice results or retain pagination state. Import the same implementation from the package root, `/controls`, or `/controls/pagination`.

```tsx
import * as React from 'react';
import { Pagination } from '@apvee/spfx-react-toolkit/controls/pagination';

export function OrdersNavigation() {
  const [pageIndex, setPageIndex] = React.useState(0);
  return <Pagination
    pagination={{ type: 'pages', pageIndex, pageSize: 25, totalItems: 100 }}
    onPaginationChange={(_event, request) => {
      if (request.type === 'pages') setPageIndex(request.pageIndex);
    }}
  />;
}
```

Supply the corresponding result rows separately, for example to a typed `createDataGrid` suite created outside rendering. For local arrays, slice using `pageIndex * pageSize`. For server data, resolve the request and replace controlled state only when the results arrive. Reset the application page/cursor and pending request when filters, page size or the dataset change. Announce the loaded page/item interval in a consumer-owned live region; the control does not know result contents.

## PaginationProps

The native root is `nav`, with an `HTMLElement` ref. Native attributes, events, `hidden`, classes and consumer `style` are forwarded; `children` and `dangerouslySetInnerHTML` are excluded. Root layout uses classes and the host renderer. Native `hidden` remains effective. A nonempty `aria-label` overrides the localized navigation name; native `aria-labelledby` retains normal browser precedence. Loading sets `aria-busy` on the navigation.

| Member | Type | Behavior/default |
| --- | --- | --- |
| `PaginationProps.pagination` | `PaginationState` | Required application-owned state. |
| `PaginationProps.onPaginationChange` | `(event: React.SyntheticEvent, data: PaginationChangeData) => void` | Optional available-action request; the application supplies updated state. Repeated requests retain the received state. A prevented click does not request navigation. |
| `PaginationProps.loading` | `boolean` | false; disables all actions and shows an accessible spinner. |
| `PaginationProps.size` | `PaginationSize` | medium; matches native Fluent Button small/medium/large sizing. |
| `PaginationProps.appearance` | `ButtonProps['appearance']` | subtle; native Fluent appearance applied to every action, including Load more and loading buttons. |
| `PaginationProps.strings` | `Partial<PaginationStrings>` | Localized text overrides with the host-selected native baseline for absent or blank labels. |

`appearance` supports `'primary' | 'secondary' | 'outline' | 'subtle' | 'transparent'`. Omit it to retain subtle navigation, or use `appearance="outline"` for bordered actions. Fluent owns theme, interaction and disabled colors; appearance does not change button size, request semantics or spinner placement.

```tsx
<Pagination
  pagination={{ type: 'loadMore', hasNextPage: hasMore }}
  appearance="secondary"
  loading={loadingMore}
  onPaginationChange={requestMore}
/>
```

## PaginationSize

`PaginationSize` follows native Fluent Button: `'small' | 'medium' | 'large'`, with medium as the default. Progress follows button density: 16px (`extra-tiny`) for small, 20px (`tiny`) for medium and 24px (`extra-small`) for large. Buttons use the requested native Fluent appearance, defaulting to subtle. The centered layout wraps within its container. Loading replaces the icon of the last activated action with a spinner in the same icon slot. There is no reserved gap between navigation actions. Load-more has no icon at rest; during loading, only the proportional spinner appears beside its visible text. Each instance remembers only its last clicked action for presentation; page state, results and loading remain controlled. Loading supplied before any action uses Next for pages/cursor and the Load more action for incremental mode. Every button remains disabled during loading, and the button containing the spinner retains its action accessible name. Logical navigation icons mirror with host RTL direction.

## PaginationState

`PaginationState` is `PaginationPagesKnownTotal | PaginationPagesUnknownTotal | PaginationCursor | PaginationLoadMore`. All modes are controlled. There are no numbered-page links, ellipsis, page-size selector or uncontrolled defaults.

## PaginationPagesKnownTotal

| Member | Type | Behavior |
| --- | --- | --- |
| `PaginationPagesKnownTotal.type` | `'pages'` | Offset pagination with an exact total. |
| `PaginationPagesKnownTotal.pageIndex` | `number` | Non-negative safe integer, zero-based. |
| `PaginationPagesKnownTotal.pageSize` | `number` | Positive safe integer, retained in requests. |
| `PaginationPagesKnownTotal.totalItems` | `number` | Non-negative safe integer exact count. |
| `PaginationPagesKnownTotal.hasNextPage` | `never` | Excluded when totalItems is supplied. |

Page count is `Math.ceil(totalItems / pageSize)`. First requests page 0; last requests `Math.max(0, pageCount - 1)`. Previous requests `pageIndex - 1`; next requests `pageIndex + 1`. First/previous disable at page 0, next/last disable at or beyond the final page. An empty result at page 0 disables all four actions. A stale index beyond the total is deliberately retained: first can recover to 0, previous requests the received index minus one, and next/last remain disabled. No implicit clamping, state mutation or correction callback occurs; reset stale state in the application.

## PaginationPagesUnknownTotal

| Member | Type | Behavior |
| --- | --- | --- |
| `PaginationPagesUnknownTotal.type` | `'pages'` | Offset pagination without an exact total. |
| `PaginationPagesUnknownTotal.pageIndex` | `number` | Non-negative safe integer, zero-based. |
| `PaginationPagesUnknownTotal.pageSize` | `number` | Positive safe integer, retained in requests. |
| `PaginationPagesUnknownTotal.hasNextPage` | `boolean` | Application-owned following-page availability. |
| `PaginationPagesUnknownTotal.totalItems` | `never` | Excluded; explicit undefined retains unknown-total behavior. |

Only previous/next render. Previous disables at page 0; next disables when `hasNextPage` is false or the index is `Number.MAX_SAFE_INTEGER`, so no unsafe index is emitted. The application decides whether more results exist.

## PaginationCursor

| Member | Type | Behavior |
| --- | --- | --- |
| `PaginationCursor.type` | `'cursor'` | Directional requests without exposing cursor values. |
| `PaginationCursor.hasPreviousPage` | `boolean` | Enables previous navigation. |
| `PaginationCursor.hasNextPage` | `boolean` | Enables next navigation. |

Cursor tokens/history remain private to the application. Resolve the requested direction and then supply new capability flags and result rows.

## PaginationLoadMore

| Member | Type | Behavior |
| --- | --- | --- |
| `PaginationLoadMore.type` | `'loadMore'` | Incremental append request. |
| `PaginationLoadMore.hasNextPage` | `boolean` | Enables the visible Load more button. |

Append results in the application. At rest, the button shows text only. During loading, a proportional spinner appears in the icon slot, retaining the text; exhaustion disables it without hiding the existing results.

## PaginationChangeData

| Member | Type | Behavior |
| --- | --- | --- |
| `PaginationChangeData.type` | `'pages' \| 'cursor' \| 'loadMore'` | Discriminates the semantic request. |
| `PaginationChangeData.pageIndex` | `number` | Present for pages: requested zero-based index. |
| `PaginationChangeData.pageSize` | `number` | Present for pages: received page size. |
| `PaginationChangeData.direction` | `'previous' \| 'next'` | Present for cursor: requested direction. |

Load-more data contains only `{ type: 'loadMore' }`. Each available click emits one request. Navigation controls have `type="button"`, so an enclosing form is not submitted.

## PaginationStrings

| Member | en-US resource example | Purpose |
| --- | --- | --- |
| `PaginationStrings.pagination` | Pagination | Navigation landmark and group accessible name. |
| `PaginationStrings.firstPage` | First page | First action accessible name. |
| `PaginationStrings.previousPage` | Previous page | Previous action accessible name. |
| `PaginationStrings.nextPage` | Next page | Next action accessible name. |
| `PaginationStrings.lastPage` | Last page | Last action accessible name. |
| `PaginationStrings.loadMore` | Load more | Visible label and accessible name. |
| `PaginationStrings.loading` | Loading | Spinner accessible name. |

SPFx selects the native Pagination resource from the host UI culture. Labels are trimmed; missing, undefined or whitespace-only overrides use that native baseline. The table shows the en-US resource values. The navigation contains one labeled group and native Fluent buttons. It does not move focus or announce fetched results automatically.

## Validation and dependencies

Malformed runtime states render an empty fragment. Numeric fields reject negatives, fractions, NaN, Infinity and unsafe integers; page size also rejects zero. Known total and next-page capability are mutually exclusive. Cursor/load-more availability must be boolean. Use undefined for absence.

No SPFx provider or new dependency is required. The control uses existing shared React 17, Fluent UI 9 Button/Spinner/tokens, Fluent Icons and Griffel peers. It inherits host theme, direction and renderer through the established styling integration. It does not create a renderer or register PnP features.

## Working sample and validation limits

The lazy Pagination tab shows one small result list and its navigation. Mode, page size, filter, size/appearance and explicit pending completion drive the current result. Configuration/status/diagnostics share one outline card. The separate enclosing-form/native/isolation comparisons remain in the [versioned regression fixture](../../../tests/fixtures/demo-control-isolation/pagination-regressions.tsx). Source/compiled runners exercise the product route before the preserved regression scenarios; local execution does not establish authenticated host behavior.

Authenticated SharePoint, Teams/iframe and screen-reader validation is **NOT EXECUTED** until separately recorded. Static demo checks do not establish rendered or tenant behavior.

## Native UI localization

Pagination requires the host to register its toolkit resource ID through the [localization configuration CLI](../../guides/localization.md). SPFx selects the UI resource; no toolkit provider is required. Outside the SPFx loader, supply an explicit resource-module adapter: the native resource is a static dependency even when passing a complete dictionary.

`strings?: Partial<PaginationStrings>` overrides labels per instance. Overrides are trimmed, and omitted, undefined or blank labels use the selected resource. A complete nonblank `PaginationStrings` dictionary replaces every owned label. This does not translate result summaries or change controlled navigation callbacks.

```tsx
<Pagination pagination={{ type: 'cursor', hasPreviousPage: false, hasNextPage: true }}
  strings={{ nextPage: 'Next batch', loading: 'Fetching results' }} />
```

# PnP Search Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

```ts
import { createSPFxPnPSearchService } from '@apvee/spfx-react-toolkit/services';
```


## `createSPFxPnPSearchService`

```ts
function createSPFxPnPSearchService<T = Record<string, string>>(
  sp: SPFI,
  defaultOptions?: SPFxPnPSearchOptions
): SPFxPnPSearchService<T>;
```

Required public types:

- `SPFxPnPSearchQueryBuilder`
- `SPFxPnPSearchOptions`
- `SPFxPnPSearchRequestOptions`
- `SPFxPnPSearchResult`
- `SPFxPnPSearchRefiner`
- `SPFxPnPSearchResponse`
- `SPFxPnPSearchService`

Returned methods:

| Method | Description |
|--------|-------------|
| `search(query, options?)` | Runs SharePoint Search from a text query or query builder |
| `suggest(queryText)` | Returns query suggestions |

Default options can set `pageSize`, `selectProperties`, and `refiners`. Per-request options can set `pageSize`, `startRow`, and refinement filters.

`SPFxPnPSearchResponse<T>` contains `results`, `totalResults`, and `refiners`.

# PnP Search Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

Import this factory and its public types from `@apvee/spfx-react-toolkit/services` or the package root. Callers supply the initialized client/context; these services require no React provider.

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

Default options can set `pageSize`, `selectProperties`, `refiners`, and `refinerOperator`. Per-request options can set `pageSize`, `startRow`, refinement filters, and a `refinerOperator` override.

`refinementFilters` maps managed-property names to arrays of server-returned string refinement tokens. Use each entry's `token`, including any FQL quote delimiters, rather than its display `value`. The service handles tokens generically:

| Input | Emitted operand |
| --- | --- |
| Quoted string or quoted hex token | Unchanged, without `equals(...)` |
| Unquoted hex token (`ǂǂ...`) | Quoted once |
| Numeric or ISO date scalar | Unchanged |
| Native FQL expression (`range(...)`, `datetime(...)`, etc.) | Unchanged; SharePoint validates it |
| Legacy unquoted text | `equals("value")`, escaping quotes and backslashes |

`refinerOperator: 'and' | 'or'` defaults to `'and'`. It controls values within each property; different properties always combine with AND. OR groups a property's operands as `Property:or(token1,token2)`, including numeric/date intervals. A request override takes precedence over the service default. Empty arrays contribute no restriction. For person facets, the sample requests `DisplayAuthor`; `Author` can still be selected as result metadata. Property names are preserved without automatic remapping.

```ts
const service = createSPFxPnPSearchService(sp, {
  refiners: 'FileType,DisplayAuthor,Size,Write', refinerOperator: 'or'
});
const response = await service.search('*');
const fileType = response.refiners.find(refiner => refiner.name === 'FileType');
const pdf = fileType?.entries.find(entry => entry.value === 'pdf');
if (pdf?.token) {
  await service.search('*', {
    refinementFilters: new Map([['FileType', [pdf.token]]])
  });
}
```

`SPFxPnPSearchResponse<T>` contains `results`, `totalResults`, and `refiners`.

## Option precedence and result types

| Public contract | Fields and behavior |
| --- | --- |
| `SPFxPnPSearchQueryBuilder` | `(builder: ISearchBuilder) => ISearchBuilder`; receives select-property/refiner defaults before caller changes. Return the configured builder. |
| `SPFxPnPSearchOptions` | Optional `pageSize: number` defaults to 50; `selectProperties?: string[]` and `refiners?: string` set builder defaults when nonempty; `refinerOperator?: 'and' \| 'or'` defaults to and. |
| `SPFxPnPSearchRequestOptions` | Optional `pageSize` overrides service size and any builder row limit; positive `startRow` overrides builder offset, while omitted/zero keeps it; `refinementFilters?: Map<string,string[]>` supplies tokens; optional `refinerOperator` overrides the service operator. |
| `SPFxPnPSearchResult<T>` | `id: string` comes from DocId, then Path, then random fallback (not stable); `data: T` and `raw: unknown` reference the same unchanged PnP result; optional `rank: number` is parsed from truthy Rank (zero is absent). |
| `SPFxPnPSearchRefiner` | `name: string` defaults to empty; `entries` contains `{ value: string; count: number; token: string }`, with absent labels/tokens becoming empty text and absent/unparseable counts zero. |
| `SPFxPnPSearchResponse<T>` | Current-page `results`, SharePoint `totalResults` (zero when absent), and returned `refiners` (empty when absent). |
| `SPFxPnPSearchService<T>` | `search(string \| SPFxPnPSearchQueryBuilder, options?)` returns the response; `suggest(queryText: string)` returns `Promise<string[]>`, empty when no Queries are returned. |

Neither `T` nor `data` converts cells or validates application fields. Search/suggestion request failures and builder errors reject; no empty-success fallback, retry, React state or authentication acquisition is added. This module registers the PnP search/suggestion feature. Construction performs no request and a caller-supplied configured/authenticated `SPFI` is required.

```ts
const search = createSPFxPnPSearchService<{ Title: string; Path: string }>(sp, {
  pageSize: 25, selectProperties: ['Title', 'Path'],
});
const page = await search.search(builder => builder.text('training'), { startRow: 25 });
const suggestions = await search.suggest('train');
console.log(page.results[0]?.data.Title, suggestions);
```

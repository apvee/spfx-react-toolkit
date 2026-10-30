# SharePoint pickers

[Documentation](../../INDEX.md) > [API reference](../INDEX.md#react-controls) > SharePoint pickers

The seven controls use Fluent UI 9 Dropdown or TagPicker, with Persona rows. They require the matching `SPFxWebPartProvider` (or extension provider), native authenticated clients and a host FluentProvider/renderer. Use `@apvee/spfx-react-toolkit/controls/pickers`; the same React exports and JSON types are available from the package root and `/controls`.

```tsx
import { SitePicker, PrincipalPicker } from '@apvee/spfx-react-toolkit/controls/pickers';

<SitePicker label="Site" mode="site" presentation="combobox" multiselect={false} />;
<PrincipalPicker label="Principals" principalTypes={['user', 'sharePointGroup', 'microsoft365Group']}
  presentation="tagPicker" multiselect userIdentity={{ mode: 'both', ensureUser: false }} />;
```

## Shared contract

### Dropdown selection and popup search

The public `presentation="combobox"` string is retained for compatibility and now renders Dropdown. Selection and search behavior is unchanged; no `"dropdown"` value is added. Its selection button always displays the selected name and media, or `N items selected` for multiple values. Empty fields show `Select an item` / `Select items`. `placeholder` and picker overrides such as SitePicker's `searchPlaceholder` customize search text, with `Type at least N characters to search` as the searchable fallback, using the configured `minimumSearchLength`; they do not replace the empty selection prompt.

With `searchable !== false`, opening focuses a SearchBox in the popup, outside the semantic list. Query clear affects only search, while native single-selection clear emits a validated clear action. `searchable=false` omits this SearchBox. TagPicker keeps its input and removable tags; selected identities are excluded from every candidate/contextual section and are not repeated as selected popup rows. `colorfulTagAvatars` applies only to selected tags.

Arrow keys navigate native options, Enter confirms, Escape closes and returns focus to the trigger, and Tab/Shift+Tab dismiss without selecting a highlighted option. Editing keys, spaces and composition stay with search. Single selection closes after asynchronous confirmation/validation accepts; rejection preserves the query and popup and emits no selection event. Pending confirmation preserves focus; a late result cannot steal outside focus or close a newer popup session. Blur validation observes exit from the entire control, including popup and actions.

Retry, explicit ID resolution and **Altri risultati** remain separate actions below the selection control. Reopening after an action retains the current query and accumulated pages.

The changed DOM and interaction are intentional: `.fui-Combobox` selectors and a permanent text input in the closed field are no longer supported assumptions. `controlClassName` targets Dropdown, while `id` and field ARIA naming target its actual button; popup search has its own name. Public imports, props, selection callbacks and JSON references remain unchanged. See the [native fields](../property-pane/pickers.md) for persistence and Apply ownership.


`PickerProps<T>` extends the native Fluent Field label, hint, required, validationMessage and validationState contracts. `PickerOption<TValue>` is `{ key, text, value }` plus domain display metadata. Keys are stable opaque renderer aliases; canonical identity comes from the validated scoped value, never a display label. Generated keys carry a versioned JSON reference and can hydrate without a display model. An arbitrary consumer key needs a model; missing models remain unresolved.

The picker additionally requires an accessible input name: a nonblank plain string `label`, or explicit `aria-label`/`aria-labelledby`. JSX and Fluent object labels retain their native slot types, but require explicit ARIA naming even when they display text. The renderer throws `TypeError` if neither naming input is supplied. A supplied `aria-label` takes precedence over the plain label, so an empty override without `aria-labelledby` also fails this check. Consumers own the referenced IDs and meaningful accessible text.

```tsx
<SitePicker label={<span>Site</span>} aria-label="Site" />
```

| Input | Purpose | Updates |
| --- | --- | --- |
| `selectedOptions` | Controlled selection keys; `[]` clears | Authoritative; omission enables uncontrolled mode |
| `defaultSelectedOptions` | Uncontrolled initial keys | Read once per mount |
| `initialOptions` | Models for selected values | Does not add blank-query suggestions; retained selections survive model-set shrinkage |
| `availableOptions` | Consumer candidates from any source | Recomposes the menu without resetting selection or remounting; equivalent identities reuse eligibility requests |
| `initialSuggestions.limit` | Site/Item/Principal pre-search cap | Integer 1–5, default 5 |
| `initialCatalogLimit` | ListPicker available catalog cap | Integer 1–200, default 30; selected rows are separate |
| `refreshKey` | Explicit cache invalidation token | Change to retry/refresh the effective source |

All controls support `presentation="combobox" | "tagPicker"`, single/multiple selection and typed search. Searchable defaults to true, minimumSearchLength to 2 (integer at least 2), debounceMs to 300 (integer at least 0), suggestionsLimit to 50 (integer 1–200). The result page size is independent of the initial suggestion/catalog cap. `selectionLimit` limits new choices; single selection always allows at most one new value. Existing conflicting controlled values are preserved and diagnosed. `showDefaultSuggestions=false` suppresses initial suggestion sources and their lazy requests while leaving typed search available.

Consumer candidates are verified read-only against their source, type and scope before blank-query display; read-only eligibility runs in batches of at most 20 concurrent requests, without truncating the supplied set. Results are cached by source/client/identity, so new arrays, labels and callbacks do not repeat equivalent requests. Consumer models precede the scoped List/Field/View/ContentType catalog and optional current-user and personal models, win equivalent display identities, and are deduplicated after filtering. Already selected identities are excluded before applying the total cap. The search placeholder displays “Type at least 2 characters to search” (or the configured larger minimum); empty or below-minimum queries do not duplicate that guidance in a status row. Typing queries the configured SharePoint/Graph source, including entities outside the consumer set; Field/View/ContentType instead filter their complete scoped cache locally. `filterOptions` is a local transformation; mandatory scope/type, selected exclusion, deduplication and the initial cap apply again afterward.

### Search feedback

All seven controls share the same feedback for Dropdown and TagPicker, in single and multiple selection. Feedback describes the **current query**, independently of selected rows/tags and SitePicker's fixed current-site/actual-hub choices. Those choices remain visible when enabled, but their presence alone does not establish a search match or hide a no-results message.

| Current query state | Default feedback |
| --- | --- |
| Nonempty text shorter than `minimumSearchLength` | No default status row; the minimum-character guidance belongs to the search placeholder |
| Blank query without available options | Minimum-character guidance in the placeholder only when searchable; “No options available.” when `searchable=false` |
| Eligible query waiting for debounce or loading | “Searching…”; no temporary no-results message |
| Settled, complete search with no eligible matches | “No results for this search” |
| Settled, complete search whose only eligible matches are already selected or pinned contextual choices | “No other results for this search” |
| Incomplete or bounded search | A partial-result warning; never a claim that the complete search has no results |

Eligible matches respect the configured source, scope and filters. The distinction between no matches and no **other** matches is evaluated before selected and pinned identities are removed from ordinary result rows. A previous query's completion, empty result or failure cannot replace the current query's feedback. Query failures and contextual-suggestion failures are shown independently, alongside any healthy rows; a failed hub suggestion must not conceal a successful empty search, and a query failure must not conceal the contextual failure.

`emptyContent`, `loadingContent` and `errorContent` retain their existing fixed inner `div` slot contracts. Omission selects the state-specific defaults above; literal content or shorthand customizes that region, and `false` suppresses its content while preserving the control-owned status shell. Explicit custom empty text still applies to initial/below-minimum guidance and complete empty/no-other-result states; partial warnings remain distinct. Retry remains an independent action. These slots change presentation, not query timing, selection, eligibility, error reporting or completion guarantees.

For example, keep the defaults to distinguish an unfinished search from a complete empty result, or supply product-specific wording through the existing slots:

```tsx
<SitePicker label="Site" minimumSearchLength={3} includeCurrentSite includeHubSite />;
<ListPicker label="List" emptyContent="Nessuna lista disponibile per i criteri correnti" />;
<PrincipalPicker label="Principals" loadingContent="Ricerca persone e gruppi…" />;
```

To observe the distinction, open SitePicker with contextual choices enabled, enter fewer than the minimum characters, then enter an eligible query with no matches. The placeholder supplies the initial guidance; an eligible query shows pending feedback and then the complete-empty message while contextual choices remain available. Search for a selected or pinned identity to observe the no-other-results state. Repeat with TagPicker or multiple selection; changing presentation does not change the query contract. Field/View/ContentType load metadata remotely once and then match locally over the complete scoped catalog. ListItemPicker retrieves remote candidate pages and matches locally within those pages, so a bounded scan can remain partial even with no matching rows. Other remote search sources retain their documented paging and permission limits.

All picker search inputs set browser `autocomplete="off"` in both Dropdown and TagPicker presentations, including Property Pane. This requests suppression of browser history/autofill suggestions; picker search and suggestions remain enabled. Browsers and password-manager extensions may override this hint.

Names and secondary metadata in Persona suggestion/selection rows truncate with an ellipsis when the control is narrow. All avatars use Fluent `shape="square"` with its native corner rounding. Popup rows use 24px avatars; the selected Dropdown value uses 20px for medium and 24px for large. Entity icons use Fluent's theme brand palette; people retain their photos or neutral initials. Selected tags also stay within the control, keeping their icon and removal action visible. Tag media avatars inherit Fluent’s native tag size, so icons, photos and initials remain proportional for the supported medium and large picker sizes. Full labels remain in the DOM/accessibility text and native hover titles; display truncation does not alter keys or saved values. This presentation is shared with Property Pane.

### Initial catalog and selection visibility

ListPicker loads the scoped remote catalog lazily when the empty-query menu opens. `initialCatalogLimit` defaults to **30** (integer 1–200), applied after consumer precedence, deduplication, filtering and selected exclusion. Selected rows do not consume this cap. A lookahead/continuation distinguishes exactly 30 results from additional results and displays the search hint. Composition reads at most five source pages of 200 and preserves partial-source warnings.

FieldPicker and ViewPicker load **all** matching fields/views, following every native REST continuation without a display or five-page limit. They show the entire available catalog, excluding selected identities and applying configured filters. **`initialCatalogLimit` is removed from both React and Property Pane contracts.** Typed search is local, case insensitive prefix matching over the complete cached catalog, even for entries beyond 30 or 200. Changing query/casing/page-size does not make additional catalog REST calls. Source identity/client/user changes, cache expiration or explicit refresh can reload the catalog; failed or cyclic paging is reported as an error rather than complete data. Read-only selection hydration/eligibility requests remain separate from search. Raw filters/order remain server-side. Native paging is retrieved before showing the complete catalog, so large metadata sets can take longer to load.

`showDefaultSuggestions=false` suppresses blank-query loading/display for List/Field/View/ContentType; typing still loads Field/View/ContentType's complete catalog once if needed, then filters locally. `initialOptions` remains selection-model seeds. Site/Item/Principal retain the five ordinary suggestion policy; Site also has the fixed contextual choices described below.
```tsx
<ListPicker label="Lists" initialCatalogLimit={30} />;
<FieldPicker label="Fields" list="Projects" multiselect />;
<ViewPicker label="Views" list="Projects" />;
```

Site/List/Field/View/ContentType/Item typed search is case insensitive: Site uses SharePoint's KQL semantics; hub catalogs and the other metadata/item sources use deterministic Unicode lowercase matching. This does not promise accent removal or transliteration. Query cache/cursors treat case-only edits as equivalent while preserving the displayed input and accumulated pages. IDs, raw consumer filters/KQL and ordering are not case-folded. Site/List/Item typed search continues through the configured source with bounded paging; Field/View/ContentType search filters the complete memory-cached catalog.

Multiple Dropdown menus always include the localized **Selected items** group, independent of the typed query and available cap. Checked rows, including unresolved fallback labels, can be clicked to remove them without searching. They remain excluded from available/search results. Single Dropdown menus show the localized **Selected item** group first, with a checked Persona row (or unresolved fallback label), independent of query and suggestion limits. This identity is excluded from available/search rows. Clicking it preserves the value without validation, resolution requests or change events; choosing another replaces it. Typing changes only the popup query and preserves the selection button text. The native clearable action still clears it. TagPicker retains removable tags. Both menus scroll within a 320px/80vh maximum height. Read-only hydration and scope checks show loading while pending, without a transient unresolved error. Definitive rejected/missing values remain diagnosed, even while another selection is pending; hydration emits no selection change.

`onOptionSelect(event, data)` runs only for accepted explicit `selection`, `remove`, `clear` and `resolve` actions. `PickerSelectionEvent` is the native Fluent event; `PickerSelectionData` contains ordered selectedOptions, resolved selectedItems/selectedValues, unresolvedOptions and the optional affected key/text. Hydration and prop updates emit no change event. Resolved React selectedValues may be a subset when unresolved keys exist: consumers must not replace persisted references with that subset. The native Property Pane bridge retains the complete reference array.

`validateSelection(items)` receives the proposed resolved models, excluding unresolved aliases. Returning `undefined` or an empty string accepts the validated action; a nonempty message vetoes it, preserving selection and emitting no `onOptionSelect`. Promises are awaited. A throw or rejection uses the failure UI and `onError`, without committing: selection/remove/clear and blur report operation `validate`, while explicit resolution reports `resolve`. Omission means no custom validation. External `validationMessage` has display precedence only; it does not override the validator's acceptance decision.

`validationTrigger` defaults to `change`, validating selection, removal and clear before committing. With `blur`, those gestures commit without custom change validation; focus leaving the whole control and its owned popup validates the retained models when no confirmation is pending, without revoking an existing selection. Explicit ID resolution, when available, always validates before committing, including in blur mode. Consumer Promises are not cancelled.

`onError(error, details)` identifies transport, operation and optional personal suggestion source. Empty responses, partial sources and failures are distinct. Source failures remain visible beside healthy rows with a retry action. Bounded-result warnings remain visible beside rows and ask for narrower criteria instead of offering retry; failed requests are retried only after an explicit retry/refresh, cache eviction or a new lifetime.

`WebPickerSource`, `SearchPickerOptions`, `PickerOrderBy` and `PickerInitialSuggestions` are the shared source/query/policy types. Native and React counterparts use the same defaults below.

## SitePicker

`SitePickerProps` / `SitePickerOption` / `SitePickerValue`: Dropdown, single selection by default. Mode defaults to `site`; `site` discovers collections, `hub` registered hubs, and `associatedSites` sites associated with a supplied hubSiteId or the current collection's hub. No association yields an empty result. Discovery is not restricted to the host collection; the configured mode and KQL filters determine candidates, subject to SharePoint permissions. `additionalQuery` is a KQL AND clause allowed only in site mode. Order defaults to Title ascending; trimDuplicates defaults to true. `searchPlaceholder` overrides placeholder.

`includeCurrentSite` and `includeHubSite` are flat boolean props, both **true by default** (also in Property Pane). The menu keeps **Sito corrente** and **Hub associato** visible before and during typed search, after the selected section. These fixed contextual choices sit outside the cap of five ordinary initial suggestions, so a full consumer set cannot hide them. Ordinary consumer suggestions follow the contextual choices. The current site is the collection root, including from subweb hosts. The hub is read from the current collection's actual `HubSiteId`/`IsHubSite` metadata and resolved via `HubSites/GetById`; the configured `hubSiteId` only controls associatedSites discovery and does not substitute the current collection's real association. A hub with no actual association does not suggest itself as an associated hub. Canonical duplicates (including current=hub and typed results) appear once, consumer models win, and selected identities stay in their checked selected section (or selected tags).

Hub metadata follows the [SharePoint hub REST API](https://learn.microsoft.com/en-us/sharepoint/dev/features/hub-site/hub-site-rest-api), including [GetById](https://learn.microsoft.com/en-us/sharepoint/dev/features/hub-site/rest-getbyid-method). The optional sources load lazily when the menu opens, including with typed text, using scoped memory-cached read-only REST metadata. False disables that source; both false restore consumer-only defaults; `showDefaultSuggestions=false` suppresses both sources. No association is an empty source. Failures are independent: healthy current/hub/consumer rows remain; errors identify `suggestionSource: 'currentSite' | 'hubSite'`, `source: 'sharePoint'`, `operation: 'load'` and support explicit retry/refresh. Cosmetic prop changes and equivalent flag values reuse cached requests. Search and hub catalog cache entries also survive these display-policy switches.

Plain site-mode defaults need no text Search. Other modes and `additionalQuery` still perform their normal source eligibility checks, which can use Search; explicit selection verification retains the configured source contract. Typed search continues using SharePoint Search (hub mode filters its hub catalog), independently of initial suggestions. No storage, provisioning or permission grants occur.

```tsx
<SitePicker label="Site" includeCurrentSite includeHubSite />;
<SitePicker label="Site" includeCurrentSite={false} includeHubSite={false} />;
```

Site/association uses the toolkit Search service without SortList: Title and Path are not sortable in the standard SharePoint schema. It reads at most five native pages of 200 results, caches and sorts the entire retrieved set by title or URL, then pages that ordered set locally. Continuations are bound to query, owner/client and page size. Hub mode uses the same five-page bound and local ordering/paging for its catalog. If native results remain beyond 1000, the control reports partial data and recommends narrowing the query/scope; ordering then covers only the retrieved set. URLs establish site collection identity even when the endpoint omits siteId.

```tsx
<SitePicker label="Associated sites" mode="associatedSites" multiselect
  orderBy={{ field: 'title', direction: 'asc' }} />
```

## ListPicker

`ListPickerProps` / `ListPickerOption` / `ListPickerValue`: Dropdown, single, current web, Title ascending. includeHidden defaults to true. baseTemplate accepts one numeric template or an OR set; `filter` overrides these convenience filters. contentTypeId matches expanded content-type ID prefixes case-insensitively. Typed search matches Title prefixes case-insensitively across bounded native REST pages. Native continuations retain filtering, projection and ordering; they cannot be reused against another query/client/source.

```tsx
<ListPicker label="Libraries" baseTemplate={101} includeHidden={false} />
```

## FieldPicker

`FieldPickerProps` / `FieldPickerOption` / `FieldPickerValue`: Dropdown, multiple by default. Omitted list queries site columns; a supplied exact title or `SPFxPnPListSelector` selects list fields. String GUIDs are titles, not implicitly converted IDs. includeHidden/includeReadOnly default true; group is exact. Raw filter overrides convenience filters. Field-specific icons distinguish text, choice, number, date, person, lookup and other kinds; metadata appears once as the secondary line. showBlankOption defaults false and enables clearing to `[]`.

```tsx
<FieldPicker label="Fields" list="Projects" includeReadOnly={false}
  presentation="tagPicker" orderBy={{ field: 'internalName' }} />
```

## ViewPicker

`ViewPickerProps` / `ViewPickerOption` / `ViewPickerValue`: Dropdown, single. list is required; source web defaults current, Title ascending. Excluded titles are exact case-sensitive matches. DefaultView and PersonalView metadata distinguish public/personal views. showBlankOption means an explicit clear action, never a fake view GUID.

```tsx
<ViewPicker label="View" list={{ kind: 'title', title: 'Projects' }} excludedViewTitles={['Archive']} />
```

## ListItemPicker

`ListItemPickerProps` / `ListItemPickerOption` / `ListItemPickerValue`: TagPicker, multiple. list and internal textField are required. keyField defaults `Id`, searchMode defaults startsWith, order defaults Id ascending, includeFolders defaults false. Existing toolkit selectors also support decoded list paths/URLs. Consumer keys remain opaque; values always retain actual itemId, even with a custom key.

For file candidates, typed search matches either the configured display field or `File.Name`, using the same case insensitive `startsWith`/`contains` mode. A nonempty display field remains the label; when empty (including whitespace), the filename is used during discovery, selection and hydration. Ordinary list records and folders retain their existing labels and search behavior. React and Property Pane use the same source.

Text, Note and calculated text (verified OutputType) match the REST EntityPropertyName locally with case-insensitive startsWith/contains; other field types match their FieldValuesAsText projection. SharePoint REST does not support `tolower`/`toupper`, and CAML does not establish the required casing contract, so generated string predicates are not sent. Raw consumer filters, folder policy and ordering remain server-side. Each search page scans at most five native REST pages (up to 1000 candidates at page size 200); a continuation and partial status allow further loading. Incomplete scans are never presented as complete. Large or threshold-limited lists may require narrower indexed filters. A custom key must have unique finite-number/string values; duplicates or an unfinished uniqueness scan leave the value unresolved. Numeric `1` differs from string `'1'`.

A library row becomes `file` or `folder` only with verified FSObjType and File/Folder UniqueId, otherwise ordinary list rows remain `item`. Icons distinguish records, folders, generic files and known extensions such as PDF. Files use the verified filename when their configured display value is absent, empty or whitespace. For ordinary records and folders, only an absent/null display value uses the Item ID fallback; defined empty or whitespace strings retain their existing value.

```tsx
<ListItemPicker label="Projects" list="Projects" textField="Title" keyField="ExternalCode" searchMode="contains" />
```

## ContentTypePicker

`ContentTypePickerProps`, `ContentTypePickerOption` and `ContentTypePickerValue` select existing SharePoint content types. Default: `scope="site"`, Dropdown, single selection, `includeHidden=true`. The complete catalog is loaded with native paging and sorted by name/ID; there is no initial 30-item cap. Typed search filters names locally by prefix, case insensitively. All standard picker selection, model/consumer-option, accessibility and visual slot contracts above apply.

```tsx
import { ContentTypePicker } from '@apvee/spfx-react-toolkit/controls/pickers';

<ContentTypePicker label="Site content type" scope="site" />;
<ContentTypePicker label="List content types" scope="list"
  list={{ kind: 'id', id: '11111111-1111-4111-8111-111111111111' }}
  presentation="tagPicker" multiselect includeHidden={false} />;
```

Site scope resolves the owning **site collection root web**, including when `webUrl` identifies a subweb; it does not enumerate only that subweb's definitions. List scope requires the existing toolkit `list` selector and returns only content types already associated with that list/library. The same content-type ID in these two scopes remains a distinct reference. The control does not create, publish, associate or synchronize content types, and does not enable content types on a list.

| Prop | Behavior |
| --- | --- |
| `scope` | `site` by default; `list` requires `list`. A list selector is not valid in site scope. |
| `webUrl` | Current provider web by default; resolves its collection root in site scope, owns the list in list scope. |
| `group` | Optional exact case insensitive group filter. |
| `includeHidden` | Defaults true; false excludes Hidden content types. ReadOnly/Sealed types remain selectable. |
| `contentTypeId` | Optional exact hierarchical SharePoint content-type ID filter. It is not a GUID. |
| `includeDescendants` | Defaults false. True includes that ID and its descendants; requires `contentTypeId`. |

Persona rows show name and group, with an icon; selected tags show name/icon. Description and Hidden/ReadOnly/Sealed metadata belong to the option model. The serializable reference contains `kind:'contentType'`, `scope`, `webUrl`, `contentTypeId`, optional snapshot `text`, and `listId` only in list scope. Site references store the verified collection root web URL. Identity includes scope, source URL, list and the complete ID; names and parent-ID prefixes never establish equality.

The source reuses `usePickerSourceContext` (and its toolkit `useSPFxSPHttpClient` client), `cachedPickerSpRead`, `resolvePickerListTarget` for lists, and `createMetadataPickerSource`. Rendering and selection reuse `usePickerRuntime`/`PickerRenderer`. There is no new HTTP client, Graph dependency or public content-type service. REST failures remain errors rather than empty catalogs; valid mounted catalogs are reused, refresh/source changes may reload, and native pane disposal releases its cache as described below.

## PrincipalPicker

`PrincipalPickerProps` / `PrincipalPickerOption` combines people and groups. `PrincipalType` supports user, sharePointGroup, microsoft365Group, securityGroup and distributionList. Default is `['user']`; an empty array disables suggestions. Default presentation is TagPicker with multiple selection. User rows use avatar/name/email once; group rows use category icon/name/type and origin. Tags show only name plus avatar/icon. Microsoft 365 classification takes precedence over security flags; mail-enabled security groups remain securityGroup, and non-security mail groups are distributionList.

`PrincipalUserScope` defaults to tenant; web scope queries registered principals in the target collection's UserInfoList and does **not** establish effective per-web permissions. Group scopes accept `PickerGroupReference` (numeric SharePoint group ID scoped to web, or Entra GUID) or exact SharePoint group name. These scopes restrict users; selectable groups retain their own source scope. Entra membership checks may include transitive relationships; SharePoint group membership is direct. Authorization services for the current user are not substituted for candidate membership checks.

`defaultPrincipals` accepts once-only `PrincipalPickerInitialPrincipal` login, email, Entra user ID, group reference or displayOnly. It cannot coexist with uncontrolled defaultSelectedOptions. displayOnly remains unresolved rather than pretending that a title identifies a user.

### User identities and ensureUser

For native user candidates without IDs, verification searches the UPN without the SharePoint membership claim wrapper, then checks the exact canonical login and configured scope. The saved login retains its original claim format; verification never provisions a user.

`UserPrincipal` / `UserPrincipalValue` may carry both ID systems. `UserIdentityResolution` defaults to `{mode:'available'}`:

| Mode | Required on new confirmation |
| --- | --- |
| available | Returned verified login and any available IDs |
| entra | Verified `{tenantId, id}` directory user identity |
| sharePoint | Verified `{siteCollectionUrl, id}` user registration |
| both | Both independent scoped identities |

The SharePoint destination defaults to userIdentity.webUrl, then PrincipalPicker.webUrl, then current web. `ensureUser` defaults false and exists only in sharePoint/both mode. False reads an existing registration; missing registration blocks confirmation. True may POST `_api/web/ensureuser` only during explicit selection or “Risolvi ID”. Search, personal suggestions, hydration, seed/config updates and preflight never provision a user. Required ID failure preserves the previous selection. A sent server write may finish after cancellation; the UI ignores stale completion but cannot promise server rollback. Changing modes preserves existing values and diagnoses missing requested IDs until explicit resolution. The resolution action appears only for users missing the requested tenant/destination-collection IDs; an ID from a previous collection does not satisfy a new destination. Enrichment cannot rebind a retained alias to conflicting verified IDs; changing collection registration requires matching verified Entra identity.

allowUnvalidatedEmail defaults false, requires user enabled and available mode, and produces an explicit UnvalidatedEmailValue rather than a user identity. `useSubstrateSearch` defaults false and configures native Microsoft 365 Substrate user search. `showHiddenPrincipals` is retained as a deprecated compatibility flag with no effect: SharePoint determines native user visibility. The unsupported ShowHiddenPrincipals payload member caused HTTP400 and is not sent. There is no verified equivalent visibility override in the [native query contract](https://learn.microsoft.com/en-us/dotnet/api/microsoft.sharepoint.applicationpages.clientpickerquery.clientpeoplepickerqueryparameters?view=sharepoint-csom). The [PnP search service](https://github.com/pnp/sp-dev-fx-controls-react/blob/master/src/services/PeopleSearchService.ts) confirms top-level UseSubstrateSearch and the modern value response envelope; the adapter also accepts the legacy named envelope.

### Personal suggestions, permissions and endpoint limits

`PrincipalInitialSuggestions` extends the shared limit with the independent current-user option. loadPersonalSuggestions defaults false; initialSuggestions.includeCurrentUser defaults false independently. Opening an empty-query menu loads only enabled sources. The order is consumer options, current user (“Tu”), “Persone suggerite”, “I tuoi gruppi”, with **five overall**, excluding selected identities. These are not selection history. Every source obeys type and user scope; roles/admin units and incomplete group classifications are excluded. Partial/failing personal sources do not remove healthy consumer options.

| REST operation | Delegated permission / limitation |
| --- | --- |
| GET `/me/people` | People.Read; relevance-ranked person objects. **Maintenance mode**, intentionally retained by requirement. Not available in China operated by 21Vianet. Only verified OrganizationUser references are promoted; person.id is never an Entra user ID. |
| GET `/me/memberOf` | User.Read for the signed-in user's direct memberships; non-transitive. Only `#microsoft.graph.group` objects are candidates. Limited-information objects may require group read permission for classification. |
| GET `/users/{id\|UPN}`, user lookup | User.ReadBasic.All for basic directory-user data; results are checked against supplied identity. Ambiguous email results are rejected. |
| GET `/groups`, group metadata/members and candidate membership checks | GroupMember.Read.All / User.Read.All as required by the configured endpoints; higher permissions or Member.Read.Hidden may be needed for protected/hidden membership. Scope and tenant policy can still deny access. |
| SharePoint REST reads/search; ensureuser | Host user's SharePoint access and tenant policies. An API being reachable does not establish that a candidate has effective permissions. |

The Graph SDK owns authentication and retry. `/me` uses delegated context; application-only execution is not supported by this UI. Personal retrieval stops after five pages of 50 and twenty fresh identity/scope lookups per open. The result is partial if the budget or a source is exhausted; selected exclusion and final cap are applied after source collection. Cache hits do not spend fresh-lookup budget. No grants are performed by the picker or sample.

Verified against [List people](https://learn.microsoft.com/en-us/graph/api/user-list-people?view=graph-rest-1.0), [List direct memberships](https://learn.microsoft.com/en-us/graph/api/user-list-memberof?view=graph-rest-1.0), [Get user](https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0), [List groups](https://learn.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0) and [Check member groups](https://learn.microsoft.com/en-us/graph/api/directoryobject-checkmembergroups?view=graph-rest-1.0) on 2026-10-14. Host/tenant validation remains separate from local mocked-transport tests.

## JSON values, storage and lifetime

`PickerValueLabel` is an optional stale display snapshot. SitePickerValue is collection URL plus optional SiteId GUID; ListPickerValue/FieldPickerValue/ViewPickerValue carry web and scoped entity GUIDs. `ListItemPickerValueBase`, `PickerCustomItemKey`, SharePointItemValue/SharePointFileValue/SharePointFolderValue retain item identity and explicit kind. `EntraObjectIdentity` and `SharePointPrincipalIdentity` scope native IDs. SharePointGroupValue, EntraGroupValue and UnvalidatedEmailValue cannot masquerade as UserPrincipalValue. `PickerValue` is their union; `PickerSelection<TValue>` is always an array for zero/single/multiple selection.

```ts
import type { SitePickerValue, PrincipalPickerValue, PickerSelection } from '@apvee/spfx-react-toolkit/controls/pickers';
interface Properties {
  sites: PickerSelection<SitePickerValue>;
  principals: PrincipalPickerValue[];
}
const principals: PrincipalPickerValue[] = [{ kind: 'user', loginName: 'alice@tenant.com',
  entra: { tenantId: '11111111-1111-4111-8111-111111111111', id: '22222222-2222-4222-8222-222222222222' },
  sharePoint: { siteCollectionUrl: 'https://tenant.sharepoint.com/sites/hr', id: 42 } }];
```

Values contain plain own JSON members only. Accessors, extra properties, prototypes, invalid URLs, non-finite numbers and invalid domain IDs are rejected at the boundary. Models may carry display fields, but they are not saved automatically. The consumer owns all persistence. There is no selection history, localStorage, sessionStorage or OneDrive writing. The PnP client is configured without persistent cache.

Requests use an owner/client/tenant/user/source-scoped memory cache: 128 settled entries, five-minute successful-result TTL, separate shared in-flight requests, and explicit retry for cached failures. TTL is checked on use, not by background refetch. Changing display props, callbacks or equivalent option arrays does not initialize a new client or repeat equivalent requests while their cache entries remain valid. Identical metadata GETs share their exact URI/projection contract across discovery, hydration and verification; different filters/projections remain separate, and POST/ensureUser is never routed through that read cache.

Field/View/ContentType verification may reuse a positive canonical match from an already cached or loading, complete, error-free catalog in the same configured scope. Verification does not start a complete catalog by itself; absent/expired/replaced/failed/partial catalogs and missing candidates fall back to the existing filtered entity read. Refresh invalidates dependent reads and an old completion cannot repopulate the new cache generation.

Within a mounted control, reopening a settled typed menu preserves accumulated result pages without a loading flash when the exact first-page cache entry is still valid. A cache entry replaced by another field is adopted rather than retaining stale pages. Source/query/page-size/refresh changes invalidate obsolete continuations. Slow initial consumer eligibility and optional contextual/personal retrieval show loading instead of a premature empty/minimum prompt; settled cosmetic/equivalent updates reuse their results. A real unmount releases an owned cache; disposal/recreation or an expired/evicted entry can query again. Independent React controls have their own runtime cache unless an internal host shares one; native fields share the owner/document lease described in the Property Pane guide. Graph initialization across independent roots is not promised to coalesce. Graph-dependent Principal operations await the current provider’s client initialization, preserving pending labels and explicit ID-resolution actions. Terminal initialization failures remain reported; SharePoint-only operations remain usable when optional Graph initialization fails. Replacing the provider cancels obsolete pending work.

## Slots and verification

suggestionsHeader is a fixed inner span; emptyContent/loadingContent/errorContent are fixed inner div content slots using the toolkit private Fluent resolver and class composer. Their status shells remain owned by the control; false suppresses content. Native Field and Tooltip slots retain Fluent contracts. Structural suggestion options, Persona avatars and selection callbacks are not generalized content slots. Single Dropdown uses the native clearable action, routed through the same validated clear event; there is no separate “Svuota selezione” button. Multiple Dropdown selection removes checked rows individually; TagPicker dismisses tags. Field/View showBlankOption retains its optional explicit localized “No selection” action. Single Dropdown keeps the native checkIcon for its selected row and hides the column on unselected candidates using the peer's required null slot sentinel; multi keeps native checkboxes. Repository styles use Griffel classes.

The sample **SharePoint pickers** panel mounts one live variant with configurable presentation, multi/search, consumer model updates, principal types/IDs and collapsed JSON diagnostics. Its ContentTypePicker example offers site/list scope, group/visibility filters and an ID preset selector (Item, Document or confirmed selected IDs), with an optional descendants toggle. Selecting All content types resets that toggle. The actual Property Pane mounts all seven native fields. Its initially expanded **Picker modes** group globally selects Dropdown/TagPicker and single/multiple selection for all seven fields. Defaults are Dropdown and multiple selection, preserving existing multiple values. Presentation changes preserve selections; changing multiple to single keeps the first value. Cascading sources require exactly one selected site; list-dependent fields also require exactly one list. ContentTypePicker requires a list only in list scope. Site changes clear its value in both scopes; list changes clear it only in list scope. Versioned browser fixtures provide the larger simulation matrix; mocked REST is labelled simulation and is not tenant proof. See [native fields](../property-pane/pickers.md) and [SharePoint validation](../../development/sharepoint-validation/INDEX.md).

## Member reference

Source contracts below supplement the examples. Inherited Fluent Field members retain native contracts.

| Member | Contract |
| --- | --- |
| `PickerOption.key` | Required stable renderer alias. It must never be reused for a different scoped entity. |
| `PickerOption.text` | Required plain display label; not an authoritative entity identifier. |
| `PickerOption.value` | Minimal JSON reference. Must agree with the option identity; other metadata is never persisted automatically. |
| `PickerSelectionData.reason` | Explicit user action; hydration/model updates emit no selection event. |
| `PickerSelectionData.selectedOptions` | Complete ordered renderer aliases after the explicit action, including unresolved selections. |
| `PickerSelectionData.selectedItems` | Resolved models in selection order; unresolved aliases have no model in this array. |
| `PickerSelectionData.selectedValues` | Serializable references of resolved models; unresolvedOptions remain explicit and are not silently discarded. |
| `PickerSelectionData.unresolvedOptions` | Selected renderer aliases whose models could not be resolved; retain them until an explicit removal. |
| `PickerSelectionData.optionValue` | Alias affected by a selection or removal; omitted for whole-selection actions. |
| `PickerSelectionData.optionText` | Display text for the affected resolved model, when available. |
| `PickerProps.presentation` | Presentation only; changing this retains selection, data and request runtime. Default varies by picker. Suggestions use Persona with name/metadata and category icon; user principals retain their user avatar. Selected tags show name and category icon (user avatar for user principals). |
| `PickerProps.colorfulTagAvatars` | Optional selected-tag avatar palette; default false. Uses stable renderer keys; Dropdown and suggestion avatars are unchanged. |
| `PickerProps.searchable` | Enable popup SearchBox in Dropdown or the TagPicker input. Defaults true. |
| `PickerProps.selectedOptions` | Controlled stable renderer keys (including retained aliases). Empty selection is []; undefined selects uncontrolled mode. |
| `PickerProps.defaultSelectedOptions` | Initial keys, read once per mount; changing them does not reset selection. |
| `PickerProps.initialOptions` | Models for displaying/resolving selected keys; does not add pre-search menu suggestions. Equivalent seeds avoid hydration; incoming models may enrich selected snapshots without changing keys. |
| `PickerProps.availableOptions` | Reactive pre-search candidates fetched from any consumer source. Defaults to []. Does not initialize selection. Updates recompose the menu without remount, selection reset or remote refetch. Models override equivalent personal/search models; removed selected models remain available to the selection. Canonical identities include tenant/site/list scope; labels and email alone are not canonical IDs. Consumer owns persistence and retrieval errors. All candidates still undergo type/source/scope filters. |
| `PickerProps.initialSuggestions` | Maximum pre-search suggestions for Site/Item/Principal. Default 5; List uses initialCatalogLimit; Field/View/ContentType show their complete catalogs. No automatic history or persistence. |
| `PickerProps.multiselect` | Matches Fluent naming. False limits new user selection to one key; selecting another replaces it. Changing to false with multiple controlled keys reports a validation conflict; existing keys are never silently dropped. |
| `PickerProps.selectionLimit` | Optional positive integer cap for new selections. Omission is unlimited in multi mode and one in single mode; conflicting existing selections remain visible. |
| `PickerProps.onOptionSelect` | Called after an accepted explicit selection, removal, clear or ID resolution. Arguments are the native Fluent event and complete selection data. Not called by hydration or prop updates; return value is ignored. |
| `PickerProps.filterOptions` | Local suggestion/search transformation; changing it recomputes the projection, without server refetch. Mandatory type/scope filters, selected exclusion, deduplication and the pre-search limit apply again afterward. |
| `PickerProps.size` | Fluent selection control size; defaults to medium. |
| `PickerProps.disabled` | Disables the control and explicit selection/resolution actions. Defaults to false. |
| `PickerProps.strings` | Optional `PickersStringsOverride`, merged by known nested group and leaf with native resources. Direct control text/slots win. |
| `PickerProps.placeholder` | Search placeholder; omitted uses “Type at least N characters to search” with the configured minimum when searchable. Empty Dropdown selection uses a separate fixed invitation. |
| `PickerProps.className` | Optional CSS class on the outer field region. |
| `PickerProps.controlClassName` | Optional CSS class on the native Dropdown or TagPicker root. |
| `PickerProps.id` | Optional stable selection trigger/input DOM ID. Omission uses a generated Fluent Field association; this is not an entity ID. |
| `PickerProps.'aria-label'` | Explicit selection trigger/input name; required with JSX/object labels unless aria-labelledby is supplied. Takes precedence over a plain string label. |
| `PickerProps.'aria-labelledby'` | Explicit selection trigger/input naming IDs; supply this or aria-label when no nonblank plain string label is available. Forwarded to the Dropdown button or TagPicker input. |
| `PickerProps.suggestionsHeader` | Literal presentation slots, fixed inner span/div; the control owns status/semantic shells. |
| `PickerProps.emptyContent` | Literal div content for explicit custom initial/below-minimum guidance and complete empty/no-other-result states; omission shows the current-query default described under Search feedback. False hides only presentation, preserving the status shell. |
| `PickerProps.loadingContent` | Literal div content during initial loading or eligible typed-query debounce/loading. Omission uses the native loading or searching message for that state. False hides its content. |
| `PickerProps.errorContent` | Literal div content for request or identity failures; omission shows the error message. Query and contextual-source failures remain independently visible. False hides content; retry remains an independent action. |
| `PickerProps.tooltip` | Optional native Fluent Tooltip content, relationship and positioning. Omission renders no tooltip; it does not replace the accessible input label. |
| `PickerProps.refreshKey` | Explicit invalidation token; unchanged token never requests an automatic refresh. |
| `PickerProps.onError` | Called with the Error and transport/operation metadata for failed asynchronous work; suggestionSource identifies a personal or contextual Site source. Return value is ignored; consumer options remain independent. |
| `PickerProps.validateSelection` | Validates resolved models; undefined/empty string accepts, a nonempty message vetoes without selection change or callback. Promises are awaited; throws/rejections use the failure channels. validationMessage overrides display only. |
| `PickerProps.validationTrigger` | change (default) validates selection/remove/clear before commit. blur validates retained models after focus leaves the whole control and owned popup when no confirmation is pending. Explicit ID resolution always validates, including in blur mode. |
| `PickerInitialSuggestions.limit` | Total pre-search suggestions, integer 1–5, default 5. Search results use suggestionsLimit instead. |
| `WebPickerSource.webUrl` | Absolute HTTP(S) source web URL. Omission, when allowed, uses the provider current web; never inferred from a display label. |
| `SearchPickerOptions.minimumSearchLength` | Defaults to 2 for all pickers; shorter typed inputs never initiate typed search; minimum-character guidance appears in the search placeholder without a duplicate default status row. Values below 2 are invalid. |
| `SearchPickerOptions.debounceMs` | Debounce for typed searches, including local Field/View/ContentType matching; default 300 ms. |
| `SearchPickerOptions.suggestionsLimit` | Remote search result page size; independent of the pre-search suggestion or catalog limit. Field/View/ContentType return all local matches and ignore this paging size. |
| `SearchPickerOptions.showDefaultSuggestions` | Shows pre-search consumer/personal sections and the List/Field/View/ContentType catalog. Defaults to true; false suppresses their lazy eligibility/retrieval, without disabling typed search. |
| `PickerOrderBy.field` | Required server sort field supported by this picker. |
| `PickerOrderBy.direction` | Optional server sort direction; defaults to asc. |
| `SitePickerOption.url` | Required normalized absolute site collection URL, agreeing with value.url. |
| `SitePickerOption.siteId` | Optional verified site collection GUID; omitted when the discovery endpoint does not return it. |
| `SitePickerOption.hubSiteId` | Optional verified hub association GUID from source metadata; omission means metadata was not returned. |
| `SitePickerProps.includeCurrentSite` | Keep the host collection root visible before/during search. Default true; outside the ordinary cap, subject to mode/KQL, consumer model priority and selected exclusion. |
| `SitePickerProps.includeHubSite` | Keep the actually associated hub visible before/during search. Default true; no association means no hub row, including when the host is itself a hub. Independent of configured hubSiteId; outside the ordinary cap, subject to source scope and selected exclusion. |
| `SitePickerProps.mode` | Site discovery mode. Defaults to site; site searches collections, hub lists registered hubs, associatedSites uses the configured or current hub. |
| `SitePickerProps.hubSiteId` | Optional hub GUID for associatedSites. Omission reads the current collection association; no association yields an empty result. |
| `SitePickerProps.searchPlaceholder` | Site query input placeholder; overrides the shared placeholder when supplied. |
| `SitePickerProps.orderBy` | Local ordering of the retrieved site collection set, default Title ascending. Reads at most 1000 Search/hub results before paging; reports partial data if more remain. |
| `SitePickerProps.trimDuplicates` | SharePoint search duplicate trimming; defaults to true. Only applies to site/associatedSites search. |
| `SitePickerProps.additionalQuery` | Additional KQL AND clause, only for site mode. |
| `ListPickerOption.webUrl` | Absolute source web URL; supplied options must match the effective query web. |
| `ListPickerOption.id` | Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. |
| `ListPickerOption.title` | Required plain metadata title from the source; a consumer display alias may use a different text. |
| `ListPickerOption.rootFolderUrl` | Optional server-relative root-folder navigation snapshot, never canonical identity. |
| `ListPickerOption.baseTemplate` | Required numeric SharePoint BaseTemplate metadata for this list; not a query filter. |
| `ListPickerOption.hidden` | Required hidden metadata flag used by convenience filters. |
| `ListPickerProps.initialCatalogLimit` | Maximum scoped initial available options, default 30, integer 1–200. Selected rows are separate; showDefaultSuggestions=false suppresses loading. |
| `ListPickerProps.baseTemplate` | Optional numeric list template or template set (OR). Omission allows all; a raw filter overrides this convenience filter. |
| `ListPickerProps.includeHidden` | Includes hidden lists/fields. Defaults to true; ignored when a raw filter overrides convenience filters. |
| `ListPickerProps.filter` | OData expression overrides includeHidden/baseTemplate, as in PnP. |
| `ListPickerProps.contentTypeId` | Optional case-insensitive content-type ID prefix. Lists without a matching expanded content type are excluded. |
| `ListPickerProps.orderBy` | Optional server ordering. Metadata defaults to Title ascending; item search defaults to Id ascending. Continuations preserve ordering. |
| `FieldPickerOption.webUrl` | Absolute source web; listId is absent only for site columns. |
| `FieldPickerOption.listId` | Scoped list GUID, agreeing with the JSON reference. For site columns this is omitted. |
| `FieldPickerOption.id` | Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. |
| `FieldPickerOption.internalName` | Required field internal name, agreeing with the JSON reference; distinct from a REST EntityPropertyName. |
| `FieldPickerOption.title` | Required plain metadata title from the source; a consumer display alias may use a different text. |
| `FieldPickerOption.typeAsString` | Required SharePoint field type name; determines the field-specific icon and secondary label. |
| `FieldPickerOption.fieldTypeKind` | Required numeric SharePoint field type discriminator. |
| `FieldPickerOption.hidden` | Required hidden metadata flag used by convenience filters. |
| `FieldPickerOption.readOnly` | Required ReadOnlyField metadata flag. |
| `FieldPickerOption.group` | Required source field group metadata used by the exact group filter. |
| `FieldPickerProps.list` | Existing toolkit selector; omitted queries site columns. A string is an exact title. |
| `FieldPickerProps.includeHidden` | Includes hidden lists/fields. Defaults to true; ignored when a raw filter overrides convenience filters. |
| `FieldPickerProps.includeReadOnly` | Includes read-only fields. Defaults to true; a raw filter overrides this convenience flag. |
| `FieldPickerProps.group` | Optional exact field group filter; omission accepts every group. Raw filter takes precedence. |
| `FieldPickerProps.filter` | Overrides includeHidden/includeReadOnly/group, as in PnP. |
| `FieldPickerProps.orderBy` | Optional server ordering. Metadata defaults to Title ascending; item search defaults to Id ascending. Continuations preserve ordering. |
| `FieldPickerProps.showBlankOption` | Shows the explicit clear action. Defaults to false; clearing emits an empty array, never a synthetic domain reference. |
| `ViewPickerOption.webUrl` | Effective source web/list identity, required for validating supplied candidates. |
| `ViewPickerOption.listId` | Required source list GUID, agreeing with the JSON reference; views always belong to a list. |
| `ViewPickerOption.id` | Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. |
| `ViewPickerOption.title` | Required plain metadata title from the source; a consumer display alias may use a different text. |
| `ViewPickerOption.defaultView` | Required DefaultView metadata flag. |
| `ViewPickerOption.personalView` | Required PersonalView metadata flag. |
| `ViewPickerOption.serverRelativeUrl` | Optional server-relative navigation snapshot. Not canonical identity; must have one leading slash and no query, fragment or traversal segments. |
| `ViewPickerProps.list` | Required: there is no meaningful view query without a list. |
| `ViewPickerProps.filter` | Optional SharePoint OData filter ANDed with typed search and exact candidate identity. Field/list convenience filters are overridden; scope restrictions still apply. |
| `ViewPickerProps.orderBy` | Optional server ordering. Metadata defaults to Title ascending; item search defaults to Id ascending. Continuations preserve ordering. |
| `ViewPickerProps.excludedViewTitles` | Optional exact case-sensitive view titles to omit from suggestions and candidate verification. Omission excludes none. |
| `ViewPickerProps.showBlankOption` | Shows the explicit clear action. Defaults to false; clearing emits an empty array, never a synthetic domain reference. |
| `ListItemPickerOption.kind` | Verified entity kind, matching value.kind. Library origin alone does not identify a file. |
| `ListItemPickerOption.fileExtension` | Optional file extension for file-specific presentation, not inferred identity. |
| `ListItemPickerOption.webUrl` | Effective source web/list identity, required for validating supplied candidates. |
| `ListItemPickerOption.listId` | Required source list GUID for the item, file or folder, agreeing with the JSON reference. |
| `ListItemPickerOption.itemId` | Required positive safe integer SharePoint item ID, scoped to the web/list. Always retained even with a custom key. |
| `ListItemPickerOption.keyField` | Effective key field, Id when omitted; validated against value.customKey or itemId. |
| `ListItemPickerOption.keyValue` | Required typed custom-key value, or numeric itemId for the default Id key. Must agree with value.customKey when present. |
| `ListItemPickerOption.fields` | Required read-only source field snapshot for consumer display logic. Never included in serialized picker values automatically. |
| `ListItemPickerProps.list` | Required exact list title or existing toolkit selector for item queries. String GUIDs are still titles. |
| `ListItemPickerProps.includeFolders` | Include folder candidates. Default false; folders stay distinguishable from files and list records. |
| `ListItemPickerProps.textField` | Required internal display/search field. Files also match their filename and use it for absent, empty or whitespace display values. Other missing values use an Item ID fallback; defined empty/whitespace record or folder values are preserved. Field metadata determines safe query handling. |
| `ListItemPickerProps.keyField` | Defaults to Id; custom fields must contain unique string/number values. |
| `ListItemPickerProps.filter` | Optional SharePoint OData filter ANDed with typed search and exact candidate identity. Field/list convenience filters are overridden; scope restrictions still apply. |
| `ListItemPickerProps.orderBy` | Optional server ordering. Metadata defaults to Title ascending; item search defaults to Id ascending. Continuations preserve ordering. |
| `ListItemPickerProps.searchMode` | Item text match mode; defaults to startsWith. Both modes compare display text case-insensitively over bounded REST pages; non-text fields use FieldValuesAsText. |
| `PickerGroupReference.source` | Required source or transport discriminator; SharePoint identities and Entra identities remain distinct. |
| `PickerGroupReference.webUrl` | Absolute HTTP(S) source web URL. Omission, when allowed, uses the provider current web; never inferred from a display label. |
| `PickerGroupReference.id` | Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. |
| `UserIdentityResolution.mode` | Required user identity mode. Available retains returned IDs; entra requires a verified directory ID; sharePoint and both require a target collection registration. |
| `UserIdentityResolution.webUrl` | Absolute HTTP(S) source web URL. Omission, when allowed, uses the provider current web; never inferred from a display label. |
| `UserIdentityResolution.ensureUser` | Allows registering a missing SharePoint user only during explicit selection or ID resolution. Defaults to false; never runs during search, hydration or configuration changes. |
| `UserPrincipal.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `UserPrincipal.loginName` | Required verified login/UPN or SharePoint membership claim. Login-only identity requires the provider verified tenant context. |
| `UserPrincipal.email` | Optional user contact email, never canonical user identity. For unvalidatedEmail values it is required and remains explicitly unverified. |
| `UserPrincipal.userPrincipalName` | Optional verified directory UPN used for identity lookup; not an Entra object GUID. |
| `UserPrincipal.entraId` | Optional verified Entra user GUID; must agree with value.entra.id. |
| `UserPrincipal.entraTenantId` | Optional verified Entra tenant GUID; must agree with value.entra.tenantId. |
| `UserPrincipal.sharePointId` | Optional verified positive SharePoint user ID; scoped to sharePointWebUrl and agreeing with value.sharePoint.id. |
| `UserPrincipal.sharePointWebUrl` | Optional verified site collection URL for the SharePoint user ID; not the requesting subweb URL. |
| `UserPrincipal.imageUrl` | Optional user avatar image URL; omission uses Fluent initials. Non-user entities use category icons. |
| `PrincipalPickerOption.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `PrincipalPickerOption.source` | Required source or transport discriminator; SharePoint identities and Entra identities remain distinct. |
| `PrincipalPickerOption.id` | Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. |
| `PrincipalPickerOption.webUrl` | Absolute HTTP(S) source web URL. Omission, when allowed, uses the provider current web; never inferred from a display label. |
| `PrincipalPickerOption.description` | Optional group description snapshot; the default secondary label uses type and source instead. |
| `PrincipalPickerOption.tenantId` | Required verified Entra tenant GUID on serialized identities; optional redundant display-model metadata must agree when provided. |
| `PrincipalPickerOption.mail` | Optional directory group mail address; never used as canonical group identity. |
| `PrincipalPickerOption.securityEnabled` | Required verified directory security flag for group classification. |
| `PrincipalPickerOption.mailEnabled` | Required verified directory mail flag for group classification. |
| `PrincipalPickerOption.email` | Contact email: optional for user models, required for the unvalidatedEmail variant. Never canonical user identity. |
| `PrincipalPickerInitialPrincipal.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `PrincipalPickerInitialPrincipal.loginName` | Required verified login/UPN or SharePoint membership claim. Login-only identity requires the provider verified tenant context. |
| `PrincipalPickerInitialPrincipal.text` | Optional fallback display label for login/email/Entra/group references; required for displayOnly. Never establishes identity. |
| `PrincipalPickerInitialPrincipal.email` | Required email reference for read-only, unambiguous directory lookup; not an Entra ID. |
| `PrincipalPickerInitialPrincipal.id` | Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. |
| `PrincipalPickerInitialPrincipal.group` | Required native group reference with its source and scope; numeric SharePoint IDs and Entra GUIDs remain distinct. |
| `PrincipalPickerInitialPrincipal.key` | Required stable renderer alias. It must never be reused for a different scoped entity. |
| `PrincipalUserScope.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `PrincipalUserScope.webUrl` | Absolute HTTP(S) source web URL. Omission, when allowed, uses the provider current web; never inferred from a display label. |
| `PrincipalUserScope.groups` | Required membership group references; user candidates must belong to at least one configured group. |
| `PrincipalUserScope.name` | Required exact SharePoint group name within the configured source web. |
| `PrincipalInitialSuggestions.includeCurrentUser` | Offer the current SPFx user if allowed by user type/scope. Default false; independent of personal suggestions. |
| `PrincipalPickerProps.initialSuggestions` | Optional current user plus shared five-item pre-search policy. showDefaultSuggestions=false suppresses all pre-search sources and their lazy retrieval. Search placeholder shows “Type at least 2 characters to search” (configured larger minimum is reflected), without a duplicate minimum-character status row. |
| `PrincipalPickerProps.loadPersonalSuggestions` | Lazy delegated Graph personal suggestions on opening an empty-query menu. Default false. Users: GET /me/people (“Persone suggerite”), currently in maintenance mode, retained intentionally. Groups: GET /me/memberOf (“I tuoi gruppi”), direct group memberships only; roles/admin units excluded. Respect principalTypes and scopes; unverified directory identities/membership are not selectable. person.id is not an Entra user ID. SharePoint groups cannot come from memberOf. Failure is reported through onError with suggestionSource; healthy/consumer sources remain usable. Permissions/limits: see docs/api/controls/pickers.md. Equivalent requests share owner-scoped memory cache only. |
| `PrincipalPickerProps.principalTypes` | Selectable categories; defaults to ['user']. |
| `PrincipalPickerProps.webUrl` | SharePoint group lookup web, also fallback identity-resolution destination; defaults to current web. |
| `PrincipalPickerProps.userScope` | User suggestion scope; defaults to tenant. |
| `PrincipalPickerProps.defaultPrincipals` | Optional uncontrolled principal references consumed once at mount. Mutually exclusive with defaultSelectedOptions; controlled selectedOptions takes precedence. |
| `PrincipalPickerProps.userIdentity` | IDs required for confirmed users; defaults to available. See UserIdentityResolution. |
| `PrincipalPickerProps.allowUnvalidatedEmail` | Defaults to false; only allowed with user enabled and available identity mode. |
| `PrincipalPickerProps.showHiddenPrincipals` | Deprecated compatibility flag; no effect. Native SharePoint user visibility applies. Not sent to REST and does not invalidate request identity. |
| `PrincipalPickerProps.useSubstrateSearch` | User search option; defaults to false. |
| `PickerValueLabel.text` | Optional plain label used until live metadata resolves. Omission means no fallback label. |
| `SitePickerValue.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `SitePickerValue.url` | Absolute normalized site collection URL. |
| `SitePickerValue.siteId` | Site collection GUID, when returned/verified. |
| `ListPickerValue.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `ListPickerValue.webUrl` | Required absolute HTTP(S) source web URL; serialized values have no provider-current-web default. |
| `ListPickerValue.listId` | Required list GUID scoped to webUrl; lists have no site-column variant. |
| `FieldPickerValue.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `FieldPickerValue.webUrl` | Required absolute HTTP(S) source web URL; serialized values have no provider-current-web default. |
| `FieldPickerValue.listId` | Scoped list GUID, agreeing with the JSON reference. For site columns this is omitted. |
| `FieldPickerValue.fieldId` | Required field GUID within the source web/list. |
| `FieldPickerValue.internalName` | Required field internal name, agreeing with the JSON reference; distinct from a REST EntityPropertyName. |
| `ViewPickerValue.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `ViewPickerValue.webUrl` | Required absolute HTTP(S) source web URL; serialized values have no provider-current-web default. |
| `ViewPickerValue.listId` | Required source list GUID; views always belong to a list in webUrl. |
| `ViewPickerValue.viewId` | Required view GUID within its source list. |
| `PickerCustomItemKey.field` | Verified internal field name; must select unique finite-number/string values. |
| `PickerCustomItemKey.value` | String/number discriminator is retained; numeric 1 differs from string '1'. |
| `ListItemPickerValueBase.webUrl` | Required absolute HTTP(S) source web URL; serialized values have no provider-current-web default. |
| `ListItemPickerValueBase.listId` | Required source list GUID for the item, file or folder, scoped to webUrl. |
| `ListItemPickerValueBase.itemId` | Required positive safe integer SharePoint item ID, scoped to the web/list. Always retained even with a custom key. |
| `ListItemPickerValueBase.customKey` | Custom key pair; omission uses authoritative itemId (Id). |
| `SharePointItemValue.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `SharePointFileValue.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `SharePointFileValue.uniqueId` | Verified File.UniqueId; required for a confirmed file value. |
| `SharePointFileValue.serverRelativeUrl` | Optional server-relative navigation snapshot. Not canonical identity; must have one leading slash and no query, fragment or traversal segments. |
| `SharePointFileValue.extension` | Optional verified file extension snapshot, without a leading dot; used only for presentation. |
| `SharePointFolderValue.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `SharePointFolderValue.uniqueId` | Required verified File/Folder UniqueId GUID. Library origin alone does not establish file or folder identity. |
| `SharePointFolderValue.serverRelativeUrl` | Optional server-relative navigation snapshot. Not canonical identity; must have one leading slash and no query, fragment or traversal segments. |
| `EntraObjectIdentity.tenantId` | Required verified Entra tenant GUID on serialized identities; optional redundant display-model metadata must agree when provided. |
| `EntraObjectIdentity.id` | Required Entra object GUID, scoped to tenantId. |
| `SharePointPrincipalIdentity.siteCollectionUrl` | Required verified absolute site collection URL owning the numeric principal ID. |
| `SharePointPrincipalIdentity.id` | Required positive safe integer principal ID scoped to its site collection. |
| `UserPrincipalValue.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `UserPrincipalValue.loginName` | Required verified login/UPN or SharePoint membership claim. Login-only identity requires the provider verified tenant context. |
| `UserPrincipalValue.entra` | Optional verified tenant-scoped Entra user identity. Required on confirmed selection in entra/both modes. |
| `UserPrincipalValue.sharePoint` | Optional verified collection-scoped SharePoint user identity. Required on confirmed selection in sharePoint/both modes. |
| `SharePointGroupValue.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `SharePointGroupValue.identity` | Required scoped native group identity; never substituted with a name or email. |
| `EntraGroupValue.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `EntraGroupValue.identity` | Required scoped native group identity; never substituted with a name or email. |
| `UnvalidatedEmailValue.kind` | Required discriminant selecting the entity or reference variant; never inferred from a display name. |
| `UnvalidatedEmailValue.email` | Required syntactically valid but unverified email address, explicitly distinct from a directory user identity. |


### Content type member reference

Inherited selection, label, visual slots and source-web behavior follow `PickerProps`, `PickerOption`, `PickerValueLabel` and `WebPickerSource` above.

| Member | Contract |
| --- | --- |
| `ContentTypePicker.props` | Required ContentTypePickerProps; toolkit provider, authenticated client and host Fluent provider are required. Returns the shared picker Field. |
| `ContentTypePickerOption.kind` | Required `contentType` discriminator, matching the reference. |
| `ContentTypePickerOption.webUrl` | Required verified catalog web URL, matching value.webUrl. |
| `ContentTypePickerOption.id` | Required hierarchical StringId, matching value.contentTypeId case insensitively. |
| `ContentTypePickerOption.name` | Required SharePoint Name; used as the option label. |
| `ContentTypePickerOption.group` | Required group string used as secondary metadata; not serialized automatically. |
| `ContentTypePickerOption.description` | Optional description shown as a native hover title; omitted when unavailable. |
| `ContentTypePickerOption.hidden` | Required Hidden flag; filtering is controlled by includeHidden. |
| `ContentTypePickerOption.readOnly` | Required ReadOnly metadata, which alone does not exclude a selection. |
| `ContentTypePickerOption.sealed` | Required Sealed metadata, which alone does not exclude a selection. |
| `ContentTypePickerProps.scope` | Site by default; list requires list. Selects a catalog, not a provisioning operation. |
| `ContentTypePickerProps.list` | Existing toolkit list selector required only in list scope; forbidden in site scope. |
| `ContentTypePickerProps.group` | Optional exact case insensitive group filter. |
| `ContentTypePickerProps.includeHidden` | True by default; false excludes hidden CTs. |
| `ContentTypePickerProps.contentTypeId` | Optional exact full hierarchical ID filter. |
| `ContentTypePickerProps.includeDescendants` | False by default; requires contentTypeId when true and includes it plus descendants. |
| `ContentTypePickerValue.kind` | Required contentType discriminator. |
| `ContentTypePickerValue.contentTypeId` | Required full hierarchical ID, including valid System root 0x; never interpreted as a GUID. |
| `ContentTypePickerValue.scope` | Required site/list discriminator determining URL and list rules. |
| `ContentTypePickerValue.webUrl` | Verified collection root for site scope, owning list web for list scope. |
| `ContentTypePickerValue.listId` | Required list GUID in list scope, forbidden in site scope. |

## Colorful selected-tag avatars

All seven picker props inherit `PickerProps.colorfulTagAvatars?: boolean`. Omitted or false preserves the default palette: theme brand colors for entity icons and neutral initials for people. With true and `presentation="tagPicker"`, resolved selected tags use Fluent Avatar `color="colorful"`; the selected renderer key supplies `idForColor`, so renaming its label preserves its color. The finite Fluent palette can assign the same color to different identities.

```tsx
<PrincipalPicker label="People" presentation="tagPicker" colorfulTagAvatars />
<SitePicker label="Sites" presentation="tagPicker" colorfulTagAvatars />
```

Existing user photos, category icons, avatar shape and inherited tag dimensions are retained. Color is most visible on initials/icons and photo fallbacks. Unresolved tags without media remain without an avatar. Suggestion/Persona rows and Dropdown presentation keep their existing appearance. This is presentation only: toggling it retains selection, typed query and requests/cache; it neither writes values nor triggers a change event. The existing host provider/theme requirements apply; Fluent tokens provide the palette in the active theme.

Every `PropertyPane*Picker` inherits the same flag. The **SharePoint pickers** demo exposes **Colorful tag avatars** in TagPicker presentation; the native **Picker modes** group has a separate toggle for all seven fields. Both default off.

## Native localization and instance strings

SPFx selects the owning Pickers resource from the web part's UI culture. The English resource is the fallback. There is no toolkit locale selector. `PickerProps.strings` is inherited by all seven pickers and by their Property Pane wrappers. Public `PickersStrings` describes a complete dictionary; `PickersStringsOverride` allows optional known groups and optional leaves. Both types are exported from the picker family entry, `/controls` and the package root.

```tsx
import { SitePicker } from '@apvee/spfx-react-toolkit/controls/pickers';
import type { PickersStringsOverride } from '@apvee/spfx-react-toolkit/controls/pickers';

const strings: PickersStringsOverride = {
  labels: {
    clearSelection: 'Clear chosen site',
    namedSearch: name => `Find ${name}`,
    selectedCount: count => `${count} choices`,
  },
  status: { noResults: 'No matching sites', partialHubs: 'Narrow the hub scope' },
  sections: { consumer: 'Suggested by this application' },
};
<SitePicker label="Project site" strings={strings} placeholder="Find a project site" />;
```

Nested groups merge only known fields. `undefined` keeps the native default, and `''` remains a deliberate text override. A complete dictionary supplies every member, including callbacks, and can be passed through the same `strings` prop. Dictionary changes update UI text without clearing selections or invalidating transport caches. Concurrent controls resolve their own messages even when they reuse cached source warnings.

| Group | Members / behavior |
| --- | --- |
| `PickersStrings.labels`, `PickersStringsOverride.labels` | `selectOne`, `selectMany`, `search`, `clearSelection`, `clearSearch`, `options`, `searchOptions`, `selectedOne`, `selectedMany`, `noSelection`, `retry`, `resolveIds`, `moreResults`; `namedOptions(name)`, `namedSearch(name)`, `selectedCount(count)`, `minimumSearchLength(count)` return plain text. The native count callbacks use the loaded resource's plural grammar. |
| `PickersStrings.status`, `PickersStringsOverride.status` | `searching`, `loading`, `partialSearch`, `noOtherResults`, `noResults`, `noOptions`, `moreItemsAvailable`, `partialSuggestions`, `unresolvedValue`, `partialHubs`, `partialSiteCollections`. |
| `PickersStrings.sections`, `PickersStringsOverride.sections` | `consumer`, `currentSite`, `hubSite`, `catalog`, `currentUser`, `people`, `memberOf`. |
| `PickersStrings.metadata`, `PickersStringsOverride.metadata` | `documentLibrary`, `sharePointList`, `personalView`, `publicView`, `field`, `file`, `folder`, `sharePointGroup`, `microsoft365Group`, `securityGroup`, `distributionList`, `unvalidatedEmail`; callbacks `fieldDescription(internalName, type)`, `itemDescription(id)`, `fileWithExtension(extension)` format literal source metadata. Extensions are supplied in uppercase. |
| `PickersStrings.validation`, `PickersStringsOverride.validation` | `selectionLimitReached`, `selectionLimitExceeded`, `unresolvedValues`, `outOfScopeValues`, `resolveRequiredUserIds`. |
| `PickersStrings.errors`, `PickersStringsOverride.errors` | `candidateOutOfScope`, `incompleteCandidateResolution`, `resolutionFailed`, `graphUnavailable`: toolkit-created fallback messages only. |

Existing direct search placeholders (including picker-specific placeholder props), `loadingContent`, `emptyContent`, `errorContent`, Field validation messages and caller section headers retain precedence and their existing slot contracts. The Dropdown's empty selection prompt is distinct from its search placeholder. Dictionary values are plain text and do not replace content slots.

Source titles, selected labels, source descriptions, internal field names/types, IDs, extensions and user data remain literal. Caller validation text, backend errors and developer diagnostics retain their original contracts. The dictionary does not translate arbitrary external errors, change query matching, reformat stored values or change callback payloads. English wording above illustrates the English native resource; the actual default follows the loaded SPFx resource.

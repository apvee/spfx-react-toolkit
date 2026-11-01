# Select and persist SharePoint references

Choose a discovery source, keep complete saved identities distinct from their resolved display subset, and treat navigation, accepted scope and remote writes as separate decisions.

[Documentation home](../INDEX.md)

## Selecting SharePoint values

The seven controls use Fluent UI 9 Dropdown or TagPicker, with Persona rows. They require the matching `SPFxWebPartProvider` (or extension provider), native authenticated clients and a host FluentProvider/renderer. Use `@apvee/spfx-react-toolkit/controls/pickers`; the same React exports and JSON types are available from the package root and `/controls`.

### Dropdown selection and popup search

The public `presentation="combobox"` string is retained for compatibility and now renders Dropdown. Selection and search behavior is unchanged; no `"dropdown"` value is added. Its selection button always displays the selected name and media, or `N items selected` for multiple values. Empty fields show `Select an item` / `Select items`. `placeholder` and picker overrides such as SitePicker's `searchPlaceholder` customize search text, with `Type at least N characters to search` as the searchable fallback, using the configured `minimumSearchLength`; they do not replace the empty selection prompt.

With `searchable !== false`, opening focuses a SearchBox in the popup, outside the semantic list. Query clear affects only search, while native single-selection clear emits a validated clear action. `searchable=false` omits this SearchBox. TagPicker keeps its input and removable tags; selected identities are excluded from every candidate/contextual section and are not repeated as selected popup rows. `colorfulTagAvatars` applies only to selected tags.

Arrow keys navigate native options, Enter confirms, Escape closes and returns focus to the trigger, and Tab/Shift+Tab dismiss without selecting a highlighted option. Editing keys, spaces and composition stay with search. Single selection closes after asynchronous confirmation/validation accepts; rejection preserves the query and popup and emits no selection event. Pending confirmation preserves focus; a late result cannot steal outside focus or close a newer popup session. Blur validation observes exit from the entire control, including popup and actions.

Retry, explicit ID resolution and **Altri risultati** remain separate actions below the selection control. Reopening after an action retains the current query and accumulated pages.

The changed DOM and interaction are intentional: `.fui-Combobox` selectors and a permanent text input in the closed field are no longer supported assumptions. `controlClassName` targets Dropdown, while `id` and field ARIA naming target its actual button; popup search has its own name. Public imports, props, selection callbacks and JSON references remain unchanged. See the [native fields](../../api/property-pane/pickers.md) for persistence and Apply ownership.

`PickerProps<T>` extends the native Fluent Field label, hint, required, validationMessage and validationState contracts. `PickerOption<TValue>` is `{ key, text, value }` plus domain display metadata. Keys are stable opaque renderer aliases; canonical identity comes from the validated scoped value, never a display label. Generated keys carry a versioned JSON reference and can hydrate without a display model. An arbitrary consumer key needs a model; missing models remain unresolved.

The picker additionally requires an accessible input name: a nonblank plain string `label`, or explicit `aria-label`/`aria-labelledby`. JSX and Fluent object labels retain their native slot types, but require explicit ARIA naming even when they display text. The renderer throws `TypeError` if neither naming input is supplied. A supplied `aria-label` takes precedence over the plain label, so an empty override without `aria-labelledby` also fails this check. Consumers own the referenced IDs and meaningful accessible text.

All controls support `presentation="combobox" | "tagPicker"`, single/multiple selection and typed search. Searchable defaults to true, minimumSearchLength to 2 (integer at least 2), debounceMs to 300 (integer at least 0), suggestionsLimit to 50 (integer 1–200). The result page size is independent of the initial suggestion/catalog cap. `selectionLimit` limits new choices; single selection always allows at most one new value. Existing conflicting controlled values are preserved and diagnosed. `showDefaultSuggestions=false` suppresses initial suggestion sources and their lazy requests while leaving typed search available.

Consumer candidates are verified read-only against their source, type and scope before blank-query display; read-only eligibility runs in batches of at most 20 concurrent requests, without truncating the supplied set. Results are cached by source/client/identity, so new arrays, labels and callbacks do not repeat equivalent requests. Consumer models precede the scoped List/Field/View/ContentType catalog and optional current-user and personal models, win equivalent display identities, and are deduplicated after filtering. Already selected identities are excluded before applying the total cap. The search placeholder displays “Type at least 2 characters to search” (or the configured larger minimum); empty or below-minimum queries do not duplicate that guidance in a status row. Typing queries the configured SharePoint/Graph source, including entities outside the consumer set; Field/View/ContentType instead filter their complete scoped cache locally. `filterOptions` is a local transformation; mandatory scope/type, selected exclusion, deduplication and the initial cap apply again afterward.

### Search feedback

All seven controls share the same feedback for Dropdown and TagPicker, in single and multiple selection. Feedback describes the **current query**, independently of selected rows/tags and SitePicker's fixed current-site/actual-hub choices. Those choices remain visible when enabled, but their presence alone does not establish a search match or hide a no-results message.

Eligible matches respect the configured source, scope and filters. The distinction between no matches and no **other** matches is evaluated before selected and pinned identities are removed from ordinary result rows. A previous query's completion, empty result or failure cannot replace the current query's feedback. Query failures and contextual-suggestion failures are shown independently, alongside any healthy rows; a failed hub suggestion must not conceal a successful empty search, and a query failure must not conceal the contextual failure.

`emptyContent`, `loadingContent` and `errorContent` retain their existing fixed inner `div` slot contracts. Omission selects the state-specific defaults above; literal content or shorthand customizes that region, and `false` suppresses its content while preserving the control-owned status shell. Explicit custom empty text still applies to initial/below-minimum guidance and complete empty/no-other-result states; partial warnings remain distinct. Retry remains an independent action. These slots change presentation, not query timing, selection, eligibility, error reporting or completion guarantees.

For example, keep the defaults to distinguish an unfinished search from a complete empty result, or supply product-specific wording through the existing slots:

<!-- toolkit:include symbol="SitePicker" section="example" title="Compare contextual search and custom feedback" -->

To observe the distinction, open SitePicker with contextual choices enabled, enter fewer than the minimum characters, then enter an eligible query with no matches. The placeholder supplies the initial guidance; an eligible query shows pending feedback and then the complete-empty message while contextual choices remain available. Search for a selected or pinned identity to observe the no-other-results state. Repeat with TagPicker or multiple selection; changing presentation does not change the query contract. Field/View/ContentType load metadata remotely once and then match locally over the complete scoped catalog. ListItemPicker retrieves remote candidate pages and matches locally within those pages, so a bounded scan can remain partial even with no matching rows. Other remote search sources retain their documented paging and permission limits.

All picker search inputs set browser `autocomplete="off"` in both Dropdown and TagPicker presentations, including Property Pane. This requests suppression of browser history/autofill suggestions; picker search and suggestions remain enabled. Browsers and password-manager extensions may override this hint.

Names and secondary metadata in Persona suggestion/selection rows truncate with an ellipsis when the control is narrow. All avatars use Fluent `shape="square"` with its native corner rounding. Popup rows use 24px avatars; the selected Dropdown value uses 20px for medium and 24px for large. Entity icons use Fluent's theme brand palette; people retain their photos or neutral initials. Selected tags also stay within the control, keeping their icon and removal action visible. Tag media avatars inherit Fluent’s native tag size, so icons, photos and initials remain proportional for the supported medium and large picker sizes. Full labels remain in the DOM/accessibility text and native hover titles; display truncation does not alter keys or saved values. This presentation is shared with Property Pane.

### Initial catalog and selection visibility

ListPicker loads the scoped remote catalog lazily when the empty-query menu opens. `initialCatalogLimit` defaults to **30** (integer 1–200), applied after consumer precedence, deduplication, filtering and selected exclusion. Selected rows do not consume this cap. A lookahead/continuation distinguishes exactly 30 results from additional results and displays the search hint. Composition reads at most five source pages of 200 and preserves partial-source warnings.

FieldPicker and ViewPicker load **all** matching fields/views, following every native REST continuation without a display or five-page limit. They show the entire available catalog, excluding selected identities and applying configured filters. **`initialCatalogLimit` is removed from both React and Property Pane contracts.** Typed search is local, case insensitive prefix matching over the complete cached catalog, even for entries beyond 30 or 200. Changing query/casing/page-size does not make additional catalog REST calls. Source identity/client/user changes, cache expiration or explicit refresh can reload the catalog; failed or cyclic paging is reported as an error rather than complete data. Read-only selection hydration/eligibility requests remain separate from search. Raw filters/order remain server-side. Native paging is retrieved before showing the complete catalog, so large metadata sets can take longer to load.

`showDefaultSuggestions=false` suppresses blank-query loading/display for List/Field/View/ContentType; typing still loads Field/View/ContentType's complete catalog once if needed, then filters locally. `initialOptions` remains selection-model seeds. Site/Item/Principal retain the five ordinary suggestion policy; Site also has the fixed contextual choices described below.

Site/List/Field/View/ContentType/Item typed search is case insensitive: Site uses SharePoint's KQL semantics; hub catalogs and the other metadata/item sources use deterministic Unicode lowercase matching. This does not promise accent removal or transliteration. Query cache/cursors treat case-only edits as equivalent while preserving the displayed input and accumulated pages. IDs, raw consumer filters/KQL and ordering are not case-folded. Site/List/Item typed search continues through the configured source with bounded paging; Field/View/ContentType search filters the complete memory-cached catalog.

Multiple Dropdown menus always include the localized **Selected items** group, independent of the typed query and available cap. Checked rows, including unresolved fallback labels, can be clicked to remove them without searching. They remain excluded from available/search results. Single Dropdown menus show the localized **Selected item** group first, with a checked Persona row (or unresolved fallback label), independent of query and suggestion limits. This identity is excluded from available/search rows. Clicking it preserves the value without validation, resolution requests or change events; choosing another replaces it. Typing changes only the popup query and preserves the selection button text. The native clearable action still clears it. TagPicker retains removable tags. Both menus scroll within a 320px/80vh maximum height. Read-only hydration and scope checks show loading while pending, without a transient unresolved error. Definitive rejected/missing values remain diagnosed, even while another selection is pending; hydration emits no selection change.

`onOptionSelect(event, data)` runs only for accepted explicit `selection`, `remove`, `clear` and `resolve` actions. `PickerSelectionEvent` is the native Fluent event; `PickerSelectionData` contains ordered selectedOptions, resolved selectedItems/selectedValues, unresolvedOptions and the optional affected key/text. Hydration and prop updates emit no change event. Resolved React selectedValues may be a subset when unresolved keys exist: consumers must not replace persisted references with that subset. The native Property Pane bridge retains the complete reference array.

`validateSelection(items)` receives the proposed resolved models, excluding unresolved aliases. Returning `undefined` or an empty string accepts the validated action; a nonempty message vetoes it, preserving selection and emitting no `onOptionSelect`. Promises are awaited. A throw or rejection uses the failure UI and `onError`, without committing: selection/remove/clear and blur report operation `validate`, while explicit resolution reports `resolve`. Omission means no custom validation. External `validationMessage` has display precedence only; it does not override the validator's acceptance decision.

`validationTrigger` defaults to `change`, validating selection, removal and clear before committing. With `blur`, those gestures commit without custom change validation; focus leaving the whole control and its owned popup validates the retained models when no confirmation is pending, without revoking an existing selection. Explicit ID resolution, when available, always validates before committing, including in blur mode. Consumer Promises are not cancelled.

`onError(error, details)` identifies transport, operation and optional personal suggestion source. Empty responses, partial sources and failures are distinct. Source failures remain visible beside healthy rows with a retry action. Bounded-result warnings remain visible beside rows and ask for narrower criteria instead of offering retry; failed requests are retried only after an explicit retry/refresh, cache eviction or a new lifetime.

`WebPickerSource`, `SearchPickerOptions`, `PickerOrderBy` and `PickerInitialSuggestions` are the shared source/query/policy types. Native and React counterparts use the same defaults below.

### SitePicker

`SitePickerProps` / `SitePickerOption` / `SitePickerValue`: Dropdown, single selection by default. Mode defaults to `site`; `site` discovers collections, `hub` registered hubs, and `associatedSites` sites associated with a supplied hubSiteId or the current collection's hub. No association yields an empty result. Discovery is not restricted to the host collection; the configured mode and KQL filters determine candidates, subject to SharePoint permissions. `additionalQuery` is a KQL AND clause allowed only in site mode. Order defaults to Title ascending; trimDuplicates defaults to true. `searchPlaceholder` overrides placeholder.

`includeCurrentSite` and `includeHubSite` are flat boolean props, both **true by default** (also in Property Pane). The menu keeps **Sito corrente** and **Hub associato** visible before and during typed search, after the selected section. These fixed contextual choices sit outside the cap of five ordinary initial suggestions, so a full consumer set cannot hide them. Ordinary consumer suggestions follow the contextual choices. The current site is the collection root, including from subweb hosts. The hub is read from the current collection's actual `HubSiteId`/`IsHubSite` metadata and resolved via `HubSites/GetById`; the configured `hubSiteId` only controls associatedSites discovery and does not substitute the current collection's real association. A hub with no actual association does not suggest itself as an associated hub. Canonical duplicates (including current=hub and typed results) appear once, consumer models win, and selected identities stay in their checked selected section (or selected tags).

Hub metadata follows the [SharePoint hub REST API](https://learn.microsoft.com/en-us/sharepoint/dev/features/hub-site/hub-site-rest-api), including [GetById](https://learn.microsoft.com/en-us/sharepoint/dev/features/hub-site/rest-getbyid-method). The optional sources load lazily when the menu opens, including with typed text, using scoped memory-cached read-only REST metadata. False disables that source; both false restore consumer-only defaults; `showDefaultSuggestions=false` suppresses both sources. No association is an empty source. Failures are independent: healthy current/hub/consumer rows remain; errors identify `suggestionSource: 'currentSite' | 'hubSite'`, `source: 'sharePoint'`, `operation: 'load'` and support explicit retry/refresh. Cosmetic prop changes and equivalent flag values reuse cached requests. Search and hub catalog cache entries also survive these display-policy switches.

Plain site-mode defaults need no text Search. Other modes and `additionalQuery` still perform their normal source eligibility checks, which can use Search; explicit selection verification retains the configured source contract. Typed search continues using SharePoint Search (hub mode filters its hub catalog), independently of initial suggestions. No storage, provisioning or permission grants occur.

Site/association uses the toolkit Search service without SortList: Title and Path are not sortable in the standard SharePoint schema. It reads at most five native pages of 200 results, caches and sorts the entire retrieved set by title or URL, then pages that ordered set locally. Continuations are bound to query, owner/client and page size. Hub mode uses the same five-page bound and local ordering/paging for its catalog. If native results remain beyond 1000, the control reports partial data and recommends narrowing the query/scope; ordering then covers only the retrieved set. URLs establish site collection identity even when the endpoint omits siteId.

### ListPicker

`ListPickerProps` / `ListPickerOption` / `ListPickerValue`: Dropdown, single, current web, Title ascending. includeHidden defaults to true. baseTemplate accepts one numeric template or an OR set; `filter` overrides these convenience filters. contentTypeId matches expanded content-type ID prefixes case-insensitively. Typed search matches Title prefixes case-insensitively across bounded native REST pages. Native continuations retain filtering, projection and ordering; they cannot be reused against another query/client/source.

### FieldPicker

`FieldPickerProps` / `FieldPickerOption` / `FieldPickerValue`: Dropdown, multiple by default. Omitted list queries site columns; a supplied exact title or `SPFxPnPListSelector` selects list fields. String GUIDs are titles, not implicitly converted IDs. includeHidden/includeReadOnly default true; group is exact. Raw filter overrides convenience filters. Field-specific icons distinguish text, choice, number, date, person, lookup and other kinds; metadata appears once as the secondary line. showBlankOption defaults false and enables clearing to `[]`.

### ViewPicker

`ViewPickerProps` / `ViewPickerOption` / `ViewPickerValue`: Dropdown, single. list is required; source web defaults current, Title ascending. Excluded titles are exact case-sensitive matches. DefaultView and PersonalView metadata distinguish public/personal views. showBlankOption means an explicit clear action, never a fake view GUID.

### ListItemPicker

`ListItemPickerProps` / `ListItemPickerOption` / `ListItemPickerValue`: TagPicker, multiple. list and internal textField are required. keyField defaults `Id`, searchMode defaults startsWith, order defaults Id ascending, includeFolders defaults false. Existing toolkit selectors also support decoded list paths/URLs. Consumer keys remain opaque; values always retain actual itemId, even with a custom key.

For file candidates, typed search matches either the configured display field or `File.Name`, using the same case insensitive `startsWith`/`contains` mode. A nonempty display field remains the label; when empty (including whitespace), the filename is used during discovery, selection and hydration. Ordinary list records and folders retain their existing labels and search behavior. React and Property Pane use the same source.

Text, Note and calculated text (verified OutputType) match the REST EntityPropertyName locally with case-insensitive startsWith/contains; other field types match their FieldValuesAsText projection. SharePoint REST does not support `tolower`/`toupper`, and CAML does not establish the required casing contract, so generated string predicates are not sent. Raw consumer filters, folder policy and ordering remain server-side. Each search page scans at most five native REST pages (up to 1000 candidates at page size 200); a continuation and partial status allow further loading. Incomplete scans are never presented as complete. Large or threshold-limited lists may require narrower indexed filters. A custom key must have unique finite-number/string values; duplicates or an unfinished uniqueness scan leave the value unresolved. Numeric `1` differs from string `'1'`.

A library row becomes `file` or `folder` only with verified FSObjType and File/Folder UniqueId, otherwise ordinary list rows remain `item`. Icons distinguish records, folders, generic files and known extensions such as PDF. Files use the verified filename when their configured display value is absent, empty or whitespace. For ordinary records and folders, only an absent/null display value uses the Item ID fallback; defined empty or whitespace strings retain their existing value.

### ContentTypePicker

`ContentTypePickerProps`, `ContentTypePickerOption` and `ContentTypePickerValue` select existing SharePoint content types. Default: `scope="site"`, Dropdown, single selection, `includeHidden=true`. The complete catalog is loaded with native paging and sorted by name/ID; there is no initial 30-item cap. Typed search filters names locally by prefix, case insensitively. All standard picker selection, model/consumer-option, accessibility and visual slot contracts above apply.

Site scope resolves the owning **site collection root web**, including when `webUrl` identifies a subweb; it does not enumerate only that subweb's definitions. List scope requires the existing toolkit `list` selector and returns only content types already associated with that list/library. The same content-type ID in these two scopes remains a distinct reference. The control does not create, publish, associate or synchronize content types, and does not enable content types on a list.

Persona rows show name and group, with an icon; selected tags show name/icon. Description and Hidden/ReadOnly/Sealed metadata belong to the option model. The serializable reference contains `kind:'contentType'`, `scope`, `webUrl`, `contentTypeId`, optional snapshot `text`, and `listId` only in list scope. Site references store the verified collection root web URL. Identity includes scope, source URL, list and the complete ID; names and parent-ID prefixes never establish equality.

The source reuses `usePickerSourceContext` (and its toolkit `useSPFxSPHttpClient` client), `cachedPickerSpRead`, `resolvePickerListTarget` for lists, and `createMetadataPickerSource`. Rendering and selection reuse `usePickerRuntime`/`PickerRenderer`. There is no new HTTP client, Graph dependency or public content-type service. REST failures remain errors rather than empty catalogs; valid mounted catalogs are reused, refresh/source changes may reload, and native pane disposal releases its cache as described below.

### PrincipalPicker

`PrincipalPickerProps` / `PrincipalPickerOption` combines people and groups. `PrincipalType` supports user, sharePointGroup, microsoft365Group, securityGroup and distributionList. Default is `['user']`; an empty array disables suggestions. Default presentation is TagPicker with multiple selection. User rows use avatar/name/email once; group rows use category icon/name/type and origin. Tags show only name plus avatar/icon. Microsoft 365 classification takes precedence over security flags; mail-enabled security groups remain securityGroup, and non-security mail groups are distributionList.

`PrincipalUserScope` defaults to tenant; web scope queries registered principals in the target collection's UserInfoList and does **not** establish effective per-web permissions. Group scopes accept `PickerGroupReference` (numeric SharePoint group ID scoped to web, or Entra GUID) or exact SharePoint group name. These scopes restrict users; selectable groups retain their own source scope. Entra membership checks may include transitive relationships; SharePoint group membership is direct. Authorization services for the current user are not substituted for candidate membership checks.

`defaultPrincipals` accepts once-only `PrincipalPickerInitialPrincipal` login, email, Entra user ID, group reference or displayOnly. It cannot coexist with uncontrolled defaultSelectedOptions. displayOnly remains unresolved rather than pretending that a title identifies a user.

### User identities and ensureUser

For native user candidates without IDs, verification searches the UPN without the SharePoint membership claim wrapper, then checks the exact canonical login and configured scope. The saved login retains its original claim format; verification never provisions a user.

`UserPrincipal` / `UserPrincipalValue` may carry both ID systems. `UserIdentityResolution` defaults to `{mode:'available'}`:

See {@link UserIdentityResolution} for every mode and its independently configured SharePoint/Entra resolution and ensureUser policy.

The SharePoint destination defaults to userIdentity.webUrl, then PrincipalPicker.webUrl, then current web. `ensureUser` defaults false and exists only in sharePoint/both mode. False reads an existing registration; missing registration blocks confirmation. True may POST `_api/web/ensureuser` only during explicit selection or “Risolvi ID”. Search, personal suggestions, hydration, seed/config updates and preflight never provision a user. Required ID failure preserves the previous selection. A sent server write may finish after cancellation; the UI ignores stale completion but cannot promise server rollback. Changing modes preserves existing values and diagnoses missing requested IDs until explicit resolution. The resolution action appears only for users missing the requested tenant/destination-collection IDs; an ID from a previous collection does not satisfy a new destination. Enrichment cannot rebind a retained alias to conflicting verified IDs; changing collection registration requires matching verified Entra identity.

allowUnvalidatedEmail defaults false, requires user enabled and available mode, and produces an explicit UnvalidatedEmailValue rather than a user identity. `useSubstrateSearch` defaults false and configures native Microsoft 365 Substrate user search. `showHiddenPrincipals` is retained as a deprecated compatibility flag with no effect: SharePoint determines native user visibility. The unsupported ShowHiddenPrincipals payload member caused HTTP400 and is not sent. There is no verified equivalent visibility override in the [native query contract](https://learn.microsoft.com/en-us/dotnet/api/microsoft.sharepoint.applicationpages.clientpickerquery.clientpeoplepickerqueryparameters?view=sharepoint-csom). The [PnP search service](https://github.com/pnp/sp-dev-fx-controls-react/blob/master/src/services/PeopleSearchService.ts) confirms top-level UseSubstrateSearch and the modern value response envelope; the adapter also accepts the legacy named envelope.

### Personal suggestions, permissions and endpoint limits

`PrincipalInitialSuggestions` extends the shared limit with the independent current-user option. loadPersonalSuggestions defaults false; initialSuggestions.includeCurrentUser defaults false independently. Opening an empty-query menu loads only enabled sources. The order is consumer options, current user (“Tu”), “Persone suggerite”, “I tuoi gruppi”, with **five overall**, excluding selected identities. These are not selection history. Every source obeys type and user scope; roles/admin units and incomplete group classifications are excluded. Partial/failing personal sources do not remove healthy consumer options.

The Graph SDK owns authentication and retry. `/me` uses delegated context; application-only execution is not supported by this UI. Personal retrieval stops after five pages of 50 and twenty fresh identity/scope lookups per open. The result is partial if the budget or a source is exhausted; selected exclusion and final cap are applied after source collection. Cache hits do not spend fresh-lookup budget. No grants are performed by the picker or sample.

Verified against [List people](https://learn.microsoft.com/en-us/graph/api/user-list-people?view=graph-rest-1.0), [List direct memberships](https://learn.microsoft.com/en-us/graph/api/user-list-memberof?view=graph-rest-1.0), [Get user](https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0), [List groups](https://learn.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0) and [Check member groups](https://learn.microsoft.com/en-us/graph/api/directoryobject-checkmembergroups?view=graph-rest-1.0) on 2026-10-14. Host/tenant validation remains separate from local mocked-transport tests.

### JSON values, storage and lifetime

`PickerValueLabel` is an optional stale display snapshot. SitePickerValue is collection URL plus optional SiteId GUID; ListPickerValue/FieldPickerValue/ViewPickerValue carry web and scoped entity GUIDs. `ListItemPickerValueBase`, `PickerCustomItemKey`, SharePointItemValue/SharePointFileValue/SharePointFolderValue retain item identity and explicit kind. `EntraObjectIdentity` and `SharePointPrincipalIdentity` scope native IDs. SharePointGroupValue, EntraGroupValue and UnvalidatedEmailValue cannot masquerade as UserPrincipalValue. `PickerValue` is their union; `PickerSelection<TValue>` is always an array for zero/single/multiple selection.

Values contain plain own JSON members only. Accessors, extra properties, prototypes, invalid URLs, non-finite numbers and invalid domain IDs are rejected at the boundary. Models may carry display fields, but they are not saved automatically. The consumer owns all persistence. There is no selection history, localStorage, sessionStorage or OneDrive writing. The PnP client is configured without persistent cache.

Requests use an owner/client/tenant/user/source-scoped memory cache: 128 settled entries, five-minute successful-result TTL, separate shared in-flight requests, and explicit retry for cached failures. TTL is checked on use, not by background refetch. Changing display props, callbacks or equivalent option arrays does not initialize a new client or repeat equivalent requests while their cache entries remain valid. Identical metadata GETs share their exact URI/projection contract across discovery, hydration and verification; different filters/projections remain separate, and POST/ensureUser is never routed through that read cache.

Field/View/ContentType verification may reuse a positive canonical match from an already cached or loading, complete, error-free catalog in the same configured scope. Verification does not start a complete catalog by itself; absent/expired/replaced/failed/partial catalogs and missing candidates fall back to the existing filtered entity read. Refresh invalidates dependent reads and an old completion cannot repopulate the new cache generation.

Within a mounted control, reopening a settled typed menu preserves accumulated result pages without a loading flash when the exact first-page cache entry is still valid. A cache entry replaced by another field is adopted rather than retaining stale pages. Source/query/page-size/refresh changes invalidate obsolete continuations. Slow initial consumer eligibility and optional contextual/personal retrieval show loading instead of a premature empty/minimum prompt; settled cosmetic/equivalent updates reuse their results. A real unmount releases an owned cache; disposal/recreation or an expired/evicted entry can query again. Independent React controls have their own runtime cache unless an internal host shares one; native fields share the owner/document lease described in the Property Pane guide. Graph initialization across independent roots is not promised to coalesce. Graph-dependent Principal operations await the current provider’s client initialization, preserving pending labels and explicit ID-resolution actions. Terminal initialization failures remain reported; SharePoint-only operations remain usable when optional Graph initialization fails. Replacing the provider cancels obsolete pending work.

### Slots and verification

suggestionsHeader is a fixed inner span; emptyContent/loadingContent/errorContent are fixed inner div content slots using the toolkit private Fluent resolver and class composer. Their status shells remain owned by the control; false suppresses content. Native Field and Tooltip slots retain Fluent contracts. Structural suggestion options, Persona avatars and selection callbacks are not generalized content slots. Single Dropdown uses the native clearable action, routed through the same validated clear event; there is no separate “Svuota selezione” button. Multiple Dropdown selection removes checked rows individually; TagPicker dismisses tags. Field/View showBlankOption retains its optional explicit localized “No selection” action. Single Dropdown keeps the native checkIcon for its selected row and hides the column on unselected candidates using the peer's required null slot sentinel; multi keeps native checkboxes. Repository styles use Griffel classes.

The sample **SharePoint pickers** panel mounts one live variant with configurable presentation, multi/search, consumer model updates, principal types/IDs and collapsed JSON diagnostics. Its ContentTypePicker example offers site/list scope, group/visibility filters and an ID preset selector (Item, Document or confirmed selected IDs), with an optional descendants toggle. Selecting All content types resets that toggle. The actual Property Pane mounts all seven native fields. Its initially expanded **Picker modes** group globally selects Dropdown/TagPicker and single/multiple selection for all seven fields. Defaults are Dropdown and multiple selection, preserving existing multiple values. Presentation changes preserve selections; changing multiple to single keeps the first value. Cascading sources require exactly one selected site; list-dependent fields also require exactly one list. ContentTypePicker requires a list only in list scope. Site changes clear its value in both scopes; list changes clear it only in list scope. Versioned browser fixtures provide the larger simulation matrix; mocked REST is labelled simulation and is not tenant proof. See [native fields](../../api/property-pane/pickers.md) and [SharePoint validation](../development/sharepoint-validation/INDEX.md).

## Picker avatar presentation

### Colorful selected-tag avatars

All seven picker props inherit `PickerProps.colorfulTagAvatars?: boolean`. Omitted or false preserves the default palette: theme brand colors for entity icons and neutral initials for people. With true and `presentation="tagPicker"`, resolved selected tags use Fluent Avatar `color="colorful"`; the selected renderer key supplies `idForColor`, so renaming its label preserves its color. The finite Fluent palette can assign the same color to different identities.

Existing user photos, category icons, avatar shape and inherited tag dimensions are retained. Color is most visible on initials/icons and photo fallbacks. Unresolved tags without media remain without an avatar. Suggestion/Persona rows and Dropdown presentation keep their existing appearance. This is presentation only: toggling it retains selection, typed query and requests/cache; it neither writes values nor triggers a change event. The existing host provider/theme requirements apply; Fluent tokens provide the palette in the active theme.

Every `PropertyPane*Picker` inherits the same flag. The **SharePoint pickers** demo exposes **Colorful tag avatars** in TagPicker presentation; the native **Picker modes** group has a separate toggle for all seven fields. Both default off.

### Native localization and instance strings

SPFx selects the owning Pickers resource from the web part's UI culture. The English resource is the fallback. There is no toolkit locale selector. `PickerProps.strings` is inherited by all seven pickers and by their Property Pane wrappers. Public `PickersStrings` describes a complete dictionary; `PickersStringsOverride` allows optional known groups and optional leaves. Both types are exported from the picker family entry, `/controls` and the package root.

Nested groups merge only known fields. `undefined` keeps the native default, and `''` remains a deliberate text override. A complete dictionary supplies every member, including callbacks, and can be passed through the same `strings` prop. Dictionary changes update UI text without clearing selections or invalidating transport caches. Concurrent controls resolve their own messages even when they reuse cached source warnings.

Existing direct search placeholders (including picker-specific placeholder props), `loadingContent`, `emptyContent`, `errorContent`, Field validation messages and caller section headers retain precedence and their existing slot contracts. The Dropdown's empty selection prompt is distinct from its search placeholder. Dictionary values are plain text and do not replace content slots.

Source titles, selected labels, source descriptions, internal field names/types, IDs, extensions and user data remain literal. Caller validation text, backend errors and developer diagnostics retain their original contracts. The dictionary does not translate arbitrary external errors, change query matching, reformat stored values or change callback payloads. English wording above illustrates the English native resource; the actual default follows the loaded SPFx resource.

## Hosted file selection

`FileFolderPicker` selects existing SharePoint or OneDrive for Business files and folders through Microsoft's hosted File Picker v8. It requires a matching SPFx provider and the existing SPFx AAD token provider. The toolkit does not install the sample-local `@pnp/picker-api` package or another MSAL client.

Use the root, `/controls`, or `/controls/file-folder-picker` entry. Native fields have a separate [Property Pane entry](../../api/property-pane/file-folder-picker.md).

### Presets: the simple configuration

`FileFolderPickerPreset` is the public `'focused' | 'extended'` union used by the `preset` prop. Omitting `preset` preserves the legacy configuration contract.

Pass `preset="focused"` or `preset="extended"` explicitly. Both use the same source and result validation; the preset changes the hosted browsing experience, not what may be accepted.

The main choices are `scope`, `preset`, `itemTypes` and `multiselect`. Additional intent props are `selectionLimit`, `extensions`, `searchable`, `initialLocation`, `uploadEnabled` and `createFolderEnabled`. Standard Field/accessibility props, `size`, selected values/models and callbacks remain available. The theme comes from the existing Fluent/SPFx theme hooks. No raw SDK options object or custom preset is exposed.

Graph handling does not add consent, grants or a Graph client, and is requested only for validated hosted authentication commands. A token obtained from the existing provider may carry its already approved scopes; the picker does not narrow those scopes or prove that Microsoft uses no other APIs internally. SharePoint tokens remain resource-specific. Handling an already approved resource does not guarantee that a service operation is authorized. The [permission matrix](permissions.md#permissions-by-enabled-file-feature) applies to both presets; actual hosted API traffic and tenant grants need separate validation.

Uploads and create-folder remain explicit independent choices. Enabling uploads also offers native device upload where supported. Such writes can happen before confirmation; cancellation cannot undo them, and the selectable-reference restriction is not a pre-write destination boundary.

Microsoft owns the available navigation, including Media/People/Meetings/Favorites; there are no individual public switches for those entries. A requested entry is not a promise of available content or successful service navigation. There is no silent fallback to another preset/source when a route or resource fails. `initialLocation` changes only the start within the allowed source.

### Migrating existing consumers

Omitting `preset` retains the legacy contract and its exact defaults. A defined legacy SDK prop cannot be combined with an explicit preset, including `false` or an empty array; `undefined` counts as omission. Types and runtime validation reject mixed configuration before hosted authentication/navigation. Remove legacy props and choose a preset intentionally rather than letting one silently override the other.

The native SDK knobs listed under [legacy flat configuration](#flat-native-configuration-and-stable-opening) remain supported only without `preset` and are deprecated for new usage. This includes Graph/claims flags, navigation flags, sort/filter/layout, internal labels and link/device configuration. Common functional options such as `uploadEnabled`, `createFolderEnabled`, `hint` and `onReady` remain supported in both paths. Saved `FileFolderPickerValue` arrays do not change and never acquire a preset field. Migration does not clear selections or infer an opt-in from saved legacy flags.

Changing the preset affects the next opening and preserves current labels, selections and the mounted Property Pane host. The current opening retains its snapshot; source/item/access policy and authoritative changes retain their existing cancellation rules.

`scope` is a discriminated union:

{@link FileFolderPickerScope} owns the complete branches and required inputs; use that canonical reference when choosing collection, library, folder or current-user OneDrive scope.

Ordinary lists are rejected. A collection is verified by its actual SharePoint site ID; URL prefixes are insufficient. The hosted start location is separate from the permitted source. `initialLocation` can narrow the starting web/library/folder inside that source; it never broadens the boundary. A SharePoint starting folder requires a library. OneDrive accepts only a starting folder.

Folder inputs preserve literal spaces, apostrophes, `#`, `%` and Unicode. They are decoded paths, not arbitrary absolute URLs. Traversal and segment-prefix collisions are rejected. The toolkit validates authoritative returned metadata before accepting an entire picked batch; changing navigation or returning a shortcut does not bypass the configured scope.

A collection begins browsing at its verified root web. This does not promise enumeration of all subweb libraries by Microsoft's UI. Multiple arbitrary source roots, personal Microsoft accounts and other tenants are unsupported.

### File categories

`FileFolderPickerFileCategory` is the public union used by `fileCategory`, a shared flat prop for React and Property Pane, with or without a preset. Omission leaves files unrestricted unless `extensions` is supplied. Choose one of `images`, `video`, `audio`, `media` or `documents`; `media` includes images, video and audio. Supply either a category or a nonempty extension list. An empty `extensions` array imposes no extra restriction and may accompany a category. Categories are incompatible with `itemTypes="folders"`; `itemTypes="all"` applies the category only to files and keeps folders selectable.

The hosted picker receives the requested category filter. Before accepting returned files, the toolkit independently checks the suffix of the authoritative metadata name against its conservative extension catalog. Unknown, missing or mismatching suffixes are rejected; supplied display labels do not prove eligibility. This catalog does not promise exact equivalence with every file Microsoft may display for a category, and it does not inspect file content or verify MIME types. Use `extensions` when a precise suffix allowlist is required. Existing selected references remain consumer-owned and visible when the category changes; the control does not silently clear them or emit a change during hydration. A changed category invalidates stale pending selections.

The exact suffix catalog belongs to {@link FileFolderPickerFileCategory}; choose an explicit extension allowlist when those categories do not fit the task.

The React demo and native Property Pane expose a **File category** menu. Choosing a category clears the demo's custom extension setting and disables its extension input. Choosing folders-only clears the category; mixed mode retains it. These policy changes preserve selected references. The existing explicit source-change/reset actions retain their selection-clearing behavior.

### Props and state

`FileFolderPickerProps` defines control configuration and selection ownership. `FileFolderPickerScope` defines selectable provenance; `FileFolderPickerInitialLocation` defines initial navigation. `FileFolderPickerChangeData` is the explicit-action callback payload. `FileFolderPickerValue` and `FileFolderPickerOption` separate persisted references from display models.

`onChange` receives the complete ordered `values`, currently resolved display `items`, `unresolvedValues` and the action `reason`. Hydration and property updates never emit. Invalid batches are rejected atomically. Selection-limit conflicts never silently truncate existing references.

Each `FileFolderPickerValue` has `kind: 'file' | 'folder'`, `tenantId`, normalized SharePoint v2 `endpoint`, opaque case-sensitive `driveId` and `itemId`. Optional `name`, `webUrl` and verified `sharePoint` metadata are snapshots. Identity does not depend on the name/path. References are intentionally separate from the existing `PickerValue` union: a drive item ID is not a numeric SharePoint list-item ID.

`FileFolderPickerOption` adds `text`, optional `extension` and `sizeInBytes` to a `value`. Display models are not proof of eligibility. Unreadable/deleted references are preserved for consumer decisions and removal; hydration does not silently rewrite saved values.

### Authentication and lifecycle

Opening creates the single Stage dialog, then resolves metadata and obtains a resource-specific token using `useSPFxAadTokenProvider`. Only explicitly allowlisted same-tenant SharePoint resources receive SharePoint tokens. With a preset, validated Graph requests are handled on demand; in the legacy path graphEnabled additionally permits the exact https://graph.microsoft.com audience with a Graph request type. The hosted `SharePoint`, observed `SharePoint_SelfIssued` and `SharePoint_MSAL` authentication hints use the existing SPFx provider for that exact SharePoint resource; the integration does not issue its own tokens. Graph tokens are never substituted for SharePoint tokens. In the legacy path claims challenges default to disabled. Set claimsChallengeEnabled to permit one validated JSON claims challenge per resource/opening through the public SPFx token provider; retries bypass its cached token. In that legacy path Graph remains independently disabled unless graphEnabled is true. Presets internally enable the same bounded claims handler. Consumers must already have the appropriate delegated grants; the control does not request consent or change tenant permissions.

Microsoft can request an optional personal-site audience while browsing a SharePoint source. With `extended`, this path is enabled independently of the names of individual pivots; `focused` does not enable it. In the legacy path, an explicit `visibleLocations` array containing myFiles, recent or shared, with `navigationVisible` other than false, enables a bounded auxiliary authentication path. Only an exact tenant personal-resource origin root requested with a supported SharePoint authentication hint can trigger read-only current-user OneDrive discovery. The existing profile URL, tenant and drive-owner checks must resolve the same requested origin before a token is acquired. Discovery is lazy, attempted at most once per opening and reuses the mounted metadata request cache. Untouched defaults, hidden navigation and explicit arrays without those personal entries do not enable it.

This auxiliary token audience is separate from the selectable metadata provenance: SharePoint scope, identity, item-policy and batch acceptance checks remain unchanged. It introduces no public prop, consent, Graph dependency or write opt-in. An unrelated or unverified resource stays denied. Auxiliary discovery/token failures and optional Graph failures remain local to their hosted command; they neither set a false configured-source Field error nor clear an existing source error. Errors for the configured resource, unsupported authentication, unavailable items, source/policy violations and invalid identities use shared sanitized messages in React and Property Pane. `onError` receives the current actionable failure without provider/native payload strings. Cancel, retry and authoritative changes clear the Field's opening error; a successful token retry clears authentication errors without hiding an unrelated selection failure. Consumers maintaining their own error state must manage its lifetime independently.

With SPFx `debug=true`, internal authentication diagnostics use bounded `console.debug` events with finite categories. They distinguish host policy decisions from token acquisition outcomes without logging tokens, claims, raw resource URLs or provider/native messages. They are diagnostic evidence, not a grants inventory or proof of the iframe's internal requests. There is no new public diagnostic prop, automatic Field hint or opening spinner.

When the hosted picker returns the symbolic item ID `root`, the integration verifies the live drive-root relationship and all normal provenance restrictions, then stores the canonical drive-item ID. Other identity mismatches remain rejected. Names and paths never replace identity checks.

The URL `/_layouts/15/FilePicker.aspx` does **not** establish which internal APIs the service uses. The v8 protocol can request SharePoint or Graph resources. Host metadata reads, token requests and the hosted service's actual network traffic are separate evidence. Document permission requirements using observed operations and Microsoft's current documentation; do not infer Graph independence from a successful SharePoint page load.

A unique channel binds each opening to its exact origin, content window and message port. Unsupported commands fail explicitly. Source/item-policy and authoritative value changes cancel stale confirmations. Equivalent settings, labels and callbacks do not restart the session. Closing or disposing removes handlers, ports and the owned iframe.

Only mounted metadata requests are coalesced in memory. No selected-item history, localStorage, OneDrive storage or token cache is managed by the control. Microsoft's within-opening multiple-folder selection persistence does not persist toolkit values between sessions. The consumer owns eventual JSON persistence.

### Presentation and limitations

Browse opens Microsoft's UI; it is not a Combobox or TagPicker dropdown. The outer Field, button and removable tags inherit the host theme; native Microsoft light/dark follows the closest effective FluentProvider at opening. Hosted iframe content is owned by Microsoft and does not inherit arbitrary Fluent classes. Authentication failures and iframe restrictions are actionable errors; retry keeps the existing references. Stage owns dialog dismissal and focus; keyboard interaction inside the remote frame is Microsoft's responsibility.

The outer control is one outlined tag surface. File/folder tags wrap, truncate long names with ellipsis and expose the complete name in a tooltip; each remains individually removable. The trailing icon opens the hosted picker: Attach for files, FolderOpen for folders, FolderDocument for mixed mode. An empty field contains a focusable selection placeholder that opens the same picker. Both opening actions preserve the Field's label, hint and validation associations. There is no editable text input and no separate Browse or clear-all button. Controlled consumers can clear the array; users can remove tags individually.

The border uses the installed Fluent Input style hook directly, including its 1px border, theme radius (4px in the standard themes), hover/invalid/disabled states and native focus animation. Empty Medium/Large fields match native 32px/40px minimum heights; selected tags can expand the field vertically. The underline remains active throughout opening and while the hosted dialog is open, even after focus leaves the field. Reduced-motion behavior follows Fluent. The mode-specific opening icon and selected labels remain visible while busy; no spinner or automatic Field description is added. `controlClassName` targets this outlined surface; consumer overrides can intentionally change border fidelity. The iframe itself retains Microsoft's presentation.

Uploads, folder creation and local-device upload are disabled until the corresponding explicit write intent is enabled. Link entry is disabled in focused and legacy defaults, and enabled as a native pivot in extended. Local-device pick, account switching and share/move/download actions remain gated. Confirming selection reads metadata and does not itself perform those operations. Restrictions are checked again when confirming results.

The private sample exposes source, start location, file/folder mode, suffixes, single/multiple selection and the single hosted dialog. Automated protocol/scope tests and simulated browser fixtures are separate from authenticated tenant validation. See [SharePoint validation](../development/sharepoint-validation/INDEX.md) for the host checklist and execution limitations.

Primary references: [Microsoft File Picker v8](https://learn.microsoft.com/en-us/onedrive/developer/controls/file-pickers/), [current schema](https://learn.microsoft.com/en-us/onedrive/developer/controls/file-pickers/v8-schema), [React sample](https://github.com/OneDrive/samples/tree/master/samples/file-picking/typescript-react).

## Hosted dialog migration and opening snapshots

### Dialog and current-user OneDrive

The only presentation is a responsive **1080 × 720 px** modal dialog. The toolkit adds no visible title, close button or internal padding: the Microsoft iframe fills the surface. A hidden semantic title names the dialog for assistive technology. Microsoft's own title, toolbar and Cancel/Select actions remain part of the hosted UI. Stage retains native dismissal and focus handling; Escape inside the remote iframe is governed by Microsoft. Smaller viewports constrain both dimensions with the existing Stage outer gap.

`presentation` and OneDrive `baseUrl` have been removed. Replace popup/dialog configuration with omission, and configure the current user's drive as follows:

<!-- toolkit:include symbol="FileFolderPicker" section="example" title="Choose current-user OneDrive root or Projects folder" -->

The control uses the existing SharePoint REST client and mounted request cache to read `/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=PersonalUrl,Email` on the current collection. It validates the returned personal-site URL against the tenant, then verifies the drive owner against the current login/profile email and validates returned item provenance. Personal-domain metadata reads use the existing SPFx token provider and SharePoint HTTP client with a token for that exact trusted resource; the memory cache stores response models, never tokens. It does not infer an URL from an email address, invoke Graph `/me/drive`, provision a personal site, or request new consent. A missing personal URL produces an actionable unavailable message; profile/drive access failures retain existing selections. The same source contract applies to Property Pane. See Microsoft's [user profile REST documentation](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/work-with-user-profiles-in-sharepoint).

### Flat native configuration and stable opening

The following SDK-facing props are legacy-only when `preset` is omitted, except for the shared functional props/callbacks explicitly listed below. They work identically in React and Property Pane; do not combine them with a preset. There is no raw SDK options bag.

Theme is automatic from the closest FluentProvider with SPFx/Teams fallback, mapped to native light/dark at Browse. There is no public theme override and no promise of arbitrary Fluent colors or high-contrast parity inside Microsoft content. The mode-specific browse icon remains visible and the animated lower border remains active while opening/open; duplicate openings are prevented.

All opening-only configuration, descriptors, arrays, theme and verified preselection are captured before async work on Browse. Updating them affects the next opening and does not close the active picker or repeat selected-reference hydration. Source/item/access policy, authoritative values/revision, disabled or disposal still invalidate stale results. Hydration resolves the accepted source independently of initial navigation.

Legacy-only example (omit `preset` deliberately):

<!-- toolkit:include symbol="FileFolderPicker" section="example" title="Configure a legacy library opening without a preset" -->

Native upload/create/device-upload may modify content before Pick. Cancel, rejected selection or Property Pane Apply/Discard cannot undo those changes. The selectable-source boundary validates accepted references; it is not a pre-write destination boundary and cannot guarantee confinement of all Microsoft-owned writes. Use a separately validated host operation if strict pre-write destination control is needed. The control never grants permissions, provisions OneDrive, logs tokens or automatically persists principal/file history.

### Permissions by enabled feature

The baseline obtains a SharePoint-audience token through AadTokenProvider. SPHttpClient metadata access and iframe bearer-token authorization are distinct; success of one does not prove the other's grants. Permissions are delegated and bounded by user rights; existing shared SPFx tenant grants may already suffice.

Microsoft also lists Graph Files.Read/Files.ReadWrite alternatives in its operation table; these do not substitute a Graph token for the SharePoint-audience token used by our hosted iframe. Do not union all scopes from the sample. Optional service pivots may request additional audiences; test the actual enabled experience before declaring a minimal deployment. The toolkit changes neither the package manifest nor tenant approvals. See the authoritative [SharePoint-audience baseline request](permissions.md#sharepoint-audience-baseline-request) for the exact consumer manifest JSON, conditional approval and alternative-scope guidance.

For a consumer choosing an all-sites read grant use AllSites.Read instead; add write or Graph requests separately only for enabled capabilities. Admin approval and user rights are still required. See [Microsoft permissions](https://learn.microsoft.com/en-us/onedrive/developer/controls/file-pickers/?view=odsp-graph-online) and [SPFx API access](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aadhttpclient).

### Explicit compatibility gates

The [current v8 schema](https://learn.microsoft.com/en-us/onedrive/developer/controls/file-pickers/v8-schema?view=odsp-graph-online) contains additional options whose presence does not establish their complete message/result contracts. Local-file pick, draft-selection payloads, custom/formatter commands, host-rendered tray/navigation, account switching and download/share/move/save results are not exposed as supported toolkit actions. No guessed callbacks or inert configuration flags are forwarded. Copy has a schema block without a copy action enum. Substrate, parent/configure protocol and SDK timeout overrides remain internal research gates. Consumer-only photos are outside SPFx Business support; site/list/listItem/Syntex domains require different result models and cannot be coerced into file/folder references. Only one file category is supported at a time; arbitrary category combinations are not exposed. Use extension filters for an explicit suffix allowlist.

Automated config/auth/lifecycle tests and browser/tenant acceptance are separate evidence. Enabling a capability does not prove service availability, tenant grants or actual internal API behavior; see the [validation guide](../development/sharepoint-validation/INDEX.md).

### Trying presets in the sample

React and Property Pane use the same preset contract and source-aware configuration. New scenarios start with explicit focused. Choose extended to exercise service-owned navigation, while retaining the same accepted source. The ordinary demo exposes functional choices rather than SDK/authentication switches; advanced diagnostics report applied settings and JSON references without becoming a second configurator. The sample Property Pane offers only focused and extended. Saved legacy sample configurations display and use focused until extended is selected; saved references and old SDK settings are retained. This sample policy does not remove the reusable control’s legacy contract. Explicit effective-source changes in the demo clear dependent starting-location settings; preset-only and cosmetic updates retain them. The reusable control leaves consumer-owned source/selection updates to the consumer.

The validation guide distinguishes route navigation, empty content, denied resources and tooling limitations. A successful token acquisition alone does not establish a working Media route or a successful file selection.

In mixed mode (`itemTypes: 'all'`), the SDK filter includes the selected semantic category **and** `folder` explicitly, so the category does not disable folder confirmation. Toolkit validation still applies the category only to files.

`FileFolderPickerProps.colorfulTagAvatars?: boolean` defaults to false. Enable it to use Fluent's colorful palette for the existing file/folder icon avatars:

<!-- toolkit:include symbol="FileFolderPicker" section="example" title="Seed file labels and show colorful selected avatars" -->

Fluent `idForColor` receives the canonical tenant/endpoint/drive/item identity, excluding display-name snapshots and enrichment. Renaming a selected file preserves its color; the finite palette does not guarantee unique colors. File/folder icons, dimensions, removal behavior and saved JSON references are retained. Changes affect only outer selected tags, preserve the open session and do not recolor Microsoft's hosted iframe, request metadata or emit selection changes. Palette colors follow the existing Fluent host theme.

`PropertyPaneFileFolderPickerProps.colorfulTagAvatars` inherits the same contract. The React demo exposes **Colorful tag avatars** under **Options**; the native **File picker options** group exposes its independent toggle. Both default off.

## Native Property Pane integration

Import the factories only from `@apvee/spfx-react-toolkit/property-pane/pickers`. React controls and JSON types remain available from `/controls/pickers`, `/controls` and the root; the root does not load native Property Pane factory runtime. The peer `@microsoft/sp-property-pane` supports SPFx >=1.18 <2, using the host's installed version. React 17 render/unmount is used, with the toolkit provider and scoped Fluent renderer/theme; no React 18 createRoot is required.

### Native pane theme

All eight native picker fields use the effective Fluent UI 9 tokens inherited by the SPFx custom-field container, including typography, spacing, borders and shadows. A private provider bridges the separate React root to body-mounted menus, tooltips and dialogs; it does not use the web-part section theme. Ordinary React pickers continue to follow their own provider.

The bridge reads from the container's owning document. Connected containers with usable opaque hex/RGB surface and foreground tokens use the native snapshot, completing missing tokens with a matching neutral light/dark Fluent theme. A temporarily unavailable snapshot retains the last valid theme for that container; before the first valid snapshot, the fallback is neutral light. No REST/Graph request or persistent cache is involved.

Native renders, host theme notifications, ancestor class/style changes, reparenting, stylesheet replacement/text changes and stylesheet loads refresh the snapshot without remounting the picker. Equivalent snapshots retain their identity. Unsignalled CSSOM-only edits such as `insertRule` in an existing stylesheet require a host notification or another native render; there is no polling. Observers, listeners and scheduled frames are released on disposal. The current anchor parser intentionally accepts opaque hex and comma-separated RGB/RGBA colors; unsupported formats retain the previous snapshot.

### Typed property values

Use the exported JSON interfaces as web-part property types. Both single and multiple fields store **arrays**; no selection is `[]`.

`PickerPropertyKey<TProperties,TValue>` accepts only direct safe array properties with the exact domain. Optional arrays are supported; scalar, mixed/untyped arrays, nested path strings and prototype-related names are rejected by types and runtime. `PropertyPanePickerProps<TProperties,TControlProps,TValue>` inherits the matching React props except renderer selected/default keys, defaultPrincipals and onOptionSelect. Native `value` is authoritative; initialOptions provides display models, availableOptions provides independent menu candidates. Models and clients are not stored in the property value.

The seven native fields inherit the [React picker naming and validation contract](../../api/controls/pickers.md#pickerprops). Supply a nonblank plain string `label`, or explicit `aria-label`/`aria-labelledby`; JSX and object labels require explicit ARIA naming. `validateSelection` validates selection, removal and clear before emission in the default change mode; a veto preserves the previous value. `validationMessage` overrides display only. `validationTrigger="blur"` lets ordinary gestures commit and validates retained models after focus exits the entire control and owned popup, without reverting those gestures. Explicit ID resolution always validates before committing. Failures use the existing error channels.

The historical `presentation: 'combobox'` value now renders Fluent Dropdown with a persistent selection button and optional popup SearchBox. Search text and query clear do not replace or clear selection; `placeholder` customizes search while the empty button uses native localized selection prompts (English: “Select an item” / “Select items”). Several selected values show a count using the loaded resource's plural grammar. TagPicker retains its input and removable tags, without duplicate selected popup rows. `colorfulTagAvatars` remains tag-only. See [Dropdown interaction and compatibility](#dropdown-selection-and-popup-search) for keyboard, asynchronous acceptance and changed DOM assumptions. Native property callbacks and Apply ownership are unchanged; acceptance in the React runtime does not acknowledge persistence by the host.

### Seven fields

Every factory is called as `Factory<TProperties>(targetProperty, props)` and returns the native `IPropertyPaneField<IPropertyPaneCustomFieldProps>` from `@microsoft/sp-property-pane`. The type parameter is the owning web part's property bag; ordinary calls infer it from `instance`. `targetProperty` must pass the matching `PickerPropertyKey<TProperties, TValue>` contract. The returned field belongs in a native `groupFields` array; rendering a factory as JSX is unsupported.

Each factory requires the initialized web-part `instance`, a stable field `key`, and `value`. Fields/View/Item source selectors and Principal scope/identity configuration have exactly the [React contracts](../../api/controls/pickers.md). Single/multiple and Dropdown/TagPicker remain configurable for every field. Persona rows/icons, consumer/personal ordering, five ordinary initial suggestions for Site/Item/Principal (Site contextual choices are separate), up to 30 initial scoped options for List (`initialCatalogLimit`, integer 1–200), complete catalogs with local case-insensitive search for Field/View/ContentType, read-only preflight and typed discovery are shared with React. Field/View/ContentType match locally over their complete remotely loaded catalogs; ListItemPicker retrieves remote candidates and matches locally within each bounded scan. User identity modes and explicit ensureUser remain independent of presentation and apply only to users.

`PropertyPanePickerProps<TProperties, TControlProps, TValue>` preserves discriminated React source branches, including ContentTypePicker's scope/list requirement. It removes `selectedOptions`, `defaultSelectedOptions`, `defaultPrincipals` and `onOptionSelect`; consumers configure saved references through `value`, while SPFx supplies the persistence notification. Other inherited props remain the matching React family's contract, including label/ARIA naming, validation, source restrictions, strings, error callbacks, consumer options and display seeds. The factory never serializes client instances, options or callbacks into the property value.

All seven factories use the [React search-feedback contract](#search-feedback) for both Dropdown and TagPicker, with single or multiple selection. Selected rows/tags and SitePicker's fixed current-site/actual-hub choices do not count as ordinary search results or suppress feedback. For empty or below-minimum queries, the default minimum-character guidance stays in the search placeholder without a duplicate status row. A blank nonsearchable menu without options shows the localized no-options message. An eligible query shows the localized searching message during debounce and loading. A settled complete search distinguishes no eligible matches from matches that are already selected or pinned. All default feedback follows the loaded native picker resource and optional instance `strings` overrides. Partial retrieval shows its warning without claiming a complete empty result.

The inherited `emptyContent`, `loadingContent` and `errorContent` slots customize the same fixed inner `div` regions; omission keeps state-specific defaults and `false` suppresses content while preserving the owned shell. Query and contextual-source errors remain independently visible beside healthy results. Stale query completions/errors cannot replace feedback for the current input. These presentation states do not emit property changes, remount the field, reset nonreactive drafts or alter native Apply/dispose behavior.

Open this field, type fewer than the configured minimum characters (two by default), then an eligible query with no matches. Observe the placeholder guidance, pending status and final no-results message while contextual choices stay visible. Matching only a selected or pinned choice shows the no-other-results message. Repeat after switching presentation or selection mode; saved references change only through the accepted explicit actions described below. Field/View/ContentType's local matching and ListItemPicker's remote retrieval with local matching retain the same completeness limitations as React.

### Changes, nonreactive drafts and lifecycle

The SPFx callback supplied to `onRender` receives `(targetProperty, freshValidatedReferences, true)` only after an accepted explicit selection/removal/clear/ID resolution. The field does not call setProperties, PropertyPane.refresh or the web part render method. The host owns reactive updates, nonreactive Apply and persistence. Missing native callback disables confirmation and displays a diagnostic.

Unresolved existing references are retained in complete order alongside new confirmed references. Ordinary model hydration cannot rewrite native IDs. Malformed external values are diagnosed and cannot overwrite the prior confirmed selection; the field does not automatically repair the property bag or globally veto the native Apply button. A consumer should validate its own external properties where needed.

In a nonreactive pane the control keeps a pending emitted draft while the host continues passing the same old property value. Equivalent old props do not reset that draft. Pass a changed `valueRevision` to explicitly reset/adopt value even when it equals the old props. Omitted/unchanged revisions preserve an equivalent-value draft; removing an existing revision to `undefined` is itself a revision change and adopts the authoritative value.

A changed authoritative value/revision cancels pending confirmation without remounting or restarting equivalent queries. Invalid/incompatible user IDs cannot rebind a retained alias or adopt a stale seed. Field creation may happen repeatedly: same owner/key/DOM updates reuse the React root, provider and source runtime, keeping trigger/input, query and selection stable while callbacks/options update. A stale field instance's dispose cannot destroy the newer host. Cache lifetime is owner/document scoped; disposing the last field releases its lease and clears memory. A true dispose/reopen may mount and query again. Separate roots can independently initialize Graph clients; no global initialization guarantee is made. Graph-dependent Principal operations await the current provider’s client initialization, preserving pending labels and explicit ID-resolution actions. Terminal initialization failures remain reported; SharePoint-only operations remain usable when optional Graph initialization fails. Replacing the provider cancels obsolete pending work.

The same scoped read cache coalesces identical metadata GETs across mounted native fields. A field reopening after another field refreshes or replaces its cached first page adopts the new result; it does not restore stale accumulated pages. Slow model hydration, consumer eligibility and optional suggestions retain saved labels/JSON and show pending feedback instead of a false empty or unresolved state. Field/View/ContentType can reuse a positive match from their already loaded complete filtered catalog without another entity GET; explicit refresh and distinct URI/projection contracts remain separate.

No picker-managed storage, history, localStorage, OneDrive or automatic principal saving exists. JSON reference persistence belongs to SPFx or the consumer. Personal REST suggestions remain opt-in, carry the same permissions/maintenance limits as React, and never run ensureUser.

### Initial hydration and AutoSkeleton

All seven native fields preserve available selection labels while resolving and verifying their references. A display model or the optional JSON reference `text` provides a provisional label; it does not bypass identity, source or scope verification, and hydration never rewrites the saved property.

If the mounted field initially contains selected references with no display labels/models, its selection-control area uses the existing AutoSkeleton during that initial hydration. Label, hint, validation and actions remain outside the skeleton. The same trigger/input stays mounted when loading finishes. Empty selections show the normal control immediately. With mixed known and unknown selections, known labels stay visible and unresolved pending entries use loading text.

The skeleton is limited to the initial mounted hydration cycle. Search/debounce, pagination, candidate verification, explicit ID resolution, cosmetic/equivalent updates, refresh and later value/source changes do not reactivate it. Definitive hydration failures end the skeleton and expose the normal unresolved/error state and retry behavior. A genuine native disposal/remount starts a new initial cycle. This presentation does not retain a disposed React root, prevent necessary REST calls or add storage/history.

Use `text` when a consumer already knows a label:

<!-- toolkit:include symbol="@apvee/spfx-react-toolkit/property-pane/pickers#PropertyPaneSitePicker" section="example" title="Supply a known native site-reference label" -->

`initialOptions` remains the separate way to provide full display models. No new public skeleton prop is required. AutoSkeleton retains its browser geometry/native-inert requirements and graceful fallback; see [AutoSkeleton](../../api/controls/auto-skeleton.md). Only the bounded selection-control region is measured; portaled menus use the picker's existing loading/search feedback.

### Observable validation

The sample web part's “Picker references (JSON)” group cascades Sites → Lists → Fields/Views/Items. ContentTypePicker depends on the selected site and, only in list scope, the selected list. Global picker modes configure presentation and single/multiple selection. Cascading sources require exactly one parent reference: the selected collection URL targets Lists, and the selected list GUID plus web URL target list-dependent fields. Lists is disabled without a selected collection; list-dependent fields appear only with a matching list reference. No manual list-title field or fabricated fallback list is used. Changing a parent identity clears dependent references, increments valueRevision and refreshes the consumer-owned pane configuration with stable field keys. Site changes clear content types in both scopes; list changes clear them only in list scope. Display-only changes preserve values. PrincipalPicker remains independent. The React “SharePoint pickers” panel exposes live source settings and diagnostic JSON. Automated native tests use actual provider/hooks/Fluent with mocked network boundaries to observe root/input preservation, callback replacement, reset cancellation, unresolved ordering, malformed data rejection, identity conflicts and final dispose cleanup. These local results do not establish authenticated Apply behavior or tenant permissions; follow [SharePoint validation](../development/sharepoint-validation/INDEX.md).

The four catalog fields share React lazy loading, consumer precedence and selected-exclusion rules. Only List uses the configurable limit/lookahead hint. Field/View/ContentType follow every REST page, show all matching options and filter their cached catalog locally without typed REST requests; none of these three accepts initialCatalogLimit. Single Dropdown fields expose the localized selected-item group first, with a checked current value even during search; clicking it preserves selection, selecting another replaces it, and clearable clears it. No duplicate of the selected identity appears among candidates. Multiple Dropdown fields expose a separate localized selected-items group, including unresolved references, for deselection without a query. Pending hydration/preflight displays loading; definitive unresolved references stay diagnosed and preserved. Site/List/Field/View/ContentType/Item searches are case insensitive, including equivalent-query cache/cursor reuse; raw filters and saved identities remain unchanged. Updating the List catalog limit, available models or cosmetic props retains the mounted field and selections, reusing equivalent memory-cached source pages.

`PropertyPaneSitePicker` forwards flat `includeCurrentSite` and `includeHubSite`, both default true, to the same React source/runtime. “Sito corrente” and an actually associated “Hub associato” remain visible before and during search, after the selected section and outside the ordinary five-item cap. Consumer models take precedence; canonical duplicates in consumer/search results and selected identities are excluded. A hub without an association is not its own associated hub. The current collection root and actual hub association use cached read-only metadata; modes/KQL still constrain eligibility. Independent flags or showDefaultSuggestions suppress retrieval; errors report their SharePoint suggestionSource. Flag/cosmetic updates preserve the mounted root, saved references and equivalent cached queries. For example: `PropertyPaneSitePicker('sites', { instance: this, key: 'sites', label: 'Sito', value: this.properties.sites ?? [], includeCurrentSite: true, includeHubSite: false })`.

### Content type field

`PropertyPaneContentTypePicker` shares the [React ContentTypePicker contract](#contenttypepicker), including site collection versus list scope, local complete-catalog search and ID/group/visibility filters.

Single selection still stores a `ContentTypePickerValue[]`. Supply `initialOptions` for selection display models and `availableOptions` for independent consumer menu choices; neither is persisted by the control. `valueRevision` can authoritatively replace a consumer draft after changing scope/site/list. Known models retain their labels; control-only AutoSkeleton applies only to the first hydration without models. Mounted equivalent updates reuse the shared runtime, and closing the last native field releases its in-memory cache.

The test web part offers a content-type scope selector. Site CTs depend on exactly one selected site; list CTs also require exactly one selected list. A source change clears the demo's dependent CT reference; presentation changes retain values. The global **Picker modes** presentation/multiple settings apply to this seventh field as well. Scope changes are consumer configuration, not a mutation of SharePoint content-type definitions.

All seven native picker factories inherit `colorfulTagAvatars?: boolean` from `PickerProps`. Default false/omitted preserves current neutral avatars. With `presentation: 'tagPicker'` and `colorfulTagAvatars: true`, resolved selected-tag avatars use Fluent's colorful palette and stable selected renderer keys. Photos/icons are retained, suggestions stay unchanged, and unresolved tags without media remain without avatars. See the [React contract](#colorful-selected-tag-avatars).

The flag is display configuration; JSON values and SPFx change/apply semantics are unchanged. Cosmetic updates preserve the mounted field, typed query and request cache. The sample **Picker modes** group provides a default-off toggle applying to all seven fields.

## Native file reference configuration

The native field follows the [Property Pane token bridge](#native-pane-theme), independently of the web-part section theme. Its body-mounted Stage dialog uses the same native tokens. The existing File Picker SDK receives `default` or `dark` from that theme on each opening; an already-open Microsoft iframe is not reskinned live. Renderer, direction, focus, selections and SDK ownership remain unchanged.

The native field renders the same outlined tag surface as the React control: individually removable tags, an empty selection action and a mode-specific icon (Attach, FolderOpen or FolderDocument). It has no separate Browse/clear-all buttons. Medium/Large use the native Fluent Input border and 32px/40px minimum height; tags can wrap in a narrow pane. The focus underline remains active while the hosted picker is open. Labels, hint/error associations, disabled/error styling, host theme and reduced motion follow the shared control. This presentation does not change JSON ownership, metadata leases or hydration behavior.

Import the native factory separately from the React control:

<!-- toolkit:include symbol="@apvee/spfx-react-toolkit/property-pane/file-folder-picker#PropertyPaneFileFolderPicker" section="example" title="Native file and folder configuration" -->

The persisted value is always a JSON array of `FileFolderPickerValue`, including in single mode. `FileFolderPickerPropertyKey<P>` accepts only safe direct properties whose non-undefined type is exactly that array domain; scalar, nested, untyped and mixed-domain properties are rejected. Drive-item references are separate from the existing SharePoint list-item picker values.

The factory call returns `IPropertyPaneField<IPropertyPaneCustomFieldProps>` for a native `groupFields` array. The array described above is the persisted property value supplied through its native change callback, rather than the factory's return type. Optional and readonly reference-array properties are supported; prototype-related target names are forbidden by types and runtime.

The factory requires `instance`, a stable `key` and the current `value`. Optional `valueRevision` marks an authoritative external reset, including resets to an array equivalent to a previous host value. Remaining options follow [FileFolderPicker](../../api/controls/file-folder-picker.md), excluding `value`, `defaultValue` and `onChange` from the React contract. SPFx supplies the native change callback; absent callbacks disable mutations with a diagnostic.

`initialOptions` supplies selected display models independently of saved references. Labels already available remain visible on reopening. Initially unresolved selections without labels use the shared AutoSkeleton boundary only until hydration settles. Browsing, subsequent updates and failed hydration do not repeatedly hide existing values.

The field composes the existing WebPart provider, native mount/update/dispose layer, renderer/theme/direction host and owner/document cache lease. Repeated native renders with equivalent source/value props retain the React root, pending draft and opened picker. A nonreactive pane can repeat its old property bag without discarding an accepted draft. Authoritative changes cancel pending confirmation and never emit a hydration write.

Closing the last native field releases its in-memory metadata cache and cancels its owned remote session. Reopening creates a fresh cache and uses the consumer's saved JSON values. No token storage, selection history or browser persistence is introduced. Use stable array references where practical, but equivalent recreated arrays must not cause an unnecessary remount.

The new factory uses a small direct-array value bridge because the older bridge requires `PickerValue` renderer aliases and Combobox selection events. Authentication, metadata clients, scope resolution and native hosting remain shared toolkit capabilities. Source, kind, category, extension, cardinality and returned-item restrictions are identical between React and Property Pane.

See [SharePoint validation](../development/sharepoint-validation/INDEX.md) for required real-host dialog, scope, authentication and reopen checks. Local simulations do not establish tenant permissions or Microsoft's internal API traffic.

### Dialog and OneDrive choices

`PropertyPaneFileFolderPickerProps` inherits the React source, item policy, display seeds, native Field slots and onError contracts. These native members replace React selection ownership:

See {@link @apvee/spfx-react-toolkit/property-pane/file-folder-picker#PropertyPaneFileFolderPickerProps | the native ownership members} for initialized instance, stable field key, authoritative JSON value and revision, while inherited React options retain their own contracts.

The native field always opens the same responsive 1080 × 720 dialog without an outer visible title, close button or padding. No `presentation` option remains. `scope: { kind: 'oneDrive' }` resolves the authenticated user’s OneDrive read-only without an URL property; optional `folder` restricts that drive. Source changes still belong to the consumer and do not rewrite saved references automatically.

### Presets and opening feedback

The native factory accepts the same explicit `preset: 'focused' | 'extended'` and functional props as [React](#presets-the-simple-configuration), with the same resolver, defaults and validation. Graph/claims handling is internal to the preset and uses existing provider/grants only. Upload and create-folder remain explicit false-by-default intents. A preset does not change accepted source or saved JSON values.

Omitting preset preserves the deprecated legacy SDK configuration. Defined legacy SDK flags cannot be mixed with a preset; choose migration explicitly. Existing web-part preferences are not silently reinterpreted, and draft/confirmed references survive presentation/preset updates. Use `hint` for the description: omission renders none; opening never adds automatic status text or replaces the browse icon with a spinner. The animated lower border and internal busy state remain. Opening-only changes apply on the next Browse without replacing the mounted root, losing the native draft or repeating hydration. Source/policy or authoritative revision changes still cancel stale work. The final native lease disposal clears mounted metadata cache; a later pane opening hydrates with a fresh cache and preserves saved JSON references/available labels.

Optional native writes are immediate: cancelling the dialog or discarding/applying pane changes does not undo uploaded files or created folders. Persist only FileFolderPickerValue arrays; transient capabilities and callbacks never belong in the web-part property value. The sample exposes preset and explicit write intents without SDK authentication switches or tenant permission changes.

### Navigation and diagnostics

The native sample uses focused/extended with the same effective recipe as React. Focused hides native navigation; extended requests the complete source-aware navigation set. Neither changes accepted provenance. The sample clears optional initial navigation when its effective site/library/folder source changes, while preserving it for equivalent source identities, preset-only updates and unrelated OneDrive dependencies. Microsoft can display additional entries or return empty content; those outcomes need real-host verification.

Extended SharePoint browsing can lazily authenticate the exact independently verified current-user OneDrive origin. Focused does not enable auxiliary personal discovery. Existing legacy callers retain their explicit pivot-based rule. Auxiliary failures stay local to the hosted command and cannot replace a source Field error. No automatic fallback broadens scope or changes preset.

With SPFx `debug=true`, the shared implementation emits bounded internal `console.debug` authentication categories without tokens, claims, raw resource URLs or provider/native payloads. These are separate from the sanitized actionable `onError` contract; optional audience/Graph failures do not become a false configured-source failure. They introduce no public configuration member, automatic hint or spinner. Empty routes, unsupported source combinations, host policy denials, server authorization failures and unavailable browser automation targets require separate validation outcomes.

Native uploads and folder creation are immediate writes. The accepted-reference scope check is not a pre-write destination boundary, and cancelling or discarding a native draft cannot undo those operations.

### Sample Preset menu

The test web part offers only **Focused** and **Extended** in its Property Pane. Saved legacy demo configurations use Focused for both the displayed choice and the native control. Changing the preset preserves selected references, the field key and its value revision; old saved SDK fields are retained without passing them alongside a preset. The reusable Property Pane factory still supports the documented legacy contract when `preset` is omitted.

### File category policy

The native factory accepts the same optional `fileCategory: 'images' | 'video' | 'audio' | 'media' | 'documents'` as [React](#file-categories), in preset and legacy configurations. Omission imposes no category restriction. A category cannot be combined with nonempty `extensions` or folders-only mode; mixed mode filters files while retaining folders. The toolkit checks authoritative metadata suffixes against a conservative catalog, independently of the hosted category filter, without content or MIME inspection. It does not promise exact equivalence with Microsoft's category catalog.

The sample's **File category** menu clears incompatible custom extensions and disables their input while a category is active. Choosing folders-only clears the category. Category changes preserve saved references and the native value revision; pending picks still follow the shared policy invalidation rules. Source changes continue to use the sample's explicit reset behavior.

`PropertyPaneFileFolderPickerProps.colorfulTagAvatars?: boolean` inherits the [React FileFolderPicker contract](#colorful-selected-tag-avatars). Default false/omitted preserves current icon avatars. True colors those existing avatars through Fluent's palette using the canonical saved identity; filenames can change without changing color. It preserves icons, sizes, the native field/session, saved references and the typed change callback. It does not style Microsoft's remote iframe or start metadata/authentication work.

The native sample has a separate **Colorful file/folder tag avatars** toggle under **File picker options**, default off.

### Complete native host configurations

The source-owned six-array recipe supplies a real initialized web part and returns native fields from groupFields. The [native file example above](#native-file-and-folder-configuration) separately configures focused PDF selection. These independent fields do not implement the sample's parent-reset policy; apply that policy deliberately as described above.

<!-- toolkit:include symbol="@apvee/spfx-react-toolkit/property-pane/pickers#PropertyPaneSitePicker" section="example" title="Configure six native reference arrays" -->

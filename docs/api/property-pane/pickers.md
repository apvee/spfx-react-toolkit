# Native Property Pane pickers

Import the factories only from `@apvee/spfx-react-toolkit/property-pane/pickers`. React controls and JSON types remain available from `/controls/pickers`, `/controls` and the root; the root does not load native Property Pane factory runtime. The peer `@microsoft/sp-property-pane` supports SPFx >=1.18 <2, using the host's installed version. React 17 render/unmount is used, with the toolkit provider and scoped Fluent renderer/theme; no React 18 createRoot is required.

## Native pane theme

All eight native picker fields use the effective Fluent UI 9 tokens inherited by the SPFx custom-field container, including typography, spacing, borders and shadows. A private provider bridges the separate React root to body-mounted menus, tooltips and dialogs; it does not use the web-part section theme. Ordinary React pickers continue to follow their own provider.

The bridge reads from the container's owning document. Connected containers with usable opaque hex/RGB surface and foreground tokens use the native snapshot, completing missing tokens with a matching neutral light/dark Fluent theme. A temporarily unavailable snapshot retains the last valid theme for that container; before the first valid snapshot, the fallback is neutral light. No REST/Graph request or persistent cache is involved.

Native renders, host theme notifications, ancestor class/style changes, reparenting, stylesheet replacement/text changes and stylesheet loads refresh the snapshot without remounting the picker. Equivalent snapshots retain their identity. Unsignalled CSSOM-only edits such as `insertRule` in an existing stylesheet require a host notification or another native render; there is no polling. Observers, listeners and scheduled frames are released on disposal. The current anchor parser intentionally accepts opaque hex and comma-separated RGB/RGBA colors; unsupported formats retain the previous snapshot.

## Typed property values

Use the exported JSON interfaces as web-part property types. Both single and multiple fields store **arrays**; no selection is `[]`.

```ts
import type {
  SitePickerValue, ListPickerValue, FieldPickerValue, ViewPickerValue,
  ListItemPickerValue, PrincipalPickerValue,
} from '@apvee/spfx-react-toolkit/controls/pickers';
import {
  PropertyPaneSitePicker, PropertyPaneListPicker, PropertyPaneFieldPicker,
  PropertyPaneViewPicker, PropertyPaneListItemPicker, PropertyPanePrincipalPicker,
} from '@apvee/spfx-react-toolkit/property-pane/pickers';

interface Properties {
  sites?: SitePickerValue[];
  lists?: ListPickerValue[];
  fields?: FieldPickerValue[];
  views?: ViewPickerValue[];
  items?: ListItemPickerValue[];
  principals?: PrincipalPickerValue[];
}
// Inside a BaseClientSideWebPart<Properties> configuration:
const groupFields = [
  PropertyPaneSitePicker('sites', { instance: this, key: 'sites-field', value: this.properties.sites ?? [], label: 'Site' }),
  PropertyPaneListPicker('lists', { instance: this, key: 'lists-field', value: this.properties.lists ?? [], label: 'List' }),
  PropertyPaneFieldPicker('fields', { instance: this, key: 'fields-field', value: this.properties.fields ?? [], list: 'Projects', label: 'Fields' }),
  PropertyPaneViewPicker('views', { instance: this, key: 'views-field', value: this.properties.views ?? [], list: 'Projects', label: 'View' }),
  PropertyPaneListItemPicker('items', { instance: this, key: 'items-field', value: this.properties.items ?? [], list: 'Projects', textField: 'Title', label: 'Items' }),
  PropertyPanePrincipalPicker('principals', { instance: this, key: 'principals-field', value: this.properties.principals ?? [], label: 'Principals',
    principalTypes: ['user', 'microsoft365Group'], userIdentity: { mode: 'both', ensureUser: false } }),
];
```

`PickerPropertyKey<TProperties,TValue>` accepts only direct safe array properties with the exact domain. Optional arrays are supported; scalar, mixed/untyped arrays, nested path strings and prototype-related names are rejected by types and runtime. `PropertyPanePickerProps<TProperties,TControlProps,TValue>` inherits the matching React props except renderer selected/default keys, defaultPrincipals and onOptionSelect. Native `value` is authoritative; initialOptions provides display models, availableOptions provides independent menu candidates. Models and clients are not stored in the property value.

## Seven fields

| Factory | JSON domain | Default presentation / selection |
| --- | --- | --- |
| PropertyPaneSitePicker | SitePickerValue | Combobox / single |
| PropertyPaneListPicker | ListPickerValue | Combobox / single |
| PropertyPaneFieldPicker | FieldPickerValue | Combobox / multiple |
| PropertyPaneViewPicker | ViewPickerValue | Combobox / single |
| PropertyPaneContentTypePicker | ContentTypePickerValue | Combobox / single |
| PropertyPaneListItemPicker | ListItemPickerValue | TagPicker / multiple |
| PropertyPanePrincipalPicker | PrincipalPickerValue | TagPicker / multiple |

Each factory requires the initialized web-part `instance`, a stable field `key`, and `value`. Fields/View/Item source selectors and Principal scope/identity configuration have exactly the [React contracts](../controls/pickers.md). Single/multiple and Combobox/TagPicker remain configurable for every field. Persona rows/icons, consumer/personal ordering, five ordinary initial suggestions for Site/Item/Principal (Site contextual choices are separate), up to 30 initial scoped options for List (`initialCatalogLimit`, integer 1–200), complete catalogs with local case-insensitive search for Field/View/ContentType, read-only preflight and typed discovery are shared with React. Field/View/ContentType match locally over their complete remotely loaded catalogs; ListItemPicker retrieves remote candidates and matches locally within each bounded scan. User identity modes and explicit ensureUser remain independent of presentation and apply only to users.

## Search feedback

All seven factories use the [React search-feedback contract](../controls/pickers.md#search-feedback) for both Combobox and TagPicker, with single or multiple selection. Selected rows/tags and SitePicker's fixed current-site/actual-hub choices do not count as ordinary search results or suppress feedback. Nonempty text below `minimumSearchLength` shows the minimum-character prompt even beside suggestions; a blank menu without options also shows that prompt when searchable. An eligible query shows “Ricerca in corso…” during debounce and loading. A settled complete search shows “Nessun risultato per questa ricerca” for no eligible matches, or “Nessun altro risultato per questa ricerca” when only selected/pinned identities match. Partial retrieval shows its warning without claiming a complete empty result.

The inherited `emptyContent`, `loadingContent` and `errorContent` slots customize the same fixed inner `div` regions; omission keeps state-specific defaults and `false` suppresses content while preserving the owned shell. Query and contextual-source errors remain independently visible beside healthy results. Stale query completions/errors cannot replace feedback for the current input. These presentation states do not emit property changes, remount the field, reset nonreactive drafts or alter native Apply/dispose behavior.

```ts
PropertyPaneSitePicker('sites', {
  instance: this, key: 'sites-field', value: this.properties.sites ?? [],
  label: 'Site', minimumSearchLength: 3, includeCurrentSite: true, includeHubSite: true,
});
```

Open this field, type fewer than three characters, then an eligible query with no matches. Observe the prompt, pending status and final no-results message while contextual choices stay visible. Matching only a selected or pinned choice shows the no-other-results message. Repeat after switching presentation or selection mode; saved references change only through the accepted explicit actions described below. Field/View/ContentType's local matching and ListItemPicker's remote retrieval with local matching retain the same completeness limitations as React.

## Changes, nonreactive drafts and lifecycle

The SPFx callback supplied to `onRender` receives `(targetProperty, freshValidatedReferences, true)` only after an accepted explicit selection/removal/clear/ID resolution. The field does not call setProperties, PropertyPane.refresh or the web part render method. The host owns reactive updates, nonreactive Apply and persistence. Missing native callback disables confirmation and displays a diagnostic.

Unresolved existing references are retained in complete order alongside new confirmed references. Ordinary model hydration cannot rewrite native IDs. Malformed external values are diagnosed and cannot overwrite the prior confirmed selection; the field does not automatically repair the property bag or globally veto the native Apply button. A consumer should validate its own external properties where needed.

In a nonreactive pane the control keeps a pending emitted draft while the host continues passing the same old property value. Equivalent old props do not reset that draft. Pass a changed `valueRevision` (string/number) to explicitly reset/adopt value even when it equals the old props:

```ts
PropertyPaneSitePicker('sites', {
  instance: this, key: 'sites-field', value: this.properties.sites ?? [],
  valueRevision: resetRevision, label: 'Site', availableOptions: consumerSites,
});
```

A changed authoritative value/revision cancels pending confirmation without remounting or restarting equivalent queries. Invalid/incompatible user IDs cannot rebind a retained alias or adopt a stale seed. Field creation may happen repeatedly: same owner/key/DOM updates reuse the React root, provider and source runtime, keeping input/query and selection stable while callbacks/options update. A stale field instance's dispose cannot destroy the newer host. Cache lifetime is owner/document scoped; disposing the last field releases its lease and clears memory. A true dispose/reopen may mount and query again. Separate roots can independently initialize Graph clients; no global initialization guarantee is made. Graph-dependent Principal operations await the current provider’s client initialization, preserving pending labels and explicit ID-resolution actions. Terminal initialization failures remain reported; SharePoint-only operations remain usable when optional Graph initialization fails. Replacing the provider cancels obsolete pending work.

The same scoped read cache coalesces identical metadata GETs across mounted native fields. A field reopening after another field refreshes or replaces its cached first page adopts the new result; it does not restore stale accumulated pages. Slow model hydration, consumer eligibility and optional suggestions retain saved labels/JSON and show pending feedback instead of a false empty or unresolved state. Field/View/ContentType can reuse a positive match from their already loaded complete filtered catalog without another entity GET; explicit refresh and distinct URI/projection contracts remain separate.

No picker-managed storage, history, localStorage, OneDrive or automatic principal saving exists. JSON reference persistence belongs to SPFx or the consumer. Personal REST suggestions remain opt-in, carry the same permissions/maintenance limits as React, and never run ensureUser.

## Initial hydration and AutoSkeleton

All seven native fields preserve available selection labels while resolving and verifying their references. A display model or the optional JSON reference `text` provides a provisional label; it does not bypass identity, source or scope verification, and hydration never rewrites the saved property.

If the mounted field initially contains selected references with no display labels/models, its input area uses the existing AutoSkeleton during that initial hydration. Label, hint, validation and actions remain outside the skeleton. The same input stays mounted when loading finishes. Empty selections show the normal input immediately. With mixed known and unknown selections, known labels stay visible and unresolved pending entries use loading text.

The skeleton is limited to the initial mounted hydration cycle. Search/debounce, pagination, candidate verification, explicit ID resolution, cosmetic/equivalent updates, refresh and later value/source changes do not reactivate it. Definitive hydration failures end the skeleton and expose the normal unresolved/error state and retry behavior. A genuine native disposal/remount starts a new initial cycle. This presentation does not retain a disposed React root, prevent necessary REST calls or add storage/history.

Use `text` when a consumer already knows a label:

```tsx
PropertyPaneSitePicker('sites', {
  instance: this,
  key: 'sites',
  label: 'Sito',
  hint: 'Seleziona una raccolta siti',
  value: [{ kind: 'site', url: 'https://contoso.sharepoint.com/sites/marketing', text: 'Marketing' }],
});
```

`initialOptions` remains the separate way to provide full display models. No new public skeleton prop is required. AutoSkeleton retains its browser geometry/native-inert requirements and graceful fallback; see [AutoSkeleton](../controls/auto-skeleton.md). Only the bounded input region is measured; portaled menus use the picker's existing loading/search feedback.

## Observable validation

The sample web part's “Picker references (JSON)” group cascades Sites → Lists → Fields/Views/Items. ContentTypePicker depends on the selected site and, only in list scope, the selected list. Global picker modes configure presentation and single/multiple selection. Cascading sources require exactly one parent reference: the selected collection URL targets Lists, and the selected list GUID plus web URL target list-dependent fields. Lists is disabled without a selected collection; list-dependent fields appear only with a matching list reference. No manual list-title field or fabricated fallback list is used. Changing a parent identity clears dependent references, increments valueRevision and refreshes the consumer-owned pane configuration with stable field keys. Site changes clear content types in both scopes; list changes clear them only in list scope. Display-only changes preserve values. PrincipalPicker remains independent. The React “SharePoint pickers” panel exposes live source settings and diagnostic JSON. Automated native tests use actual provider/hooks/Fluent with mocked network boundaries to observe root/input preservation, callback replacement, reset cancellation, unresolved ordering, malformed data rejection, identity conflicts and final dispose cleanup. These local results do not establish authenticated Apply behavior or tenant permissions; follow [SharePoint validation](../../SHAREPOINT-VALIDATION.md).

The four catalog fields share React lazy loading, consumer precedence and selected-exclusion rules. Only List uses the configurable limit/lookahead hint. Field/View/ContentType follow every REST page, show all matching options and filter their cached catalog locally without typed REST requests; none of these three accepts initialCatalogLimit. Single Combobox fields expose **Selezionato** first, with a checked current value even during search; clicking it preserves selection, selecting another replaces it, and clearable clears it. No duplicate of the selected identity appears among candidates. Multiple Combobox fields expose a separate **Selezionati** group, including unresolved references, for deselection without a query. Pending hydration/preflight displays loading; definitive unresolved references stay diagnosed and preserved. Site/List/Field/View/ContentType/Item searches are case insensitive, including equivalent-query cache/cursor reuse; raw filters and saved identities remain unchanged. Updating the List catalog limit, available models or cosmetic props retains the mounted field and selections, reusing equivalent memory-cached source pages.


`PropertyPaneSitePicker` forwards flat `includeCurrentSite` and `includeHubSite`, both default true, to the same React source/runtime. “Sito corrente” and an actually associated “Hub associato” remain visible before and during search, after the selected section and outside the ordinary five-item cap. Consumer models take precedence; canonical duplicates in consumer/search results and selected identities are excluded. A hub without an association is not its own associated hub. The current collection root and actual hub association use cached read-only metadata; modes/KQL still constrain eligibility. Independent flags or showDefaultSuggestions suppress retrieval; errors report their SharePoint suggestionSource. Flag/cosmetic updates preserve the mounted root, saved references and equivalent cached queries. For example: `PropertyPaneSitePicker('sites', { instance: this, key: 'sites', label: 'Sito', value: this.properties.sites, includeCurrentSite: true, includeHubSite: false })`.


## Content type field

`PropertyPaneContentTypePicker` shares the [React ContentTypePicker contract](../controls/pickers.md#contenttypepicker), including site collection versus list scope, local complete-catalog search and ID/group/visibility filters.

```ts
import type { ContentTypePickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
import { PropertyPaneContentTypePicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';

interface Properties { contentTypes?: ContentTypePickerValue[]; }
// Inside BaseClientSideWebPart<Properties>:
PropertyPaneContentTypePicker('contentTypes', {
  instance: this, key: 'content-types', value: this.properties.contentTypes ?? [],
  label: 'Content types', scope: 'site', presentation: 'combobox', multiselect: false,
});
PropertyPaneContentTypePicker('contentTypes', {
  instance: this, key: 'list-content-types', value: this.properties.contentTypes ?? [],
  label: 'List content types', scope: 'list', list: 'Documents',
});
```

Single selection still stores a `ContentTypePickerValue[]`. Supply `initialOptions` for selection display models and `availableOptions` for independent consumer menu choices; neither is persisted by the control. `valueRevision` can authoritatively replace a consumer draft after changing scope/site/list. Known models retain their labels; input-only AutoSkeleton applies only to the first hydration without models. Mounted equivalent updates reuse the shared runtime, and closing the last native field releases its in-memory cache.

The test web part offers a content-type scope selector. Site CTs depend on exactly one selected site; list CTs also require exactly one selected list. A source change clears the demo's dependent CT reference; presentation changes retain values. The global **Picker modes** presentation/multiple settings apply to this seventh field as well. Scope changes are consumer configuration, not a mutation of SharePoint content-type definitions.

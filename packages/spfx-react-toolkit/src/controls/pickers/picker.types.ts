import type { PickersStringsOverride } from './pickers-strings.types';
import type { FieldProps, ComboboxProps, TagPickerProps, TooltipProps } from '@fluentui/react-components';
import type { SPFxPnPListSelector } from '../../services/spfx-pnp-list.service';
import type { ContentSlot } from '../content-slot.types.internal';

import type { PickerValue, SitePickerValue, ListPickerValue, FieldPickerValue, ViewPickerValue, ContentTypePickerValue, ListItemPickerValue, PrincipalPickerValue } from './picker-values.types';

/** Stable renderer alias plus plain display label and authoritative JSON value.
 * Generated aliases are scoped; verified canonical identity is derived separately from value.  * @example const option: PickerOption<SitePickerValue> = { key: 'hr', text: 'HR', value: { kind: 'site', url: 'https://tenant.sharepoint.com/sites/hr' } };
 */
export interface PickerOption<TValue extends PickerValue = PickerValue> {
  /** Required stable renderer alias. It must never be reused for a different scoped entity. */
  readonly key: string;
  /** Required plain display label; not an authoritative entity identifier. */
  readonly text: string;
  /** Minimal JSON reference. Must agree with the option identity; other metadata is never persisted automatically. */
  readonly value: TValue;
}
/** Native Fluent picker event (historical Combobox-compatible union) passed for explicit picker actions.
 * @example const handle = (event: PickerSelectionEvent) => console.log(event.type);
 */
export type PickerSelectionEvent =
  | Parameters<NonNullable<ComboboxProps['onOptionSelect']>>[0]
  | Parameters<NonNullable<TagPickerProps['onOptionSelect']>>[0];
/** Selected models may be unresolved: selectedOptions remains authoritative and is never silently dropped.  * @example const handle = (_event: PickerSelectionEvent, data: PickerSelectionData<SitePickerOption>) => console.log(data.selectedOptions, data.unresolvedOptions);
 */
export interface PickerSelectionData<T extends PickerOption> {
  /** Explicit user action; hydration/model updates emit no selection event. */
  readonly reason: 'selection' | 'remove' | 'clear' | 'resolve';
  /** Complete ordered renderer aliases after the explicit action, including unresolved selections. */
  readonly selectedOptions: readonly string[];
  /** Resolved models in selection order; unresolved aliases have no model in this array. */
  readonly selectedItems: readonly T[];
  /** Serializable references of resolved models; unresolvedOptions remain explicit and are not silently discarded. */
  readonly selectedValues: readonly T['value'][];
  /** Selected renderer aliases whose models could not be resolved; retain them until an explicit removal. */
  readonly unresolvedOptions: readonly string[];
  /** Alias affected by a selection or removal; omitted for whole-selection actions. */
  readonly optionValue?: string;
  /** Display text for the affected resolved model, when available. */
  readonly optionText?: string;
}
/** Native Fluent Field label/hint/validation slots keep their peer contracts.
 * Rendering requires a nonblank plain string label or explicit `aria-label`/`aria-labelledby`;
 * JSX and object labels require explicit ARIA naming even when they display text.
 * The renderer throws TypeError when neither naming input is supplied.
 * Single Dropdown offers native clearable selection through the validated clear event; multiple Dropdown deselects options and TagPicker dismisses tags. No default separate bulk-clear button is rendered.
 * @example const props: PickerProps<SitePickerOption> = { label: 'Site', presentation: 'combobox', availableOptions: [] };
 * @example
 * ```tsx
 * <SitePicker label={<span>Site</span>} aria-label="Site" />
 * ```
 */
export interface PickerProps<T extends PickerOption> extends SearchPickerOptions, Pick<FieldProps,
  'label' | 'hint' | 'required' | 'validationMessage' | 'validationState'> {
  /** Optional instance UI dictionary. Known nested leaves merge with native SPFx resources; direct labels/content/Field messages win. Undefined retains native defaults; empty strings remain explicit. */
  readonly strings?: PickersStringsOverride;
  /** Presentation only; changing this retains selection, data and request runtime. Default varies by picker.
   * Suggestions use Persona with name/metadata and category icon; user principals retain their user avatar.
   * The historical combobox value renders Fluent Dropdown: a persistent selection button and, when searchable,
   * a separate popup SearchBox. One selection displays its name/media; multiple selections display a native localized count.
   * Dropdown menus retain checked selections in native localized single/multiple selected-item groups, independent of query.
   * Clicking the current single value keeps it; native clearable clears it. Accepted single selection closes the popup;
   * asynchronous refusal preserves it. TagPicker keeps removable tags and its input, excluding selected identities from popup rows.
   * No dropdown presentation value is exposed. */
  readonly presentation?: 'combobox' | 'tagPicker';
  /** Color resolved selected-tag avatars with Fluent's colorful palette. Defaults to false.
   * The default palette uses the theme brand color for entity icons and neutral user initials.
   * Colors use the selected renderer key, retaining their identity across label changes.
   * Photos/icons and avatar dimensions are preserved; suggestion avatars, Dropdown and unresolved tags are unchanged.
   * Property Pane counterparts inherit this option. It never changes selection, queries or saved values.
   * @example <PrincipalPicker label="People" presentation="tagPicker" colorfulTagAvatars />
   */
  readonly colorfulTagAvatars?: boolean;
  /** Enable popup SearchBox in Dropdown or the TagPicker query input. Defaults to true; false omits popup search. */
  readonly searchable?: boolean;
  /** Controlled stable renderer keys (including retained aliases). Empty selection is []; undefined selects uncontrolled mode. */
  readonly selectedOptions?: readonly string[];
  /** Initial keys, read once per mount; changing them does not reset selection. */
  readonly defaultSelectedOptions?: readonly string[];
  /** Models for displaying/resolving selected keys; does not add pre-search menu suggestions.
   * Equivalent seeds avoid hydration; incoming models may enrich selected snapshots without changing keys. */
  readonly initialOptions?: readonly T[];
  /** Reactive pre-search candidates fetched from any consumer source. Defaults to [].
   * Does not initialize selection. Updates recompose the menu without remount, selection reset or remote refetch.
   * Models override equivalent personal/search models; removed selected models remain available to the selection.
   * Canonical identities include tenant/site/list scope; labels and email alone are not canonical IDs.
   * Consumer owns persistence and retrieval errors. All candidates still undergo type/source/scope filters.
   * @example availableOptions={[{ key: 'https://tenant.sharepoint.com/sites/hr', text: 'HR', url: 'https://tenant.sharepoint.com/sites/hr', value: {kind: 'site', url: 'https://tenant.sharepoint.com/sites/hr'} }]}
   */
  readonly availableOptions?: readonly T[];
  /** Maximum pre-search suggestions for Site/Item/Principal. Default 5; List uses initialCatalogLimit; Field/View/ContentType show their complete catalogs. No automatic history or persistence. */
  readonly initialSuggestions?: PickerInitialSuggestions;
  /** Matches Fluent naming. False limits new user selection to one key; selecting another replaces it.
   * Changing to false with multiple controlled keys reports a validation conflict; existing keys are never silently dropped. */
  readonly multiselect?: boolean;
  /** Optional positive integer cap for new selections. Omission is unlimited in multi mode and one in single mode; conflicting existing selections remain visible. */
  readonly selectionLimit?: number;
  /** Called after an accepted explicit selection, removal, clear or ID resolution. Arguments are the native Fluent event and complete selection data. Not called by hydration or prop updates; return value is ignored. */
  readonly onOptionSelect?: (event: PickerSelectionEvent, data: PickerSelectionData<T>) => void;
  /** Local suggestion/search transformation; changing it recomputes the projection, without server refetch.
   * Mandatory type/scope filters, selected exclusion, deduplication and the pre-search limit apply again afterward. */
  readonly filterOptions?: (options: readonly T[]) => readonly T[];
  /** Fluent selection control size; defaults to medium. */
  readonly size?: 'medium' | 'large';
  /** Disables the control and explicit selection/resolution actions. Defaults to false. */
  readonly disabled?: boolean;
  /** Search placeholder (including picker-specific overrides); when searchable, omission uses strings.labels.minimumSearchLength(minimumSearchLength). Dropdown uses a separate native localized empty selection prompt from strings.labels.selectOne/selectMany. */
  readonly placeholder?: string;
  /** Optional CSS class on the outer field region. */
  readonly className?: string;
  /** Optional CSS class on the native Dropdown or TagPicker root. */
  readonly controlClassName?: string;
  /** Optional stable selection trigger/input DOM ID. Omission uses a generated Fluent Field association; this is not an entity ID. */
  readonly id?: string;
  /** Explicit accessible selection trigger/input name, required with JSX/object labels unless aria-labelledby is supplied.
   * Omission uses a plain string label; a supplied aria-label takes precedence over that label. */
  readonly 'aria-label'?: string;
  /** DOM ID references naming the Dropdown trigger or TagPicker input; popup search has its own name.
   * Supply this or aria-label when no nonblank plain string label is available. */
  readonly 'aria-labelledby'?: string;
  /** Literal presentation slots, fixed inner span/div; the control owns status/semantic shells. */
  readonly suggestionsHeader?: ContentSlot<'span'>;
  /** Literal div content for custom initial/below-minimum guidance and complete searches with no new choices. Omission keeps minimum-character guidance in the search placeholder only, then distinguishes no matches from matches already selected or pinned; contextual rows never suppress query feedback. False hides only presentation, preserving the status shell. */
  readonly emptyContent?: ContentSlot<'div'>;
  /** Literal div content during loading, including eligible typed-query debounce. Omission uses strings.status.searching for typed search and strings.status.loading for initial loading. False hides its content. */
  readonly loadingContent?: ContentSlot<'div'>;
  /** Literal div content for request or identity failures; omission shows the error message. Typed-query and contextual-source failures remain independent alongside healthy rows. False hides content; retry remains an independent action. */
  readonly errorContent?: ContentSlot<'div'>;
  /** Optional native Fluent Tooltip content, relationship and positioning. Omission renders no tooltip; it does not replace the accessible input label. */
  readonly tooltip?: Pick<TooltipProps, 'content' | 'relationship' | 'positioning'>;
  /** Explicit cache invalidation token; changing it retries/reloads the effective source. Omission or an unchanged token does not poll; expired or evicted entries may reload on use. */
  readonly refreshKey?: string | number;
  /** Called with the Error and transport/operation metadata for failed asynchronous work; suggestionSource identifies a personal or contextual Site source. Return value is ignored; consumer options remain independent. */
  readonly onError?: (error: Error, details: {
  /** Required source or transport discriminator; SharePoint identities and Entra identities remain distinct. */
  readonly source: 'sharePoint' | 'graph';
  /** Failed picker operation: load, search, resolve or selection validation. */
  readonly operation: 'load' | 'search' | 'resolve' | 'validate';
  /** Optional failed personal or contextual Site suggestion source; absent for ordinary REST query failures. */
  readonly suggestionSource?: 'people' | 'memberOf' | 'currentUser' | 'currentSite' | 'hubSite' }) => void;
  /** Validates resolved models before an action commits, according to validationTrigger.
   * Undefined/empty string accepts; a nonempty message blocks the action without changing selection
   * or calling onOptionSelect. Promises are awaited. Throws/rejections report through onError
   * and the failure UI without committing. Omission adds no custom validation.
   * External validationMessage overrides only the displayed message, not this acceptance decision. */
  readonly validateSelection?: (items: readonly T[]) => string | undefined | Promise<string | undefined>;
  /** Defaults to change: validates selection, removal and clear before committing.
   * Blur instead validates retained models after focus exits the entire control and owned popup when no confirmation is pending;
   * it does not revoke an existing selection. Explicit ID resolution, when available, always
   * validates before committing, including in blur mode. Consumer Promises are not cancelled. */
  readonly validationTrigger?: 'change' | 'blur';
}
/** Shared pre-search policy for every picker and Property Pane counterpart.
 * @example const policy: PickerInitialSuggestions = { limit: 5 };
 */
export interface PickerInitialSuggestions {
  /** Total pre-search suggestions, integer 1–5, default 5. Search results use suggestionsLimit instead. */
  readonly limit?: 1 | 2 | 3 | 4 | 5;
}
/** Optional source web, inheriting the current SPFx provider web when omitted.
 * @example const source: WebPickerSource = { webUrl: 'https://tenant.sharepoint.com/sites/hr' };
 */
export interface WebPickerSource {
  /** Absolute HTTP(S) source web URL. Omission, when allowed, uses the provider current web; never inferred from a display label. */
  readonly webUrl?: string; }
/** Typed search and pre-search visibility policy, shared across React and native fields. Field/View/ContentType match locally over complete cached catalogs; other domains use remote discovery.
 * @example const search: SearchPickerOptions = { minimumSearchLength: 2, debounceMs: 300, suggestionsLimit: 50 };
 */
export interface SearchPickerOptions {
  /** Defaults to 2 for all pickers; shorter nonempty inputs show the minimum-character prompt and never initiate typed search. Values below 2 are invalid. */
  readonly minimumSearchLength?: number;
  /** Debounce for typed searches, including local Field/View/ContentType matching; default 300 ms. */
  readonly debounceMs?: number;
  /** Remote search result page size; independent of the pre-search suggestion or catalog limit. Field/View/ContentType return all local matches and ignore this paging size. */
  readonly suggestionsLimit?: number;
  /** Shows pre-search consumer/personal sections and the List/Field/View/ContentType catalog. Defaults to true; false suppresses their lazy eligibility/retrieval, without disabling typed search. */
  readonly showDefaultSuggestions?: boolean;
}
/** Typed server sort field and optional direction; direction defaults to ascending.
 * @example const order: PickerOrderBy<'title'> = { field: 'title', direction: 'asc' };
 */
export interface PickerOrderBy<T extends string> {
  /** Required server sort field supported by this picker. */
  readonly field: T;
  /** Optional server sort direction; defaults to asc. */
  readonly direction?: 'asc' | 'desc'; }

/** Key: normalized absolute site collection URL, including hub-mode results without siteId.  * @example const option: SitePickerOption = { key: 'hr', text: 'HR', url: 'https://tenant.sharepoint.com/sites/hr', value: { kind: 'site', url: 'https://tenant.sharepoint.com/sites/hr' } };
 */
export interface SitePickerOption extends PickerOption<SitePickerValue> {
  /** Required normalized absolute site collection URL, agreeing with value.url. */
  readonly url: string;
  /** Optional verified site collection GUID; omitted when the discovery endpoint does not return it. */
  readonly siteId?: string;
  /** Optional verified hub association GUID from source metadata; omission means metadata was not returned. */
  readonly hubSiteId?: string;
}
/** Verified SitePickerProps contract for picker consumers.
 * @example <SitePicker label="Site" mode="site" presentation="combobox" />
 */
export interface SitePickerProps extends PickerProps<SitePickerOption>, SearchPickerOptions {
  /** Keep the current site collection root visible before and during typed search. Defaults to true. Outside the ordinary initial suggestion cap; honors mode/KQL, selected exclusion and consumer model priority. showDefaultSuggestions=false suppresses loading/display. */
  readonly includeCurrentSite?: boolean;
  /** Keep the current collection's actually associated hub visible before and during search. Defaults to true. A hub without an association is not its own associated hub. Independent of hubSiteId; no association is an empty source. Outside the ordinary initial suggestion cap; honors mode/KQL, selected exclusion, consumer model priority and showDefaultSuggestions. */
  readonly includeHubSite?: boolean;
  /** Site discovery mode. Defaults to site; site searches collections, hub lists registered hubs, associatedSites uses the configured or current hub. */
  readonly mode?: 'site' | 'hub' | 'associatedSites';
  /** Optional hub GUID for associatedSites. Omission reads the current collection association; no association yields an empty result. */
  readonly hubSiteId?: string;
  /** Site query input placeholder; overrides the shared placeholder when supplied. */
  readonly searchPlaceholder?: string;
  /** Local ordering of the retrieved site collection set, default Title ascending. Reads at most 1000 Search/hub results before paging; reports partial data if more remain. */
  readonly orderBy?: PickerOrderBy<'title' | 'url'>;
  /** SharePoint search duplicate trimming; defaults to true. Only applies to site/associatedSites search. */
  readonly trimDuplicates?: boolean;
  /** Additional KQL AND clause, only for site mode. */
  readonly additionalQuery?: string;
}

/** Generated key includes normalized web and list GUID; caller aliases remain stable.  * @example const option: ListPickerOption = { key: 'projects', text: 'Projects', webUrl: 'https://tenant.sharepoint.com/sites/hr', id: '11111111-1111-4111-8111-111111111111', title: 'Projects', baseTemplate: 100, hidden: false, value: { kind: 'list', webUrl: 'https://tenant.sharepoint.com/sites/hr', listId: '11111111-1111-4111-8111-111111111111' } };
 */
export interface ListPickerOption extends PickerOption<ListPickerValue> {
  /** Absolute source web URL; supplied options must match the effective query web. */
  readonly webUrl: string;
  /** Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. */
  readonly id: string;
  /** Required plain metadata title from the source; a consumer display alias may use a different text. */
  readonly title: string;
  /** Optional server-relative root-folder navigation snapshot, never canonical identity. */
  readonly rootFolderUrl?: string;
  /** Required numeric SharePoint BaseTemplate metadata for this list; not a query filter. */
  readonly baseTemplate: number;
  /** Required hidden metadata flag used by convenience filters. */
  readonly hidden: boolean;
}
/** Verified ListPickerProps contract for picker consumers.
 * @example <ListPicker label="List" includeHidden={false} />
 */
export interface ListPickerProps extends PickerProps<ListPickerOption>, WebPickerSource {
  /** Maximum available options shown from the scoped initial catalog. Defaults to 30; integer 1–200. Selected rows are separate. Disabled by showDefaultSuggestions=false; typed search remains independent. */
  readonly initialCatalogLimit?: number;
  /** Optional numeric list template or template set (OR). Omission allows all; a raw filter overrides this convenience filter. */
  readonly baseTemplate?: number | readonly number[];
  /** Includes hidden lists/fields. Defaults to true; ignored when a raw filter overrides convenience filters. */
  readonly includeHidden?: boolean;
  /** OData expression overrides includeHidden/baseTemplate, as in PnP. */
  readonly filter?: string;
  /** Optional case-insensitive content-type ID prefix. Lists without a matching expanded content type are excluded. */
  readonly contentTypeId?: string;
  /** Optional server ordering. Metadata defaults to Title ascending; item search defaults to Id ascending. Continuations preserve ordering. */
  readonly orderBy?: PickerOrderBy<'id' | 'title'>;
}

/** Generated key includes effective web/list scope and field GUID; aliases remain stable.  * @example const handle = (option: FieldPickerOption) => console.log(option.internalName, option.typeAsString);
 */
export interface FieldPickerOption extends PickerOption<FieldPickerValue> {
  /** Absolute source web; listId is absent only for site columns. */
  readonly webUrl: string;
  /** Scoped list GUID, agreeing with the JSON reference. For site columns this is omitted. */
  readonly listId?: string;
  /** Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. */
  readonly id: string;
  /** Required field internal name, agreeing with the JSON reference; distinct from a REST EntityPropertyName. */
  readonly internalName: string;
  /** Required plain metadata title from the source; a consumer display alias may use a different text. */
  readonly title: string;
  /** Required SharePoint field type name; determines the field-specific icon and secondary label. */
  readonly typeAsString: string;
  /** Required numeric SharePoint field type discriminator. */
  readonly fieldTypeKind: number;
  /** Required hidden metadata flag used by convenience filters. */
  readonly hidden: boolean;
  /** Required ReadOnlyField metadata flag. */
  readonly readOnly: boolean;
  /** Required source field group metadata used by the exact group filter. */
  readonly group: string;
}
/** Verified FieldPickerProps contract for picker consumers.
 * @example <FieldPicker label="Field" list="Projects" includeReadOnly={false} />
 */
export interface FieldPickerProps extends PickerProps<FieldPickerOption>, WebPickerSource {
  /** Existing toolkit selector; omitted queries site columns. A string is an exact title. */
  readonly list?: string | SPFxPnPListSelector;
  /** Includes hidden lists/fields. Defaults to true; ignored when a raw filter overrides convenience filters. */
  readonly includeHidden?: boolean;
  /** Includes read-only fields. Defaults to true; a raw filter overrides this convenience flag. */
  readonly includeReadOnly?: boolean;
  /** Optional exact field group filter; omission accepts every group. Raw filter takes precedence. */
  readonly group?: string;
  /** Overrides includeHidden/includeReadOnly/group, as in PnP. */
  readonly filter?: string;
  /** Optional server ordering. Metadata defaults to Title ascending; item search defaults to Id ascending. Continuations preserve ordering. */
  readonly orderBy?: PickerOrderBy<'title' | 'internalName'>;
  /** Shows the explicit clear action. Defaults to false; clearing emits an empty array, never a synthetic domain reference. */
  readonly showBlankOption?: boolean;
}

/** Key: view GUID, scoped to the effective web/list.  * @example const handle = (option: ViewPickerOption) => console.log(option.title, option.personalView);
 */
export interface ViewPickerOption extends PickerOption<ViewPickerValue> {
  /** Effective source web/list identity, required for validating supplied candidates. */
  readonly webUrl: string;
  /** Required source list GUID, agreeing with the JSON reference; views always belong to a list. */
  readonly listId: string;
  /** Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. */
  readonly id: string;
  /** Required plain metadata title from the source; a consumer display alias may use a different text. */
  readonly title: string;
  /** Required DefaultView metadata flag. */
  readonly defaultView: boolean;
  /** Required PersonalView metadata flag. */
  readonly personalView: boolean;
  /** Optional server-relative navigation snapshot. Not canonical identity; must have one leading slash and no query, fragment or traversal segments. */
  readonly serverRelativeUrl?: string;
}
/** Verified ViewPickerProps contract for picker consumers.
 * @example <ViewPicker label="View" list="Projects" />
 */
export interface ViewPickerProps extends PickerProps<ViewPickerOption>, WebPickerSource {
  /** Required: there is no meaningful view query without a list. */
  readonly list: string | SPFxPnPListSelector;
  /** Optional SharePoint OData filter ANDed with typed search and exact candidate identity. Field/list convenience filters are overridden; scope restrictions still apply. */
  readonly filter?: string;
  /** Optional server ordering. Metadata defaults to Title ascending; item search defaults to Id ascending. Continuations preserve ordering. */
  readonly orderBy?: PickerOrderBy<'id' | 'title'>;
  /** Optional exact case-sensitive view titles to omit from suggestions and candidate verification. Omission excludes none. */
  readonly excludedViewTitles?: readonly string[];
  /** Shows the explicit clear action. Defaults to false; clearing emits an empty array, never a synthetic domain reference. */
  readonly showBlankOption?: boolean;
}

/** Key: encoded keyValue (number/string discriminator), scoped to effective web/list/keyField.  * @example const handle = (option: ListItemPickerOption) => console.log(option.kind, option.value.itemId);
 */
export interface ListItemPickerOption extends PickerOption<ListItemPickerValue> {
  /** Verified entity kind, matching value.kind. Library origin alone does not identify a file. */
  readonly kind: 'item' | 'file' | 'folder';
  /** Optional file extension for file-specific presentation, not inferred identity. */
  readonly fileExtension?: string;
  /** Effective source web/list identity, required for validating supplied candidates. */
  readonly webUrl: string;
  /** Required source list GUID for the item, file or folder, agreeing with the JSON reference. */
  readonly listId: string;
  /** Required positive safe integer SharePoint item ID, scoped to the web/list. Always retained even with a custom key. */
  readonly itemId: number;
  /** Effective key field, Id when omitted; validated against value.customKey or itemId. */
  readonly keyField?: string;
  /** Required typed custom-key value, or numeric itemId for the default Id key. Must agree with value.customKey when present. */
  readonly keyValue: string | number;
  /** Required read-only source field snapshot for consumer display logic. Never included in serialized picker values automatically. */
  readonly fields: Readonly<Record<string, unknown>>;
}
/** Verified ListItemPickerProps contract for picker consumers.
 * @example <ListItemPicker label="Item" list="Projects" textField="Title" searchMode="contains" />
 */
export interface ListItemPickerProps extends PickerProps<ListItemPickerOption>, WebPickerSource, SearchPickerOptions {
  /** Required exact list title or existing toolkit selector for item queries. String GUIDs are still titles. */
  readonly list: string | SPFxPnPListSelector;
  /** Include folder candidates. Default false; folders stay distinguishable from files and list records. */
  readonly includeFolders?: boolean;
  /** Required internal name of the item display/search field. Files also match their filename and use it as the label when this field is empty or whitespace; other missing values use an Item ID fallback. Field metadata determines safe query handling. */
  readonly textField: string;
  /** Defaults to Id; custom fields must contain unique string/number values. */
  readonly keyField?: string;
  /** Optional SharePoint OData filter ANDed with typed search and exact candidate identity. Field/list convenience filters are overridden; scope restrictions still apply. */
  readonly filter?: string;
  /** Optional server ordering. Metadata defaults to Title ascending; item search defaults to Id ascending. Continuations preserve ordering. */
  readonly orderBy?: readonly PickerOrderBy<string>[];
  /** Item text match mode; defaults to startsWith. Both modes compare display text and file names case-insensitively over bounded REST pages; non-text fields use FieldValuesAsText. */
  readonly searchMode?: 'startsWith' | 'contains';
}

/** References used to restrict user membership; SharePoint IDs and Entra IDs remain distinct.  * @example const group: PickerGroupReference = { source: 'entra', id: '11111111-1111-4111-8111-111111111111' };
 */
export type PickerGroupReference =
  | {
  /** Required source or transport discriminator; SharePoint identities and Entra identities remain distinct. */
  readonly source: 'sharePoint';
  /** Absolute HTTP(S) source web URL. Omission, when allowed, uses the provider current web; never inferred from a display label. */
  readonly webUrl: string;
  /** Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. */
  readonly id: number }
  | {
  /** Required source or transport discriminator; SharePoint identities and Entra identities remain distinct. */
  readonly source: 'entra';
  /** Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. */
  readonly id: string };
/** Supported principal categories; unvalidated email is an explicit user-input capability, not a directory group type.  * @example const types: PrincipalType[] = ['user', 'microsoft365Group'];
 */
export type PrincipalType = 'user' | 'sharePointGroup' | 'microsoft365Group' | 'securityGroup' | 'distributionList';
/** Required user identities; existing IDs are retained, never filtered from results by this mode.
 * Omission of userIdentity means available. Entra resolution is read-only Graph work.
 * SharePoint resolution without ensureUser reads only an existing registration; missing registration is an error.
 * ensureUser defaults to false, may register users, and runs only on confirmation/explicit resolution.
 * Destination defaults to PrincipalPickerProps.webUrl, then the provider's current web.
 * A requested identity failure prevents confirmation of a new selection. Preserve existing controlled keys.
 * Changing this configuration does not silently provision users or delete selections; resolve on an explicit action.
  * @example const resolution: UserIdentityResolution = { mode: 'both', ensureUser: false };
 */
export type UserIdentityResolution =
  | {
  /** Required user identity mode. Available retains returned IDs; entra requires a verified directory ID; sharePoint and both require a target collection registration. */
  readonly mode: 'available' | 'entra' }
  | {
      /** Required user identity mode. Available retains returned IDs; entra requires a verified directory ID; sharePoint and both require a target collection registration. */
  readonly mode: 'sharePoint' | 'both';
      /** Absolute HTTP(S) source web URL. Omission, when allowed, uses the provider current web; never inferred from a display label. */
  readonly webUrl?: string;
      /** Allows registering a missing SharePoint user only during explicit selection or ID resolution. Defaults to false; never runs during search, hydration or configuration changes. */
  readonly ensureUser?: boolean;
    };
/** Directory/SharePoint user snapshot; optional IDs allow unresolved initial/controlled values.
 * sharePointId is scoped to the collection containing sharePointWebUrl; Entra ID is scoped to entraTenantId.
 * Confirmed new selections satisfy the configured userIdentity requirements.
  * @example const user: UserPrincipal = { kind: 'user', key: 'alice', text: 'Alice', loginName: 'alice@tenant.com', value: { kind: 'user', loginName: 'alice@tenant.com' } };
 */
export interface UserPrincipal extends PickerOption<Extract<PrincipalPickerValue, {/** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  kind:'user'}>> {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'user';
  /** Required verified login/UPN or SharePoint membership claim. Login-only identity requires the provider verified tenant context. */
  readonly loginName: string;
  /** Optional user contact email, never canonical user identity. For unvalidatedEmail values it is required and remains explicitly unverified. */
  readonly email?: string;
  /** Optional verified directory UPN used for identity lookup; not an Entra object GUID. */
  readonly userPrincipalName?: string;
  /** Optional verified Entra user GUID; must agree with value.entra.id. */
  readonly entraId?: string;
  /** Optional verified Entra tenant GUID; must agree with value.entra.tenantId. */
  readonly entraTenantId?: string;
  /** Optional verified positive SharePoint user ID; scoped to sharePointWebUrl and agreeing with value.sharePoint.id. */
  readonly sharePointId?: number;
  /** Optional verified site collection URL for the SharePoint user ID; not the requesting subweb URL. */
  readonly sharePointWebUrl?: string;
  /** Optional user avatar image URL; omission uses Fluent initials. Non-user entities use category icons. */
  readonly imageUrl?: string;
}
/** Native identities of users/groups plus optional unvalidated email. Key is namespaced and stable across ID enrichment.  * @example const principal: PrincipalPickerOption = { kind: 'unvalidatedEmail', key: 'external', text: 'external@example.com', email: 'external@example.com', value: { kind: 'unvalidatedEmail', email: 'external@example.com' } };
 */
export type PrincipalPickerOption =
  | UserPrincipal
  | (PickerOption<Extract<PrincipalPickerValue, {/** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  kind:'sharePointGroup'}>> & {
      /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'sharePointGroup';
      /** Required source or transport discriminator; SharePoint identities and Entra identities remain distinct. */
  readonly source: 'sharePoint';
      /** Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. */
  readonly id: number;
      /** Absolute HTTP(S) source web URL. Omission, when allowed, uses the provider current web; never inferred from a display label. */
  readonly webUrl: string;
      /** Optional group description snapshot; the default secondary label uses type and source instead. */
  readonly description?: string;
    })
  | (PickerOption<Extract<PrincipalPickerValue, {/** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  kind:'microsoft365Group' | 'securityGroup' | 'distributionList'}>> & {
      /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'microsoft365Group' | 'securityGroup' | 'distributionList';
      /** Required source or transport discriminator; SharePoint identities and Entra identities remain distinct. */
  readonly source: 'entra';
      /** Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. */
  readonly id: string;
      /** Required verified Entra tenant GUID on serialized identities; optional redundant display-model metadata must agree when provided. */
  readonly tenantId?: string;
      /** Optional directory group mail address; never used as canonical group identity. */
  readonly mail?: string;
      /** Optional group description snapshot; the default secondary label uses type and source instead. */
  readonly description?: string;
      /** Required verified directory security flag for group classification. */
  readonly securityEnabled: boolean;
      /** Required verified directory mail flag for group classification. */
  readonly mailEnabled: boolean;
    })
  | (PickerOption<Extract<PrincipalPickerValue, {/** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  kind:'unvalidatedEmail'}>> & {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'unvalidatedEmail';
  /** Contact email: optional for user models, required for the unvalidatedEmail variant. Never canonical user identity. */
  readonly email: string });
/** Initial references are consumed once; subsequent controlled updates use selectedOptions/initialOptions.
 * displayOnly preserves the old PnP title-only scenario as an unresolved key, not a verified identity.
  * @example const initial: PrincipalPickerInitialPrincipal = { kind: 'email', email: 'alice@tenant.com' };
 */
export type PrincipalPickerInitialPrincipal =
  | {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'login';
  /** Required verified login/UPN or SharePoint membership claim. Login-only identity requires the provider verified tenant context. */
  readonly loginName: string;
  /** Optional fallback display label for login/email/Entra/group references; required for displayOnly. Never establishes identity. */
  readonly text?: string }
  | {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'email';
  /** Required email reference for read-only, unambiguous directory lookup; not an Entra ID. */
  readonly email: string;
  /** Optional fallback display label for login/email/Entra/group references; required for displayOnly. Never establishes identity. */
  readonly text?: string }
  | {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'entraUser';
  /** Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. */
  readonly id: string;
  /** Optional fallback display label for login/email/Entra/group references; required for displayOnly. Never establishes identity. */
  readonly text?: string }
  | {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'group';
  /** Required native group reference with its source and scope; numeric SharePoint IDs and Entra GUIDs remain distinct. */
  readonly group: PickerGroupReference;
  /** Optional fallback display label for login/email/Entra/group references; required for displayOnly. Never establishes identity. */
  readonly text?: string }
  | {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'displayOnly';
  /** Required stable renderer alias. It must never be reused for a different scoped entity. */
  readonly key: string;
  /** Optional fallback display label for login/email/Entra/group references; required for displayOnly. Never establishes identity. */
  readonly text: string };
/** Scope of user suggestions only; selectable group lookup is not restricted to these membership groups.  * @example const scope: PrincipalUserScope = { kind: 'sharePointGroupName', name: 'Project Readers' };
 */
export type PrincipalUserScope =
  | {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'tenant' }
  | {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'web';
  /** Absolute HTTP(S) source web URL. Omission, when allowed, uses the provider current web; never inferred from a display label. */
  readonly webUrl?: string }
  | {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'groups';
  /** Required membership group references; user candidates must belong to at least one configured group. */
  readonly groups: readonly PickerGroupReference[] }
  | {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'sharePointGroupName';
  /** Required exact SharePoint group name within the configured source web. */
  readonly name: string;
  /** Absolute HTTP(S) source web URL. Omission, when allowed, uses the provider current web; never inferred from a display label. */
  readonly webUrl?: string };
/** Principal pre-search policy. Consumer candidates precede the independent current user,
 * then “Persone suggerite” and “I tuoi gruppi”. Later sections may be truncated by the shared limit.
 * Current user is filtered/deduplicated like other sources; it does not displace consumer candidates.
 * No ensureUser call runs for suggestions. No selection history or automatic persistence exists.
 * @example const policy: PrincipalInitialSuggestions = { includeCurrentUser: true, limit: 5 };
 */
export interface PrincipalInitialSuggestions extends PickerInitialSuggestions {
  /** Offer the current SPFx user if allowed by user type/scope. Default false; independent of personal suggestions. */
  readonly includeCurrentUser?: boolean;
}
/** Unified user/group selection with common Fluent presentation, search and Property Pane lifecycle.
 * All principal suggestions use Persona; selected tags use name plus user avatar/group icon.
 * principalTypes changes affect suggestions without silently dropping incompatible existing selections.
 * Empty principalTypes disables suggestions. userIdentity applies only to users, never groups.
 * Graph/SharePoint permissions and available identities govern actual resolution; a matching SPFx provider is required.
 * @example
 * <PrincipalPicker principalTypes={['user', 'microsoft365Group']} presentation="tagPicker" multiselect
 *   userIdentity={{ mode: 'both', webUrl: 'https://tenant.sharepoint.com/sites/projects', ensureUser: true }} />
 */
export interface PrincipalPickerProps extends PickerProps<PrincipalPickerOption> {
  /** Optional current user plus shared five-item pre-search policy.
   * showDefaultSuggestions=false suppresses all pre-search sources and their lazy retrieval.
   * Search placeholder shows “Digita almeno 2 caratteri per cercare” (configured larger minimum is reflected); there is no duplicate minimum-character status row.
   */
  readonly initialSuggestions?: PrincipalInitialSuggestions;
  /** Lazy delegated Graph personal suggestions on opening an empty-query menu. Default false.
   * Users: GET /me/people (“Persone suggerite”), currently in maintenance mode, retained intentionally.
   * Groups: GET /me/memberOf (“I tuoi gruppi”), direct group memberships only; roles/admin units excluded.
   * Respect principalTypes and scopes; unverified directory identities/membership are not selectable.
   * person.id is not an Entra user ID. SharePoint groups cannot come from memberOf.
   * Failure is reported through onError with suggestionSource; healthy/consumer sources remain usable.
   * Permissions/limits: see docs/api/controls/pickers.md. Equivalent requests share owner-scoped memory cache only.
   */
  readonly loadPersonalSuggestions?: boolean;
  /** Selectable categories; defaults to ['user']. */
  readonly principalTypes?: readonly PrincipalType[];
  /** SharePoint group lookup web, also fallback identity-resolution destination; defaults to current web. */
  readonly webUrl?: string;
  /** User suggestion scope; defaults to tenant. */
  readonly userScope?: PrincipalUserScope;
  /** Optional uncontrolled principal references consumed once at mount. Mutually exclusive with defaultSelectedOptions; controlled selectedOptions takes precedence. */
  readonly defaultPrincipals?: readonly PrincipalPickerInitialPrincipal[];
  /** IDs required for confirmed users; defaults to available. See UserIdentityResolution. */
  readonly userIdentity?: UserIdentityResolution;
  /** Defaults to false; only allowed with user enabled and available identity mode. */
  readonly allowUnvalidatedEmail?: boolean;
  /** Compatibility-only flag with no effect. Native SharePoint user search determines visibility.
   * @deprecated The people-picker endpoint does not support a ShowHiddenPrincipals parameter; this flag is not forwarded or used as request identity.
   */
  readonly showHiddenPrincipals?: boolean;
  /** Enables Microsoft 365 Substrate user search in the native SharePoint query; defaults to false. */
  readonly useSubstrateSearch?: boolean;
}


/** Verified content type display model. Metadata is not serialized automatically.
 * Inherits PickerOption stable aliases, plain text labels and authoritative reference values.
 * @example const option: ContentTypePickerOption = {kind:'contentType',key:'document',text:'Document',webUrl:'https://tenant.sharepoint.com/sites/hr',id:'0x0101',name:'Document',group:'Document Content Types',hidden:false,readOnly:false,sealed:false,value:{kind:'contentType',scope:'site',webUrl:'https://tenant.sharepoint.com/sites/hr',contentTypeId:'0x0101'}};
 */
export interface ContentTypePickerOption extends PickerOption<ContentTypePickerValue> {
  /** Required model domain, agreeing with value.kind. */
  readonly kind: 'contentType';
  /** Required verified catalog web URL, agreeing with value.webUrl. */
  readonly webUrl: string;
  /** Required hierarchical StringId, agreeing with value.contentTypeId; never a GUID. */
  readonly id: string;
  /** Required source name. A consumer text alias may differ. */
  readonly name: string;
  /** Required source group, displayed as secondary Persona text. */
  readonly group: string;
  /** Optional source description; omission has no description snapshot. */
  readonly description?: string;
  /** Required hidden flag; included by default. */
  readonly hidden: boolean;
  /** Required read-only flag; never automatically excluded. */
  readonly readOnly: boolean;
  /** Required sealed flag; never automatically excluded. */
  readonly sealed: boolean;
}
/** Read-only collection or list content type picker. Requires an SPFx toolkit provider with SPHttpClient.
 * Inherits PickerProps selection arrays, model seeds, independent consumer candidates, callbacks,
 * accessibility and slots. Defaults to single Dropdown (historical combobox presentation), complete Name/ID ascending catalog, local
 * case-insensitive name-prefix search, and includeHidden=true. No personal/recent source is added.
 * webUrl inherits the provider web; site scope resolves its collection root even from a subweb.
 * List strings are exact titles; use the existing typed selector for IDs/paths/URLs.
 * @example const site: ContentTypePickerProps = {label:'Content type',webUrl:'https://tenant.sharepoint.com/sites/hr/sub'};
 * @example const list: ContentTypePickerProps = {scope:'list',list:'Documents',contentTypeId:'0x0101',includeDescendants:true};
 */
export type ContentTypePickerProps = PickerProps<ContentTypePickerOption> & WebPickerSource & {
  /** Optional exact group match, case insensitive. Omission accepts all groups. */
  readonly group?: string;
  /** Include hidden types. Defaults to true; readOnly/sealed do not affect eligibility. */
  readonly includeHidden?: boolean;
  /** Optional valid hierarchical ID filter; exact match unless includeDescendants=true. */
  readonly contentTypeId?: string;
  /** Include the filtered ID and all hierarchical descendants. Defaults to false; true requires contentTypeId. */
  readonly includeDescendants?: boolean;
} & ({
  /** Catalog scope; omission selects the owning site collection root. */
  readonly scope?: 'site';
  /** Forbidden for collection scope. */
  readonly list?: never;
} | {
  /** Required list scope discriminator when a list is supplied. */
  readonly scope: 'list';
  /** Required exact existing toolkit list selector in its owning web. */
  readonly list: string | SPFxPnPListSelector;
});

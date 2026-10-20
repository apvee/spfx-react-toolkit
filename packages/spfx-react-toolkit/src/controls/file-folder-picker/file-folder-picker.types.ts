import type { FieldProps } from '@fluentui/react-components';
import type { SPFxPnPListSelector } from '../../services/spfx-pnp-list.service';

/** Selectable provenance, independently verified after hosted navigation. Defaults to the current site collection.
 * Uses a single responsive 1080×720 hosted dialog without an outer visible header or padding.
 * Personal Microsoft accounts and other tenants are unsupported. Library selectors resolve actual document-library roots.
 * @example const scope: FileFolderPickerScope = { kind: 'folder', webUrl: 'https://tenant.sharepoint.com/sites/hr', list: 'Documents', folder: 'Policies' };
 */
export type FileFolderPickerScope =
  | {
    /** Restrict results to one verified site collection. */
    readonly kind: 'siteCollection';
    /** Absolute collection URL; omitted means pageContext.site.absoluteUrl, never the current subweb. */
    readonly siteUrl?: string;
  }
  | {
    /** Restrict results to one document library. */
    readonly kind: 'library';
    /** Required absolute owning web URL, without query or fragment. */
    readonly webUrl: string;
    /** Required title or explicit toolkit list selector; ordinary lists are rejected. */
    readonly list: string | SPFxPnPListSelector;
  }
  | {
    /** Restrict results to one folder and its verified descendants. */
    readonly kind: 'folder';
    /** Required absolute owning web URL, without query or fragment. */
    readonly webUrl: string;
    /** Required document-library title or explicit toolkit list selector. */
    readonly list: string | SPFxPnPListSelector;
    /** Required decoded server-relative or library-relative folder path; traversal and absolute URLs are rejected. */
    readonly folder: string;
  }
  | {
    /** Restrict results to the current authenticated user's same-tenant OneDrive for Business. The personal site is discovered read-only; unavailable drives are never provisioned. */
    readonly kind: 'oneDrive';
    /** Optional decoded folder path within the verified drive; omission starts at its root. */
    readonly folder?: string;
  };

/** Initial navigation only; this never broadens selectable scope. SharePoint folders require a list.
 * OneDrive accepts folder only; web/list combinations and containment are validated at runtime.
 * @example const location: FileFolderPickerInitialLocation = { list: 'Documents', folder: 'Policies' };
 */
export interface FileFolderPickerInitialLocation {
  /** Optional absolute SharePoint web URL; omission uses the resolved scope web. */
  readonly webUrl?: string;
  /** Optional document-library title or explicit selector; required when a SharePoint folder is given. */
  readonly list?: string | SPFxPnPListSelector;
  /** Optional decoded folder path; omission starts at the resolved library/web/drive root. */
  readonly folder?: string;
}

/** Serializable drive reference; distinct from legacy SharePointFileValue, with no fabricated numeric IDs.
 * Tenant and endpoint provenance require contextual verification before accepting new picks. Snapshots never determine identity.
 * @example const value: FileFolderPickerValue = { kind: 'file', tenantId: '11111111-1111-4111-8111-111111111111', endpoint: 'https://tenant.sharepoint.com/sites/hr/_api/v2.0', driveId: 'b!drive', itemId: '01ITEM' };
 */
export interface FileFolderPickerValue {
  /** Required verified item facet; files and folders remain distinguishable. */
  readonly kind: 'file' | 'folder';
  /** Required tenant GUID; canonical identity compares GUID case insensitively. */
  readonly tenantId: string;
  /** Required HTTPS Microsoft SharePoint Online API endpoint ending /_api/v2.0, optionally preceded by web path.
   * Host case and a trailing slash normalize; path case is preserved. No credentials, port, query, fragment, Graph or consumer endpoints.
   * Syntactic validation alone does not establish same-tenant trust.
   */
  readonly endpoint: string;
  /** Required opaque case-sensitive drive ID; never coerced into a GUID. */
  readonly driveId: string;
  /** Required opaque case-sensitive drive item ID; never coerced into a numeric list item ID. */
  readonly itemId: string;
  /** Optional display-name snapshot; omission leaves resolution/display fallback to the control. */
  readonly name?: string;
  /** Optional absolute HTTPS navigation URL snapshot; not an identity or scope authority. */
  readonly webUrl?: string;
  /** Optional verified SharePoint enrichment; omitted when these identifiers are unavailable. */
  readonly sharePoint?: {
    /** Required site collection GUID when enrichment is present. */
    readonly siteCollectionId: string;
    /** Required web GUID when enrichment is present. */
    readonly webId: string;
    /** Required document-library GUID when enrichment is present. */
    readonly listId: string;
    /** Optional positive safe-integer list item ID; omission is never replaced with a fabricated ID. */
    readonly listItemId?: number;
    /** Optional SharePoint unique item GUID. */
    readonly uniqueId?: string;
  };
}

/** Resolved selected display model; a seed is never an available candidate or persistence authority.
 * @example const option: FileFolderPickerOption = { value: { kind: 'folder', tenantId: '11111111-1111-4111-8111-111111111111', endpoint: 'https://tenant.sharepoint.com/_api/v2.0', driveId: 'b!drive', itemId: '01FOLDER' }, text: 'Policies' };
 */
export interface FileFolderPickerOption {
  /** Required authoritative scoped reference represented by this model. */
  readonly value: FileFolderPickerValue;
  /** Required plain display label. */
  readonly text: string;
  /** Optional file suffix used for presentation; absent for folders or unavailable metadata. */
  readonly extension?: string;
  /** Optional nonnegative safe-integer byte size; omitted when unavailable. */
  readonly sizeInBytes?: number;
}

/** Complete selection after an accepted explicit action. Hydration and prop updates never emit this event.
 * @example const handle = (data: FileFolderPickerChangeData): void => { console.log(data.reason, data.values); };
 */
export interface FileFolderPickerChangeData {
  /** Explicit accepted action: pick adds/replaces according to mode, remove deletes one, clear removes all. */
  readonly reason: 'pick' | 'remove' | 'clear';
  /** Complete ordered references, including retained unresolved values; fresh validated copies. */
  readonly values: readonly FileFolderPickerValue[];
  /** Resolved selected models only, in selection order. */
  readonly items: readonly FileFolderPickerOption[];
  /** Retained references without a resolved model; never fabricated into remote source items. */
  readonly unresolvedValues: readonly FileFolderPickerValue[];
}

interface FileFolderPickerCommonProps extends Pick<FieldProps,
  'label' | 'hint' | 'required' | 'validationMessage' | 'validationState'> {
  /** Optional root ID; owned controls retain their accessible associations. */
  readonly id?: string;
  /** Optional root CSS class. */
  readonly className?: string;
  /** Optional CSS class for the outlined tag/input surface; consumer overrides may change native Fluent border fidelity. */
  readonly controlClassName?: string;
  /** Optional accessible name when a visible label is absent. */
  readonly 'aria-label'?: string;
  /** Optional IDs of accessible labels. */
  readonly 'aria-labelledby'?: string;
  /** Disable browsing/removal; defaults to false. */
  readonly disabled?: boolean;
  /** Native tag/avatar size; defaults to medium. Small is unsupported. */
  readonly size?: 'medium' | 'large';
  /** Verified selection restriction; defaults to the current site collection. */
  readonly scope?: FileFolderPickerScope;
  /** Initial hosted navigation within scope; omission starts at the scope root. */
  readonly initialLocation?: FileFolderPickerInitialLocation;
  /** Selectable kinds; defaults to files. Folders-only cannot use extensions. */
  readonly itemTypes?: 'files' | 'folders' | 'all';
  /** Allow multiple selections; defaults to false. Incompatible existing values stay visible with validation. */
  readonly multiselect?: boolean;
  /** Optional positive integer maximum; effective single-selection maximum is one. */
  readonly selectionLimit?: number;
  /** Optional case-insensitive file suffixes such as .pdf; empty/omitted means unrestricted. */
  readonly extensions?: readonly string[];
  /** Enable hosted search; defaults to true. */
  readonly searchable?: boolean;
  /** Offer immediate native uploads; defaults false. Requires delegated write grants and user rights.
   * Uploads may complete before confirmation; Cancel/Property Pane Discard cannot undo them.
   * Accepted-reference scope validation does not constrain every native write destination.
   */
  readonly uploadEnabled?: boolean;
  /** Offer immediate native folder creation; defaults false. Same permissions and no-rollback constraints as uploadEnabled. */
  readonly createFolderEnabled?: boolean;
  /** Called once when the service reports page-loaded for this opening. No selection is committed. */
  readonly onReady?: () => void;
  /** Plain opening-action tooltip/accessibility label. Defaults to Sfoglia file…, Sfoglia cartelle… or Sfoglia file e cartelle… according to itemTypes. */
  readonly browseLabel?: string;
  /** Optional reactive selected-model seeds; never candidates, scope authority or persistence. */
  readonly initialOptions?: readonly FileFolderPickerOption[];
  /** Called with complete selection only after accepted explicit pick/remove/clear; never during hydration. */
  readonly onChange?: (data: FileFolderPickerChangeData) => void;
  /** Called for current actionable failures with sanitized user-facing messages. Optional out-of-scope token requests are denied inside the hosted protocol without reporting a configured-source failure. Cancellation/retry clears the Field message; this callback is not an error history. */
  readonly onError?: (error: Error) => void;
}

/** Temporary SDK-facing options used only when preset is omitted. */
interface FileFolderPickerLegacyOptions {
  /** Microsoft internal title; omitted uses the service title. This never adds an outer dialog header.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly title?: string;
  /** Initial native layout; omitted leaves the Microsoft service default.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly layout?: 'details' | 'compact-details' | 'tiles';
  /** Sticky native column headers; omitted leaves the service default.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly stickyHeaders?: boolean;
  /** Localized confirmation text; omitted uses the service label. Blank labels are rejected.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly selectLabel?: string;
  /** Localized native cancel text; omitted uses the service label. Blank labels are rejected.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly cancelLabel?: string;
  /** Initial view, default location. Other views require current-user OneDrive Business scope and no initialLocation.
   * @deprecated Use initialLocation with an explicit preset. Discovery entry views require a consumer choice to remain on legacy.
   */
  readonly initialView?: 'location' | 'myFiles' | 'recent' | 'shared' | 'site' | 'organization';
  /** Missing initial folder behavior, default error. scopeRoot falls back only on 404 to the verified source root;
   * it never broadens folder scope or hides invalid paths, permission failures or a missing source.
   *
   * @deprecated Presets fail on missing initial locations. Retaining scopeRoot fallback requires a consumer choice to remain on legacy.
   */
  readonly initialLocationFallback?: 'error' | 'scopeRoot';
  /** Initial SharePoint internal field sort. Omitted means service order; ascending defaults true.
   * @deprecated Presets use service order. Retaining an initial sort requires a consumer choice to remain on legacy.
   */
  readonly initialSort?: { readonly fieldName: string; readonly ascending?: boolean };
  /** Initial native field/value filter, not an enforced acceptance boundary. Omitted means no initial filter.
   * @deprecated Presets use no initial filter. Retaining this presentation filter requires a consumer choice to remain on legacy.
   */
  readonly initialFilter?: { readonly fieldName: string; readonly value: string };
  /** Explicit native pivots; omission derives source defaults. Pivots never broaden accepted scope.
   * Availability and token requirements depend on the hosted service; Graph is not enabled implicitly.
   * With visible SharePoint navigation, explicitly requesting myFiles/recent/shared permits lazy
   * authentication of the independently verified current user's OneDrive resource when Microsoft
   * requests it. Default navigation performs no auxiliary discovery. This does not broaden the
   * selectable source, grant permissions or enable Graph; discovery failures remain optional.
   *
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly visibleLocations?: readonly ('myFiles' | 'recent' | 'shared' | 'quickAccess' | 'site' | 'organization')[];
  /** Show native left navigation; defaults true. Hiding it is not a security restriction.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly navigationVisible?: boolean;
  /** Initial responsive navigation modality; omitted leaves the service default.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly navigationInitialState?: 'modal' | 'hidden';
  /** Native navigation preset; omitted derives oneDrive/current-site from scope.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly navigationPreset?: 'oneDrive' | 'current-site';
  /** Site pivot URL: resolved scope web, collection root or resolved initial web only. Omitted uses the scope web.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly siteNavigationUrl?: string;
  /** Retain the native draft across folders, default true in multiple mode and false in single mode.
   * Only the current opening is affected; no selection history or storage is created.
   *
   * @deprecated Explicit presets derive retention from multiselect. A different rule requires a consumer choice to remain on legacy.
   */
  readonly retainSelectionOnNavigate?: boolean;
  /** Native access filter; defaults read. This does not grant rights or enable writes.
   * @deprecated Presets use read access. Retaining read-write filtering requires a consumer choice to remain on legacy.
   */
  readonly requiredAccess?: 'read' | 'read-write';
  /** Permit exact Microsoft Graph token requests; defaults false. Does not request consent or initialize Graph clients.
   * @deprecated Use an explicit preset for validated on-demand Graph requests. To retain a Graph opt-out, keep the legacy path; there is no preset equivalent.
   */
  readonly graphEnabled?: boolean;
  /** Handle validated claims via the existing public SPFx token provider; defaults false. One challenge per resource/opening.
   * @deprecated Use an explicit preset for bounded validated claims retries. To retain a claims opt-out, keep the legacy path; there is no preset equivalent.
   */
  readonly claimsChallengeEnabled?: boolean;
  /** Link navigation/pivot, default disabled. Returned remote references still undergo complete scope validation.
   * @deprecated Use preset="extended" for link entry or preset="focused" to disable it. Navigation mode requires a consumer choice; it has no direct preset equivalent.
   */
  readonly fromLinkMode?: 'disabled' | 'navigation' | 'pivot';
  /** Offer device upload, default disabled. Upload requires uploadEnabled; local pick payloads are deliberately unsupported.
   * @deprecated Use uploadEnabled to control device upload in an explicit preset.
   */
  readonly localFilesMode?: 'disabled' | 'upload';
  /** Localized device-upload action label; omitted uses the service default.
   * @deprecated Preset device labels use service localization. A custom label requires a consumer choice to remain on legacy.
   */
  readonly localFilesLabel?: string;
  /** Native column filter affordance; omitted leaves the service default.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly columnFilterMode?: 'panel' | 'menu';
  /** Native selection tray prompt; defaults selection-summary. Destination/write prompts require a separate action contract.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly trayPrompt?: 'selection-summary' | 'selection-editor' | 'none';
}

/** Coherent hosted navigation recipe; availability of Microsoft pivots depends on the service.
 * focused starts at the resolved source with hidden navigation and no link entry. extended adds source-specific
 * discovery pivots and link entry; every accepted result still requires the original scope and policy checks.
 * Neither recipe enables upload or folder creation. Both support validated on-demand Graph requests and bounded
 * claims retries through the existing SPFx token provider, without requesting consent.
 * @example const preset: FileFolderPickerPreset = 'focused';
 */
export type FileFolderPickerPreset = 'focused' | 'extended';

/** Resolved internal options contain common props and SDK settings without consumer discriminants. */
export type FileFolderPickerResolvedOptions = FileFolderPickerCommonProps & FileFolderPickerLegacyOptions;

type FileFolderPickerExperienceProps =
  | ({
    /** Required on the simple path. Omission/undefined keeps exact temporary legacy behavior.
     * Presets cannot be combined with defined legacy SDK options, including false or empty arrays.
     */
    readonly preset: FileFolderPickerPreset;
  } & { readonly [Key in keyof FileFolderPickerLegacyOptions]?: never })
  | ({
    /** Omission/undefined selects the temporary legacy contract and its existing defaults. */
    readonly preset?: undefined;
  } & FileFolderPickerLegacyOptions);

/** Hosted File Picker v8 props. Requires an SPFx provider for context and AadTokenProvider authentication.
 * Native Fluent Field label/hint/required/validation slots retain their exact FieldProps contracts and defaults.
 * Omitted slots render no content. Medium is the default size; value/defaultValue are mutually exclusive arrays.
 * Controlled/uncontrolled mode is fixed for a mounted instance. Uncontrolled omission starts empty; defaultValue is read once.
 * No remote session opens until explicit Browse. Defaults to read-only same-tenant business sources; writes require explicit opt-in. Presets permit validated on-demand Graph authentication; legacy Graph defaults false.
 * Opening-only props are captured on Browse; updating them never cancels an active opening. Theme derives from the nearest Fluent provider.
 * The Field description uses inherited hint; omission renders no description. No automatic opening spinner/status is shown.
 * @example const props: FileFolderPickerProps = { preset: 'focused', label: 'Documents', defaultValue: [], scope: { kind: 'library', webUrl: 'https://tenant.sharepoint.com/sites/hr', list: 'Documents' }, itemTypes: 'files', multiselect: true, extensions: ['.pdf', '.docx'], onChange: data => console.log(data.values) };
 */
export type FileFolderPickerProps = FileFolderPickerCommonProps & FileFolderPickerExperienceProps & (
  | {
    /** Authoritative ordered references; [] explicitly controls an empty selection. */
    readonly value: readonly FileFolderPickerValue[];
    /** Disallowed in controlled mode. */
    readonly defaultValue?: never;
  }
  | {
    /** Omitted in uncontrolled mode; switching mode after mount is unsupported. */
    readonly value?: never;
    /** Initial ordered references, read once; omission defaults to []. */
    readonly defaultValue?: readonly FileFolderPickerValue[];
  }
);

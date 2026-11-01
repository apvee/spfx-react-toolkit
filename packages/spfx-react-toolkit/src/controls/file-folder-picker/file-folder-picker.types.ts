/**
 * @packageDocumentation
 * @docPage controls/file-folder-picker
 */
import type { FileFolderPickerStringsOverride } from './file-folder-picker-strings.types';
/**
 * File category mapped to the Microsoft v8 semantic filter. When omitted, leaves files unrestricted.
 * Toolkit acceptance checks a documented, case-insensitive suffix catalog against the authoritative metadata name;
 * it does not inspect bytes or promise Microsoft's complete internal classification. Media combines images, video and audio.
 * @remarks
 * ## Accepted suffixes
 * The case-insensitive authoritative-name catalog contains these suffixes without a leading dot:
 * - `images`: jpg, jpeg, jfif, png, gif, bmp, tif, tiff, webp, svg, ico, heic, heif, avif.
 * - `video`: mp4, m4v, mov, wmv, avi, mkv, webm, mpg, mpeg, 3gp, 3g2, m2ts, ts, mts, ogv.
 * - `audio`: mp3, wav, m4a, aac, flac, ogg, oga, wma, aif, aiff, opus, mid, midi.
 * - `media`: the union of images, video and audio above.
 * - `documents`: pdf, doc, docx, docm, dot, dotx, dotm, xls, xlsx, xlsm, xlsb, xlt, xltx, xltm, ppt, pptx, pptm, pot, potx, potm, pps, ppsx, ppsm, odt, ods, odp, rtf, txt, csv.
 * This is a suffix acceptance policy, not MIME/content inspection or an exact Microsoft catalog guarantee.
 * @example
 * ```ts
 * import { FileFolderPickerFileCategory } from '@apvee/spfx-react-toolkit';
 * const category: FileFolderPickerFileCategory = 'images';
 * ```
 */
export type FileFolderPickerFileCategory = 'images' | 'video' | 'audio' | 'media' | 'documents';

import type { FieldProps } from '@fluentui/react-components';
import type { SPFxPnPListSelector } from '../../services/spfx-pnp-list.service';

/**
 * Selectable provenance, independently verified after hosted navigation. Defaults to the current site collection.
 * Uses a single responsive 1080×720 hosted dialog without an outer visible header or padding.
 * Personal Microsoft accounts and other tenants are unsupported. Library selectors resolve actual document-library roots.
 * @example
 * ```ts
 * import { FileFolderPickerScope } from '@apvee/spfx-react-toolkit';
 * const scope: FileFolderPickerScope = { kind: 'folder', webUrl: 'https://tenant.sharepoint.com/sites/hr', list: 'Documents', folder: 'Policies' };
 * ```
 */
export type FileFolderPickerScope =
  | {
    /**
     * Restrict results to one verified site collection.
     */
    readonly kind: 'siteCollection';
    /**
     * Absolute collection URL; omitted means pageContext.site.absoluteUrl, never the current subweb.
     */
    readonly siteUrl?: string;
  }
  | {
    /**
     * Restrict results to one document library.
     */
    readonly kind: 'library';
    /**
     * Required absolute owning web URL, without query or fragment.
     */
    readonly webUrl: string;
    /**
     * Required title or explicit toolkit list selector; ordinary lists are rejected.
     */
    readonly list: string | SPFxPnPListSelector;
  }
  | {
    /**
     * Restrict results to one folder and its verified descendants.
     */
    readonly kind: 'folder';
    /**
     * Required absolute owning web URL, without query or fragment.
     */
    readonly webUrl: string;
    /**
     * Required document-library title or explicit toolkit list selector.
     */
    readonly list: string | SPFxPnPListSelector;
    /**
     * Required decoded server-relative or library-relative folder path; traversal and absolute URLs are rejected.
     */
    readonly folder: string;
  }
  | {
    /**
     * Restrict results to the current authenticated user's same-tenant OneDrive for Business. The personal site is discovered read-only; unavailable drives are never provisioned.
     */
    readonly kind: 'oneDrive';
    /**
     * Optional decoded folder path within the verified drive; when omitted, starts at its root.
     */
    readonly folder?: string;
  };

/**
 * Initial navigation only; this never broadens selectable scope. SharePoint folders require a list.
 * OneDrive accepts folder only; web/list combinations and containment are validated at runtime.
 * @example
 * ```ts
 * import { FileFolderPickerInitialLocation } from '@apvee/spfx-react-toolkit';
 * const location: FileFolderPickerInitialLocation = { list: 'Documents', folder: 'Policies' };
 * ```
 */
export interface FileFolderPickerInitialLocation {
  /**
   * Optional absolute SharePoint web URL; when omitted, uses the resolved scope web.
   */
  readonly webUrl?: string;
  /**
   * Optional document-library title or explicit selector; required when a SharePoint folder is given.
   * When omitted without a folder, SharePoint starts at the scoped source root; a folder requires an explicit library.
   */
  readonly list?: string | SPFxPnPListSelector;
  /**
   * Optional decoded folder path; when omitted, starts at the resolved library/web/drive root.
   */
  readonly folder?: string;
}

/**
 * Serializable drive reference; distinct from legacy SharePointFileValue, with no fabricated numeric IDs.
 * Tenant and endpoint provenance require contextual verification before accepting new picks. Snapshots never determine identity.
 * @example
 * ```ts
 * import { FileFolderPickerValue } from '@apvee/spfx-react-toolkit';
 * const value: FileFolderPickerValue = { kind: 'file', tenantId: '11111111-1111-4111-8111-111111111111', endpoint: 'https://tenant.sharepoint.com/sites/hr/_api/v2.0', driveId: 'b!drive', itemId: '01ITEM' };
 * ```
 */
export interface FileFolderPickerValue {
  /**
   * Required verified item facet; files and folders remain distinguishable.
   */
  readonly kind: 'file' | 'folder';
  /**
   * Required tenant GUID; canonical identity compares GUID case insensitively.
   */
  readonly tenantId: string;
  /**
   * Required HTTPS Microsoft SharePoint Online API endpoint ending /_api/v2.0, optionally preceded by web path.
   * Host case and a trailing slash normalize; path case is preserved. No credentials, port, query, fragment, Graph or consumer endpoints.
   * Syntactic validation alone does not establish same-tenant trust.
   */
  readonly endpoint: string;
  /**
   * Required opaque case-sensitive drive ID; never coerced into a GUID.
   */
  readonly driveId: string;
  /**
   * Required opaque case-sensitive drive item ID; never coerced into a numeric list item ID.
   */
  readonly itemId: string;
  /**
   * Optional display-name snapshot; when omitted, leaves resolution/display fallback to the control.
   */
  readonly name?: string;
  /**
   * Optional absolute HTTPS navigation URL snapshot; not an identity or scope authority.
   * When omitted, no snapshot for this optional metadata is supplied; canonical endpoint/drive/item identity remains authoritative.
   */
  readonly webUrl?: string;
  /**
   * Optional verified SharePoint enrichment; omitted when these identifiers are unavailable.
   */
  readonly sharePoint?: {
    /**
     * Required site collection GUID when enrichment is present.
     */
    readonly siteCollectionId: string;
    /**
     * Required web GUID when enrichment is present.
     */
    readonly webId: string;
    /**
     * Required document-library GUID when enrichment is present.
     */
    readonly listId: string;
    /**
     * Optional positive safe-integer list item ID; when omitted, is never replaced with a fabricated ID.
     */
    readonly listItemId?: number;
    /**
     * Optional SharePoint unique item GUID.
     * When omitted, no snapshot for this optional metadata is supplied; canonical endpoint/drive/item identity remains authoritative.
     */
    readonly uniqueId?: string;
  };
}

/**
 * Resolved selected display model; a seed is never an available candidate or persistence authority.
 * @example
 * ```ts
 * import { FileFolderPickerOption } from '@apvee/spfx-react-toolkit';
 * const option: FileFolderPickerOption = { value: { kind: 'folder', tenantId: '11111111-1111-4111-8111-111111111111', endpoint: 'https://tenant.sharepoint.com/_api/v2.0', driveId: 'b!drive', itemId: '01FOLDER' }, text: 'Policies' };
 * ```
 */
export interface FileFolderPickerOption {
  /**
   * Required authoritative scoped reference represented by this model.
   */
  readonly value: FileFolderPickerValue;
  /**
   * Required plain display label.
   */
  readonly text: string;
  /**
   * Optional file suffix used for presentation; absent for folders or unavailable metadata.
   * When omitted, no snapshot for this optional metadata is supplied; canonical endpoint/drive/item identity remains authoritative.
   */
  readonly extension?: string;
  /**
   * Optional nonnegative safe-integer byte size; omitted when unavailable.
   */
  readonly sizeInBytes?: number;
}

/**
 * Complete selection after an accepted explicit action. Hydration and prop updates never emit this event.
 * @example
 * ```ts
 * import { FileFolderPickerChangeData } from '@apvee/spfx-react-toolkit';
 * const handle = (data: FileFolderPickerChangeData): void => { console.log(data.reason, data.values); };
 * ```
 */
export interface FileFolderPickerChangeData {
  /**
   * Explicit accepted action: pick adds/replaces according to mode, remove deletes one, clear removes all.
   */
  readonly reason: 'pick' | 'remove' | 'clear';
  /**
   * Complete ordered references, including retained unresolved values; fresh validated copies.
   */
  readonly values: readonly FileFolderPickerValue[];
  /**
   * Resolved selected models only, in selection order.
   */
  readonly items: readonly FileFolderPickerOption[];
  /**
   * Retained references without a resolved model; never fabricated into remote source items.
   */
  readonly unresolvedValues: readonly FileFolderPickerValue[];
}

/**
 * Shared hosted picker Field presentation, source constraints and explicit selection events.
 * Native Field slots retain their Fluent contracts; authentication requires a matching SPFx provider.
 * Selection references and opening-only options are described by FileFolderPickerProps.
 * @example
 * ```ts
 * import { FileFolderPickerProps } from '@apvee/spfx-react-toolkit';
 * const picker: FileFolderPickerProps = { label: 'Documents', preset: 'focused', size: 'medium' };
 * ```
 */
interface FileFolderPickerCommonProps extends Pick<FieldProps,
  'label' | 'hint' | 'required' | 'validationMessage' | 'validationState'> {
  /**
   * Optional instance UI dictionary. Known nested leaves merge with native SPFx resources; direct labels/content/Field messages win. Undefined retains native defaults; empty strings remain explicit.
   */
  readonly strings?: FileFolderPickerStringsOverride;
  /**
   * Optional opening-button ID; the native Field label targets that button and owned controls retain their accessible associations.
   * When omitted, Fluent Field generates the opening-button association.
   */
  readonly id?: string;
  /**
   * Optional root CSS class.
   * When omitted, no consumer class is added to the root.
   */
  readonly className?: string;
  /**
   * Optional CSS class for the outlined tag/input surface; consumer overrides may change native Fluent border fidelity.
   * When omitted, no consumer class is added to the outlined surface.
   */
  readonly controlClassName?: string;
  /**
   * Optional accessible name when a visible label is absent.
   * When omitted, the visible Field label or aria-labelledby must provide the accessible name.
   */
  readonly 'aria-label'?: string;
  /**
   * Optional IDs of accessible labels.
   * When omitted, the visible Field label or aria-label supplies naming.
   */
  readonly 'aria-labelledby'?: string;
  /**
   * Disable browsing/removal; when omitted, defaults to false.
   * @defaultValue false
   */
  readonly disabled?: boolean;
  /**
   * Native tag/avatar size; when omitted, defaults to medium. Small is unsupported.
   * @defaultValue 'medium'
   */
  readonly size?: 'medium' | 'large';
  /**
   * Color selected file/folder tag avatars with Fluent's colorful palette. when omitted, defaults to false.
   * Uses the canonical tenant/endpoint/drive/item identity, so renames preserve color. Icons and sizes are unchanged.
   * Inherited by PropertyPaneFileFolderPicker; does not style Microsoft's hosted iframe or alter saved references.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { FileFolderPicker } from '@apvee/spfx-react-toolkit';
   * <FileFolderPicker label="Documents" colorfulTagAvatars />
   * ```
   *
   * @defaultValue false
   */
  readonly colorfulTagAvatars?: boolean;
  /**
   * Verified selection restriction; when omitted, defaults to the current site collection.
   * @defaultValue The current site collection.
   */
  readonly scope?: FileFolderPickerScope;
  /**
   * Initial hosted navigation within scope; when omitted, starts at the scope root.
   */
  readonly initialLocation?: FileFolderPickerInitialLocation;
  /**
   * Selectable kinds; when omitted, defaults to files. Folders-only cannot use extensions.
   * @defaultValue 'files'
   */
  readonly itemTypes?: 'files' | 'folders' | 'all';
  /**
   * Allow multiple selections; when omitted, defaults to false. Incompatible existing values stay visible with validation.
   * @defaultValue false
   */
  readonly multiselect?: boolean;
  /**
   * Optional positive integer maximum; effective single-selection maximum is one.
   * When omitted, multiple mode adds no custom count cap; single mode still allows one value.
   */
  readonly selectionLimit?: number;
  /**
   * Optional semantic file category; omitted means unrestricted. Incompatible with nonempty extensions or folders-only mode.
   * In mixed mode folders remain selectable. A filter change retains existing values and reports incompatible files.
   * Confirmed files must match the toolkit's documented suffix catalog using authoritative metadata; no content sniffing.
   */
  readonly fileCategory?: FileFolderPickerFileCategory;
  /**
   * Optional case-insensitive file suffixes such as .pdf; empty/omitted means unrestricted. Cannot combine nonempty with fileCategory.
   */
  readonly extensions?: readonly string[];
  /**
   * Enable hosted search; when omitted, defaults to true.
   * @defaultValue true
   */
  readonly searchable?: boolean;
  /**
   * Offer immediate native uploads; defaults false. Requires delegated write grants and user rights.
   * Uploads may complete before confirmation; Cancel/Property Pane Discard cannot undo them.
   * Accepted-reference scope validation does not constrain every native write destination.
   * When omitted, defaults to false.
   * @defaultValue false
   */
  readonly uploadEnabled?: boolean;
  /**
   * Offer immediate native folder creation; defaults false. Same permissions and no-rollback constraints as uploadEnabled.
   * When omitted, defaults to false.
   * @defaultValue false
   */
  readonly createFolderEnabled?: boolean;
  /**
   * Called once when the service reports page-loaded for this opening. No selection is committed.
   * When omitted, no consumer readiness notification is delivered.
   */
  readonly onReady?: () => void;
  /**
   * Plain opening-action tooltip/accessibility label. When omitted, uses the native localized strings.browse entry for itemTypes; a defined browseLabel takes precedence and also names the owned dialog/iframe.
   */
  readonly browseLabel?: string;
  /**
   * Optional reactive selected-model seeds; never candidates, scope authority or persistence.
   * When omitted, existing value references are resolved without consumer-supplied model seeds.
   */
  readonly initialOptions?: readonly FileFolderPickerOption[];
  /**
   * Called with complete selection only after accepted explicit pick/remove/clear; never during hydration.
   * When omitted, no consumer receives accepted selection changes; controlled values still require external updates.
   * @param data - Complete accepted pick/remove/clear selection with resolved and unresolved references.
   */
  readonly onChange?: (data: FileFolderPickerChangeData) => void;
  /**
   * Called for current actionable failures with sanitized user-facing messages. Optional out-of-scope token requests are denied inside the hosted protocol without reporting a configured-source failure. Cancellation/retry clears the Field message; this callback is not an error history.
   * When omitted, actionable errors remain visible in the Field without a consumer error notification.
   * @param error - Current actionable Error with a sanitized user-facing message.
   */
  readonly onError?: (error: Error) => void;
}

/**
 * Temporary SDK-facing options used only when preset is omitted.
 */
interface FileFolderPickerLegacyOptions {
  /**
   * Microsoft internal title; omitted uses the service title. This never adds an outer dialog header.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly title?: string;
  /**
   * Initial native layout; omitted leaves the Microsoft service default.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly layout?: 'details' | 'compact-details' | 'tiles';
  /**
   * Sticky native column headers; omitted leaves the service default.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly stickyHeaders?: boolean;
  /**
   * Localized confirmation text; omitted uses the service label. Blank labels are rejected.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly selectLabel?: string;
  /**
   * Localized native cancel text; omitted uses the service label. Blank labels are rejected.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly cancelLabel?: string;
  /**
   * Initial view, default location. Other views require current-user OneDrive Business scope with neither a scope folder nor initialLocation.
   *
   * When omitted, defaults to location.
   * @deprecated Use initialLocation with an explicit preset. Discovery entry views require a consumer choice to remain on legacy.
   * @defaultValue 'location'
   */
  readonly initialView?: 'location' | 'myFiles' | 'recent' | 'shared' | 'site' | 'organization';
  /**
   * Missing initial folder behavior, default error. scopeRoot falls back only on 404 to the verified source root;
   * it never broadens folder scope or hides invalid paths, permission failures or a missing source.
   *
   *
   * When omitted, defaults to error.
   * @deprecated Presets fail on missing initial locations. Retaining scopeRoot fallback requires a consumer choice to remain on legacy.
   * @defaultValue 'error'
   */
  readonly initialLocationFallback?: 'error' | 'scopeRoot';
  /**
   * Initial SharePoint internal field sort. Omitted means service order; ascending defaults true.
   * @deprecated Presets use service order. Retaining an initial sort requires a consumer choice to remain on legacy.
   */
  readonly initialSort?: {
    /**
     * Required SharePoint internal field name for the initial native sort.
     */
    readonly fieldName: string;
    /**
     * Initial sort direction; when omitted, defaults to true (ascending).
     */
    readonly ascending?: boolean
  };
  /**
   * Initial native field/value filter, not an enforced acceptance boundary. Omitted means no initial filter.
   * @deprecated Presets use no initial filter. Retaining this presentation filter requires a consumer choice to remain on legacy.
   */
  readonly initialFilter?: {
    /**
     * Required SharePoint internal field name for the initial native UI filter, not an acceptance boundary.
     */
    readonly fieldName: string;
    /**
     * Required initial native filter value; changes the initial UI without enforcing accepted-reference scope.
     */
    readonly value: string
  };
  /**
   * Explicit native pivots; when omitted, derives source defaults. Pivots never broaden accepted scope.
   * Availability and token requirements depend on the hosted service; Graph is not enabled implicitly.
   * With visible SharePoint navigation, explicitly requesting myFiles/recent/shared permits lazy
   * authentication of the independently verified current user's OneDrive resource when Microsoft
   * requests it. Default navigation performs no auxiliary discovery. This does not broaden the
   * selectable source, grant permissions or enable Graph; discovery failures remain optional.
   *
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly visibleLocations?: readonly ('myFiles' | 'recent' | 'shared' | 'quickAccess' | 'site' | 'organization')[];
  /**
   * Show native left navigation; defaults true. Hiding it is not a security restriction.
   *
   * When omitted, defaults to true.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   * @defaultValue true
   */
  readonly navigationVisible?: boolean;
  /**
   * Initial responsive navigation modality; omitted leaves the service default.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly navigationInitialState?: 'modal' | 'hidden';
  /**
   * Native navigation preset; omitted derives oneDrive/current-site from scope.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly navigationPreset?: 'oneDrive' | 'current-site';
  /**
   * Site pivot URL: resolved scope web, collection root or resolved initial web only. Omitted uses the scope web.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly siteNavigationUrl?: string;
  /**
   * Retain the native draft across folders. When omitted in legacy multiple mode,
   * the toolkit sets the SDK persistence flag to true. Legacy single mode leaves
   * that flag unset, so the hosted service's default applies. Explicit true/false
   * values are forwarded. Presets instead set the flag from multiselect: true for
   * multiple mode and false for single mode, without accepting this legacy override.
   * Only the current opening is affected; no selection history or storage is created.
   * @deprecated Explicit presets derive retention from multiselect. A different rule requires a consumer choice to remain on legacy.
   */
  readonly retainSelectionOnNavigate?: boolean;
  /**
   * Native access filter; defaults read. This does not grant rights or enable writes.
   *
   * When omitted, defaults to read.
   * @deprecated Presets use read access. Retaining read-write filtering requires a consumer choice to remain on legacy.
   * @defaultValue 'read'
   */
  readonly requiredAccess?: 'read' | 'read-write';
  /**
   * Permit exact Microsoft Graph token requests; defaults false. Does not request consent or initialize Graph clients.
   *
   * When omitted, defaults to false.
   * @deprecated Use an explicit preset for validated on-demand Graph requests. To retain a Graph opt-out, keep the legacy path; there is no preset equivalent.
   * @defaultValue false
   */
  readonly graphEnabled?: boolean;
  /**
   * Handle validated claims via the existing public SPFx token provider; defaults false. One challenge per resource/opening.
   *
   * When omitted, defaults to false.
   * @deprecated Use an explicit preset for bounded validated claims retries. To retain a claims opt-out, keep the legacy path; there is no preset equivalent.
   * @defaultValue false
   */
  readonly claimsChallengeEnabled?: boolean;
  /**
   * Link navigation/pivot, default disabled. Returned remote references still undergo complete scope validation.
   *
   * When omitted, defaults to disabled.
   * @deprecated Use preset="extended" for link entry or preset="focused" to disable it. Navigation mode requires a consumer choice; it has no direct preset equivalent.
   * @defaultValue 'disabled'
   */
  readonly fromLinkMode?: 'disabled' | 'navigation' | 'pivot';
  /**
   * Offer device upload, default disabled. Upload requires uploadEnabled; local pick payloads are deliberately unsupported.
   *
   * When omitted, defaults to disabled.
   * @deprecated Use uploadEnabled to control device upload in an explicit preset.
   * @defaultValue 'disabled'
   */
  readonly localFilesMode?: 'disabled' | 'upload';
  /**
   * Localized device-upload action label; omitted uses the service default.
   * @deprecated Preset device labels use service localization. A custom label requires a consumer choice to remain on legacy.
   */
  readonly localFilesLabel?: string;
  /**
   * Native column filter affordance; omitted leaves the service default.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   */
  readonly columnFilterMode?: 'panel' | 'menu';
  /**
   * Native selection tray prompt; when omitted, uses selection-summary. Destination/write prompts require a separate action contract.
   * @deprecated Choose focused or extended; presets own native presentation/navigation defaults. A custom override requires a consumer choice to remain on legacy.
   * @defaultValue 'selection-summary'
   */
  readonly trayPrompt?: 'selection-summary' | 'selection-editor' | 'none';
}

/**
 * Coherent hosted navigation recipe; availability of Microsoft pivots depends on the service.
 * focused starts at the resolved source with hidden navigation and no link entry. extended adds source-specific
 * discovery pivots and link entry; every accepted result still requires the original scope and policy checks.
 * Neither recipe enables upload or folder creation. Both support validated on-demand Graph requests and bounded
 * claims retries through the existing SPFx token provider, without requesting consent.
 * @example
 * ```ts
 * import { FileFolderPickerPreset } from '@apvee/spfx-react-toolkit';
 * const preset: FileFolderPickerPreset = 'focused';
 * ```
 */
export type FileFolderPickerPreset = 'focused' | 'extended';

/**
 * Resolved internal options contain common props and SDK settings without consumer discriminants.
 */
export type FileFolderPickerResolvedOptions = FileFolderPickerCommonProps & FileFolderPickerLegacyOptions;

/**
 * Mutually exclusive hosted experience configuration: a preset owns its native SDK
 * options, while omitted/undefined preset retains the temporary legacy option contract.
 * Presets reject every defined legacy override, including false and empty arrays.
 * @example
 * ```ts
 * import { FileFolderPickerProps } from '@apvee/spfx-react-toolkit';
 * const focused: FileFolderPickerProps = { label: 'Documents', preset: 'focused' };
 * const legacy: FileFolderPickerProps = { label: 'Documents', navigationVisible: false };
 * ```
 */
type FileFolderPickerExperienceProps =
  | ({
    /**
     * Required focused/extended discriminator for this preset branch. Omit preset entirely to select the separate temporary legacy branch.
     * Presets cannot be combined with defined legacy SDK options, including false or empty arrays.
     */
    readonly preset: FileFolderPickerPreset;
  } & { readonly [Key in keyof FileFolderPickerLegacyOptions]?: never })
  | ({
    /**
     * When omitted or undefined, selects the temporary legacy contract and its existing defaults.
     */
    readonly preset?: undefined;
  } & FileFolderPickerLegacyOptions);

/**
 * Hosted File Picker v8 props. Requires an SPFx provider for context and AadTokenProvider authentication.
 * Native Fluent Field label/hint/required/validation slots retain their exact FieldProps contracts and defaults.
 * Omitted slots render no content. Medium is the default size; value/defaultValue are mutually exclusive arrays.
 * Controlled/uncontrolled mode is fixed for a mounted instance. An omitted uncontrolled default starts empty; defaultValue is read once.
 * No remote session opens until explicit Browse. Defaults to read-only same-tenant business sources; writes require explicit opt-in. Presets permit validated on-demand Graph authentication; legacy Graph defaults false.
 * Opening-only props are captured on Browse; updating them never cancels an active opening. Theme derives from the nearest Fluent provider.
 * The Field description uses inherited hint; when omitted, renders no description. No automatic opening spinner/status is shown.
 * @example
 * ```ts
 * import { FileFolderPickerProps } from '@apvee/spfx-react-toolkit';
 * const props: FileFolderPickerProps = { preset: 'focused', label: 'Documents', defaultValue: [], scope: { kind: 'library', webUrl: 'https://tenant.sharepoint.com/sites/hr', list: 'Documents' }, itemTypes: 'files', multiselect: true, extensions: ['.pdf', '.docx'], onChange: data => console.log(data.values) };
 * ```
 */
export type FileFolderPickerProps = FileFolderPickerCommonProps & FileFolderPickerExperienceProps & (
  | {
    /**
     * Authoritative ordered references; [] explicitly controls an empty selection.
     */
    readonly value: readonly FileFolderPickerValue[];
    /**
     * Disallowed in controlled mode.
     * Must be omitted or undefined in this branch.
     */
    readonly defaultValue?: never;
  }
  | {
    /**
     * Omitted in uncontrolled mode; switching mode after mount is unsupported.
     */
    readonly value?: never;
    /**
     * Initial ordered references, read once; when omitted, defaults to [].
     */
    readonly defaultValue?: readonly FileFolderPickerValue[];
  }
);

# FileFolderPicker

`FileFolderPicker` selects existing SharePoint or OneDrive for Business files and folders through Microsoft's hosted File Picker v8. It requires a matching SPFx provider and the existing SPFx AAD token provider. The toolkit does not install the sample-local `@pnp/picker-api` package or another MSAL client.

```tsx
import * as React from 'react';
import { FileFolderPicker } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';
import type { FileFolderPickerValue } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';

function AttachmentReferences(): React.ReactElement {
  const [value, setValue] = React.useState<readonly FileFolderPickerValue[]>([]);
  return <FileFolderPicker preset="focused" label="Documents" value={value}
    scope={{ kind: 'siteCollection' }} itemTypes="files" multiselect
    extensions={['.pdf', '.docx']} onChange={data => setValue(data.values)} />;
}
```

Use the root, `/controls`, or `/controls/file-folder-picker` entry. Native fields have a separate [Property Pane entry](../property-pane/file-folder-picker.md).

## Presets: the simple configuration

`FileFolderPickerPreset` is the public `'focused' | 'extended'` union used by the `preset` prop. Omitting `preset` preserves the legacy configuration contract.

Pass `preset="focused"` or `preset="extended"` explicitly. Both use the same source and result validation; the preset changes the hosted browsing experience, not what may be accepted.

| Setting | `focused` | `extended` |
| --- | --- | --- |
| Initial location | Resolved scope or valid `initialLocation` | Same |
| Sidebar | Hidden | Visible |
| SharePoint navigation requests | None | Recent, Shared, Quick access, My organization, Site |
| Current-user OneDrive navigation requests | Only the required entry | My files, Recent, Shared, My organization, Site |
| From a link | Disabled | Native pivot; returned references still require scope verification |
| Auxiliary personal-resource authentication from SharePoint | Disabled | Lazy, independently verified current-user OneDrive only |
| Graph / claims | Validated on-demand handling through the existing provider/grants | Same |
| Upload / create-folder | Disabled unless explicitly enabled | Same |

The main choices are `scope`, `preset`, `itemTypes` and `multiselect`. Additional intent props are `selectionLimit`, `extensions`, `searchable`, `initialLocation`, `uploadEnabled` and `createFolderEnabled`. Standard Field/accessibility props, `size`, selected values/models and callbacks remain available. The theme comes from the existing Fluent/SPFx theme hooks. No raw SDK options object or custom preset is exposed.

Graph handling does not add consent, grants or a Graph client, and is requested only for validated hosted authentication commands. A token obtained from the existing provider may carry its already approved scopes; the picker does not narrow those scopes or prove that Microsoft uses no other APIs internally. SharePoint tokens remain resource-specific. Handling an already approved resource does not guarantee that a service operation is authorized. The [permission matrix](#permissions-by-enabled-feature) applies to both presets; actual hosted API traffic and tenant grants need separate validation.

Uploads and create-folder remain explicit independent choices. Enabling uploads also offers native device upload where supported. Such writes can happen before confirmation; cancellation cannot undo them, and the selectable-reference restriction is not a pre-write destination boundary.

Microsoft owns the available navigation, including Media/People/Meetings/Favorites; there are no individual public switches for those entries. A requested entry is not a promise of available content or successful service navigation. There is no silent fallback to another preset/source when a route or resource fails. `initialLocation` changes only the start within the allowed source.

### Migrating existing consumers

Omitting `preset` retains the legacy contract and its exact defaults. A defined legacy SDK prop cannot be combined with an explicit preset, including `false` or an empty array; `undefined` counts as omission. Types and runtime validation reject mixed configuration before hosted authentication/navigation. Remove legacy props and choose a preset intentionally rather than letting one silently override the other.

The native SDK knobs listed under [legacy flat configuration](#flat-native-configuration-and-stable-opening) remain supported only without `preset` and are deprecated for new usage. This includes Graph/claims flags, navigation flags, sort/filter/layout, internal labels and link/device configuration. Common functional options such as `uploadEnabled`, `createFolderEnabled`, `hint` and `onReady` remain supported in both paths. Saved `FileFolderPickerValue` arrays do not change and never acquire a preset field. Migration does not clear selections or infer an opt-in from saved legacy flags.

Changing the preset affects the next opening and preserves current labels, selections and the mounted Property Pane host. The current opening retains its snapshot; source/item/access policy and authoritative changes retain their existing cancellation rules.

## Sources and restrictions

`scope` is a discriminated union:

| Kind | Configuration |
| --- | --- |
| `siteCollection` | Optional `siteUrl`; omission uses `pageContext.site.absoluteUrl`, even in a subweb. |
| `library` | Required `webUrl` and `list`, using the toolkit's title, GUID, URL or `SPFxPnPListSelector` conventions. |
| `folder` | Required `webUrl`, `list` and decoded library-relative or server-relative `folder`. |
| `oneDrive` | Current authenticated user's OneDrive Business, discovered read-only; optional `folder`. No URL input or automatic provisioning. |

Ordinary lists are rejected. A collection is verified by its actual SharePoint site ID; URL prefixes are insufficient. The hosted start location is separate from the permitted source. `initialLocation` can narrow the starting web/library/folder inside that source; it never broadens the boundary. A SharePoint starting folder requires a library. OneDrive accepts only a starting folder.

Folder inputs preserve literal spaces, apostrophes, `#`, `%` and Unicode. They are decoded paths, not arbitrary absolute URLs. Traversal and segment-prefix collisions are rejected. The toolkit validates authoritative returned metadata before accepting an entire picked batch; changing navigation or returning a shortcut does not bypass the configured scope.

A collection begins browsing at its verified root web. This does not promise enumeration of all subweb libraries by Microsoft's UI. Multiple arbitrary source roots, personal Microsoft accounts and other tenants are unsupported.

## Props and state

`FileFolderPickerProps` defines control configuration and selection ownership. `FileFolderPickerScope` defines selectable provenance; `FileFolderPickerInitialLocation` defines initial navigation. `FileFolderPickerChangeData` is the explicit-action callback payload. `FileFolderPickerValue` and `FileFolderPickerOption` separate persisted references from display models.

| Prop | Behavior when omitted |
| --- | --- |
| `preset` | Omission/undefined retains legacy. Explicit `focused` or `extended` selects the simple contract and rejects defined legacy SDK options. |
| `value` / `defaultValue` | Controlled or initially uncontrolled JSON arrays, mutually exclusive; initial selection defaults to empty. Choose ownership once per mount. |
| `initialOptions` | Optional display models for selected references; no suggested-menu source or authorization. |
| `itemTypes` | `files`; also `folders` and `all`. |
| `multiselect` | `false`; single picks replace, multiple picks append and deduplicate. |
| `selectionLimit` | Optional positive integer; single mode always permits at most one. Existing conflicting values remain visible for removal. |
| `extensions` | Unrestricted file suffixes; case-insensitive, e.g. `.pdf`. Invalid in folder-only mode. |
| `searchable` | `true`, controlling hosted search. |
| `browseLabel` | Plain opening-action tooltip/accessibility label. Default: `Sfoglia file…`, `Sfoglia cartelle…` or `Sfoglia file e cartelle…`, according to `itemTypes`. |
| `size` | `medium`; also `large`. Selected tag media follow native Fluent sizing. |
| `disabled` | `false`; prevents opening and selection mutation. |
| `label`, `hint`, `required`, `validationMessage`, `validationState` | Native Fluent Field contracts; `id`, `className`, `controlClassName` and accessible-name attributes apply to the owned surface. |
| `onChange` | Optional callback for explicit accepted `pick`, `remove` or `clear` actions. |
| `onError` | Optional current actionable error callback with a sanitized message; optional audience/Graph failures remain isolated from the configured source. No raw token, claims or message payload is exposed. |

`onChange` receives the complete ordered `values`, currently resolved display `items`, `unresolvedValues` and the action `reason`. Hydration and property updates never emit. Invalid batches are rejected atomically. Selection-limit conflicts never silently truncate existing references.

Each `FileFolderPickerValue` has `kind: 'file' | 'folder'`, `tenantId`, normalized SharePoint v2 `endpoint`, opaque case-sensitive `driveId` and `itemId`. Optional `name`, `webUrl` and verified `sharePoint` metadata are snapshots. Identity does not depend on the name/path. References are intentionally separate from the existing `PickerValue` union: a drive item ID is not a numeric SharePoint list-item ID.

`FileFolderPickerOption` adds `text`, optional `extension` and `sizeInBytes` to a `value`. Display models are not proof of eligibility. Unreadable/deleted references are preserved for consumer decisions and removal; hydration does not silently rewrite saved values.

## Authentication and lifecycle

Opening creates the single Stage dialog, then resolves metadata and obtains a resource-specific token using `useSPFxAadTokenProvider`. Only explicitly allowlisted same-tenant SharePoint resources receive SharePoint tokens. With a preset, validated Graph requests are handled on demand; in the legacy path graphEnabled additionally permits the exact https://graph.microsoft.com audience with a Graph request type. The hosted `SharePoint`, observed `SharePoint_SelfIssued` and `SharePoint_MSAL` authentication hints use the existing SPFx provider for that exact SharePoint resource; the integration does not issue its own tokens. Graph tokens are never substituted for SharePoint tokens. In the legacy path claims challenges default to disabled. Set claimsChallengeEnabled to permit one validated JSON claims challenge per resource/opening through the public SPFx token provider; retries bypass its cached token. In that legacy path Graph remains independently disabled unless graphEnabled is true. Presets internally enable the same bounded claims handler. Consumers must already have the appropriate delegated grants; the control does not request consent or change tenant permissions.

Microsoft can request an optional personal-site audience while browsing a SharePoint source. With `extended`, this path is enabled independently of the names of individual pivots; `focused` does not enable it. In the legacy path, an explicit `visibleLocations` array containing myFiles, recent or shared, with `navigationVisible` other than false, enables a bounded auxiliary authentication path. Only an exact tenant personal-resource origin root requested with a supported SharePoint authentication hint can trigger read-only current-user OneDrive discovery. The existing profile URL, tenant and drive-owner checks must resolve the same requested origin before a token is acquired. Discovery is lazy, attempted at most once per opening and reuses the mounted metadata request cache. Untouched defaults, hidden navigation and explicit arrays without those personal entries do not enable it.

This auxiliary token audience is separate from the selectable metadata provenance: SharePoint scope, identity, item-policy and batch acceptance checks remain unchanged. It introduces no public prop, consent, Graph dependency or write opt-in. An unrelated or unverified resource stays denied. Auxiliary discovery/token failures and optional Graph failures remain local to their hosted command; they neither set a false configured-source Field error nor clear an existing source error. Errors for the configured resource, unsupported authentication, unavailable items, source/policy violations and invalid identities use shared sanitized messages in React and Property Pane. `onError` receives the current actionable failure without provider/native payload strings. Cancel, retry and authoritative changes clear the Field's opening error; a successful token retry clears authentication errors without hiding an unrelated selection failure. Consumers maintaining their own error state must manage its lifetime independently.

With SPFx `debug=true`, internal authentication diagnostics use bounded `console.debug` events with finite categories. They distinguish host policy decisions from token acquisition outcomes without logging tokens, claims, raw resource URLs or provider/native messages. They are diagnostic evidence, not a grants inventory or proof of the iframe's internal requests. There is no new public diagnostic prop, automatic Field hint or opening spinner.

When the hosted picker returns the symbolic item ID `root`, the integration verifies the live drive-root relationship and all normal provenance restrictions, then stores the canonical drive-item ID. Other identity mismatches remain rejected. Names and paths never replace identity checks.

The URL `/_layouts/15/FilePicker.aspx` does **not** establish which internal APIs the service uses. The v8 protocol can request SharePoint or Graph resources. Host metadata reads, token requests and the hosted service's actual network traffic are separate evidence. Document permission requirements using observed operations and Microsoft's current documentation; do not infer Graph independence from a successful SharePoint page load.

A unique channel binds each opening to its exact origin, content window and message port. Unsupported commands fail explicitly. Source/item-policy and authoritative value changes cancel stale confirmations. Equivalent settings, labels and callbacks do not restart the session. Closing or disposing removes handlers, ports and the owned iframe.

Only mounted metadata requests are coalesced in memory. No selected-item history, localStorage, OneDrive storage or token cache is managed by the control. Microsoft's within-opening multiple-folder selection persistence does not persist toolkit values between sessions. The consumer owns eventual JSON persistence.

## Presentation and limitations

Browse opens Microsoft's UI; it is not a Combobox or TagPicker dropdown. The outer Field, button and removable tags inherit the host theme; native Microsoft light/dark follows the closest effective FluentProvider at opening. Hosted iframe content is owned by Microsoft and does not inherit arbitrary Fluent classes. Authentication failures and iframe restrictions are actionable errors; retry keeps the existing references. Stage owns dialog dismissal and focus; keyboard interaction inside the remote frame is Microsoft's responsibility.

The outer control is one outlined tag surface. File/folder tags wrap, truncate long names with ellipsis and expose the complete name in a tooltip; each remains individually removable. The trailing icon opens the hosted picker: Attach for files, FolderOpen for folders, FolderDocument for mixed mode. An empty field contains a focusable selection placeholder that opens the same picker. Both opening actions preserve the Field's label, hint and validation associations. There is no editable text input and no separate Browse or clear-all button. Controlled consumers can clear the array; users can remove tags individually.

The border uses the installed Fluent Input style hook directly, including its 1px border, theme radius (4px in the standard themes), hover/invalid/disabled states and native focus animation. Empty Medium/Large fields match native 32px/40px minimum heights; selected tags can expand the field vertically. The underline remains active throughout opening and while the hosted dialog is open, even after focus leaves the field. Reduced-motion behavior follows Fluent. The mode-specific opening icon and selected labels remain visible while busy; no spinner or automatic Field description is added. `controlClassName` targets this outlined surface; consumer overrides can intentionally change border fidelity. The iframe itself retains Microsoft's presentation.

Uploads, folder creation and local-device upload are disabled until the corresponding explicit write intent is enabled. Link entry is disabled in focused and legacy defaults, and enabled as a native pivot in extended. Local-device pick, account switching and share/move/download actions remain gated. Confirming selection reads metadata and does not itself perform those operations. Restrictions are checked again when confirming results.

The private sample exposes source, start location, file/folder mode, suffixes, single/multiple selection and the single hosted dialog. Automated protocol/scope tests and simulated browser fixtures are separate from authenticated tenant validation. See [SharePoint validation](../../SHAREPOINT-VALIDATION.md) for the host checklist and execution limitations.

Primary references: [Microsoft File Picker v8](https://learn.microsoft.com/en-us/onedrive/developer/controls/file-pickers/), [current schema](https://learn.microsoft.com/en-us/onedrive/developer/controls/file-pickers/v8-schema), [React sample](https://github.com/OneDrive/samples/tree/master/samples/file-picking/typescript-react).

## Member reference

| Member | Contract |
| --- | --- |
| `FileFolderPickerProps.value` | Authoritative controlled JSON array; omit defaultValue. |
| `FileFolderPickerProps.defaultValue` | Initial uncontrolled array only; defaults to empty. |
| `FileFolderPickerScope.kind` | Selects collection, library, folder or OneDrive Business provenance. |
| `FileFolderPickerScope.siteUrl` | Optional collection URL; defaults to the provider collection root. |
| `FileFolderPickerScope.webUrl` | Required owning web for library/folder source. |
| `FileFolderPickerScope.list` | Required document-library title or explicit toolkit selector. |
| `FileFolderPickerScope.folder` | Required folder-source boundary; optional OneDrive boundary. |
| `FileFolderPickerInitialLocation.webUrl` | Optional initial web inside the allowed collection. |
| `FileFolderPickerInitialLocation.list` | Initial library; required with a SharePoint starting folder. |
| `FileFolderPickerInitialLocation.folder` | Initial decoded folder inside the boundary; OneDrive accepts this member alone. |
| `FileFolderPickerValue.kind` | Authoritatively confirmed file or folder. |
| `FileFolderPickerValue.tenantId` | Tenant GUID; compared case insensitively. |
| `FileFolderPickerValue.endpoint` | Normalized trusted SharePoint HTTPS API base ending in /_api/v2.0. |
| `FileFolderPickerValue.driveId` | Opaque case-sensitive drive identifier. |
| `FileFolderPickerValue.itemId` | Opaque case-sensitive drive-item identifier. |
| `FileFolderPickerValue.name` | Optional display snapshot; never identity or authorization. |
| `FileFolderPickerValue.webUrl` | Optional HTTPS navigation snapshot, including ordinary query/fragment links. |
| `FileFolderPickerValue.sharePoint` | Optional verified siteCollectionId, webId and listId GUIDs with optional positive listItemId and uniqueId; absent facts are not invented. |
| `FileFolderPickerOption.value` | Scoped serializable reference represented by the display model. |
| `FileFolderPickerOption.text` | Required plain display label. |
| `FileFolderPickerOption.extension` | Optional file suffix for presentation; independent from a cosmetic label. |
| `FileFolderPickerOption.sizeInBytes` | Optional nonnegative safe integer byte size. |
| `FileFolderPickerChangeData.reason` | Explicit pick, remove or clear action. |
| `FileFolderPickerChangeData.values` | Complete ordered fresh JSON references, including unresolved retained values. |
| `FileFolderPickerChangeData.items` | Resolved selected models only, in selection order. |
| `FileFolderPickerChangeData.unresolvedValues` | Retained references without a resolved model; never fabricated as native preselection. |


### Dialog and current-user OneDrive

The only presentation is a responsive **1080 × 720 px** modal dialog. The toolkit adds no visible title, close button or internal padding: the Microsoft iframe fills the surface. A hidden semantic title names the dialog for assistive technology. Microsoft's own title, toolbar and Cancel/Select actions remain part of the hosted UI. Stage retains native dismissal and focus handling; Escape inside the remote iframe is governed by Microsoft. Smaller viewports constrain both dimensions with the existing Stage outer gap.

`presentation` and OneDrive `baseUrl` have been removed. Replace popup/dialog configuration with omission, and configure the current user's drive as follows:

```tsx
<FileFolderPicker preset="focused" scope={{ kind: 'oneDrive' }} itemTypes="all" multiselect />
<FileFolderPicker preset="focused" scope={{ kind: 'oneDrive', folder: 'Projects' }} itemTypes="files" />
```

The control uses the existing SharePoint REST client and mounted request cache to read `/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=PersonalUrl,Email` on the current collection. It validates the returned personal-site URL against the tenant, then verifies the drive owner against the current login/profile email and validates returned item provenance. Personal-domain metadata reads use the existing SPFx token provider and SharePoint HTTP client with a token for that exact trusted resource; the memory cache stores response models, never tokens. It does not infer an URL from an email address, invoke Graph `/me/drive`, provision a personal site, or request new consent. A missing personal URL produces an actionable unavailable message; profile/drive access failures retain existing selections. The same source contract applies to Property Pane. See Microsoft's [user profile REST documentation](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/work-with-user-profiles-in-sharepoint).


## Flat native configuration and stable opening

The following SDK-facing props are legacy-only when `preset` is omitted, except for the shared functional props/callbacks explicitly listed below. They work identically in React and Property Pane; do not combine them with a preset. There is no raw SDK options bag.

| Props | Default and contract |
| --- | --- |
| `title`, `selectLabel`, `cancelLabel` | Omit for Microsoft defaults; supplied strings cannot be blank. title changes only Microsoft's internal title. |
| `layout`, `stickyHeaders` | Omit for service defaults; layout is details, compact-details or tiles. |
| `initialView` | location. myFiles/recent/shared/site/organization require OneDrive Business scope and no initialLocation. Shared is service-dependent; a configured entry does not establish availability or nonempty results. |
| `initialSort`, `initialFilter` | Omit for service defaults. Use SharePoint internal fieldName; sort ascending defaults true. Filter is initial UI state, not an acceptance boundary. |
| `initialLocationFallback` | error. scopeRoot falls back only for a missing initial folder (404), to the verified source root. Does not hide permission failures, invalid paths or a missing source. |
| `visibleLocations` | Omission produces no requested pivots for SharePoint scopes and only myFiles for OneDrive. An explicit array, including [], replaces that default. Supported keys: myFiles/recent/shared/quickAccess/site/organization. Entries depend on source compatibility, Microsoft service availability and authentication policy; entry can keep a targeted pivot visible. |
| `navigationVisible`, `navigationPreset`, `navigationInitialState` | true; omitted preset is oneDrive for OneDrive and current-site for SharePoint; modality omitted. Explicit preset accepts oneDrive/current-site; modality modal/hidden. Hiding navigation grants no security guarantee. |
| `siteNavigationUrl` | Resolved scope web. Explicit URL must be the already resolved web, collection root or initial web; arbitrary unverified web URLs are rejected. |
| `retainSelectionOnNavigate` | true in multiple mode, false in single mode; current opening only, no history. |
| `requiredAccess` | read. read-write filters the native experience; does not grant rights, enable writes or establish an independent host permission proof. |
| `graphEnabled`, `claimsChallengeEnabled` | Independently false. Exact resource checks remain; validated JSON claims are ephemeral, with one attempt per audience per opening. |
| `uploadEnabled`, `createFolderEnabled` | Independently false. Require existing write grants and user destination rights. |
| `localFilesMode`, `localFilesLabel` | disabled; upload requires uploadEnabled. Optional localized label. Returning local device files is not supported. |
| `fromLinkMode` | disabled; navigation or pivot enables native link resolution. Only authoritative remote references passing the existing scope checks are accepted. |
| `columnFilterMode` | Omit for service default; panel or menu. |
| `trayPrompt` | selection-summary; also selection-editor or none. Destination/save/share prompts require a separate result contract. |
| `onReady` | Optional no-argument callback, once per opening on page-loaded, not on channel activation. Never commits selection. |
| `hint` | Existing Fluent Field description. Omit for no description. Supplied content stays unchanged during opening; no automatic opening text or spinner is added. |

Theme is automatic from the closest FluentProvider with SPFx/Teams fallback, mapped to native light/dark at Browse. There is no public theme override and no promise of arbitrary Fluent colors or high-contrast parity inside Microsoft content. The mode-specific browse icon remains visible and the animated lower border remains active while opening/open; duplicate openings are prevented.

All opening-only configuration, descriptors, arrays, theme and verified preselection are captured before async work on Browse. Updating them affects the next opening and does not close the active picker or repeat selected-reference hydration. Source/item/access policy, authoritative values/revision, disabled or disposal still invalidate stale results. Hydration resolves the accepted source independently of initial navigation.

Legacy-only example (omit `preset` deliberately):

```tsx
<FileFolderPicker label="Documents" hint="Select a project document"
  scope={{ kind: 'library', webUrl: 'https://tenant.sharepoint.com/sites/projects', list: 'Documents' }}
  title="Project documents" layout="details" stickyHeaders selectLabel="Use document" cancelLabel="Cancel"
  initialSort={{ fieldName: 'FileLeafRef', ascending: true }} navigationVisible
  graphEnabled={false} uploadEnabled={false} createFolderEnabled={false}
  onReady={() => console.info('Picker ready')} />
```

Native upload/create/device-upload may modify content before Pick. Cancel, rejected selection or Property Pane Apply/Discard cannot undo those changes. The selectable-source boundary validates accepted references; it is not a pre-write destination boundary and cannot guarantee confinement of all Microsoft-owned writes. Use a separately validated host operation if strict pre-write destination control is needed. The control never grants permissions, provisions OneDrive, logs tokens or automatically persists principal/file history.

## Permissions by enabled feature

The baseline obtains a SharePoint-audience token through AadTokenProvider. SPHttpClient metadata access and iframe bearer-token authorization are distinct; success of one does not prove the other's grants. Permissions are delegated and bounded by user rights; existing shared SPFx tenant grants may already suffice.

| Feature | SharePoint delegated permission | Optional Graph dependency |
| --- | --- | --- |
| SharePoint read selection | MyFiles.Read or AllSites.Read | None required by baseline |
| Current-user OneDrive read | MyFiles.Read | None required by baseline |
| SharePoint native upload/create | MyFiles.Write or AllSites.Write | No inherent Graph requirement |
| OneDrive native upload/create | MyFiles.Write | No inherent Graph requirement |
| Teams-channel discovery when requested by service | AllSites.Read; AllSites.Write for write | ChannelSettings.Read.All |
| Presentation/filter/selection settings | No extra grant solely for configuration | None |
| Claims challenge | Same approved resource and extra claims | Does not itself grant any new permission |

Microsoft also lists Graph Files.Read/Files.ReadWrite alternatives in its operation table; these do not substitute a Graph token for the SharePoint-audience token used by our hosted iframe. Do not union all scopes from the sample. Optional service pivots may request additional audiences; test the actual enabled experience before declaring a minimal deployment. The toolkit changes neither the package manifest nor tenant approvals. Example consumer baseline request (only if not already approved):

```json
{ "webApiPermissionRequests": [{ "resource": "SharePoint", "scope": "MyFiles.Read" }] }
```

For a consumer choosing an all-sites read grant use AllSites.Read instead; add write or Graph requests separately only for enabled capabilities. Admin approval and user rights are still required. See [Microsoft permissions](https://learn.microsoft.com/en-us/onedrive/developer/controls/file-pickers/?view=odsp-graph-online) and [SPFx API access](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aadhttpclient).

## Explicit compatibility gates

The [current v8 schema](https://learn.microsoft.com/en-us/onedrive/developer/controls/file-pickers/v8-schema?view=odsp-graph-online) contains additional options whose presence does not establish their complete message/result contracts. Local-file pick, draft-selection payloads, custom/formatter commands, host-rendered tray/navigation, account switching and download/share/move/save results are not exposed as supported toolkit actions. No guessed callbacks or inert configuration flags are forwarded. Copy has a schema block without a copy action enum. Substrate, parent/configure protocol and SDK timeout overrides remain internal research gates. Consumer-only photos are outside SPFx Business support; site/list/listItem/Syntex domains require different result models and cannot be coerced into file/folder references. File/media category combinations need verified filter semantics; use existing extension filters for enforced file suffixes.

Automated config/auth/lifecycle tests and browser/tenant acceptance are separate evidence. Enabling a capability does not prove service availability, tenant grants or actual internal API behavior; see the [validation guide](../../SHAREPOINT-VALIDATION.md).


### Trying presets in the sample

React and Property Pane use the same preset contract and source-aware configuration. New scenarios start with explicit focused. Choose extended to exercise service-owned navigation, while retaining the same accepted source. The ordinary demo exposes functional choices rather than SDK/authentication switches; advanced diagnostics report applied settings and JSON references without becoming a second configurator. The sample Property Pane offers only focused and extended. Saved legacy sample configurations display and use focused until extended is selected; saved references and old SDK settings are retained. This sample policy does not remove the reusable control’s legacy contract. Explicit effective-source changes in the demo clear dependent starting-location settings; preset-only and cosmetic updates retain them. The reusable control leaves consumer-owned source/selection updates to the consumer.

The validation guide distinguishes route navigation, empty content, denied resources and tooling limitations. A successful token acquisition alone does not establish a working Media route or a successful file selection.

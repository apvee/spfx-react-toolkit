/** Serializable reference contracts for the toolkit pickers. */

/** Optional display snapshot, never authoritative identity. It may become stale after a rename.
 * @example const label: PickerValueLabel = { text: 'Marketing' };
 */
export interface PickerValueLabel {
  /** Optional plain label used until live metadata resolves. Omission means no fallback label. */
  readonly text?: string;
}
/** Serializable site reference. The absolute collection URL also supports modes without a site GUID.
 * @example const site: SitePickerValue = { kind: 'site', url: 'https://tenant.sharepoint.com/sites/hr' };
 */
export interface SitePickerValue extends PickerValueLabel {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'site';
  /** Absolute normalized site collection URL. */
  readonly url: string;
  /** Site collection GUID, when returned/verified. */
  readonly siteId?: string;
}
/** Serializable list reference scoped to its source web.
 * @example const list: ListPickerValue = { kind: 'list', webUrl: 'https://tenant.sharepoint.com/sites/hr', listId: '11111111-1111-4111-8111-111111111111' };
 */
export interface ListPickerValue extends PickerValueLabel {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'list';
  /** Required absolute HTTP(S) source web URL; serialized values have no provider-current-web default. */
  readonly webUrl: string;
  /** Required list GUID scoped to webUrl; lists have no site-column variant. */
  readonly listId: string;
}
/** Serializable field reference; listId is absent only for site columns.
 * @example const field: FieldPickerValue = { kind: 'field', webUrl: 'https://tenant.sharepoint.com/sites/hr', fieldId: '22222222-2222-4222-8222-222222222222', internalName: 'Title' };
 */
export interface FieldPickerValue extends PickerValueLabel {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'field';
  /** Absolute HTTP(S) source web URL. Omission, when allowed, uses the provider current web; never inferred from a display label. */
  readonly webUrl: string;
  /** Scoped list GUID, agreeing with the JSON reference. For site columns this is omitted. */
  readonly listId?: string;
  /** Required field GUID within the source web/list. */
  readonly fieldId: string;
  /** Required field internal name, agreeing with the JSON reference; distinct from a REST EntityPropertyName. */
  readonly internalName: string;
}
/** Serializable view reference scoped to its list/web.
 * @example const view: ViewPickerValue = { kind: 'view', webUrl: 'https://tenant.sharepoint.com/sites/hr', listId: '11111111-1111-4111-8111-111111111111', viewId: '33333333-3333-4333-8333-333333333333' };
 */
export interface ViewPickerValue extends PickerValueLabel {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'view';
  /** Required absolute HTTP(S) source web URL; serialized values have no provider-current-web default. */
  readonly webUrl: string;
  /** Required source list GUID; views always belong to a list in webUrl. */
  readonly listId: string;
  /** Required view GUID within its source list. */
  readonly viewId: string;
}
/** Serializable content type reference. The hierarchical SharePoint ID is not a GUID.
 * Site values identify the verified collection root; list values retain the owning web and list copy ID.
 * JSON contains only this reference and the optional inherited display snapshot, never group/flags.
 * @example const site: ContentTypePickerValue = {kind:'contentType',scope:'site',webUrl:'https://tenant.sharepoint.com/sites/hr',contentTypeId:'0x0101'};
 * @example const list: ContentTypePickerValue = {kind:'contentType',scope:'list',webUrl:'https://tenant.sharepoint.com/sites/hr/sub',listId:'11111111-1111-4111-8111-111111111111',contentTypeId:'0x0101'};
 */
export type ContentTypePickerValue = PickerValueLabel & {
  /** Required content type domain discriminator. */
  readonly kind: 'contentType';
  /** Required hierarchical ID: System root 0x or 0x followed by nonzero hex bytes or 00 plus 32 hex GUID digits; comparison ignores case. */
  readonly contentTypeId: string;
} & ({
  /** Collection catalog scope; never inferred from listId. */
  readonly scope: 'site';
  /** Required verified absolute collection root web URL. */
  readonly webUrl: string;
  /** Forbidden for site content types. */
  readonly listId?: never;
} | {
  /** List catalog scope; preserves the list copy rather than its site ancestor. */
  readonly scope: 'list';
  /** Required absolute URL of the web owning the list. */
  readonly webUrl: string;
  /** Required GUID identifying the owning list. */
  readonly listId: string;
});
/** Atomic custom item key, both name and typed value required.
 * @example const custom: PickerCustomItemKey = { field: 'ExternalCode', value: 'A-001' };
 */
export interface PickerCustomItemKey {
  /** Verified internal field name; must select unique finite-number/string values. */
  readonly field: string;
  /** String/number discriminator is retained; numeric 1 differs from string '1'. */
  readonly value: string | number;
}
/** Serializable item identity; custom key field/value coexist with authoritative SharePoint item ID.
 * @example const item: ListItemPickerValueBase = { webUrl: 'https://tenant.sharepoint.com/sites/hr', listId: '11111111-1111-4111-8111-111111111111', itemId: 42 };
 */
export interface ListItemPickerValueBase extends PickerValueLabel {
  /** Required absolute HTTP(S) source web URL; serialized values have no provider-current-web default. */
  readonly webUrl: string;
  /** Required source list GUID for the item, file or folder, scoped to webUrl. */
  readonly listId: string;
  /** Required positive safe integer SharePoint item ID, scoped to the web/list. Always retained even with a custom key. */
  readonly itemId: number;
  /** Custom key pair; omission uses authoritative itemId (Id). */
  readonly customKey?: PickerCustomItemKey;
}
/** Ordinary list record; no file/folder semantics.
 * @example const item: SharePointItemValue = { kind: 'item', webUrl: 'https://tenant.sharepoint.com/sites/hr', listId: '11111111-1111-4111-8111-111111111111', itemId: 42 };
 */
export interface SharePointItemValue extends ListItemPickerValueBase {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'item'; }
/** Document-library file; extension/path are optional display/navigation snapshots, not identity.
 * @example const file: SharePointFileValue = { kind: 'file', webUrl: 'https://tenant.sharepoint.com/sites/hr', listId: '11111111-1111-4111-8111-111111111111', itemId: 7, uniqueId: '44444444-4444-4444-8444-444444444444', extension: 'pdf' };
 */
export interface SharePointFileValue extends ListItemPickerValueBase {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'file';
  /** Verified File.UniqueId; required for a confirmed file value. */
  readonly uniqueId: string;
  /** Optional server-relative navigation snapshot. Not canonical identity; must have one leading slash and no query, fragment or traversal segments. */
  readonly serverRelativeUrl?: string;
  /** Optional verified file extension snapshot, without a leading dot; used only for presentation. */
  readonly extension?: string;
}
/** SharePoint folder item; only offered when configured to include folders.
 * @example const folder: SharePointFolderValue = { kind: 'folder', webUrl: 'https://tenant.sharepoint.com/sites/hr', listId: '11111111-1111-4111-8111-111111111111', itemId: 9, uniqueId: '55555555-5555-4555-8555-555555555555' };
 */
export interface SharePointFolderValue extends ListItemPickerValueBase {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'folder';
  /** Required verified File/Folder UniqueId GUID. Library origin alone does not establish file or folder identity. */
  readonly uniqueId: string;
  /** Optional server-relative navigation snapshot. Not canonical identity; must have one leading slash and no query, fragment or traversal segments. */
  readonly serverRelativeUrl?: string;
}
/** Item/file/folder discrimination drives rendering and downstream behavior.
 * @example const values: ListItemPickerValue[] = [{kind:'item',webUrl:'https://tenant.sharepoint.com/sites/hr',listId:'11111111-1111-4111-8111-111111111111',itemId:42}];
 */
export type ListItemPickerValue = SharePointItemValue | SharePointFileValue | SharePointFolderValue;
/** Verified Entra object identity; never use Graph person.id as this ID.
 * @example const identity: EntraObjectIdentity = { tenantId: '66666666-6666-4666-8666-666666666666', id: '77777777-7777-4777-8777-777777777777' };
 */
export interface EntraObjectIdentity {
  /** Required verified Entra tenant GUID on serialized identities; optional redundant display-model metadata must agree when provided. */
  readonly tenantId: string;
  /** Required native scoped identifier: GUID for Entra/list/field/view, positive safe integer for SharePoint principals. Must agree with the JSON reference. */
  readonly id: string; }
/** SharePoint user/group ID belongs to one site collection, not to the tenant.
 * @example const identity: SharePointPrincipalIdentity = { siteCollectionUrl: 'https://tenant.sharepoint.com/sites/hr', id: 42 };
 */
export interface SharePointPrincipalIdentity {
  /** Required verified absolute site collection URL owning the numeric principal ID. */
  readonly siteCollectionUrl: string;
  /** Required positive safe integer principal ID scoped to its site collection. */
  readonly id: number; }
/** Serializable user reference. available mode may retain verified login only; requested ID modes require the respective identities on confirmation.
 * @example const user: UserPrincipalValue = { kind: 'user', loginName: 'user@tenant.com', entra: {tenantId:'66666666-6666-4666-8666-666666666666',id:'77777777-7777-4777-8777-777777777777'}, sharePoint:{siteCollectionUrl:'https://tenant.sharepoint.com/sites/hr',id:42} };
 */
export interface UserPrincipalValue extends PickerValueLabel {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'user';
  /** Required verified login/UPN or SharePoint membership claim. Login-only identity requires the provider verified tenant context. */
  readonly loginName: string;
  /** Optional verified tenant-scoped Entra user identity. Required on confirmed selection in entra/both modes. */
  readonly entra?: EntraObjectIdentity;
  /** Optional verified collection-scoped SharePoint user identity. Required on confirmed selection in sharePoint/both modes. */
  readonly sharePoint?: SharePointPrincipalIdentity;
}
/** Serializable SharePoint group reference.
 * @example const group: SharePointGroupValue = {kind:'sharePointGroup',identity:{siteCollectionUrl:'https://tenant.sharepoint.com/sites/hr',id:7}};
 */
export interface SharePointGroupValue extends PickerValueLabel {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'sharePointGroup';
  /** Required scoped native group identity; never substituted with a name or email. */
  readonly identity: SharePointPrincipalIdentity;
}
/** Serializable Entra group reference; type remains explicit.
 * @example const group: EntraGroupValue = {kind:'microsoft365Group',identity:{tenantId:'66666666-6666-4666-8666-666666666666',id:'77777777-7777-4777-8777-777777777777'}};
 */
export interface EntraGroupValue extends PickerValueLabel {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'microsoft365Group' | 'securityGroup' | 'distributionList';
  /** Required scoped native group identity; never substituted with a name or email. */
  readonly identity: EntraObjectIdentity;
}
/** Explicit unvalidated email value; cannot masquerade as a verified user.
 * @example const email: UnvalidatedEmailValue = {kind:'unvalidatedEmail',email:'external@example.com'};
 */
export interface UnvalidatedEmailValue extends PickerValueLabel {
  /** Required discriminant selecting the entity or reference variant; never inferred from a display name. */
  readonly kind: 'unvalidatedEmail';
  /** Required syntactically valid but unverified email address, explicitly distinct from a directory user identity. */
  readonly email: string; }
/** Serializable principal values; users may retain both scoped ID systems.
 * @example const principals: PrincipalPickerValue[] = [{kind:'user',loginName:'user@tenant.com'}];
 */
export type PrincipalPickerValue = UserPrincipalValue | SharePointGroupValue | EntraGroupValue | UnvalidatedEmailValue;
/** All reference value domains. Fixed JSON primitives only; runtime validation is still required for external input.
 * @example const value: PickerValue = {kind:'site',url:'https://tenant.sharepoint.com/sites/hr'};
 */
export type PickerValue = SitePickerValue | ListPickerValue | FieldPickerValue | ViewPickerValue | ContentTypePickerValue | ListItemPickerValue | PrincipalPickerValue;
/** Stable array shape for both single and multi selection. Single allows zero/one new value; existing conflicts are preserved for explicit resolution.
 * @example interface WebPartProperties { sites: PickerSelection<SitePickerValue>; }
 */
export type PickerSelection<TValue extends PickerValue> = readonly TValue[];

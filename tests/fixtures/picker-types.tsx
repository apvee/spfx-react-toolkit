import type { PickerPropertyKey } from '../../packages/spfx-react-toolkit/src/property-pane/pickers/property-pane-picker.types';
import type { PrincipalPickerProps, PrincipalPickerOption, UserIdentityResolution } from '../../packages/spfx-react-toolkit/src/controls/pickers/picker.types';
const mixed: PrincipalPickerProps = {
  presentation: 'tagPicker', multiselect: true,
  principalTypes: ['user', 'sharePointGroup', 'microsoft365Group', 'securityGroup', 'distributionList'],
  userIdentity: { mode: 'both', ensureUser: true, webUrl: 'https://tenant.sharepoint.com/sites/projects' },
  onOptionSelect: (_, data) => data.selectedItems.forEach(item => {
    if (item.kind === 'user') console.log(item.entraId, item.sharePointId);
    if (item.kind === 'sharePointGroup') console.log(item.id, item.webUrl);
  }),
};
const entra: UserIdentityResolution = { mode: 'entra' };
// @ts-expect-error ensureUser is not valid in Entra-only mode.
const invalid: UserIdentityResolution = { mode: 'entra', ensureUser: true };
const sharePointGroup: PrincipalPickerOption = { key: 'demo', text: 'Visitors', kind: 'sharePointGroup', source: 'sharePoint', id: 7, webUrl: 'https://tenant.sharepoint.com/sites/projects', value: {kind:'sharePointGroup',identity:{siteCollectionUrl:'https://tenant.sharepoint.com/sites/projects',id:7}} };
// @ts-expect-error SharePoint group identity must include its web scope.
const unscoped: PrincipalPickerOption = { key: 'demo', text: 'Visitors', kind: 'sharePointGroup', source: 'sharePoint', id: 7, value:{kind:'sharePointGroup',identity:{siteCollectionUrl:'https://tenant.sharepoint.com/sites/projects',id:7}} };
void [mixed, entra, invalid, sharePointGroup, unscoped];

const suggestions: PrincipalPickerProps = { loadPersonalSuggestions: true, availableOptions: [sharePointGroup], initialSuggestions: { includeCurrentUser: true, limit: 5 } };
// @ts-expect-error The total initial suggestion limit cannot exceed five.
const tooMany: PrincipalPickerProps = { initialSuggestions: { limit: 6 } };
void [suggestions, tooMany];

import type { SitePickerProps, ListPickerProps, FieldPickerProps, ViewPickerProps, ListItemPickerProps } from '../../packages/spfx-react-toolkit/src/controls/pickers/picker.types';
const shared: Pick<SitePickerProps, 'availableOptions' | 'initialOptions' | 'defaultSelectedOptions' | 'initialSuggestions' | 'colorfulTagAvatars'> = {
  colorfulTagAvatars: true,
  defaultSelectedOptions: ['https://tenant.sharepoint.com/sites/hr'],
  initialOptions: [{ key: 'https://tenant.sharepoint.com/sites/hr', text: 'HR', url: 'https://tenant.sharepoint.com/sites/hr', value:{kind:'site',url:'https://tenant.sharepoint.com/sites/hr'} }],
  availableOptions: [], initialSuggestions: { limit: 5 }
};
const others: [ListPickerProps, FieldPickerProps, ViewPickerProps, ListItemPickerProps] = [
  {availableOptions: [], initialSuggestions: {limit: 5}}, {availableOptions: []}, {availableOptions: [], list: 'Projects'}, {availableOptions: [], list: 'Projects', textField: 'Title'}
];
// @ts-expect-error Automatic selection history is removed.
const noHistory: PrincipalPickerProps = {initialSuggestions: {recentSelections: true}};
// @ts-expect-error Site history is also removed.
const noSiteHistory: SitePickerProps = {recentSelections: {enabled: true}};
void [shared, others, noHistory, noSiteHistory];
// @ts-expect-error Avatar coloring is a boolean opt-in, not an appearance string.
const invalidAvatarColor: PrincipalPickerProps = { colorfulTagAvatars: 'colorful' };
void invalidAvatarColor;

import type { SitePickerValue, ListPickerValue, FieldPickerValue, ViewPickerValue, ListItemPickerValue, PrincipalPickerValue } from '../../packages/spfx-react-toolkit/src/controls/pickers/picker-values.types';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
interface WebPartProperties {
  sites: SitePickerValue[];
  lists?: readonly ListPickerValue[];
  fields: FieldPickerValue[];
  views: ViewPickerValue[];
  items: ListItemPickerValue[];
  principals: PrincipalPickerValue[];
  title: string;
  mixed: (SitePickerValue | ListPickerValue)[];
}
declare const instance: BaseClientSideWebPart<WebPartProperties>;
const item: ListItemPickerValue = {kind:'item',webUrl:'https://tenant.sharepoint.com/sites/hr',listId:'11111111-1111-4111-8111-111111111111',itemId:42};
const file: ListItemPickerValue = {kind:'file',webUrl:item.webUrl,listId:item.listId,itemId:7,uniqueId:'44444444-4444-4444-8444-444444444444',extension:'pdf'};
// @ts-expect-error Confirmed file values require verified uniqueId.
const missingFileIdentity: ListItemPickerValue = {kind:'file',webUrl:item.webUrl,listId:item.listId,itemId:7};
// @ts-expect-error Minimal stored references do not contain arbitrary fields or callbacks.
const invalidStoredData: ListItemPickerValue = {kind:'item',webUrl:item.webUrl,listId:item.listId,itemId:7,fields:{onClick:()=>{}}};
const userValue: PrincipalPickerValue = {kind:'user',loginName:'user@tenant.com',entra:{tenantId:'66666666-6666-4666-8666-666666666666',id:'77777777-7777-4777-8777-777777777777'},sharePoint:{siteCollectionUrl:item.webUrl,id:42}};
// @ts-expect-error Entra identities require tenant scope.
const unscopedEntra: PrincipalPickerValue = {kind:'user',loginName:'user@tenant.com',entra:{id:'77777777-7777-4777-8777-777777777777'}};
type ValidSiteKey = PickerPropertyKey<WebPartProperties,SitePickerValue>;
const validSiteKey: ValidSiteKey = 'sites';
void [item,file,missingFileIdentity,invalidStoredData,userValue,unscopedEntra,validSiteKey];
// Unsafe/untyped property domains cannot be silently selected.
type UnknownKey = PickerPropertyKey<{unsafe: unknown}, SitePickerValue>;
// @ts-expect-error unknown properties are not eligible.
const unknownKey: UnknownKey = 'unsafe';
type UntypedKey = PickerPropertyKey<{unsafe: ReturnType<typeof JSON.parse>}, SitePickerValue>;
// @ts-expect-error any properties are not eligible.
const untypedKey: UntypedKey = 'unsafe';
void [unknownKey, untypedKey];

type UnsafeArrays = PickerPropertyKey<{anyArray: ReturnType<typeof JSON.parse>[]; readonlyAny: readonly ReturnType<typeof JSON.parse>[]; unknownArray: unknown[]; impossible: never[]; optionalUnsafe?: ReturnType<typeof JSON.parse>[]; correct?: readonly SitePickerValue[]}, SitePickerValue>;
const correctArray: UnsafeArrays = 'correct';
// @ts-expect-error any array elements are rejected.
const unsafeArray: UnsafeArrays = 'anyArray';
// @ts-expect-error readonly any arrays are rejected.
const unsafeReadonlyArray: UnsafeArrays = 'readonlyAny';
// @ts-expect-error unknown array elements are rejected.
const unknownArray: UnsafeArrays = 'unknownArray';
// @ts-expect-error never element type is not a meaningful property domain.
const neverArray: UnsafeArrays = 'impossible';
// @ts-expect-error optional any arrays are rejected.
const optionalUnsafeArray: UnsafeArrays = 'optionalUnsafe';
const customItem: ListItemPickerValue = {...item,customKey:{field:'ExternalCode',value:'A-001'}};
// @ts-expect-error Custom field/value pair is atomic.
const incompleteCustomKey: ListItemPickerValue = {...item,customKey:{field:'ExternalCode'}};
void [correctArray,unsafeArray,unsafeReadonlyArray,unknownArray,neverArray,optionalUnsafeArray,customItem,incompleteCustomKey];

import type {UserPrincipalValue} from '../../packages/spfx-react-toolkit/src/controls/pickers/picker-values.types';
type NarrowPrincipalKey = PickerPropertyKey<{users:UserPrincipalValue[]},PrincipalPickerValue>;
// @ts-expect-error The factory may emit all principal categories, not only users.
const narrowPrincipalKey: NarrowPrincipalKey = 'users';
type TupleKey = PickerPropertyKey<{one:[SitePickerValue]; readonlyOne:readonly [SitePickerValue]; nonEmpty:[SitePickerValue,...SitePickerValue[]]},SitePickerValue>;
// @ts-expect-error Factory can emit empty/multiple arrays, not a fixed one-value tuple.
const fixedTupleKey: TupleKey = 'one';
// @ts-expect-error Readonly fixed tuples are also unsafe targets.
const readonlyTupleKey: TupleKey = 'readonlyOne';
// @ts-expect-error Required prefix in a variadic tuple cannot accept empty selections.
const nonEmptyTupleKey: TupleKey = 'nonEmpty';
type ExtraKey = PickerPropertyKey<{sites:(SitePickerValue & {requiredExtra:string})[]},SitePickerValue>;
// @ts-expect-error The factory cannot emit consumer-required additional fields.
const extraKey: ExtraKey = 'sites';
void [narrowPrincipalKey,fixedTupleKey,readonlyTupleKey,nonEmptyTupleKey,extraKey];

// Literal property names that would be interpreted as nested SDK paths are also unsafe.
type UnsafePaths = PickerPropertyKey<{'site.values':SitePickerValue[]; 'sites[0]':SitePickerValue[]; '__proto__':SitePickerValue[]; sites:SitePickerValue[]},SitePickerValue>;
// @ts-expect-error SDK treats dots as paths.
const dotted: UnsafePaths = 'site.values';
// @ts-expect-error Array navigation is not a direct property.
const indexed: UnsafePaths = 'sites[0]';
// @ts-expect-error Prototype-sensitive property names are not field targets.
const dangerous: UnsafePaths = '__proto__';
void [dotted,indexed,dangerous];

// @ts-expect-error SitePicker selects collections, never the former web mode.
const invalidSiteMode: import('../../packages/spfx-react-toolkit/src/controls/pickers/picker.types').SitePickerProps = { mode: 'web' };
void invalidSiteMode;

const catalogProps: [ListPickerProps, FieldPickerProps, ViewPickerProps] = [{ initialCatalogLimit: 30 }, {}, { list: 'Projects' }];
// @ts-expect-error Site has suggestions, not the metadata catalog.
const noSiteCatalog: SitePickerProps = { initialCatalogLimit: 30 };
// @ts-expect-error Item has suggestions, not the metadata catalog.
const noItemCatalog: ListItemPickerProps = { list: 'Projects', textField: 'Title', initialCatalogLimit: 30 };
// @ts-expect-error Principal keeps the five-item suggestion policy.
const noPrincipalCatalog: PrincipalPickerProps = { initialCatalogLimit: 30 };
void [catalogProps, noSiteCatalog, noItemCatalog, noPrincipalCatalog];

// @ts-expect-error Field catalogs are complete; the display limit is removed.
const noFieldLimit: FieldPickerProps = { initialCatalogLimit: 30 };
// @ts-expect-error View catalogs are complete; the display limit is removed.
const noViewLimit: ViewPickerProps = { list: 'Projects', initialCatalogLimit: 30 };
void [noFieldLimit, noViewLimit];

// @ts-expect-error SitePicker no longer exposes the current-collection restriction.
const removedCollectionLimit: SitePickerProps = { limitToCurrentSiteCollection: true };
void removedCollectionLimit;

const siteContext: SitePickerProps = { includeCurrentSite: true, includeHubSite: false };
void siteContext;

import type { ContentTypePickerProps, ContentTypePickerOption } from '../../packages/spfx-react-toolkit/src/controls/pickers/picker.types';
import type { ContentTypePickerValue } from '../../packages/spfx-react-toolkit/src/controls/pickers/picker-values.types';
const siteContentType: ContentTypePickerValue = { kind: 'contentType', scope: 'site', webUrl: item.webUrl, contentTypeId: '0x0101' };
const listContentType: ContentTypePickerValue = { kind: 'contentType', scope: 'list', webUrl: item.webUrl, listId: item.listId, contentTypeId: '0x0101' };
const ctProps: ContentTypePickerProps[] = [{}, { scope: 'site', group: 'Documents' }, { scope: 'list', list: { kind: 'id', id: item.listId }, includeHidden: false, contentTypeId: '0x01', includeDescendants: true }];
// @ts-expect-error List scope needs an exact toolkit list selector.
const ctMissingList: ContentTypePickerProps = { scope: 'list' };
// @ts-expect-error Site scope cannot receive a list.
const ctSiteList: ContentTypePickerProps = { scope: 'site', list: 'Documents' };
// @ts-expect-error Site reference cannot carry a list GUID.
const ctSiteListValue: ContentTypePickerValue = { ...siteContentType, listId: item.listId };
const ctOption: ContentTypePickerOption = { kind: 'contentType', key: 'document', text: 'Document', webUrl: item.webUrl, id: '0x0101', name: 'Document', group: 'Documents', hidden: false, readOnly: false, sealed: false, value: siteContentType };
type CtPropertyKey = PickerPropertyKey<{ contentTypes: ContentTypePickerValue[]; fields: FieldPickerValue[] }, ContentTypePickerValue>;
const ctKey: CtPropertyKey = 'contentTypes';
// @ts-expect-error Field references are not content type references.
const ctWrongKey: CtPropertyKey = 'fields';
void [siteContentType, listContentType, ctProps, ctMissingList, ctSiteList, ctSiteListValue, ctOption, ctKey, ctWrongKey];

import { ContentTypePicker } from '../../packages/spfx-react-toolkit/src/controls/pickers';
import * as React from 'react';
const siteContentTypeControl = <ContentTypePicker label="Site content type" />;
const listContentTypeControl = <ContentTypePicker label="List content type" scope="list" list={{ kind: 'id', id: item.listId }} />;
// @ts-expect-error The public component retains the list-scope selector requirement.
const invalidContentTypeControl = <ContentTypePicker label="Content type" scope="list" />;
void [siteContentTypeControl, listContentTypeControl, invalidContentTypeControl];

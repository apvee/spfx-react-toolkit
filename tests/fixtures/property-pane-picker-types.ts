import { PropertyPaneContentTypePicker } from '../../packages/spfx-react-toolkit/src/property-pane/pickers/property-pane-content-type-picker';
import { PropertyPaneSitePicker } from '../../packages/spfx-react-toolkit/src/property-pane/pickers/property-pane-site-picker';
import { PropertyPaneListPicker } from '../../packages/spfx-react-toolkit/src/property-pane/pickers/property-pane-list-picker';
import { PropertyPaneFieldPicker } from '../../packages/spfx-react-toolkit/src/property-pane/pickers/property-pane-field-picker';
import { PropertyPaneViewPicker } from '../../packages/spfx-react-toolkit/src/property-pane/pickers/property-pane-view-picker';
import { PropertyPaneListItemPicker } from '../../packages/spfx-react-toolkit/src/property-pane/pickers/property-pane-list-item-picker';
import { PropertyPanePrincipalPicker } from '../../packages/spfx-react-toolkit/src/property-pane/pickers/property-pane-principal-picker';
import type { ContentTypePickerValue, SitePickerValue, ListPickerValue, FieldPickerValue, ViewPickerValue, ListItemPickerValue, PrincipalPickerValue } from '../../packages/spfx-react-toolkit/src/controls/pickers/picker-values.types';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
interface WebPartProperties {
  contentTypes: ContentTypePickerValue[];
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
PropertyPaneSitePicker('sites', {instance,key:'sites-field',value:[],colorfulTagAvatars:true});
PropertyPaneListPicker('lists', {instance,key:'lists-field',value:[],initialCatalogLimit:30});
PropertyPanePrincipalPicker('principals', {instance,key:'principals-field',value:[],colorfulTagAvatars:false,userIdentity:{mode:'both',ensureUser:false}});
// @ts-expect-error Wrong property domain is rejected.
PropertyPaneSitePicker('principals', {instance,key:'bad-field',value:[]});
// @ts-expect-error Scalar property is rejected.
PropertyPaneListPicker('title', {instance,key:'bad-field',value:[]});
// @ts-expect-error Mixed domains cannot be used for a site-only field.
PropertyPaneSitePicker('mixed', {instance,key:'bad-field',value:[]});
// @ts-expect-error Only value references are accepted, not raw string keys.
PropertyPaneSitePicker('sites', {instance,key:'bad-field',value:['site-key']});

PropertyPaneFieldPicker('fields', {instance,key:'fields',value:[],list:'Projects'});
PropertyPaneViewPicker('views', {instance,key:'views',value:[],list:'Projects'});
PropertyPaneListItemPicker('items', {instance,key:'items',value:[],list:'Projects',textField:'Title'});
// @ts-expect-error Views require a list.
PropertyPaneViewPicker('views', {instance,key:'views',value:[]});
// @ts-expect-error Native values never accept unresolved UI aliases.
PropertyPanePrincipalPicker('principals', {instance,key:'principals',value:['ui-key']});

// @ts-expect-error Complete native Field catalog has no limit property.
PropertyPaneFieldPicker('fields', {instance,key:'fields',value:[],initialCatalogLimit:30});
// @ts-expect-error Complete native View catalog has no limit property.
PropertyPaneViewPicker('views', {instance,key:'views',value:[],list:'Projects',initialCatalogLimit:30});

// @ts-expect-error The native SitePicker shares the removed source restriction contract.
PropertyPaneSitePicker('sites', {instance,key:'sites',value:[],limitToCurrentSiteCollection:true});

PropertyPaneSitePicker('sites', {instance,key:'site-context',value:[],includeCurrentSite:true,includeHubSite:false});

PropertyPaneContentTypePicker('contentTypes', {instance,key:'content-types',value:[]});
PropertyPaneContentTypePicker('contentTypes', {instance,key:'content-types',value:[],scope:'list',list:{kind:'id',id:'11111111-1111-4111-8111-111111111111'}});
// @ts-expect-error List-scoped content types require a list.
PropertyPaneContentTypePicker('contentTypes', {instance,key:'content-types',value:[],scope:'list'});
// @ts-expect-error Site scope cannot carry a list.
PropertyPaneContentTypePicker('contentTypes', {instance,key:'content-types',value:[],scope:'site',list:'Projects'});
// @ts-expect-error Omitted site scope cannot carry a list either.
PropertyPaneContentTypePicker('contentTypes', {instance,key:'content-types',value:[],list:'Projects'});
// @ts-expect-error Content types must target their own property domain.
PropertyPaneContentTypePicker('views', {instance,key:'content-types',value:[]});
// @ts-expect-error List references require listId in serialized content type values.
PropertyPaneContentTypePicker('contentTypes', {instance,key:'content-types',value:[{kind:'contentType',scope:'list',webUrl:'https://tenant.test/sites/hr',contentTypeId:'0x01'}]});

import * as React from 'react';
import type { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import { makeStyles } from '@fluentui/react-components';
import { SitePicker, ListPicker, FieldPicker, ViewPicker, ContentTypePicker, ListItemPicker, PrincipalPicker } from '@apvee/spfx-react-toolkit/controls/pickers';
import { PropertyPaneSitePicker, PropertyPaneListPicker, PropertyPaneFieldPicker, PropertyPaneViewPicker, PropertyPaneContentTypePicker, PropertyPaneListItemPicker, PropertyPanePrincipalPicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
import { sites, lists, fields, views, contentTypes, items, principals, names, guid, tenant, type Name } from './data';
import type { PrincipalPickerOption, SitePickerValue, ListPickerValue, FieldPickerValue, ViewPickerValue, ContentTypePickerValue, ListItemPickerValue, PrincipalPickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
import { host } from './transport';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
interface MediaProperties { sites: SitePickerValue[]; lists: ListPickerValue[]; fields: FieldPickerValue[]; views: ViewPickerValue[]; contentTypes: ContentTypePickerValue[]; items: ListItemPickerValue[]; principals: PrincipalPickerValue[]; }
// Simulated owner only; the actual native factory still validates each domain value.
const instance = host as unknown as BaseClientSideWebPart<MediaProperties>;

const sizes = ['medium', 'large'] as const;
type Size = typeof sizes[number];
// A deterministic decoded image exercises the photo path without external requests or personal data.
const photo = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=';
const people: readonly PrincipalPickerOption[] = principals.slice(0, 3).map((option, index) => ({ ...option, imageUrl: index === 1 ? photo : index === 2 ? 'data:image/png;base64,broken' : undefined }));
const group: PrincipalPickerOption = { kind: 'microsoft365Group', key: 'group-media', text: 'Marketing group', source: 'entra', id: guid(90), tenantId: tenant, securityEnabled: false, mailEnabled: true,
  value: { kind: 'microsoft365Group', identity: { tenantId: tenant, id: guid(90) } } };
const principalModels = [...people, group];
const useStyles = makeStyles({ matrix: { display: 'grid', gridTemplateColumns: 'repeat(3, minmax(0, 1fr))', gap: '16px' }, cell: { padding: '12px', minWidth: 0, border: '1px solid #ccc' } });

function MediaCase({ name, size, native }: { name: Name; size: Size; native: boolean }) {
  const styles = useStyles(), target = React.useRef<HTMLDivElement>(null);
  const id = `${native ? 'native' : 'react'}-${name}-${size}`;
  const common = { label: id, size, presentation: 'tagPicker' as const, multiselect: true };
  React.useEffect(() => {
    if (!native || target.current === null) return;
    const element = target.current;
    const nativeProps = { ...common, instance, key: id };
    let field: IPropertyPaneField<IPropertyPaneCustomFieldProps>;
    switch (name) {
      case 'SitePicker': field = PropertyPaneSitePicker('sites', { ...nativeProps, value: sites.slice(0, 1).map(o => o.value), initialOptions: sites.slice(0, 1), mode: 'hub' }); break;
      case 'ListPicker': field = PropertyPaneListPicker('lists', { ...nativeProps, value: lists.slice(0, 2).map(o => o.value), initialOptions: lists.slice(0, 2) }); break;
      case 'FieldPicker': field = PropertyPaneFieldPicker('fields', { ...nativeProps, value: fields.slice(0, 7).map(o => o.value), initialOptions: fields.slice(0, 7), list: 'Projects' }); break;
      case 'ViewPicker': field = PropertyPaneViewPicker('views', { ...nativeProps, value: views.slice(0, 1).map(o => o.value), initialOptions: views.slice(0, 1), list: 'Projects' }); break;
      case 'ContentTypePicker': field = PropertyPaneContentTypePicker('contentTypes', { ...nativeProps, value: contentTypes.slice(0, 1).map(o => o.value), initialOptions: contentTypes.slice(0, 1) }); break;
      case 'ListItemPicker': field = PropertyPaneListItemPicker('items', { ...nativeProps, value: items.slice(0, 2).map(o => o.value), initialOptions: items.slice(0, 2), list: 'Projects', textField: 'Title' }); break;
      case 'PrincipalPicker': field = PropertyPanePrincipalPicker('principals', { ...nativeProps, value: principalModels.map(o => o.value), initialOptions: principalModels, principalTypes: ['user', 'microsoft365Group'] }); break;
    }
    field.properties.onRender(element, undefined, () => undefined);
    return () => field.properties.onDispose?.(element);
  }, [name, size, native]);
  let control: React.ReactNode;
  switch (name) {
    case 'SitePicker': control = <SitePicker {...common} selectedOptions={[sites[0].key]} initialOptions={sites.slice(0, 1)} mode="hub" />; break;
    case 'ListPicker': control = <ListPicker {...common} selectedOptions={lists.slice(0, 2).map(o => o.key)} initialOptions={lists.slice(0, 2)} />; break;
    case 'FieldPicker': control = <FieldPicker {...common} selectedOptions={fields.slice(0, 7).map(o => o.key)} initialOptions={fields.slice(0, 7)} list="Projects" />; break;
    case 'ViewPicker': control = <ViewPicker {...common} selectedOptions={[views[0].key]} initialOptions={views.slice(0, 1)} list="Projects" />; break;
    case 'ContentTypePicker': control = <ContentTypePicker {...common} selectedOptions={[contentTypes[0].key]} initialOptions={contentTypes.slice(0, 1)} />; break;
    case 'ListItemPicker': control = <ListItemPicker {...common} selectedOptions={items.slice(0, 2).map(o => o.key)} initialOptions={items.slice(0, 2)} list="Projects" textField="Title" />; break;
    case 'PrincipalPicker': control = <PrincipalPicker {...common} selectedOptions={principalModels.map(o => o.key)} initialOptions={principalModels} principalTypes={['user', 'microsoft365Group']} />; break;
  }
  return <section className={styles.cell} data-tag-media-case={id}><h3>{id}</h3>{native ? <div ref={target} /> : control}</section>;
}

/** Browser-only visual verification of actual public React controls and native factories. */
export function TagMediaMatrix() {
  const styles = useStyles();
  return <><h1>Tag media size matrix</h1><p>Actual controls and native fields; all REST, principals and photo pixels are simulated. User 0: initials; User 1: decoded image; User 2: broken-image fallback; Marketing: group icon.</p>
    <main className={styles.matrix}>{[false, true].flatMap(native => names.flatMap(name => sizes.map(size => <MediaCase key={`${native}-${name}-${size}`} name={name} size={size} native={native} />)))}</main></>;
}

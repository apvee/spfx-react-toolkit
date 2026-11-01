/**
 * Native SPFx picker fields with JSON selection ownership and shared pane lifecycle.
 *
 * @packageDocumentation
 * @docPage property-pane/pickers
 */
import type { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import type { SitePickerValue } from '../../controls/pickers/picker-values.types';
import type { SitePickerProps, SitePickerOption } from '../../controls/pickers/picker.types';
import { SitePicker } from '../../controls/pickers/site-picker';
import { validateSitePickerValues } from '../../controls/pickers/picker-values-validation.internal';
import type { PickerPropertyKey, PropertyPanePickerProps } from './property-pane-picker.types';
import { createPropertyPanePicker } from './property-pane-picker-factory.internal';

/**
 * Native Property Pane Site picker, sharing the React control's options, source and presentation.
 * The host owns apply/persistence; arrays contain validated JSON references, including in single selection.
 * Same owner/key/DOM updates preserve root, selection/query state and equivalent request cache.
 * Initial unlabeled selections use selection-control AutoSkeleton during hydration; available labels stay visible. Search/refresh and later updates do not restart the skeleton.
 * The native field supplies its own matching WebPart provider, renderer and pane theme. Do not wrap it in JSX or mount it under a React provider yourself. Configure native resource mappings for the inherited React control. Existing host authentication and delegated permissions still govern queries.
 * Equivalent native renders preserve pending nonreactive drafts. Changing `valueRevision` explicitly adopts `value`; hydration never emits property changes. Invalid values or a missing native change callback disable mutations and display a diagnostic. Latest-factory disposal releases resources; stale disposal cannot unmount its replacement.
 * Inherited option defaults, source restrictions and asynchronous validation/error callbacks follow {@link SitePicker}; native persistence notification remains separate.
 * @useWhen A native SPFx configuration pane must persist SitePickerValue JSON references using the same query and validation options as the React picker.
 * @avoidWhen Rendering a picker in ordinary React content; use the corresponding React control and its selection callback instead.
 * @typeParam TProperties - Web-part property bag with a compatible direct SitePickerValue array.
 * @param targetProperty - Required eligible direct array property name, with no nested path or untyped element domain.
 * @param props - Required initialized web-part instance, stable field key, authoritative value and inherited control options.
 * @returns Native custom field for `groupFields`; the latest SPFx change callback receives targetProperty, a fresh validated array and true after an accepted explicit action. The factory does not save or apply properties.
 * @throws Error - The target property is empty, nested or prototype-related, or the field key is blank.
 * @example Native Site picker
 * ```ts
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
 * import { PropertyPaneSitePicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 * import type { SitePickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
 *
 * interface Properties { values?: SitePickerValue[]; }
 * export default class SelectionWebPart extends BaseClientSideWebPart<Properties> {
 *   public render(): void {
 *     this.domElement.textContent = `${this.properties.values?.length ?? 0} saved references`;
 *   }
 *   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
 *     return { pages: [{ groups: [{ groupName: 'Selection', groupFields: [
 *       PropertyPaneSitePicker('values', {
 *         instance: this, key: 'site-field', value: this.properties.values ?? [],
 *         label: 'Site',
 *       }),
 *     ] }] }] };
 *   }
 * }
 * ```
 *
 *
 * @example Configure six native reference arrays
 * ```tsx
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
 * import type {
 *   SitePickerValue, ListPickerValue, FieldPickerValue, ViewPickerValue,
 *   ListItemPickerValue, PrincipalPickerValue,
 * } from '@apvee/spfx-react-toolkit/controls/pickers';
 * import {
 *   PropertyPaneSitePicker, PropertyPaneListPicker, PropertyPaneFieldPicker,
 *   PropertyPaneViewPicker, PropertyPaneListItemPicker, PropertyPanePrincipalPicker,
 * } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 *
 * interface Properties {
 *   sites?: SitePickerValue[];
 *   lists?: ListPickerValue[];
 *   fields?: FieldPickerValue[];
 *   views?: ViewPickerValue[];
 *   items?: ListItemPickerValue[];
 *   principals?: PrincipalPickerValue[];
 * }
 * export default class ReferenceWebPart extends BaseClientSideWebPart<Properties> {
 *   public render(): void { this.domElement.textContent = 'Configure references in the Property Pane'; }
 *   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
 * return { pages: [{ groups: [{ groupName: 'References', groupFields: [
 *   PropertyPaneSitePicker('sites', { instance: this, key: 'sites-field', value: this.properties.sites ?? [], label: 'Site' }),
 *   PropertyPaneListPicker('lists', { instance: this, key: 'lists-field', value: this.properties.lists ?? [], label: 'List' }),
 *   PropertyPaneFieldPicker('fields', { instance: this, key: 'fields-field', value: this.properties.fields ?? [], list: 'Projects', label: 'Fields' }),
 *   PropertyPaneViewPicker('views', { instance: this, key: 'views-field', value: this.properties.views ?? [], list: 'Projects', label: 'View' }),
 *   PropertyPaneListItemPicker('items', { instance: this, key: 'items-field', value: this.properties.items ?? [], list: 'Projects', textField: 'Title', label: 'Items' }),
 *   PropertyPanePrincipalPicker('principals', { instance: this, key: 'principals-field', value: this.properties.principals ?? [], label: 'Principals',
 *     principalTypes: ['user', 'microsoft365Group'], userIdentity: { mode: 'both', ensureUser: false } }),
 * ] }] }] };
 *   }
 * }
 * ```
 *
 *
 * @example Supply a known native site-reference label
 * ```tsx
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import { PropertyPaneSitePicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 * import type { SitePickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
 * interface Properties { sites?: SitePickerValue[] }
 * // Place the returned field in the initialized host's groupFields. The host still owns persistence.
 * function knownSiteField(instance: BaseClientSideWebPart<Properties>, values: SitePickerValue[]) {
 *   return PropertyPaneSitePicker('sites', {
 *     instance, key: 'sites', label: 'Sito', hint: 'Seleziona una raccolta siti',
 *     value: values.map(site => site.url === 'https://contoso.sharepoint.com/sites/marketing'
 *       ? { ...site, text: 'Marketing' } : site),
 *   });
 * }
 * ```
 */
export function PropertyPaneSitePicker<TProperties extends object>(targetProperty: PickerPropertyKey<TProperties, SitePickerValue>,
  props: PropertyPanePickerProps<TProperties, SitePickerProps, SitePickerValue>): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  const { instance, key, value, valueRevision, ...controlProps } = props;
  return createPropertyPanePicker<TProperties, SitePickerValue, SitePickerOption, SitePickerProps>(targetProperty, { instance, key, value, valueRevision }, SitePicker, validateSitePickerValues, controlProps);
}

import type { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import type { SitePickerValue } from '../../controls/pickers/picker-values.types';
import type { SitePickerProps, SitePickerOption } from '../../controls/pickers/picker.types';
import { SitePicker } from '../../controls/pickers/site-picker';
import { validateSitePickerValues } from '../../controls/pickers/picker-values-validation.internal';
import type { PickerPropertyKey, PropertyPanePickerProps } from './property-pane-picker.types';
import { createPropertyPanePicker } from './property-pane-picker-factory.internal';

/** Native Property Pane Site picker, sharing the React control's options, source and presentation.
 * The host owns apply/persistence; arrays contain validated JSON references, including in single selection.
 * Same owner/key/DOM updates preserve root, selection/query state and equivalent request cache.
 * Initial unlabeled selections use selection-control AutoSkeleton during hydration; available labels stay visible. Search/refresh and later updates do not restart the skeleton.
 * @typeParam TProperties Web-part property bag with a compatible direct SitePickerValue array.
 * @param targetProperty Eligible direct array property name, with no nested path or untyped element domain.
 * @param props Initialized web-part instance, stable field key, authoritative value and control options.
 * @returns Native custom field; its latest change callback receives targetProperty, fresh values and true.
 * @example PropertyPaneSitePicker('values', { instance: this, key: 'site-field', value: this.properties.values, label: 'Site' });
 */
export function PropertyPaneSitePicker<TProperties extends object>(targetProperty: PickerPropertyKey<TProperties, SitePickerValue>,
  props: PropertyPanePickerProps<TProperties, SitePickerProps, SitePickerValue>): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  const { instance, key, value, valueRevision, ...controlProps } = props;
  return createPropertyPanePicker<TProperties, SitePickerValue, SitePickerOption, SitePickerProps>(targetProperty, { instance, key, value, valueRevision }, SitePicker, validateSitePickerValues, controlProps);
}

import type { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import type { ViewPickerValue } from '../../controls/pickers/picker-values.types';
import type { ViewPickerProps, ViewPickerOption } from '../../controls/pickers/picker.types';
import { ViewPicker } from '../../controls/pickers/view-picker';
import { validateViewPickerValues } from '../../controls/pickers/picker-values-validation.internal';
import type { PickerPropertyKey, PropertyPanePickerProps } from './property-pane-picker.types';
import { createPropertyPanePicker } from './property-pane-picker-factory.internal';

/** Native Property Pane View picker, sharing the React control's options, source and presentation.
 * The host owns apply/persistence; arrays contain validated JSON references, including in single selection.
 * Same owner/key/DOM updates preserve root, selection/query state and equivalent request cache.
 * Initial unlabeled selections use selection-control AutoSkeleton during hydration; available labels stay visible. Search/refresh and later updates do not restart the skeleton.
 * @typeParam TProperties Web-part property bag with a compatible direct ViewPickerValue array.
 * @param targetProperty Eligible direct array property name, with no nested path or untyped element domain.
 * @param props Initialized web-part instance, stable field key, authoritative value and control options.
 * @returns Native custom field; its latest change callback receives targetProperty, fresh values and true.
 * @example PropertyPaneViewPicker('values', { instance: this, key: 'view-field', value: this.properties.values, label: 'View', list: 'Projects' });
 */
export function PropertyPaneViewPicker<TProperties extends object>(targetProperty: PickerPropertyKey<TProperties, ViewPickerValue>,
  props: PropertyPanePickerProps<TProperties, ViewPickerProps, ViewPickerValue>): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  const { instance, key, value, valueRevision, ...controlProps } = props;
  return createPropertyPanePicker<TProperties, ViewPickerValue, ViewPickerOption, ViewPickerProps>(targetProperty, { instance, key, value, valueRevision }, ViewPicker, validateViewPickerValues, controlProps);
}

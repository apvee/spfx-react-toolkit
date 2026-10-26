import type { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import type { ListPickerValue } from '../../controls/pickers/picker-values.types';
import type { ListPickerProps, ListPickerOption } from '../../controls/pickers/picker.types';
import { ListPicker } from '../../controls/pickers/list-picker';
import { validateListPickerValues } from '../../controls/pickers/picker-values-validation.internal';
import type { PickerPropertyKey, PropertyPanePickerProps } from './property-pane-picker.types';
import { createPropertyPanePicker } from './property-pane-picker-factory.internal';

/** Native Property Pane List picker, sharing the React control's options, source and presentation.
 * The host owns apply/persistence; arrays contain validated JSON references, including in single selection.
 * Same owner/key/DOM updates preserve root, selection/query state and equivalent request cache.
 * Initial unlabeled selections use selection-control AutoSkeleton during hydration; available labels stay visible. Search/refresh and later updates do not restart the skeleton.
 * @typeParam TProperties Web-part property bag with a compatible direct ListPickerValue array.
 * @param targetProperty Eligible direct array property name, with no nested path or untyped element domain.
 * @param props Initialized web-part instance, stable field key, authoritative value and control options.
 * @returns Native custom field; its latest change callback receives targetProperty, fresh values and true.
 * @example PropertyPaneListPicker('values', { instance: this, key: 'list-field', value: this.properties.values, label: 'List' });
 */
export function PropertyPaneListPicker<TProperties extends object>(targetProperty: PickerPropertyKey<TProperties, ListPickerValue>,
  props: PropertyPanePickerProps<TProperties, ListPickerProps, ListPickerValue>): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  const { instance, key, value, valueRevision, ...controlProps } = props;
  return createPropertyPanePicker<TProperties, ListPickerValue, ListPickerOption, ListPickerProps>(targetProperty, { instance, key, value, valueRevision }, ListPicker, validateListPickerValues, controlProps);
}

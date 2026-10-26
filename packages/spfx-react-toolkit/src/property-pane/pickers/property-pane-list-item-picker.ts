import type { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import type { ListItemPickerValue } from '../../controls/pickers/picker-values.types';
import type { ListItemPickerProps, ListItemPickerOption } from '../../controls/pickers/picker.types';
import { ListItemPicker } from '../../controls/pickers/list-item-picker';
import { validateListItemPickerValues } from '../../controls/pickers/picker-values-validation.internal';
import type { PickerPropertyKey, PropertyPanePickerProps } from './property-pane-picker.types';
import { createPropertyPanePicker } from './property-pane-picker-factory.internal';

/** Native Property Pane ListItem picker, sharing the React control's options, source and presentation.
 * The host owns apply/persistence; arrays contain validated JSON references, including in single selection.
 * Same owner/key/DOM updates preserve root, selection/query state and equivalent request cache.
 * Initial unlabeled selections use selection-control AutoSkeleton during hydration; available labels stay visible. Search/refresh and later updates do not restart the skeleton.
 * @typeParam TProperties Web-part property bag with a compatible direct ListItemPickerValue array.
 * @param targetProperty Eligible direct array property name, with no nested path or untyped element domain.
 * @param props Initialized web-part instance, stable field key, authoritative value and control options.
 * @returns Native custom field; its latest change callback receives targetProperty, fresh values and true.
 * @example PropertyPaneListItemPicker('values', { instance: this, key: 'list-item-field', value: this.properties.values, label: 'ListItem', list: 'Projects', textField: 'Title' });
 */
export function PropertyPaneListItemPicker<TProperties extends object>(targetProperty: PickerPropertyKey<TProperties, ListItemPickerValue>,
  props: PropertyPanePickerProps<TProperties, ListItemPickerProps, ListItemPickerValue>): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  const { instance, key, value, valueRevision, ...controlProps } = props;
  return createPropertyPanePicker<TProperties, ListItemPickerValue, ListItemPickerOption, ListItemPickerProps>(targetProperty, { instance, key, value, valueRevision }, ListItemPicker, validateListItemPickerValues, controlProps);
}

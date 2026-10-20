import type { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import type { FieldPickerValue } from '../../controls/pickers/picker-values.types';
import type { FieldPickerProps, FieldPickerOption } from '../../controls/pickers/picker.types';
import { FieldPicker } from '../../controls/pickers/field-picker';
import { validateFieldPickerValues } from '../../controls/pickers/picker-values-validation.internal';
import type { PickerPropertyKey, PropertyPanePickerProps } from './property-pane-picker.types';
import { createPropertyPanePicker } from './property-pane-picker-factory.internal';

/** Native Property Pane Field picker, sharing the React control's options, source and presentation.
 * The host owns apply/persistence; arrays contain validated JSON references, including in single selection.
 * Same owner/key/DOM updates preserve root, selection/query state and equivalent request cache.
 * Initial unlabeled selections use input-only AutoSkeleton during hydration; available labels stay visible. Search/refresh and later updates do not restart the skeleton.
 * @typeParam TProperties Web-part property bag with a compatible direct FieldPickerValue array.
 * @param targetProperty Eligible direct array property name, with no nested path or untyped element domain.
 * @param props Initialized web-part instance, stable field key, authoritative value and control options.
 * @returns Native custom field; its latest change callback receives targetProperty, fresh values and true.
 * @example PropertyPaneFieldPicker('values', { instance: this, key: 'field-field', value: this.properties.values, label: 'Field' });
 */
export function PropertyPaneFieldPicker<TProperties extends object>(targetProperty: PickerPropertyKey<TProperties, FieldPickerValue>,
  props: PropertyPanePickerProps<TProperties, FieldPickerProps, FieldPickerValue>): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  const { instance, key, value, valueRevision, ...controlProps } = props;
  return createPropertyPanePicker<TProperties, FieldPickerValue, FieldPickerOption, FieldPickerProps>(targetProperty, { instance, key, value, valueRevision }, FieldPicker, validateFieldPickerValues, controlProps);
}

import type { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import type { PrincipalPickerValue } from '../../controls/pickers/picker-values.types';
import type { PrincipalPickerProps, PrincipalPickerOption } from '../../controls/pickers/picker.types';
import { PrincipalPicker } from '../../controls/pickers/principal-picker';
import { validatePrincipalPickerValues } from '../../controls/pickers/picker-values-validation.internal';
import type { PickerPropertyKey, PropertyPanePickerProps } from './property-pane-picker.types';
import { createPropertyPanePicker } from './property-pane-picker-factory.internal';

/** Native Property Pane Principal picker, sharing the React control's options, source and presentation.
 * The host owns apply/persistence; arrays contain validated JSON references, including in single selection.
 * Same owner/key/DOM updates preserve root, selection/query state and equivalent request cache.
 * Initial unlabeled selections use input-only AutoSkeleton during hydration; available labels stay visible. Search/refresh and later updates do not restart the skeleton.
 * @typeParam TProperties Web-part property bag with a compatible direct PrincipalPickerValue array.
 * @param targetProperty Eligible direct array property name, with no nested path or untyped element domain.
 * @param props Initialized web-part instance, stable field key, authoritative value and control options.
 * @returns Native custom field; its latest change callback receives targetProperty, fresh values and true.
 * @example PropertyPanePrincipalPicker('values', { instance: this, key: 'principal-field', value: this.properties.values, label: 'Principal' });
 */
export function PropertyPanePrincipalPicker<TProperties extends object>(targetProperty: PickerPropertyKey<TProperties, PrincipalPickerValue>,
  props: PropertyPanePickerProps<TProperties, PrincipalPickerProps, PrincipalPickerValue>): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  const { instance, key, value, valueRevision, ...controlProps } = props;
  return createPropertyPanePicker<TProperties, PrincipalPickerValue, PrincipalPickerOption, PrincipalPickerProps>(targetProperty, { instance, key, value, valueRevision }, PrincipalPicker, validatePrincipalPickerValues, controlProps);
}

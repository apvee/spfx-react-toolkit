import type { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import type { ContentTypePickerValue } from '../../controls/pickers/picker-values.types';
import type { ContentTypePickerProps, ContentTypePickerOption } from '../../controls/pickers/picker.types';
import { ContentTypePicker } from '../../controls/pickers/content-type-picker';
import { validateContentTypePickerValues } from '../../controls/pickers/picker-values-validation.internal';
import type { PickerPropertyKey, PropertyPanePickerProps } from './property-pane-picker.types';
import { createPropertyPanePicker } from './property-pane-picker-factory.internal';

/** Native Property Pane content type picker, sharing the React control's options, source and presentation.
 * Omitted scope reads the owning site collection root; list scope requires a toolkit list selector.
 * Forwards the React control's filters, local complete-catalog search and defaults. Selection never provisions a content type or enables list content types.
 * The host owns apply/persistence; arrays contain validated JSON references, including in single selection.
 * Same owner/key/DOM updates preserve root, selection/query state and equivalent request cache.
 * Initial unlabeled selections use selection-control AutoSkeleton during hydration; available labels stay visible. Search/refresh and later updates do not restart the skeleton.
 * @typeParam TProperties Web-part property bag with a compatible direct ContentTypePickerValue array.
 * @param targetProperty Eligible direct array property name, with no nested path or untyped element domain.
 * @param props Initialized web-part instance, stable field key, authoritative value and control options.
 * @returns Native custom field; its latest change callback receives targetProperty, fresh values and true.
 * @example PropertyPaneContentTypePicker('values', { instance: this, key: 'content-type-field', value: this.properties.values, label: 'Content type', scope: 'list', list: 'Projects' });
 */
export function PropertyPaneContentTypePicker<TProperties extends object>(targetProperty: PickerPropertyKey<TProperties, ContentTypePickerValue>,
  props: PropertyPanePickerProps<TProperties, ContentTypePickerProps, ContentTypePickerValue>): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  const { instance, key, value, valueRevision, ...controlProps } = props;
  return createPropertyPanePicker<TProperties, ContentTypePickerValue, ContentTypePickerOption, ContentTypePickerProps>(targetProperty, { instance, key, value, valueRevision }, ContentTypePicker, validateContentTypePickerValues, controlProps);
}

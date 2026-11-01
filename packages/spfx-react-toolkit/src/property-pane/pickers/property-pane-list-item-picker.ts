/**
 * Native SPFx picker fields with JSON selection ownership and shared pane lifecycle.
 *
 * @packageDocumentation
 * @docPage property-pane/pickers
 */
import type { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import type { ListItemPickerValue } from '../../controls/pickers/picker-values.types';
import type { ListItemPickerProps, ListItemPickerOption } from '../../controls/pickers/picker.types';
import { ListItemPicker } from '../../controls/pickers/list-item-picker';
import { validateListItemPickerValues } from '../../controls/pickers/picker-values-validation.internal';
import type { PickerPropertyKey, PropertyPanePickerProps } from './property-pane-picker.types';
import { createPropertyPanePicker } from './property-pane-picker-factory.internal';

/**
 * Native Property Pane ListItem picker, sharing the React control's options, source and presentation.
 * The host owns apply/persistence; arrays contain validated JSON references, including in single selection.
 * Same owner/key/DOM updates preserve root, selection/query state and equivalent request cache.
 * Initial unlabeled selections use selection-control AutoSkeleton during hydration; available labels stay visible. Search/refresh and later updates do not restart the skeleton.
 * The native field supplies its own matching WebPart provider, renderer and pane theme. Do not wrap it in JSX or mount it under a React provider yourself. Configure native resource mappings for the inherited React control. Existing host authentication and delegated permissions still govern queries.
 * Equivalent native renders preserve pending nonreactive drafts. Changing `valueRevision` explicitly adopts `value`; hydration never emits property changes. Invalid values or a missing native change callback disable mutations and display a diagnostic. Latest-factory disposal releases resources; stale disposal cannot unmount its replacement.
 * Inherited option defaults, source restrictions and asynchronous validation/error callbacks follow {@link ListItemPicker}; native persistence notification remains separate.
 * @useWhen A native SPFx configuration pane must persist ListItemPickerValue JSON references using the same query and validation options as the React picker.
 * @avoidWhen Rendering a picker in ordinary React content; use the corresponding React control and its selection callback instead.
 * @typeParam TProperties - Web-part property bag with a compatible direct ListItemPickerValue array.
 * @param targetProperty - Required eligible direct array property name, with no nested path or untyped element domain.
 * @param props - Required initialized web-part instance, stable field key, authoritative value and inherited control options.
 * @returns Native custom field for `groupFields`; the latest SPFx change callback receives targetProperty, a fresh validated array and true after an accepted explicit action. The factory does not save or apply properties.
 * @throws Error - The target property is empty, nested or prototype-related, or the field key is blank.
 * @example Native ListItem picker
 * ```ts
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
 * import { PropertyPaneListItemPicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 * import type { ListItemPickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
 *
 * interface Properties { values?: ListItemPickerValue[]; }
 * export default class SelectionWebPart extends BaseClientSideWebPart<Properties> {
 *   public render(): void {
 *     this.domElement.textContent = `${this.properties.values?.length ?? 0} saved references`;
 *   }
 *   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
 *     return { pages: [{ groups: [{ groupName: 'Selection', groupFields: [
 *       PropertyPaneListItemPicker('values', {
 *         instance: this, key: 'list-item-field', value: this.properties.values ?? [],
 *         label: 'Items', list: 'Projects', textField: 'Title',
 *       }),
 *     ] }] }] };
 *   }
 * }
 * ```
 */
export function PropertyPaneListItemPicker<TProperties extends object>(targetProperty: PickerPropertyKey<TProperties, ListItemPickerValue>,
  props: PropertyPanePickerProps<TProperties, ListItemPickerProps, ListItemPickerValue>): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  const { instance, key, value, valueRevision, ...controlProps } = props;
  return createPropertyPanePicker<TProperties, ListItemPickerValue, ListItemPickerOption, ListItemPickerProps>(targetProperty, { instance, key, value, valueRevision }, ListItemPicker, validateListItemPickerValues, controlProps);
}

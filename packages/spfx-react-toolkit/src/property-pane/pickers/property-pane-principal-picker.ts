/**
 * Native SPFx picker fields with JSON selection ownership and shared pane lifecycle.
 *
 * @packageDocumentation
 * @docPage property-pane/pickers
 */
import type { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import type { PrincipalPickerValue } from '../../controls/pickers/picker-values.types';
import type { PrincipalPickerProps, PrincipalPickerOption } from '../../controls/pickers/picker.types';
import { PrincipalPicker } from '../../controls/pickers/principal-picker';
import { validatePrincipalPickerValues } from '../../controls/pickers/picker-values-validation.internal';
import type { PickerPropertyKey, PropertyPanePickerProps } from './property-pane-picker.types';
import { createPropertyPanePicker } from './property-pane-picker-factory.internal';

/**
 * Native Property Pane Principal picker, sharing the React control's options, source and presentation.
 * The host owns apply/persistence; arrays contain validated JSON references, including in single selection.
 * Same owner/key/DOM updates preserve root, selection/query state and equivalent request cache.
 * Initial unlabeled selections use selection-control AutoSkeleton during hydration; available labels stay visible. Search/refresh and later updates do not restart the skeleton.
 * The native field supplies its own matching WebPart provider, renderer and pane theme. Do not wrap it in JSX or mount it under a React provider yourself. Configure native resource mappings for the inherited React control. Existing host authentication and delegated permissions still govern queries.
 * Equivalent native renders preserve pending nonreactive drafts. Changing `valueRevision` explicitly adopts `value`; hydration never emits property changes. Invalid values or a missing native change callback disable mutations and display a diagnostic. Latest-factory disposal releases resources; stale disposal cannot unmount its replacement.
 * Inherited option defaults, source restrictions and asynchronous validation/error callbacks follow {@link PrincipalPicker}; native persistence notification remains separate.
 * @useWhen A native SPFx configuration pane must persist PrincipalPickerValue JSON references using the same query and validation options as the React picker.
 * @avoidWhen Rendering a picker in ordinary React content; use the corresponding React control and its selection callback instead.
 * @typeParam TProperties - Web-part property bag with a compatible direct PrincipalPickerValue array.
 * @param targetProperty - Required eligible direct array property name, with no nested path or untyped element domain.
 * @param props - Required initialized web-part instance, stable field key, authoritative value and inherited control options.
 * @returns Native custom field for `groupFields`; the latest SPFx change callback receives targetProperty, a fresh validated array and true after an accepted explicit action. The factory does not save or apply properties.
 * @throws Error - The target property is empty, nested or prototype-related, or the field key is blank.
 * @example Native Principal picker
 * ```ts
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
 * import { PropertyPanePrincipalPicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 * import type { PrincipalPickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
 *
 * interface Properties { values?: PrincipalPickerValue[]; }
 * export default class SelectionWebPart extends BaseClientSideWebPart<Properties> {
 *   public render(): void {
 *     this.domElement.textContent = `${this.properties.values?.length ?? 0} saved references`;
 *   }
 *   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
 *     return { pages: [{ groups: [{ groupName: 'Selection', groupFields: [
 *       PropertyPanePrincipalPicker('values', {
 *         instance: this, key: 'principal-field', value: this.properties.values ?? [],
 *         label: 'People',
 *       }),
 *     ] }] }] };
 *   }
 * }
 * ```
 */
export function PropertyPanePrincipalPicker<TProperties extends object>(targetProperty: PickerPropertyKey<TProperties, PrincipalPickerValue>,
  props: PropertyPanePickerProps<TProperties, PrincipalPickerProps, PrincipalPickerValue>): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  const { instance, key, value, valueRevision, ...controlProps } = props;
  return createPropertyPanePicker<TProperties, PrincipalPickerValue, PrincipalPickerOption, PrincipalPickerProps>(targetProperty, { instance, key, value, valueRevision }, PrincipalPicker, validatePrincipalPickerValues, controlProps);
}

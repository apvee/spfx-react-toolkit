/**
 * Native SPFx picker fields with JSON selection ownership and shared pane lifecycle.
 *
 * @packageDocumentation
 * @docPage property-pane/pickers
 */
import type { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import type { FieldPickerValue } from '../../controls/pickers/picker-values.types';
import type { FieldPickerProps, FieldPickerOption } from '../../controls/pickers/picker.types';
import { FieldPicker } from '../../controls/pickers/field-picker';
import { validateFieldPickerValues } from '../../controls/pickers/picker-values-validation.internal';
import type { PickerPropertyKey, PropertyPanePickerProps } from './property-pane-picker.types';
import { createPropertyPanePicker } from './property-pane-picker-factory.internal';

/**
 * Native Property Pane Field picker, sharing the React control's options, source and presentation.
 * The host owns apply/persistence; arrays contain validated JSON references, including in single selection.
 * Same owner/key/DOM updates preserve root, selection/query state and equivalent request cache.
 * Initial unlabeled selections use selection-control AutoSkeleton during hydration; available labels stay visible. Search/refresh and later updates do not restart the skeleton.
 * The native field supplies its own matching WebPart provider, renderer and pane theme. Do not wrap it in JSX or mount it under a React provider yourself. Configure native resource mappings for the inherited React control. Existing host authentication and delegated permissions still govern queries.
 * Equivalent native renders preserve pending nonreactive drafts. Changing `valueRevision` explicitly adopts `value`; hydration never emits property changes. Invalid values or a missing native change callback disable mutations and display a diagnostic. Latest-factory disposal releases resources; stale disposal cannot unmount its replacement.
 * Inherited option defaults, source restrictions and asynchronous validation/error callbacks follow {@link FieldPicker}; native persistence notification remains separate.
 * @useWhen A native SPFx configuration pane must persist FieldPickerValue JSON references using the same query and validation options as the React picker.
 * @avoidWhen Rendering a picker in ordinary React content; use the corresponding React control and its selection callback instead.
 * @typeParam TProperties - Web-part property bag with a compatible direct FieldPickerValue array.
 * @param targetProperty - Required eligible direct array property name, with no nested path or untyped element domain.
 * @param props - Required initialized web-part instance, stable field key, authoritative value and inherited control options.
 * @returns Native custom field for `groupFields`; the latest SPFx change callback receives targetProperty, a fresh validated array and true after an accepted explicit action. The factory does not save or apply properties.
 * @throws Error - The target property is empty, nested or prototype-related, or the field key is blank.
 * @example Native Field picker
 * ```ts
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
 * import { PropertyPaneFieldPicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 * import type { FieldPickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
 *
 * interface Properties { values?: FieldPickerValue[]; }
 * export default class SelectionWebPart extends BaseClientSideWebPart<Properties> {
 *   public render(): void {
 *     this.domElement.textContent = `${this.properties.values?.length ?? 0} saved references`;
 *   }
 *   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
 *     return { pages: [{ groups: [{ groupName: 'Selection', groupFields: [
 *       PropertyPaneFieldPicker('values', {
 *         instance: this, key: 'field-field', value: this.properties.values ?? [],
 *         label: 'Fields', list: 'Projects',
 *       }),
 *     ] }] }] };
 *   }
 * }
 * ```
 *
 * @example Select native field references
 * ```tsx
 * import * as React from 'react';
 * import { Button, Input } from '@fluentui/react-components';
 * import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
 * import { PropertyPaneFieldPicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 * import type { FieldPickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
 *
 * interface SourceProperties { pickerFields?: FieldPickerValue[]; }
 * interface NativePropertyPaneEnvironment<P extends object> {
 *   readonly instance: BaseClientSideWebPart<P>;
 *   readonly readProperties: () => P;
 *   readonly refresh: () => void;
 * }
 * interface NativePropertyPaneController {
 *   readonly getConfiguration: () => IPropertyPaneConfiguration;
 *   readonly fieldChanged: (propertyPath: string, oldValue: unknown, newValue: unknown) => void;
 *   readonly configurationComplete: () => void;
 *   readonly changesApplied: () => void;
 *   readonly dispose: () => void;
 * }
 * interface NativePropertyPanePreparation { readonly release: () => void; }
 * type NativePropertyPanePreparationResult =
 *   | { readonly status: 'prepared'; readonly lease: NativePropertyPanePreparation }
 *   | { readonly status: 'blocked'; readonly reason: string };
 * interface DemoPropertyPanePort<P extends object> {
 *   readonly prepare: (create: (host: NativePropertyPaneEnvironment<P>) => NativePropertyPaneController) => NativePropertyPanePreparationResult;
 * }
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 *   children?: React.ReactNode;
 * }
 * function createFieldController(host: NativePropertyPaneEnvironment<SourceProperties>, list: string): NativePropertyPaneController {
 *   return {
 *     getConfiguration: () => ({ pages: [{
 *       header: { description: 'Select saved field references, then use native Close. Reset is not native Discard.' },
 *       groups: [{ isGroupNameHidden: true, groupFields: [
 *         PropertyPaneFieldPicker('pickerFields', {
 *           instance: host.instance, key: 'picker-fields', value: host.readProperties().pickerFields ?? [], label: 'Fields', list,
 *         }),
 *       ] }],
 *     }] }),
 *     fieldChanged: () => {},
 *     configurationComplete: () => {},
 *     changesApplied: () => {},
 *     dispose: () => {},
 *   };
 * }
 *
 * // The initialized host owns actual native Edit/Close and SDK persistence.
 * // This structural port is private example tooling. Configure the ordinary native
 * // recipe's localization resources; existing host authentication and permissions
 * // govern queries. The factory installs its own WebPart provider and renderer.
 * // This example deliberately chooses list fields. The Field API also permits
 * // omitted list for site columns; the prerequisite here is not an API requirement.
 * export function Demo({ render, resetRevision = 0, propertyPane }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 *   propertyPane?: DemoPropertyPanePort<SourceProperties>;
 * }): React.ReactElement {
 *   const [status, setStatus] = React.useState('Prepare before opening this WebPart’s actual native property pane.');
 *   const [listTitle, setListTitle] = React.useState('');
 *   const [saved, setSaved] = React.useState<string>();
 *   const lease = React.useRef<NativePropertyPanePreparation>();
 *   const readProperties = React.useRef<(() => SourceProperties)>();
 *   const lifetime = React.useMemo(() => ({ active: true }), [propertyPane]);
 *   const previousRevision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (previousRevision.current === resetRevision) return;
 *     previousRevision.current = resetRevision;
 *     lease.current?.release();
 *     lease.current = undefined;
 *     readProperties.current = undefined;
 *     setListTitle('');
 *     setSaved(undefined);
 *     setStatus('Local state reset. A produced native session survives until its real native lifecycle ends; Reset is not Discard or rollback.');
 *   }, [resetRevision]);
 *   React.useEffect(() => () => {
 *     lifetime.active = false;
 *     lease.current?.release();
 *     lease.current = undefined;
 *     readProperties.current = undefined;
 *   }, [lifetime]);
 *   const prepare = (): void => {
 *     if (!propertyPane || !lifetime.active || resetRevision !== previousRevision.current) return;
 *     const list = listTitle.trim();
 *     if (!list) { setStatus('Enter a real exact list title before preparing.'); return; }
 *     // Freeze this chosen configuration; later draft edits do not alter the native session.
 *     lease.current?.release();
 *     const result = propertyPane.prepare(host => {
 *       readProperties.current = host.readProperties;
 *       return createFieldController(host, list);
 *     });
 *     if (result.status === 'blocked') {
 *       lease.current = undefined;
 *       setStatus(result.reason);
 *       return;
 *     }
 *     lease.current = result.lease;
 *     setStatus('Prepared. Open this WebPart’s actual native Edit property pane, select Fields, then use native Close. Preparing does not open the pane or query.');
 *   };
 *   const observe = (): void => {
 *     if (!lifetime.active || resetRevision !== previousRevision.current) return;
 *     const read = readProperties.current;
 *     if (!read) { setStatus('Open the prepared native scenario before inspecting its saved values.'); return; }
 *     setSaved(JSON.stringify(read().pickerFields ?? [], undefined, 2));
 *   };
 *   const regions: DemoRegions = propertyPane ? {
 *     controls: <>
 *       <label>Exact list title<Input aria-label="Exact list title" value={listTitle} onChange={event => setListTitle(event.target.value)} /></label>
 *       <p>Enter an existing exact list title in the initialized current web; even a GUID-looking string is a title. Draft collection and preparation perform no queries.</p>
 *     </>,
 *     commands: <><Button disabled={!listTitle.trim()} onClick={prepare}>Prepare native scenario</Button><Button onClick={observe}>Inspect saved field references</Button></>,
 *     status: <p role="status">{status}</p>,
 *     children: <><p>Use the actual native Edit and Close controls. Selection, preparation and Reset do not open or close the pane. Reset cancels only an unproduced preparation and local observations; it cannot discard accepted native changes.</p><p>If the SDK, React or host fails after native configuration was returned, reload before further native scenario work; automatic recovery is unsupported.</p></>,
 *     diagnostics: <pre>{saved ?? 'Use Inspect saved field references after an actual native selection.'}</pre>,
 *   } : {
 *     status: <p role="status">The initialized native scenario host is not connected. This selected native scenario is not active.</p>,
 *     children: <p>This WebPart’s actual native Edit still opens the existing full integrated configuration independently of the selected entry until the final native cutover. There is no enabled Prepare action here. Selection and Reset do not query or open the pane.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', hostAdapter: 'propertyPane', observe: "Prepare the native Field picker, open this WebPart's actual Edit property pane, select field references and use native Close and inspect the saved JSON; selection and Reset do not open the pane or query." };
 * ```
 */
export function PropertyPaneFieldPicker<TProperties extends object>(targetProperty: PickerPropertyKey<TProperties, FieldPickerValue>,
  props: PropertyPanePickerProps<TProperties, FieldPickerProps, FieldPickerValue>): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  const { instance, key, value, valueRevision, ...controlProps } = props;
  return createPropertyPanePicker<TProperties, FieldPickerValue, FieldPickerOption, FieldPickerProps>(targetProperty, { instance, key, value, valueRevision }, FieldPicker, validateFieldPickerValues, controlProps);
}

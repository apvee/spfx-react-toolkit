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
 *
 * @example Select native principal references
 * ```tsx
 * import * as React from 'react';
 * import { Button } from '@fluentui/react-components';
 * import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
 * import { PropertyPanePrincipalPicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 * import type { PrincipalPickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
 *
 * interface SourceProperties { pickerPrincipals?: PrincipalPickerValue[]; }
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
 * function createPrincipalController(host: NativePropertyPaneEnvironment<SourceProperties>): NativePropertyPaneController {
 *   return {
 *     getConfiguration: () => ({ pages: [{
 *       header: { description: 'Select saved principal references, then use native Close. Reset is not native Discard.' },
 *       groups: [{ isGroupNameHidden: true, groupFields: [
 *         PropertyPanePrincipalPicker('pickerPrincipals', {
 *           instance: host.instance, key: 'picker-principals', value: host.readProperties().pickerPrincipals ?? [], label: 'Principals',
 *           principalTypes: ['user'], userIdentity: { mode: 'available' },
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
 * // Search users explicitly; no personal/current suggestions. Available mode
 * // has no ensureUser option and never provisions a SharePoint user.
 * // The saved array retains the complete PrincipalPickerValue domain.
 * export function Demo({ render, resetRevision = 0, propertyPane }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 *   propertyPane?: DemoPropertyPanePort<SourceProperties>;
 * }): React.ReactElement {
 *   const [status, setStatus] = React.useState('Prepare before opening this WebPart’s actual native property pane.');
 *   const [saved, setSaved] = React.useState<string>();
 *   const lease = React.useRef<NativePropertyPanePreparation>();
 *   const readProperties = React.useRef<(() => SourceProperties)>();
 *   const lifetime = React.useMemo(() => ({ active: true }), [propertyPane, resetRevision]);
 *   const previousRevision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (previousRevision.current === resetRevision) return;
 *     previousRevision.current = resetRevision;
 *     lease.current?.release();
 *     lease.current = undefined;
 *     readProperties.current = undefined;
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
 *     lease.current?.release();
 *     const result = propertyPane.prepare(host => {
 *       readProperties.current = host.readProperties;
 *       return createPrincipalController(host);
 *     });
 *     if (result.status === 'blocked') {
 *       lease.current = undefined;
 *       setStatus(result.reason);
 *       return;
 *     }
 *     lease.current = result.lease;
 *     setStatus('Prepared. Open this WebPart’s actual native Edit property pane, select Principals, then use native Close. Preparing does not open the pane or query.');
 *   };
 *   const observe = (): void => {
 *     if (!lifetime.active || resetRevision !== previousRevision.current) return;
 *     const read = readProperties.current;
 *     if (!read) { setStatus('Open the prepared native scenario before inspecting its saved values.'); return; }
 *     setSaved(JSON.stringify(read().pickerPrincipals ?? [], undefined, 2));
 *   };
 *   const regions: DemoRegions = propertyPane ? {
 *     commands: <><Button onClick={prepare}>Prepare native scenario</Button><Button onClick={observe}>Inspect saved principal references</Button></>,
 *     status: <p role="status">{status}</p>,
 *     children: <><p>Use the actual native Edit and Close controls. Selection, preparation and Reset do not open or close the pane. Reset cancels only an unproduced preparation and local observations; it cannot discard accepted native changes.</p><p>If the SDK, React or host fails after native configuration was returned, reload before further native scenario work; automatic recovery is unsupported.</p></>,
 *     diagnostics: <pre>{saved ?? 'Use Inspect saved principal references after an actual native selection.'}</pre>,
 *   } : {
 *     status: <p role="status">The initialized native scenario host is not connected. This selected native scenario is not active.</p>,
 *     children: <p>Connect an initialized native host to enable Prepare, then use this WebPart’s actual native Edit and Close. Selection and Reset do not query or open the pane.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', hostAdapter: 'propertyPane', observe: "Prepare the native Principal picker, open this WebPart's actual Edit property pane, select principal references and use native Close and inspect the saved JSON; selection and Reset do not open the pane or query." };
 * ```
 */
export function PropertyPanePrincipalPicker<TProperties extends object>(targetProperty: PickerPropertyKey<TProperties, PrincipalPickerValue>,
  props: PropertyPanePickerProps<TProperties, PrincipalPickerProps, PrincipalPickerValue>): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  const { instance, key, value, valueRevision, ...controlProps } = props;
  return createPropertyPanePicker<TProperties, PrincipalPickerValue, PrincipalPickerOption, PrincipalPickerProps>(targetProperty, { instance, key, value, valueRevision }, PrincipalPicker, validatePrincipalPickerValues, controlProps);
}

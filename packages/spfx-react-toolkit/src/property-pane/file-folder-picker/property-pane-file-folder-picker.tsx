/**
 * Native SPFx file/folder fields with hosted selection and caller-owned JSON persistence.
 *
 * @packageDocumentation
 * @docPage property-pane/file-folder-picker
 */
import * as React from 'react';
import { PropertyPaneFieldType } from '@microsoft/sp-property-pane';
import type { IPropertyPaneCustomFieldProps, IPropertyPaneField } from '@microsoft/sp-property-pane';
import { SPFxWebPartProvider } from '../../core/provider-webpart';
import { PropertyPanePickerHost } from '../pickers/property-pane-picker-host.internal';
import { disposePickerField, mountOrUpdatePickerField } from '../pickers/property-pane-picker.internal';
import { PropertyPaneFileFolderPickerValueBridge } from './property-pane-file-folder-picker-value.internal';
import type { FileFolderPickerPropertyKey, PropertyPaneFileFolderPickerProps } from './property-pane-file-folder-picker.types';

/** Native SPFx property-pane field for existing files/folders through the hosted Microsoft File Picker v8.
 * Requires an initialized web part and existing delegated permissions. React FileFolderPicker options/defaults apply,
 * while SPFx owns controlled JSON arrays and apply/persistence. Every accepted action emits targetProperty, fresh
 * validated FileFolderPickerValue[] and true; hydration, display-model changes and authoritative resets never emit.
 * Equivalent renders retain provider/root/cache and nonreactive draft. valueRevision resets the draft and cancels
 * pending confirmation. Without a native change callback the field displays saved values with mutation disabled.
 * Initial unlabeled selections use the existing input-only AutoSkeleton boundary; known labels remain visible.
 * Latest-factory disposal closes owned picker resources and releases its metadata lease; stale disposal is ignored.
 * @typeParam TProperties - Web-part property bag containing an exact file/folder reference array.
 * @param targetProperty - Required safe direct array property key; nested paths and prototype names are rejected at runtime too.
 * @param props - Required initialized instance, stable field key, required JSON array, optional revision and React control options.
 * @returns Native custom field using the owning document's renderer, native pane token theme and SPFxWebPartProvider.
 * @throws Invalid targetProperty or blank field key.
 * @useWhen A native SPFx pane must persist existing drive-item references through the hosted file/folder picker while preserving native Apply ownership.
 * @avoidWhen Rendering ordinary React content; use FileFolderPicker and its onChange contract instead.
 * @example Native file and folder configuration
 * ```ts
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
 * import { PropertyPaneFileFolderPicker } from '@apvee/spfx-react-toolkit/property-pane/file-folder-picker';
 * import type { FileFolderPickerValue } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';
 *
 * interface Properties {
 *   documents?: FileFolderPickerValue[];
 * }
 * export default class DocumentsWebPart extends BaseClientSideWebPart<Properties> {
 *   public render(): void {
 *     this.domElement.textContent = `${this.properties.documents?.length ?? 0} references`;
 *   }
 *   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
 *     return {
 *       pages: [
 *         {
 *           groups: [
 *             {
 *               groupName: 'Documents',
 *               groupFields: [
 *                 PropertyPaneFileFolderPicker('documents', {
 *                   instance: this,
 *                   key: 'documents-picker',
 *                   value: this.properties.documents ?? [],
 *                   label: 'Documents',
 *                   preset: 'focused',
 *                   multiselect: true,
 *                   scope: { kind: 'siteCollection' },
 *                   itemTypes: 'files',
 *                   extensions: ['.pdf']
 *                 })
 *               ]
 *             }
 *           ]
 *         }
 *       ]
 *     };
 *   }
 * }
 * ```
 *
 * @scenario Select native file and folder references
 * ```tsx
 * import * as React from 'react';
 * import { Button } from '@fluentui/react-components';
 * import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
 * import { PropertyPaneFileFolderPicker } from '@apvee/spfx-react-toolkit/property-pane/file-folder-picker';
 * import type { FileFolderPickerValue } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';
 *
 * interface SourceProperties { fileFolderValues?: FileFolderPickerValue[]; }
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
 * function createFileFolderController(host: NativePropertyPaneEnvironment<SourceProperties>): NativePropertyPaneController {
 *   return {
 *     getConfiguration: () => ({ pages: [{
 *       header: { description: 'Select saved file and folder references, then use native Close. Reset is not native Discard.' },
 *       groups: [{ isGroupNameHidden: true, groupFields: [
 *         PropertyPaneFileFolderPicker('fileFolderValues', {
 *           instance: host.instance, key: 'file-folder-values', value: host.readProperties().fileFolderValues ?? [], label: 'Files and folders',
 *           preset: 'focused', scope: { kind: 'siteCollection' }, itemTypes: 'all',
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
 * // The focused scope uses pageContext.site.absoluteUrl from the initialized current
 * // collection, including from a subweb. Upload and folder creation are omitted
 * // and default to false. Existing delegated access remains a host prerequisite.
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
 *       return createFileFolderController(host);
 *     });
 *     if (result.status === 'blocked') {
 *       lease.current = undefined;
 *       setStatus(result.reason);
 *       return;
 *     }
 *     lease.current = result.lease;
 *     setStatus('Prepared. Open this WebPart’s actual native Edit property pane, use Browse and explicitly confirm files or folders or cancel, then use native Close. Preparing does not open the pane or query.');
 *   };
 *   const observe = (): void => {
 *     if (!lifetime.active || resetRevision !== previousRevision.current) return;
 *     const read = readProperties.current;
 *     if (!read) { setStatus('Open the prepared native scenario before inspecting its saved values.'); return; }
 *     setSaved(JSON.stringify(read().fileFolderValues ?? [], undefined, 2));
 *   };
 *   const regions: DemoRegions = propertyPane ? {
 *     commands: <><Button onClick={prepare}>Prepare native scenario</Button><Button onClick={observe}>Inspect saved file and folder references</Button></>,
 *     status: <p role="status">{status}</p>,
 *     children: <><p>Use the actual native Edit and Close controls. Selection, preparation and Reset do not open or close the pane. Reset cancels only an unproduced preparation and local observations; it cannot discard accepted native changes.</p><p>If the SDK, React or host fails after native configuration was returned, reload before further native scenario work; automatic recovery is unsupported.</p></>,
 *     diagnostics: <pre>{saved ?? 'Use Inspect saved file and folder references after an actual native selection.'}</pre>,
 *   } : {
 *     status: <p role="status">The initialized native scenario host is not connected. This selected native scenario is not active.</p>,
 *     children: <p>Connect an initialized native host to enable Prepare, then use this WebPart’s actual native Edit and Close. Selection and Reset do not query or open the pane.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', hostAdapter: 'propertyPane', observe: "Prepare without querying, open this WebPart's actual native Edit property pane, use hosted Browse and explicitly confirm files or folders or cancel, then use native Close and inspect saved JSON; selection and Reset do not open the pane or query." };
 * ```
 */
export function PropertyPaneFileFolderPicker<TProperties extends object>(
  targetProperty: FileFolderPickerPropertyKey<TProperties>, props: PropertyPaneFileFolderPickerProps<TProperties>,
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  if (typeof targetProperty !== 'string' || !targetProperty || /[.[\]]/.test(targetProperty) || ['__proto__', 'constructor', 'prototype'].includes(targetProperty)) {
    throw new Error('FileFolderPicker targetProperty must be a safe direct property name.');
  }
  if (typeof props.key !== 'string' || !props.key.trim()) throw new Error('FileFolderPicker key must be nonempty and stable.');
  const { instance, key, value, valueRevision, ...controlProps } = props;
  const token = {};
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty,
    properties: {
      key,
      onRender(domElement, _context, changeCallback) {
        mountOrUpdatePickerField(domElement, { owner: instance, key, token, render: runtime =>
          <SPFxWebPartProvider instance={instance}>
            <PropertyPanePickerHost container={domElement} runtime={runtime}>
              <PropertyPaneFileFolderPickerValueBridge controlProps={controlProps} value={value} valueRevision={valueRevision}
                targetProperty={targetProperty} onChange={changeCallback} />
            </PropertyPanePickerHost>
          </SPFxWebPartProvider>,
        });
      },
      onDispose(domElement) { disposePickerField(token, domElement); },
    },
  };
}

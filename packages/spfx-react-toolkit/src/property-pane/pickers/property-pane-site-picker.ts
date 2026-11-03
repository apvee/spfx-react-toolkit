/**
 * Native SPFx picker fields with JSON selection ownership and shared pane lifecycle.
 *
 * @packageDocumentation
 * @docPage property-pane/pickers
 */
import type { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import type { SitePickerValue } from '../../controls/pickers/picker-values.types';
import type { SitePickerProps, SitePickerOption } from '../../controls/pickers/picker.types';
import { SitePicker } from '../../controls/pickers/site-picker';
import { validateSitePickerValues } from '../../controls/pickers/picker-values-validation.internal';
import type { PickerPropertyKey, PropertyPanePickerProps } from './property-pane-picker.types';
import { createPropertyPanePicker } from './property-pane-picker-factory.internal';

/**
 * Native Property Pane Site picker, sharing the React control's options, source and presentation.
 * The host owns apply/persistence; arrays contain validated JSON references, including in single selection.
 * Same owner/key/DOM updates preserve root, selection/query state and equivalent request cache.
 * Initial unlabeled selections use selection-control AutoSkeleton during hydration; available labels stay visible. Search/refresh and later updates do not restart the skeleton.
 * The native field supplies its own matching WebPart provider, renderer and pane theme. Do not wrap it in JSX or mount it under a React provider yourself. Configure native resource mappings for the inherited React control. Existing host authentication and delegated permissions still govern queries.
 * Equivalent native renders preserve pending nonreactive drafts. Changing `valueRevision` explicitly adopts `value`; hydration never emits property changes. Invalid values or a missing native change callback disable mutations and display a diagnostic. Latest-factory disposal releases resources; stale disposal cannot unmount its replacement.
 * Inherited option defaults, source restrictions and asynchronous validation/error callbacks follow {@link SitePicker}; native persistence notification remains separate.
 * @useWhen A native SPFx configuration pane must persist SitePickerValue JSON references using the same query and validation options as the React picker.
 * @avoidWhen Rendering a picker in ordinary React content; use the corresponding React control and its selection callback instead.
 * @typeParam TProperties - Web-part property bag with a compatible direct SitePickerValue array.
 * @param targetProperty - Required eligible direct array property name, with no nested path or untyped element domain.
 * @param props - Required initialized web-part instance, stable field key, authoritative value and inherited control options.
 * @returns Native custom field for `groupFields`; the latest SPFx change callback receives targetProperty, a fresh validated array and true after an accepted explicit action. The factory does not save or apply properties.
 * @throws Error - The target property is empty, nested or prototype-related, or the field key is blank.
 * @example Native Site picker
 * ```ts
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
 * import { PropertyPaneSitePicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 * import type { SitePickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
 *
 * interface Properties { values?: SitePickerValue[]; }
 * export default class SelectionWebPart extends BaseClientSideWebPart<Properties> {
 *   public render(): void {
 *     this.domElement.textContent = `${this.properties.values?.length ?? 0} saved references`;
 *   }
 *   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
 *     return { pages: [{ groups: [{ groupName: 'Selection', groupFields: [
 *       PropertyPaneSitePicker('values', {
 *         instance: this, key: 'site-field', value: this.properties.values ?? [],
 *         label: 'Site',
 *       }),
 *     ] }] }] };
 *   }
 * }
 * ```
 *
 *
 * @example Configure six native reference arrays
 * ```tsx
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
 * import type {
 *   SitePickerValue, ListPickerValue, FieldPickerValue, ViewPickerValue,
 *   ListItemPickerValue, PrincipalPickerValue,
 * } from '@apvee/spfx-react-toolkit/controls/pickers';
 * import {
 *   PropertyPaneSitePicker, PropertyPaneListPicker, PropertyPaneFieldPicker,
 *   PropertyPaneViewPicker, PropertyPaneListItemPicker, PropertyPanePrincipalPicker,
 * } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 *
 * interface Properties {
 *   sites?: SitePickerValue[];
 *   lists?: ListPickerValue[];
 *   fields?: FieldPickerValue[];
 *   views?: ViewPickerValue[];
 *   items?: ListItemPickerValue[];
 *   principals?: PrincipalPickerValue[];
 * }
 * export default class ReferenceWebPart extends BaseClientSideWebPart<Properties> {
 *   public render(): void { this.domElement.textContent = 'Configure references in the Property Pane'; }
 *   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
 * return { pages: [{ groups: [{ groupName: 'References', groupFields: [
 *   PropertyPaneSitePicker('sites', { instance: this, key: 'sites-field', value: this.properties.sites ?? [], label: 'Site' }),
 *   PropertyPaneListPicker('lists', { instance: this, key: 'lists-field', value: this.properties.lists ?? [], label: 'List' }),
 *   PropertyPaneFieldPicker('fields', { instance: this, key: 'fields-field', value: this.properties.fields ?? [], list: 'Projects', label: 'Fields' }),
 *   PropertyPaneViewPicker('views', { instance: this, key: 'views-field', value: this.properties.views ?? [], list: 'Projects', label: 'View' }),
 *   PropertyPaneListItemPicker('items', { instance: this, key: 'items-field', value: this.properties.items ?? [], list: 'Projects', textField: 'Title', label: 'Items' }),
 *   PropertyPanePrincipalPicker('principals', { instance: this, key: 'principals-field', value: this.properties.principals ?? [], label: 'Principals',
 *     principalTypes: ['user', 'microsoft365Group'], userIdentity: { mode: 'both', ensureUser: false } }),
 * ] }] }] };
 *   }
 * }
 * ```
 *
 *
 * @example Supply a known native site-reference label
 * ```tsx
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import { PropertyPaneSitePicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 * import type { SitePickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
 * interface Properties { sites?: SitePickerValue[] }
 * // Place the returned field in the initialized host's groupFields. The host still owns persistence.
 * function knownSiteField(instance: BaseClientSideWebPart<Properties>, values: SitePickerValue[]) {
 *   return PropertyPaneSitePicker('sites', {
 *     instance, key: 'sites', label: 'Sito', hint: 'Seleziona una raccolta siti',
 *     value: values.map(site => site.url === 'https://contoso.sharepoint.com/sites/marketing'
 *       ? { ...site, text: 'Marketing' } : site),
 *   });
 * }
 * ```
 *
 * @example Select native site references
 * ```tsx
 * import * as React from 'react';
 * import { Button } from '@fluentui/react-components';
 * import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
 * import { PropertyPaneSitePicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 * import type { SitePickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
 *
 * interface SourceProperties { pickerSites?: SitePickerValue[]; }
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
 * function createSiteController(host: NativePropertyPaneEnvironment<SourceProperties>): NativePropertyPaneController {
 *   return {
 *     getConfiguration: () => ({ pages: [{
 *       header: { description: 'Select saved site references, then use native Close. Reset is not native Discard.' },
 *       groups: [{ isGroupNameHidden: true, groupFields: [
 *         PropertyPaneSitePicker('pickerSites', {
 *           instance: host.instance, key: 'picker-sites', value: host.readProperties().pickerSites ?? [], label: 'Sites',
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
 * // The initialized host owns actual native Edit/Close and SDK property persistence.
 * // This private structural port is not a public toolkit export. Native resources
 * // are configured as for the ordinary native Site recipe; host authentication and
 * // existing delegated permissions govern discovery. Omitted Site options retain
 * // the public SitePicker defaults, including the current-site source behavior.
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
 *       return createSiteController(host);
 *     });
 *     if (result.status === 'blocked') {
 *       lease.current = undefined;
 *       setStatus(result.reason);
 *       return;
 *     }
 *     lease.current = result.lease;
 *     setStatus('Prepared. Open this WebPart’s actual native Edit property pane, select Sites, then use native Close. Preparing does not open the pane or query.');
 *   };
 *   const observe = (): void => {
 *     if (!lifetime.active || resetRevision !== previousRevision.current) return;
 *     const read = readProperties.current;
 *     if (!read) { setStatus('Open the prepared native scenario before inspecting its saved values.'); return; }
 *     setSaved(JSON.stringify(read().pickerSites ?? [], undefined, 2));
 *   };
 *   const regions: DemoRegions = propertyPane ? {
 *     commands: <><Button onClick={prepare}>Prepare native scenario</Button><Button onClick={observe}>Inspect saved site references</Button></>,
 *     status: <p role="status">{status}</p>,
 *     children: <><p>Use the actual native Edit and Close controls. Selection, preparation and Reset do not open or close the pane. Reset cancels only an unproduced preparation and local observations; it cannot discard accepted native changes.</p><p>If the SDK, React or host fails after native configuration was returned, reload before further native scenario work; automatic recovery is unsupported.</p></>,
 *     diagnostics: <pre>{saved ?? 'Use Inspect saved site references after an actual native selection.'}</pre>,
 *   } : {
 *     status: <p role="status">The initialized native scenario host is not connected. This selected native scenario is not active.</p>,
 *     children: <p>Connect an initialized native host to enable Prepare, then use this WebPart’s actual native Edit and Close. Selection and Reset do not query or open the pane.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', hostAdapter: 'propertyPane', observe: "Prepare the native Site picker, open this WebPart's actual property pane, select site references and inspect the saved values; selection and Reset do not open the pane or query." };
 * ```
 *
 * @example Configure native metadata and file cascade
 * ```tsx
 * import * as React from 'react';
 * import { Button } from '@fluentui/react-components';
 * import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import { type IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneLabel, PropertyPaneToggle, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
 * import { PropertyPaneSitePicker, PropertyPaneListPicker, PropertyPaneFieldPicker, PropertyPaneViewPicker, PropertyPaneContentTypePicker, PropertyPaneListItemPicker, PropertyPanePrincipalPicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 * import { PropertyPaneFileFolderPicker } from '@apvee/spfx-react-toolkit/property-pane/file-folder-picker';
 * import type { SitePickerValue, ListPickerValue, FieldPickerValue, ViewPickerValue, ContentTypePickerValue, ListItemPickerValue, PrincipalPickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
 * import type { FileFolderPickerValue, FileFolderPickerScope, FileFolderPickerProps, FileFolderPickerPreset, FileFolderPickerInitialLocation } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';
 *
 * // Saved SDK-facing fields are read only for compatibility, never exposed as configurator switches.
 * type SavedNavigationProperty = 'fileFolderMyFiles' | 'fileFolderRecent' | 'fileFolderShared' | 'fileFolderQuickAccess' | 'fileFolderOrganization' | 'fileFolderSite';
 * const savedNavigation: readonly (readonly ['myFiles' | 'recent' | 'shared' | 'quickAccess' | 'organization' | 'site', SavedNavigationProperty])[] = [
 *   ['myFiles', 'fileFolderMyFiles'], ['recent', 'fileFolderRecent'], ['shared', 'fileFolderShared'],
 *   ['quickAccess', 'fileFolderQuickAccess'], ['organization', 'fileFolderOrganization'], ['site', 'fileFolderSite'],
 * ];
 *
 * type LegacyProps = Extract<FileFolderPickerProps, { readonly preset?: undefined }>;
 * interface FileFolderDemoSettings {
 *   fileFolderMyFiles?: boolean;
 *   fileFolderRecent?: boolean;
 *   fileFolderShared?: boolean;
 *   fileFolderQuickAccess?: boolean;
 *   fileFolderOrganization?: boolean;
 *   fileFolderSite?: boolean;
 *   fileFolderPreset?: FileFolderPickerPreset | 'legacy';
 *   fileFolderTitle?: string;
 *   fileFolderLayout?: LegacyProps['layout'];
 *   fileFolderGraph?: boolean;
 *   fileFolderNavigationVisible?: boolean;
 *   fileFolderInitialView?: LegacyProps['initialView'];
 *   fileFolderNavigationPreset?: LegacyProps['navigationPreset'] | 'auto';
 *   fileFolderSiteUrl?: string;
 *   fileFolderFromLink?: LegacyProps['fromLinkMode'];
 *   fileFolderDeviceUpload?: boolean;
 *   fileFolderDeviceLabel?: string;
 *   fileFolderUpload?: boolean;
 *   fileFolderCreateFolder?: boolean;
 * }
 * const savedSdkProperties: readonly (keyof FileFolderDemoSettings)[] = [
 *   ...savedNavigation.map(([, property]) => property), 'fileFolderTitle', 'fileFolderLayout', 'fileFolderGraph',
 *   'fileFolderNavigationVisible', 'fileFolderInitialView', 'fileFolderNavigationPreset', 'fileFolderSiteUrl',
 *   'fileFolderFromLink', 'fileFolderDeviceUpload', 'fileFolderDeviceLabel',
 * ];
 *
 * // Defined false and empty saved settings are meaningful compatibility preferences. Writes remain common intents.
 * function fileFolderHasLegacySettings(settings: FileFolderDemoSettings): boolean {
 *   return savedSdkProperties.some(property => settings[property] !== undefined);
 * }
 * function fileFolderDemoPreset(settings: FileFolderDemoSettings): FileFolderPickerPreset | 'legacy' {
 *   return settings.fileFolderPreset ?? (fileFolderHasLegacySettings(settings) ? 'legacy' : 'focused');
 * }
 *
 * type LegacyExperience = Pick<LegacyProps, 'title' | 'layout' | 'graphEnabled' | 'visibleLocations' | 'initialView' | 'navigationVisible' | 'navigationPreset' | 'siteNavigationUrl' | 'fromLinkMode' | 'localFilesMode' | 'localFilesLabel'> & { readonly preset?: undefined };
 * // Select one library contract; recipe generation stays in the library resolver. Original saved fields remain restorable.
 * function fileFolderDemoExperience(settings: FileFolderDemoSettings, scope: FileFolderPickerScope, initialLocation?: FileFolderPickerInitialLocation): LegacyExperience | { readonly preset: FileFolderPickerPreset } {
 *   const preset = fileFolderDemoPreset(settings);
 *   if (preset !== 'legacy') return { preset };
 *   const visibleLocations = savedNavigation.every(([, property]) => settings[property] === undefined) ? undefined
 *     : savedNavigation.filter(([key, property]) => {
 *       const incompatibleSource = (key === 'myFiles' && scope.kind !== 'oneDrive') || (key === 'quickAccess' && scope.kind === 'oneDrive');
 *       return !incompatibleSource && (settings[property] ?? (key === 'myFiles' && scope.kind === 'oneDrive'));
 *     }).map(([key]) => key);
 *   return {
 *     title: settings.fileFolderTitle || undefined, layout: settings.fileFolderLayout, graphEnabled: settings.fileFolderGraph,
 *     visibleLocations, navigationVisible: settings.fileFolderNavigationVisible,
 *     initialView: scope.kind === 'oneDrive' && !scope.folder && initialLocation === undefined ? settings.fileFolderInitialView ?? 'location' : 'location',
 *     navigationPreset: settings.fileFolderNavigationPreset === 'oneDrive' || settings.fileFolderNavigationPreset === 'current-site' ? settings.fileFolderNavigationPreset : undefined,
 *     siteNavigationUrl: settings.fileFolderSiteUrl?.trim() || undefined, fromLinkMode: settings.fileFolderFromLink,
 *     localFilesMode: settings.fileFolderUpload && settings.fileFolderDeviceUpload ? 'upload' : 'disabled', localFilesLabel: settings.fileFolderDeviceLabel || undefined,
 *   };
 * }
 *
 * interface SourceProperties extends FileFolderDemoSettings {
 *   pickerMultiselect?: boolean;
 *   fileFolderValues?: FileFolderPickerValue[];
 *   fileFolderScope?: FileFolderPickerScope['kind'];
 *   fileFolderItemTypes?: 'files' | 'folders' | 'all';
 *   fileFolderMultiselect?: boolean;
 *   fileFolderSize?: 'medium' | 'large';
 *   fileFolderColorfulTagAvatars?: boolean;
 *   fileFolderFolder?: string;
 *   fileFolderHint?: string;
 *   fileFolderSearchable?: boolean;
 *   fileFolderExtensions?: string;
 *   fileFolderCategory?: FileFolderPickerProps['fileCategory'] | '';
 *   fileFolderSelectionLimit?: string;
 *   fileFolderInitialWeb?: string;
 *   fileFolderInitialLibrary?: string;
 *   fileFolderInitialFolder?: string;
 *   pickerPresentation?: 'combobox' | 'tagPicker';
 *   pickerColorfulTagAvatars?: boolean;
 *
 *   pickerSites?: SitePickerValue[];
 *   pickerLists?: ListPickerValue[];
 *   pickerFields?: FieldPickerValue[];
 *   pickerViews?: ViewPickerValue[];
 *   pickerContentTypes?: ContentTypePickerValue[];
 *   pickerContentTypeScope?: 'site' | 'list';
 *   pickerItems?: ListItemPickerValue[];
 *   pickerPrincipals?: PrincipalPickerValue[];
 * }
 *
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
 *
 * const pickerValueKeys: readonly ('pickerSites' | 'pickerLists' | 'pickerFields' | 'pickerViews' | 'pickerContentTypes' | 'pickerItems' | 'pickerPrincipals')[] = ['pickerSites', 'pickerLists', 'pickerFields', 'pickerViews', 'pickerContentTypes', 'pickerItems', 'pickerPrincipals'];
 *
 * // Compare reference identities only: a label enrichment must not clear dependent values.
 * function pickerCascadeIdentity(value: unknown): string {
 *   const references: readonly unknown[] = Array.isArray(value) ? value : [];
 *   return JSON.stringify(references.map(reference => {
 *     if (!reference || typeof reference !== 'object' || !('kind' in reference)) return undefined;
 *     if (reference.kind === 'site' && 'url' in reference && typeof reference.url === 'string') return ['site', reference.url.toLowerCase().replace(/\/+$/, '')];
 *     if (reference.kind === 'list' && 'webUrl' in reference && typeof reference.webUrl === 'string' && 'listId' in reference && typeof reference.listId === 'string') return ['list', reference.webUrl.toLowerCase().replace(/\/+$/, ''), reference.listId.toLowerCase()];
 *     return undefined;
 *   }));
 * }
 *
 *
 * function clearFileFolderInitialLocation(properties: SourceProperties): void {
 *     delete properties.fileFolderInitialWeb;
 *     delete properties.fileFolderInitialLibrary;
 *     delete properties.fileFolderInitialFolder;
 *   }
 *
 *
 * function resolveFileFolderSettings(properties: SourceProperties, fallbackWeb: string): {
 *   readonly site: SitePickerValue | undefined;
 *   readonly list: ListPickerValue | undefined;
 *   readonly listTarget: { kind: 'id'; id: string } | undefined;
 *   readonly fileFolderKind: FileFolderPickerScope['kind'];
 *   readonly fileFolderScope: FileFolderPickerScope;
 *   readonly fileFolderCategory: FileFolderPickerProps['fileCategory'];
 *   readonly fileFolderDisabled: boolean;
 *   readonly savedFileFolderPreset: FileFolderPickerPreset | 'legacy';
 *   readonly fileFolderPreset: FileFolderPickerPreset;
 *   readonly fileFolderInitialLocation: FileFolderPickerInitialLocation | undefined;
 * } {
 *     const site = properties.pickerSites?.length === 1 ? properties.pickerSites[0] : undefined;
 *     const candidate = properties.pickerLists?.length === 1 ? properties.pickerLists[0] : undefined;
 *     const list = site !== undefined && candidate?.webUrl.toLowerCase().replace(/\/+$/, '') === site.url.toLowerCase().replace(/\/+$/, '') ? candidate : undefined;
 *     const listTarget: { kind: 'id'; id: string } | undefined = list === undefined ? undefined : { kind: 'id', id: list.listId };
 *     const fileFolderKind = properties.fileFolderScope ?? 'siteCollection';
 *     const fileFolderScope: FileFolderPickerScope = fileFolderKind === 'oneDrive' ? { kind: 'oneDrive', folder: properties.fileFolderFolder || undefined } :
 *       fileFolderKind === 'folder' ? { kind: 'folder', webUrl: list?.webUrl ?? fallbackWeb, list: listTarget ?? '', folder: properties.fileFolderFolder ?? '' } :
 *       fileFolderKind === 'library' ? { kind: 'library', webUrl: list?.webUrl ?? fallbackWeb, list: listTarget ?? '' } : { kind: 'siteCollection', siteUrl: site?.url };
 *     const fileFolderCategory = properties.fileFolderItemTypes === 'folders' ? undefined : properties.fileFolderCategory || undefined;
 *     const fileFolderDisabled = (fileFolderKind === 'library' || fileFolderKind === 'folder') && listTarget === undefined;
 *     const savedFileFolderPreset = fileFolderDemoPreset(properties);
 *     const fileFolderPreset = savedFileFolderPreset === 'legacy' ? 'focused' : savedFileFolderPreset;
 *     const fileFolderInitialLocation = properties.fileFolderInitialWeb || properties.fileFolderInitialLibrary || properties.fileFolderInitialFolder ? {
 *       webUrl: properties.fileFolderInitialWeb?.trim() || undefined,
 *       list: properties.fileFolderInitialLibrary?.trim() || undefined,
 *       folder: properties.fileFolderInitialFolder || undefined,
 *     } : undefined;
 *     return { site, list, listTarget, fileFolderKind, fileFolderScope, fileFolderCategory, fileFolderDisabled, savedFileFolderPreset, fileFolderPreset, fileFolderInitialLocation };
 * }
 *
 * function createIntegratedController(host: NativePropertyPaneEnvironment<SourceProperties>): NativePropertyPaneController {
 *   let pickerSiteRevision = 0;
 *   let pickerListRevision = 0;
 *   let pickerContentTypeRevision = 0;
 *   let fileFolderRevision = 0;
 *   const fieldChanged = (propertyPath: string, oldValue: unknown, newValue: unknown): void => {
 *     const properties = host.readProperties();
 *     if (['fileFolderScope', 'fileFolderFolder'].includes(propertyPath)) {
 *       if (oldValue !== newValue) {
 *         properties.fileFolderValues = []; fileFolderRevision++;
 *         const folderBound = properties.fileFolderScope === 'folder' || properties.fileFolderScope === 'oneDrive';
 *         if (propertyPath === 'fileFolderScope' || folderBound) clearFileFolderInitialLocation(properties);
 *       }
 *       host.refresh(); return;
 *     }
 *     if (propertyPath === 'fileFolderCategory' && newValue) delete properties.fileFolderExtensions;
 *     if (propertyPath === 'fileFolderItemTypes' && newValue === 'folders') delete properties.fileFolderCategory;
 *     if (propertyPath.startsWith('fileFolder')) {
 *       // Preset and functional options update the same field descriptor without resetting values or its revision.
 *       host.refresh(); return;
 *     }
 *     if (propertyPath === 'pickerMultiselect' || propertyPath === 'pickerPresentation' || propertyPath === 'pickerColorfulTagAvatars') {
 *       if (propertyPath === 'pickerMultiselect' && newValue === false) {
 *         const previousSites = properties.pickerSites;
 *         const previousLists = properties.pickerLists;
 *         for (const key of pickerValueKeys) {
 *           const values = properties[key];
 *           if (values !== undefined && values.length > 1) Object.assign(properties, { [key]: values.slice(0, 1) });
 *         }
 *         if (pickerCascadeIdentity(previousSites) !== pickerCascadeIdentity(properties.pickerSites)) {
 *           fieldChanged('pickerSites', previousSites, properties.pickerSites);
 *           return;
 *         }
 *         if (pickerCascadeIdentity(previousLists) !== pickerCascadeIdentity(properties.pickerLists)) {
 *           fieldChanged('pickerLists', previousLists, properties.pickerLists);
 *           return;
 *         }
 *       }
 *       host.refresh();
 *       return;
 *     }
 *     if (propertyPath === 'pickerContentTypeScope') {
 *       if ((oldValue ?? 'site') === (newValue ?? 'site')) return;
 *       properties.pickerContentTypes = [];
 *       pickerContentTypeRevision++;
 *       host.refresh();
 *       return;
 *     }
 *     if (propertyPath !== 'pickerSites' && propertyPath !== 'pickerLists') return;
 *     if (pickerCascadeIdentity(oldValue) === pickerCascadeIdentity(newValue)) return;
 *     const fileFolderKind = properties.fileFolderScope ?? 'siteCollection';
 *     const sourceDependsOnReference = propertyPath === 'pickerSites' ? fileFolderKind !== 'oneDrive'
 *       : fileFolderKind === 'library' || fileFolderKind === 'folder';
 *     if (sourceDependsOnReference) clearFileFolderInitialLocation(properties);
 *     properties.fileFolderValues = []; fileFolderRevision++;
 *     if (propertyPath === 'pickerSites') {
 *       properties.pickerLists = [];
 *       pickerSiteRevision++;
 *     }
 *     if (propertyPath === 'pickerSites' || properties.pickerContentTypeScope === 'list') {
 *       properties.pickerContentTypes = [];
 *       pickerContentTypeRevision++;
 *     }
 *     properties.pickerFields = [];
 *     properties.pickerViews = [];
 *     properties.pickerItems = [];
 *     pickerListRevision++;
 *     // Consumer-owned dependency updates need new field descriptors; stable keys retain mounted picker roots.
 *     host.refresh();
 *   };
 *   const getConfiguration = (): IPropertyPaneConfiguration => {
 *     const properties = host.readProperties();
 *     const modeProps: { multiselect: boolean; presentation: 'combobox' | 'tagPicker'; colorfulTagAvatars: boolean } = { multiselect: properties.pickerMultiselect ?? true,
 *       presentation: properties.pickerPresentation ?? 'combobox',
 *       colorfulTagAvatars: properties.pickerColorfulTagAvatars ?? false };
 *     const { site, list, listTarget, fileFolderKind, fileFolderScope, fileFolderCategory, fileFolderDisabled, fileFolderPreset, fileFolderInitialLocation } = resolveFileFolderSettings(properties, host.instance.context.pageContext.web.absoluteUrl);
 *     const contentTypeScope = properties.pickerContentTypeScope ?? 'site';
 *     const contentTypeProps = { instance: host.instance, ...modeProps, key: 'picker-content-types', value: properties.pickerContentTypes ?? [], valueRevision: pickerContentTypeRevision, webUrl: site?.url, label: 'Content types' };
 *     const fileFolderExperience = fileFolderDemoExperience({ ...properties, fileFolderPreset }, fileFolderScope, fileFolderInitialLocation);
 *     return {
 *       pages: [
 *         {
 *           displayGroupsAsAccordion: true,
 *           header: {
 *             description: 'Configure metadata and file references, inspect saved JSON, then use actual native Close. Reset is not Discard.'
 *           },
 *           groups: [
 *             { isGroupNameHidden: true, groupFields: [PropertyPaneSitePicker('pickerSites', { instance: host.instance, ...modeProps, key: 'picker-sites', value: properties.pickerSites ?? [], label: 'Sites' })] },
 *             {
 *               groupName: 'Picker references (JSON)',
 *               groupFields: [
 *                 PropertyPaneListPicker('pickerLists', { instance: host.instance, ...modeProps, key: 'picker-lists', value: properties.pickerLists ?? [], valueRevision: pickerSiteRevision, webUrl: site?.url, disabled: site === undefined, label: 'Lists', hint: site === undefined ? 'Select a site first.' : undefined }),
 *                 ...(listTarget === undefined ? [PropertyPaneLabel('picker-cascade-hint', { text: 'Select a site and a list to configure fields, views and items.' })] : [
 *                   PropertyPaneFieldPicker('pickerFields', { instance: host.instance, ...modeProps, key: 'picker-fields', value: properties.pickerFields ?? [], valueRevision: pickerListRevision, webUrl: list?.webUrl, list: listTarget, label: 'Fields' }),
 *                   PropertyPaneViewPicker('pickerViews', { instance: host.instance, ...modeProps, key: 'picker-views', value: properties.pickerViews ?? [], valueRevision: pickerListRevision, webUrl: list?.webUrl, list: listTarget, label: 'Views' }),
 *                   PropertyPaneListItemPicker('pickerItems', { instance: host.instance, ...modeProps, key: 'picker-items', value: properties.pickerItems ?? [], valueRevision: pickerListRevision, webUrl: list?.webUrl, list: listTarget, textField: 'Title', label: 'Items / files' }),
 *                 ]),
 *                 ...(site === undefined ? [PropertyPaneLabel('picker-content-type-hint', { text: 'Select a site to configure content types.' })] : contentTypeScope === 'site' ? [
 *                   PropertyPaneContentTypePicker('pickerContentTypes', { ...contentTypeProps, scope: 'site' }),
 *                 ] : listTarget === undefined ? [PropertyPaneLabel('picker-content-type-hint', { text: 'Select a site and a list to configure list content types.' })] : [
 *                   PropertyPaneContentTypePicker('pickerContentTypes', { ...contentTypeProps, webUrl: list?.webUrl, scope: 'list', list: listTarget }),
 *                 ]),
 *                 PropertyPanePrincipalPicker('pickerPrincipals', { instance: host.instance, ...modeProps, key: 'picker-principals', value: properties.pickerPrincipals ?? [], label: 'Principals', principalTypes: ['user', 'sharePointGroup', 'microsoft365Group', 'securityGroup', 'distributionList'], userIdentity: { mode: 'available' } }),
 *               ],
 *             },
 *             {
 *               groupName: 'File and folder picker',
 *               isCollapsed: false,
 *               groupFields: [
 *                 PropertyPaneLabel('file-folder-help', { text: 'Microsoft hosted v8. SharePoint scopes use the selected site/library above; only existing content is selected.' }),
 *                 PropertyPaneDropdown('fileFolderScope', { label: 'File picker source', selectedKey: fileFolderKind, options: [
 *                   { key: 'siteCollection', text: 'Site collection' }, { key: 'library', text: 'Selected document library' }, { key: 'folder', text: 'Folder in selected library' }, { key: 'oneDrive', text: 'Current user’s OneDrive Business' },
 *                 ] }),
 *                 ...(fileFolderKind === 'folder' || fileFolderKind === 'oneDrive' ? [PropertyPaneTextField('fileFolderFolder', { label: 'Existing folder path' })] : []),
 *                 PropertyPaneDropdown('fileFolderPreset', { label: 'Preset', selectedKey: fileFolderPreset, options: [
 *                   { key: 'focused', text: 'Focused' }, { key: 'extended', text: 'Extended' },
 *                 ] }),
 *                 PropertyPaneDropdown('fileFolderItemTypes', { label: 'Selectable content', selectedKey: properties.fileFolderItemTypes ?? 'files', options: [{ key: 'files', text: 'Files' }, { key: 'folders', text: 'Folders' }, { key: 'all', text: 'Files and folders' }] }),
 *                 PropertyPaneToggle('fileFolderMultiselect', { label: 'Multiple file/folder selection', checked: properties.fileFolderMultiselect ?? false }),
 *                 PropertyPaneFileFolderPicker('fileFolderValues', { instance: host.instance, key: 'file-folder-values', label: 'Selected files / folders', value: properties.fileFolderValues ?? [], valueRevision: fileFolderRevision,
 *                   size: properties.fileFolderSize ?? 'medium', colorfulTagAvatars: properties.fileFolderColorfulTagAvatars ?? false, scope: fileFolderScope, itemTypes: properties.fileFolderItemTypes ?? 'files', multiselect: properties.fileFolderMultiselect ?? false, disabled: fileFolderDisabled,
 *                   ...fileFolderExperience, initialLocation: fileFolderInitialLocation,
 *                   searchable: properties.fileFolderSearchable ?? true,
 *                   fileCategory: fileFolderCategory,
 *                   extensions: properties.fileFolderItemTypes === 'folders' || fileFolderCategory !== undefined ? undefined : properties.fileFolderExtensions?.split(',').map(extension => extension.trim()).filter(Boolean),
 *                   selectionLimit: properties.fileFolderMultiselect && properties.fileFolderSelectionLimit?.trim() ? Number(properties.fileFolderSelectionLimit) : undefined,
 *                   uploadEnabled: properties.fileFolderUpload ?? false, createFolderEnabled: properties.fileFolderCreateFolder ?? false,
 *                   hint: fileFolderDisabled ? 'Select exactly one site and one document library above.' : properties.fileFolderHint || undefined }),
 *               ],
 *             },
 *             {
 *               groupName: 'File picker options',
 *               isCollapsed: true,
 *               groupFields: [
 *                 ...(properties.fileFolderMultiselect ? [PropertyPaneTextField('fileFolderSelectionLimit', { label: 'Selection limit (empty: default)' })] : []),
 *                 ...(properties.fileFolderItemTypes !== 'folders' ? [
 *                   PropertyPaneDropdown('fileFolderCategory', { label: 'File category', selectedKey: fileFolderCategory ?? '', options: [
 *                     { key: '', text: 'Unrestricted / custom extensions' }, { key: 'images', text: 'Images' }, { key: 'video', text: 'Video' },
 *                     { key: 'audio', text: 'Audio' }, { key: 'media', text: 'Media' }, { key: 'documents', text: 'Documents' },
 *                   ] }),
 *                   PropertyPaneTextField('fileFolderExtensions', { label: 'File extensions (comma separated)', disabled: fileFolderCategory !== undefined }),
 *                   PropertyPaneLabel('file-folder-category-help', { text: 'A category replaces custom extensions and checks file suffixes. Mixed mode keeps folders selectable. Policy changes preserve selected references.' }),
 *                 ] : []),
 *                 PropertyPaneLabel('file-folder-initial-help', { text: 'Changing source clears the optional starting location. OneDrive accepts a folder only.' }),
 *                 PropertyPaneToggle('fileFolderSearchable', { label: 'Hosted search', checked: properties.fileFolderSearchable ?? true }),
 *                 ...(fileFolderKind !== 'oneDrive' ? [
 *                   PropertyPaneTextField('fileFolderInitialWeb', { label: 'Initial web URL (navigation only)' }),
 *                   PropertyPaneTextField('fileFolderInitialLibrary', { label: 'Initial library title (navigation only)' }),
 *                 ] : []),
 *                 PropertyPaneTextField('fileFolderInitialFolder', { label: 'Initial folder (navigation only)' }),
 *                 PropertyPaneDropdown('fileFolderSize', { label: 'File/folder tag size', selectedKey: properties.fileFolderSize ?? 'medium', options: [{ key: 'medium', text: 'Medium' }, { key: 'large', text: 'Large' }] }),
 *                 PropertyPaneToggle('fileFolderColorfulTagAvatars', { label: 'Colorful file/folder tag avatars', checked: properties.fileFolderColorfulTagAvatars ?? false }),
 *                 PropertyPaneTextField('fileFolderHint', { label: 'Field description (hint)' }),
 *                 PropertyPaneLabel('file-folder-writes', { text: 'Optional actions write immediately. Cancel and Property Pane Discard cannot undo those writes.' }),
 *                 PropertyPaneToggle('fileFolderUpload', { label: 'Enable uploads', checked: properties.fileFolderUpload ?? false }),
 *                 PropertyPaneToggle('fileFolderCreateFolder', { label: 'Enable folder creation', checked: properties.fileFolderCreateFolder ?? false }),
 *               ],
 *             },
 *             {
 *               groupName: 'File picker diagnostics',
 *               isCollapsed: true,
 *               groupFields: [
 *                 PropertyPaneLabel('file-folder-diagnostics', { text: `Source: ${JSON.stringify(fileFolderScope)}. Preset: ${fileFolderPreset}. Selection: ${JSON.stringify(properties.fileFolderValues ?? [])}` }),
 *               ],
 *             },
 *             {
 *               groupName: 'Picker modes',
 *               isCollapsed: false,
 *               groupFields: [
 *                 PropertyPaneLabel('picker-modes-help', { text: 'Global settings apply to all seven pickers. Single mode keeps the first selected value. Cascading fields require exactly one site and one list.' }),
 *                 PropertyPaneDropdown('pickerPresentation', { label: 'Presentation for all pickers', selectedKey: modeProps.presentation,
 *                   options: [{ key: 'combobox', text: 'Dropdown' }, { key: 'tagPicker', text: 'TagPicker' }] }),
 *                 PropertyPaneDropdown('pickerContentTypeScope', { label: 'Content type scope', selectedKey: contentTypeScope,
 *                   options: [{ key: 'site', text: 'Site collection' }, { key: 'list', text: 'List' }] }),
 *                 PropertyPaneToggle('pickerMultiselect', { label: 'Multiple selection for all pickers', checked: modeProps.multiselect }),
 *                 PropertyPaneToggle('pickerColorfulTagAvatars', { label: 'Colorful tag avatars for all pickers', checked: modeProps.colorfulTagAvatars }),
 *               ],
 *             },
 *           ]
 *         }
 *       ]
 *     };
 *   };
 *   return { getConfiguration, fieldChanged, configurationComplete: () => {}, changesApplied: () => {}, dispose: () => {} };
 * }
 *
 * // One produced native session owns business state independently of this Demo's commands.
 * export function Demo({ render, resetRevision = 0, propertyPane }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 *   propertyPane?: DemoPropertyPanePort<SourceProperties>;
 * }): React.ReactElement {
 *   const [status, setStatus] = React.useState('Prepare before opening this WebPart’s actual native property pane.');
 *   const [saved, setSaved] = React.useState<string>();
 *   const [compatibility, setCompatibility] = React.useState<string>();
 *   const lease = React.useRef<NativePropertyPanePreparation>();
 *   const readProperties = React.useRef<(() => SourceProperties)>();
 *   const fallbackWeb = React.useRef<string>();
 *   const lifetime = React.useMemo(() => ({ active: true }), [propertyPane, resetRevision]);
 *   const previousLifetime = React.useRef(lifetime);
 *   React.useEffect(() => {
 *     if (previousLifetime.current === lifetime) return;
 *     previousLifetime.current = lifetime;
 *     lease.current?.release();
 *     lease.current = undefined;
 *     readProperties.current = undefined;
 *     fallbackWeb.current = undefined;
 *     setSaved(undefined);
 *     setCompatibility(undefined);
 *     setStatus('Local state reset. A produced native session survives until its real native lifecycle ends; Reset is not Discard or rollback.');
 *   }, [lifetime]);
 *   React.useEffect(() => () => {
 *     lifetime.active = false;
 *     lease.current?.release();
 *     lease.current = undefined;
 *     readProperties.current = undefined;
 *     fallbackWeb.current = undefined;
 *   }, [lifetime]);
 *   const prepare = (): void => {
 *     if (!propertyPane || !lifetime.active || lifetime !== previousLifetime.current) return;
 *     lease.current?.release();
 *     const result = propertyPane.prepare(host => {
 *       readProperties.current = host.readProperties;
 *       fallbackWeb.current = host.instance.context.pageContext.web.absoluteUrl;
 *       return createIntegratedController(host);
 *     });
 *     if (result.status === 'blocked') { lease.current = undefined; setStatus(result.reason); return; }
 *     lease.current = result.lease;
 *     setStatus('Prepared. Open this WebPart’s actual native Edit property pane, configure metadata and files, then use native Close. Preparing does not open the pane or query.');
 *   };
 *   const observe = (): void => {
 *     if (!lifetime.active || lifetime !== previousLifetime.current) return;
 *     const read = readProperties.current;
 *     if (!read) { setStatus('Open the prepared native scenario before inspecting its saved values.'); return; }
 *     setSaved(JSON.stringify(read(), undefined, 2));
 *   };
 *   const inspectCompatibility = (): void => {
 *     if (!lifetime.active || lifetime !== previousLifetime.current) return;
 *     const read = readProperties.current, web = fallbackWeb.current;
 *     if (!read || web === undefined) { setStatus('Open the prepared native scenario before inspecting saved compatibility mapping.'); return; }
 *     const properties = read();
 *     const { fileFolderScope, fileFolderInitialLocation, savedFileFolderPreset, fileFolderPreset } = resolveFileFolderSettings(properties, web);
 *     setCompatibility(JSON.stringify({ hasLegacySettings: fileFolderHasLegacySettings(properties), savedPreset: savedFileFolderPreset,
 *       compatibility: fileFolderDemoExperience(properties, fileFolderScope, fileFolderInitialLocation),
 *       effective: fileFolderDemoExperience({ ...properties, fileFolderPreset }, fileFolderScope, fileFolderInitialLocation) }, undefined, 2));
 *   };
 *   const regions: DemoRegions = propertyPane ? {
 *     commands: <><Button onClick={prepare}>Prepare native scenario</Button><Button onClick={observe}>Inspect saved JSON</Button><Button onClick={inspectCompatibility}>Inspect saved compatibility mapping</Button></>,
 *     status: <p role="status">{status}</p>,
 *     children: <><p>Use actual native Edit and Close. Site and list changes clear their dependent references. Global single mode keeps the first of each of seven metadata arrays; file selection has independent settings.</p><p>Selection, preparation and Reset do not open or close the pane. Reset cancels only unproduced preparation and local observations. Accepted optional uploads and folder creation write immediately; Cancel and Property Pane Discard cannot undo those writes.</p><p>If the SDK, React or host fails after native configuration was returned, reload before further native scenario work; automatic recovery is unsupported.</p></>,
 *     diagnostics: <><pre aria-label="Saved native JSON">{saved ?? 'Inspect saved JSON after the actual native getter.'}</pre><pre aria-label="Saved compatibility mapping">{compatibility ?? 'Inspect saved compatibility mapping after the actual native getter. Saved legacy preferences remain stored; the native menu uses focused or extended.'}</pre></>,
 *   } : {
 *     status: <p role="status">The initialized native scenario host is not connected. This selected native scenario is not active.</p>,
 *     children: <p>Connect an initialized native host to enable Prepare, then use this WebPart’s actual native Edit and Close. Selection and Reset do not query or open the pane.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', hostAdapter: 'propertyPane', observe: "Prepare, use actual native Edit to configure metadata and file cascades, then inspect live saved JSON and compatibility policy; use native Close and observe that Reset does not discard native changes." };
 * ```
 */
export function PropertyPaneSitePicker<TProperties extends object>(targetProperty: PickerPropertyKey<TProperties, SitePickerValue>,
  props: PropertyPanePickerProps<TProperties, SitePickerProps, SitePickerValue>): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  const { instance, key, value, valueRevision, ...controlProps } = props;
  return createPropertyPanePicker<TProperties, SitePickerValue, SitePickerOption, SitePickerProps>(targetProperty, { instance, key, value, valueRevision }, SitePicker, validateSitePickerValues, controlProps);
}

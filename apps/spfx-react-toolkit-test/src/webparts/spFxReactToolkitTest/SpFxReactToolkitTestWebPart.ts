import { fileFolderDemoExperience, fileFolderDemoPreset, fileFolderHasLegacySettings } from './components/shared/fileFolderPickerNavigation';
import type { FileFolderDemoSettings } from './components/shared/fileFolderPickerNavigation';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SpFxReactToolkitTestWebPartStrings';
import { getWebPartShellPropertiesMetadata, type IWebPartShellProperties } from '@apvee/spfx-react-toolkit';
import { WebPartShellDemoHost } from './components/WebPartShellDemoHost';
import SpFxReactToolkitTest from './components/SpFxReactToolkitTest';

import { PropertyPaneFileFolderPicker } from '@apvee/spfx-react-toolkit/property-pane/file-folder-picker';
import type { FileFolderPickerValue, FileFolderPickerScope } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';
import { PropertyPaneSitePicker, PropertyPaneListPicker, PropertyPaneFieldPicker, PropertyPaneViewPicker, PropertyPaneContentTypePicker, PropertyPaneListItemPicker, PropertyPanePrincipalPicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
import type { SitePickerValue, ListPickerValue, FieldPickerValue, ViewPickerValue, ContentTypePickerValue, ListItemPickerValue, PrincipalPickerValue } from '@apvee/spfx-react-toolkit';

export interface ISpFxReactToolkitTestWebPartProps extends IWebPartShellProperties, FileFolderDemoSettings {
  description: string;
  pickerMultiselect?: boolean;
  fileFolderValues?: FileFolderPickerValue[];
  fileFolderScope?: FileFolderPickerScope['kind'];
  fileFolderItemTypes?: 'files' | 'folders' | 'all';
  fileFolderMultiselect?: boolean;
  fileFolderSize?: 'medium' | 'large';
  fileFolderFolder?: string;
  fileFolderHint?: string;
  fileFolderSearchable?: boolean;
  fileFolderExtensions?: string;
  fileFolderSelectionLimit?: string;
  fileFolderInitialWeb?: string;
  fileFolderInitialLibrary?: string;
  fileFolderInitialFolder?: string;
  pickerPresentation?: 'combobox' | 'tagPicker';

  pickerSites?: SitePickerValue[];
  pickerLists?: ListPickerValue[];
  pickerFields?: FieldPickerValue[];
  pickerViews?: ViewPickerValue[];
  pickerContentTypes?: ContentTypePickerValue[];
  pickerContentTypeScope?: 'site' | 'list';
  pickerItems?: ListItemPickerValue[];
  pickerPrincipals?: PrincipalPickerValue[];
}

const pickerValueKeys = ['pickerSites', 'pickerLists', 'pickerFields', 'pickerViews', 'pickerContentTypes', 'pickerItems', 'pickerPrincipals'] as const;

// Compare reference identities only: a label enrichment must not clear dependent values.
function pickerCascadeIdentity(value: unknown): string {
  const references: readonly unknown[] = Array.isArray(value) ? value : [];
  return JSON.stringify(references.map(reference => {
    if (!reference || typeof reference !== 'object' || !('kind' in reference)) return undefined;
    if (reference.kind === 'site' && 'url' in reference && typeof reference.url === 'string') return ['site', reference.url.toLowerCase().replace(/\/+$/, '')];
    if (reference.kind === 'list' && 'webUrl' in reference && typeof reference.webUrl === 'string' && 'listId' in reference && typeof reference.listId === 'string') return ['list', reference.webUrl.toLowerCase().replace(/\/+$/, ''), reference.listId.toLowerCase()];
    return undefined;
  }));
}

export default class SpFxReactToolkitTestWebPart extends BaseClientSideWebPart<ISpFxReactToolkitTestWebPartProps> {

  private _pickerSiteRevision = 0;
  private _pickerListRevision = 0;
  private _pickerContentTypeRevision = 0;
  private _fileFolderRevision = 0;

  private _clearFileFolderInitialLocation(): void {
    delete this.properties.fileFolderInitialWeb;
    delete this.properties.fileFolderInitialLibrary;
    delete this.properties.fileFolderInitialFolder;
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    if (['fileFolderScope', 'fileFolderFolder'].includes(propertyPath)) {
      if (oldValue !== newValue) {
        this.properties.fileFolderValues = []; this._fileFolderRevision++;
        const folderBound = this.properties.fileFolderScope === 'folder' || this.properties.fileFolderScope === 'oneDrive';
        if (propertyPath === 'fileFolderScope' || folderBound) this._clearFileFolderInitialLocation();
      }
      this.context.propertyPane.refresh(); return;
    }
    if (propertyPath.startsWith('fileFolder')) {
      // Preset and functional options update the same field descriptor without resetting values or its revision.
      this.context.propertyPane.refresh(); return;
    }
    if (propertyPath === 'pickerMultiselect' || propertyPath === 'pickerPresentation') {
      if (propertyPath === 'pickerMultiselect' && newValue === false) {
        const previousSites = this.properties.pickerSites;
        const previousLists = this.properties.pickerLists;
        for (const key of pickerValueKeys) {
          const values = this.properties[key];
          if (values !== undefined && values.length > 1) Object.assign(this.properties, { [key]: values.slice(0, 1) });
        }
        if (pickerCascadeIdentity(previousSites) !== pickerCascadeIdentity(this.properties.pickerSites)) {
          this.onPropertyPaneFieldChanged('pickerSites', previousSites, this.properties.pickerSites);
          return;
        }
        if (pickerCascadeIdentity(previousLists) !== pickerCascadeIdentity(this.properties.pickerLists)) {
          this.onPropertyPaneFieldChanged('pickerLists', previousLists, this.properties.pickerLists);
          return;
        }
      }
      this.context.propertyPane.refresh();
      return;
    }
    if (propertyPath === 'pickerContentTypeScope') {
      if ((oldValue ?? 'site') === (newValue ?? 'site')) return;
      this.properties.pickerContentTypes = [];
      this._pickerContentTypeRevision++;
      this.context.propertyPane.refresh();
      return;
    }
    if (propertyPath !== 'pickerSites' && propertyPath !== 'pickerLists') return;
    if (pickerCascadeIdentity(oldValue) === pickerCascadeIdentity(newValue)) return;
    const fileFolderKind = this.properties.fileFolderScope ?? 'siteCollection';
    const sourceDependsOnReference = propertyPath === 'pickerSites' ? fileFolderKind !== 'oneDrive'
      : fileFolderKind === 'library' || fileFolderKind === 'folder';
    if (sourceDependsOnReference) this._clearFileFolderInitialLocation();
    this.properties.fileFolderValues = []; this._fileFolderRevision++;
    if (propertyPath === 'pickerSites') {
      this.properties.pickerLists = [];
      this._pickerSiteRevision++;
    }
    if (propertyPath === 'pickerSites' || this.properties.pickerContentTypeScope === 'list') {
      this.properties.pickerContentTypes = [];
      this._pickerContentTypeRevision++;
    }
    this.properties.pickerFields = [];
    this.properties.pickerViews = [];
    this.properties.pickerItems = [];
    this._pickerListRevision++;
    // Consumer-owned dependency updates need new field descriptors; stable keys retain mounted picker roots.
    this.context.propertyPane.refresh();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(
      WebPartShellDemoHost,
      { instance: this },
      React.createElement(SpFxReactToolkitTest)
    );
    
    ReactDom.render(element, this.domElement);
  }

  protected get propertiesMetadata(): ReturnType<typeof getWebPartShellPropertiesMetadata> {
    return getWebPartShellPropertiesMetadata();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const modeProps = { multiselect: this.properties.pickerMultiselect ?? true,
      presentation: this.properties.pickerPresentation ?? 'combobox' as const };
    const site = this.properties.pickerSites?.length === 1 ? this.properties.pickerSites[0] : undefined;
    const candidate = this.properties.pickerLists?.length === 1 ? this.properties.pickerLists[0] : undefined;
    const list = site !== undefined && candidate?.webUrl.toLowerCase().replace(/\/+$/, '') === site.url.toLowerCase().replace(/\/+$/, '') ? candidate : undefined;
    const listTarget = list === undefined ? undefined : { kind: 'id' as const, id: list.listId };
    const contentTypeScope = this.properties.pickerContentTypeScope ?? 'site';
    const contentTypeProps = { instance: this, ...modeProps, key: 'picker-content-types', value: this.properties.pickerContentTypes ?? [], valueRevision: this._pickerContentTypeRevision, webUrl: site?.url, label: 'Content types' };
    const fileFolderKind = this.properties.fileFolderScope ?? 'siteCollection';
    const fileFolderScope: FileFolderPickerScope = fileFolderKind === 'oneDrive' ? { kind: 'oneDrive', folder: this.properties.fileFolderFolder || undefined } :
      fileFolderKind === 'folder' ? { kind: 'folder', webUrl: list?.webUrl ?? this.context.pageContext.web.absoluteUrl, list: listTarget ?? '', folder: this.properties.fileFolderFolder ?? '' } :
      fileFolderKind === 'library' ? { kind: 'library', webUrl: list?.webUrl ?? this.context.pageContext.web.absoluteUrl, list: listTarget ?? '' } : { kind: 'siteCollection', siteUrl: site?.url };
    const fileFolderDisabled = (fileFolderKind === 'library' || fileFolderKind === 'folder') && listTarget === undefined;
    const fileFolderPreset = fileFolderDemoPreset(this.properties);
    const fileFolderInitialLocation = this.properties.fileFolderInitialWeb || this.properties.fileFolderInitialLibrary || this.properties.fileFolderInitialFolder ? {
      webUrl: this.properties.fileFolderInitialWeb?.trim() || undefined,
      list: this.properties.fileFolderInitialLibrary?.trim() || undefined,
      folder: this.properties.fileFolderInitialFolder || undefined,
    } : undefined;
    const fileFolderExperience = fileFolderDemoExperience(this.properties, fileFolderScope, fileFolderInitialLocation);
    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: 'Picker references (JSON)',
              groupFields: [
                PropertyPaneSitePicker('pickerSites', { instance: this, ...modeProps, key: 'picker-sites', value: this.properties.pickerSites ?? [], label: 'Sites' }),
                PropertyPaneListPicker('pickerLists', { instance: this, ...modeProps, key: 'picker-lists', value: this.properties.pickerLists ?? [], valueRevision: this._pickerSiteRevision, webUrl: site?.url, disabled: site === undefined, label: 'Lists', hint: site === undefined ? 'Select a site first.' : undefined }),
                ...(listTarget === undefined ? [PropertyPaneLabel('picker-cascade-hint', { text: 'Select a site and a list to configure fields, views and items.' })] : [
                  PropertyPaneFieldPicker('pickerFields', { instance: this, ...modeProps, key: 'picker-fields', value: this.properties.pickerFields ?? [], valueRevision: this._pickerListRevision, webUrl: list?.webUrl, list: listTarget, label: 'Fields' }),
                  PropertyPaneViewPicker('pickerViews', { instance: this, ...modeProps, key: 'picker-views', value: this.properties.pickerViews ?? [], valueRevision: this._pickerListRevision, webUrl: list?.webUrl, list: listTarget, label: 'Views' }),
                  PropertyPaneListItemPicker('pickerItems', { instance: this, ...modeProps, key: 'picker-items', value: this.properties.pickerItems ?? [], valueRevision: this._pickerListRevision, webUrl: list?.webUrl, list: listTarget, textField: 'Title', label: 'Items / files' }),
                ]),
                ...(site === undefined ? [PropertyPaneLabel('picker-content-type-hint', { text: 'Select a site to configure content types.' })] : contentTypeScope === 'site' ? [
                  PropertyPaneContentTypePicker('pickerContentTypes', { ...contentTypeProps, scope: 'site' }),
                ] : listTarget === undefined ? [PropertyPaneLabel('picker-content-type-hint', { text: 'Select a site and a list to configure list content types.' })] : [
                  PropertyPaneContentTypePicker('pickerContentTypes', { ...contentTypeProps, webUrl: list?.webUrl, scope: 'list', list: listTarget }),
                ]),
                PropertyPanePrincipalPicker('pickerPrincipals', { instance: this, ...modeProps, key: 'picker-principals', value: this.properties.pickerPrincipals ?? [], label: 'Principals', principalTypes: ['user', 'sharePointGroup', 'microsoft365Group', 'securityGroup', 'distributionList'], userIdentity: { mode: 'available' } }),
              ],
            },
            {
              groupName: 'File and folder picker',
              isCollapsed: false,
              groupFields: [
                PropertyPaneLabel('file-folder-help', { text: 'Microsoft hosted v8. SharePoint scopes use the selected site/library above; only existing content is selected.' }),
                PropertyPaneDropdown('fileFolderScope', { label: 'File picker source', selectedKey: fileFolderKind, options: [
                  { key: 'siteCollection', text: 'Site collection' }, { key: 'library', text: 'Selected document library' }, { key: 'folder', text: 'Folder in selected library' }, { key: 'oneDrive', text: 'Current user’s OneDrive Business' },
                ] }),
                ...(fileFolderKind === 'folder' || fileFolderKind === 'oneDrive' ? [PropertyPaneTextField('fileFolderFolder', { label: 'Existing folder path' })] : []),
                PropertyPaneDropdown('fileFolderPreset', { label: 'Preset', selectedKey: fileFolderPreset, options: [
                  ...(fileFolderHasLegacySettings(this.properties) || fileFolderPreset === 'legacy' ? [{ key: 'legacy', text: 'Existing configuration (legacy)' }] : []),
                  { key: 'focused', text: 'Focused' }, { key: 'extended', text: 'Extended' },
                ] }),
                PropertyPaneDropdown('fileFolderItemTypes', { label: 'Selectable content', selectedKey: this.properties.fileFolderItemTypes ?? 'files', options: [{ key: 'files', text: 'Files' }, { key: 'folders', text: 'Folders' }, { key: 'all', text: 'Files and folders' }] }),
                PropertyPaneToggle('fileFolderMultiselect', { label: 'Multiple file/folder selection', checked: this.properties.fileFolderMultiselect ?? false }),
                PropertyPaneFileFolderPicker('fileFolderValues', { instance: this, key: 'file-folder-values', label: 'Selected files / folders', value: this.properties.fileFolderValues ?? [], valueRevision: this._fileFolderRevision,
                  size: this.properties.fileFolderSize ?? 'medium', scope: fileFolderScope, itemTypes: this.properties.fileFolderItemTypes ?? 'files', multiselect: this.properties.fileFolderMultiselect ?? false, disabled: fileFolderDisabled,
                  ...fileFolderExperience, initialLocation: fileFolderInitialLocation,
                  searchable: this.properties.fileFolderSearchable ?? true,
                  extensions: this.properties.fileFolderItemTypes === 'folders' ? undefined : this.properties.fileFolderExtensions?.split(',').map(extension => extension.trim()).filter(Boolean),
                  selectionLimit: this.properties.fileFolderMultiselect && this.properties.fileFolderSelectionLimit?.trim() ? Number(this.properties.fileFolderSelectionLimit) : undefined,
                  uploadEnabled: this.properties.fileFolderUpload ?? false, createFolderEnabled: this.properties.fileFolderCreateFolder ?? false,
                  hint: fileFolderDisabled ? 'Select exactly one site and one document library above.' : this.properties.fileFolderHint || undefined }),
              ],
            },
            {
              groupName: 'File picker options',
              isCollapsed: true,
              groupFields: [
                ...(this.properties.fileFolderMultiselect ? [PropertyPaneTextField('fileFolderSelectionLimit', { label: 'Selection limit (empty: default)' })] : []),
                ...(this.properties.fileFolderItemTypes !== 'folders' ? [PropertyPaneTextField('fileFolderExtensions', { label: 'File extensions (comma separated)' })] : []),
                PropertyPaneLabel('file-folder-initial-help', { text: 'Changing source clears the optional starting location. OneDrive accepts a folder only.' }),
                PropertyPaneToggle('fileFolderSearchable', { label: 'Hosted search', checked: this.properties.fileFolderSearchable ?? true }),
                ...(fileFolderKind !== 'oneDrive' ? [
                  PropertyPaneTextField('fileFolderInitialWeb', { label: 'Initial web URL (navigation only)' }),
                  PropertyPaneTextField('fileFolderInitialLibrary', { label: 'Initial library title (navigation only)' }),
                ] : []),
                PropertyPaneTextField('fileFolderInitialFolder', { label: 'Initial folder (navigation only)' }),
                PropertyPaneDropdown('fileFolderSize', { label: 'File/folder tag size', selectedKey: this.properties.fileFolderSize ?? 'medium', options: [{ key: 'medium', text: 'Medium' }, { key: 'large', text: 'Large' }] }),
                PropertyPaneTextField('fileFolderHint', { label: 'Field description (hint)' }),
                PropertyPaneLabel('file-folder-writes', { text: 'Optional actions write immediately. Cancel and Property Pane Discard cannot undo those writes.' }),
                PropertyPaneToggle('fileFolderUpload', { label: 'Enable uploads', checked: this.properties.fileFolderUpload ?? false }),
                PropertyPaneToggle('fileFolderCreateFolder', { label: 'Enable folder creation', checked: this.properties.fileFolderCreateFolder ?? false }),
              ],
            },
            {
              groupName: 'File picker diagnostics',
              isCollapsed: true,
              groupFields: [
                PropertyPaneLabel('file-folder-diagnostics', { text: `Source: ${JSON.stringify(fileFolderScope)}. Preset: ${fileFolderPreset}. Selection: ${JSON.stringify(this.properties.fileFolderValues ?? [])}` }),
              ],
            },
            {
              groupName: 'Picker modes',
              isCollapsed: false,
              groupFields: [
                PropertyPaneLabel('picker-modes-help', { text: 'Global settings apply to all seven pickers. Single mode keeps the first selected value. Cascading fields require exactly one site and one list.' }),
                PropertyPaneDropdown('pickerPresentation', { label: 'Presentation for all pickers', selectedKey: modeProps.presentation,
                  options: [{ key: 'combobox', text: 'Combobox' }, { key: 'tagPicker', text: 'TagPicker' }] }),
                PropertyPaneDropdown('pickerContentTypeScope', { label: 'Content type scope', selectedKey: contentTypeScope,
                  options: [{ key: 'site', text: 'Site collection' }, { key: 'list', text: 'List' }] }),
                PropertyPaneToggle('pickerMultiselect', { label: 'Multiple selection for all pickers', checked: modeProps.multiselect }),
              ],
            },
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneToggle('showTitle', { label: 'Show title', checked: this.properties.showTitle ?? true }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

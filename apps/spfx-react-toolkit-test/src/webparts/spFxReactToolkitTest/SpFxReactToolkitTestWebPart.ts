import type { FileFolderDemoSettings } from './components/shared/fileFolderPickerNavigation';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SpFxReactToolkitTestWebPartStrings';
import { getWebPartShellPropertiesMetadata, type IWebPartShellProperties } from '@apvee/spfx-react-toolkit';
import { WebPartShellDemoHost } from './components/WebPartShellDemoHost';
import SpFxReactToolkitTest from './components/SpFxReactToolkitTest';
import { NativePropertyPaneDemoHost } from './components/shared/NativePropertyPaneDemoHost';
import { NativePropertyPaneContext, NativePropertyPaneObservations } from './components/shared/NativePropertyPaneContext';

import type { FileFolderPickerValue, FileFolderPickerScope, FileFolderPickerProps } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';
import type { SitePickerValue, ListPickerValue, FieldPickerValue, ViewPickerValue, ContentTypePickerValue, ListItemPickerValue, PrincipalPickerValue } from '@apvee/spfx-react-toolkit';

export interface ISpFxReactToolkitTestWebPartProps extends IWebPartShellProperties, FileFolderDemoSettings {
  description: string;
  pickerMultiselect?: boolean;
  fileFolderValues?: FileFolderPickerValue[];
  fileFolderScope?: FileFolderPickerScope['kind'];
  fileFolderItemTypes?: 'files' | 'folders' | 'all';
  fileFolderMultiselect?: boolean;
  fileFolderSize?: 'medium' | 'large';
  fileFolderColorfulTagAvatars?: boolean;
  fileFolderFolder?: string;
  fileFolderHint?: string;
  fileFolderSearchable?: boolean;
  fileFolderExtensions?: string;
  fileFolderCategory?: FileFolderPickerProps['fileCategory'] | '';
  fileFolderSelectionLimit?: string;
  fileFolderInitialWeb?: string;
  fileFolderInitialLibrary?: string;
  fileFolderInitialFolder?: string;
  pickerPresentation?: 'combobox' | 'tagPicker';
  pickerColorfulTagAvatars?: boolean;

  pickerSites?: SitePickerValue[];
  pickerLists?: ListPickerValue[];
  pickerFields?: FieldPickerValue[];
  pickerViews?: ViewPickerValue[];
  pickerContentTypes?: ContentTypePickerValue[];
  pickerContentTypeScope?: 'site' | 'list';
  pickerItems?: ListItemPickerValue[];
  pickerPrincipals?: PrincipalPickerValue[];
}

export default class SpFxReactToolkitTestWebPart extends BaseClientSideWebPart<ISpFxReactToolkitTestWebPartProps> {
  private readonly nativeObservations = new NativePropertyPaneObservations();
  private readonly nativeHost = new NativePropertyPaneDemoHost({
    instance: this,
    readProperties: () => this.properties,
    isOpen: () => this.context.propertyPane.isPropertyPaneOpen(),
    refresh: () => this.context.propertyPane.refresh()
  }, (event, fieldKey) => this.nativeObservations.record(event, this.context.propertyPane.isPropertyPaneOpen(), fieldKey));
  private readonly nativeContext = { port: this.nativeHost.port, observations: this.nativeObservations };

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    this.nativeHost.fieldChanged(propertyPath, oldValue, newValue);
  }
  protected onPropertyPaneConfigurationStart(): void { /* Native tokenless Start grants no adapter ownership or refresh rights. */ }
  protected onPropertyPaneRendered(): void {
    this.nativeObservations.record('parent-rendered', this.context.propertyPane.isPropertyPaneOpen());
    this.nativeHost.rendered();
  }
  protected onPropertyPaneConfigurationComplete(): void {
    this.nativeObservations.record('complete', this.context.propertyPane.isPropertyPaneOpen());
    this.nativeHost.configurationComplete();
  }
  protected onAfterPropertyPaneChangesApplied(): void { this.nativeHost.changesApplied(); }

  public render(): void {
    const element: React.ReactElement = React.createElement(
      NativePropertyPaneContext.Provider,
      { value: this.nativeContext },
      React.createElement(WebPartShellDemoHost, { instance: this }, React.createElement(SpFxReactToolkitTest))
    );
    ReactDom.render(element, this.domElement);
  }

  protected get propertiesMetadata(): ReturnType<typeof getWebPartShellPropertiesMetadata> {
    return getWebPartShellPropertiesMetadata();
  }

  protected onDispose(): void {
    this.nativeObservations.record('terminal-dispose', this.context.propertyPane.isPropertyPaneOpen());
    this.nativeHost.dispose();
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    this.nativeObservations.record('getter', this.context.propertyPane.isPropertyPaneOpen());
    const configuration = this.nativeHost.getConfiguration(() => ({ pages: [{ header: {
      description: 'Select a native scenario and Prepare while the native pane is closed, then use actual Edit and Close. If an external SDK or React failure occurs after configuration, reload before further native work; automatic recovery is unsupported.'
    }, groups: [] }] }));
    // Clone plain native configuration; source owns its first subject and business groups.
    return { ...configuration, pages: configuration.pages.map(page => ({ ...page,
      header: { ...page.header, description: [strings.PropertyPaneDescription, page.header?.description].filter(Boolean).join(' ') },
      groups: [...page.groups, { groupName: strings.BasicGroupName, groupFields: [
        PropertyPaneToggle('showTitle', { label: 'Show title', checked: this.properties.showTitle ?? true }),
        PropertyPaneTextField('description', { label: strings.DescriptionFieldLabel })
      ] }]
    })) };
  }
}

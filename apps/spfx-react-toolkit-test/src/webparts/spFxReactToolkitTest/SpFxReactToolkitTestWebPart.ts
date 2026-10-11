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

export interface ISpFxReactToolkitTestWebPartProps extends IWebPartShellProperties {
  description: string;
}

export default class SpFxReactToolkitTestWebPart extends BaseClientSideWebPart<ISpFxReactToolkitTestWebPartProps> {

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
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
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

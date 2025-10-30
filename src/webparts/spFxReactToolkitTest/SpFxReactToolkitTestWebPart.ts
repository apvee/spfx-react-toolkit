import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SpFxReactToolkitTestWebPartStrings';
import { SPFxWebPartProvider } from '../../core';
import SpFxReactToolkitTest from './components/SpFxReactToolkitTest';

export interface ISpFxReactToolkitTestWebPartProps {
  description: string;
}

export default class SpFxReactToolkitTestWebPart extends BaseClientSideWebPart<ISpFxReactToolkitTestWebPartProps> {

  public render(): void {
    const element: React.ReactElement = React.createElement(
      SPFxWebPartProvider,
      { instance: this },
      React.createElement(SpFxReactToolkitTest)
    );
    
    ReactDom.render(element, this.domElement);
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

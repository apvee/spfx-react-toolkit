import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from '@microsoft/sp-application-base';

import * as strings from 'SpFxReactToolkitTestApplicationCustomizerStrings';
import { SPFxApplicationCustomizerProvider } from '../../core';
import {
  useSPFxInstanceInfo,
  useSPFxPageContext,
  useSPFxProperties,
} from '../../hooks';

/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount -- PlaceholderContent owns a separate domElement lifecycle. */

const LOG_SOURCE: string = 'SpFxReactToolkitTestApplicationCustomizer';

export interface ISpFxReactToolkitTestApplicationCustomizerProperties {
  testMessage: string;
}

const ExtensionProviderProbe: React.FC = () => {
  const { id, kind } = useSPFxInstanceInfo();
  const pageContext = useSPFxPageContext();
  const { properties } = useSPFxProperties<ISpFxReactToolkitTestApplicationCustomizerProperties>();

  return React.createElement(
    'div',
    {
      style: {
        background: '#f3f2f1',
        borderBottom: '1px solid #edebe9',
        color: '#323130',
        fontSize: 12,
        padding: '6px 12px',
      },
    },
    `SPFxApplicationCustomizerProvider active: ${kind} ${id} | ${pageContext.web.title} | ${properties?.testMessage ?? 'No message'}`
  );
};

export default class SpFxReactToolkitTestApplicationCustomizer
  extends BaseApplicationCustomizer<ISpFxReactToolkitTestApplicationCustomizerProperties> {

  private topPlaceholder?: PlaceholderContent;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this.renderTopPlaceholder);
    this.renderTopPlaceholder();

    return Promise.resolve();
  }

  public onDispose(): void {
    this.context.placeholderProvider.changedEvent.remove(this, this.renderTopPlaceholder);

    if (this.topPlaceholder) {
      ReactDom.unmountComponentAtNode(this.topPlaceholder.domElement);
      this.topPlaceholder.dispose();
      this.topPlaceholder = undefined;
    }
  }

  private renderTopPlaceholder = (): void => {
    if (!this.topPlaceholder) {
      this.topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        {
          onDispose: () => {
            if (this.topPlaceholder) {
              ReactDom.unmountComponentAtNode(this.topPlaceholder.domElement);
            }
          },
        }
      );
    }

    if (!this.topPlaceholder) {
      return;
    }

    const element = React.createElement(
      SPFxApplicationCustomizerProvider,
      { instance: this },
      React.createElement(ExtensionProviderProbe)
    );

    ReactDom.render(element, this.topPlaceholder.domElement);
  };
}

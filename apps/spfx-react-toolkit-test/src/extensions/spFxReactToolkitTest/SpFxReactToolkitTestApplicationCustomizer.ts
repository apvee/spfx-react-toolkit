import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IdPrefixProvider } from '@fluentui/react-components';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from '@microsoft/sp-application-base';

import * as strings from 'SpFxReactToolkitTestApplicationCustomizerStrings';
import { ApplicationCustomizerExampleHost } from '../../webparts/spFxReactToolkitTest/components/shared/ApplicationCustomizerExampleHost';

/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount -- PlaceholderContent owns a separate domElement lifecycle. */

const LOG_SOURCE: string = 'SpFxReactToolkitTestApplicationCustomizer';

interface ISpFxReactToolkitTestApplicationCustomizerProperties {
  testMessage: string;
}

export default class SpFxReactToolkitTestApplicationCustomizer
  extends BaseApplicationCustomizer<ISpFxReactToolkitTestApplicationCustomizerProperties> {

  private topPlaceholder?: PlaceholderContent;
  // Pinned BaseComponent sets isDisposed only AFTER onDispose and deletes context.
  // Close this native root before teardown so queued/reentrant callbacks cannot recreate it.
  private terminal: boolean = false;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this.renderTopPlaceholder);
    this.renderTopPlaceholder();

    return Promise.resolve();
  }

  public onDispose(): void {
    if (this.terminal) return;
    this.terminal = true;
    this.context.placeholderProvider.changedEvent.remove(this, this.renderTopPlaceholder);

    if (this.topPlaceholder) {
      ReactDom.unmountComponentAtNode(this.topPlaceholder.domElement);
      this.topPlaceholder.dispose();
      this.topPlaceholder = undefined;
    }
  }

  private renderTopPlaceholder = (): void => {
    if (this.terminal) return;
    if (!this.topPlaceholder) {
      const createdPlaceholder: PlaceholderContent | undefined = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        {
          onDispose: () => {
            if (createdPlaceholder) {
              ReactDom.unmountComponentAtNode(createdPlaceholder.domElement);
              if (this.topPlaceholder === createdPlaceholder) {
                this.topPlaceholder = undefined;
              }
            }
          },
        }
      );
      this.topPlaceholder = createdPlaceholder;
    }

    if (!this.topPlaceholder) {
      return;
    }

    const element = React.createElement(
      IdPrefixProvider,
      // Read the initialized SDK identity only after the terminal/placeholder guards.
      { value: `spfx-toolkit-top-${this.instanceId}-` },
      React.createElement(ApplicationCustomizerExampleHost, { instance: this })
    );

    ReactDom.render(element, this.topPlaceholder.domElement);
  };
}

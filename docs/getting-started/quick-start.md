# Build a first SPFx web part

[Documentation home](../INDEX.md) · [Installation](installation.md) · [Provider selection](choosing-a-provider.md)

Use `WebPartShell` in a web part `.tsx` file. It supplies the SPFx and Fluent providers, so its children can use the toolkit hooks:

```tsx
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart, type IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';
import { PropertyPaneToggle, type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import {
  WebPartShell,
  getWebPartShellPropertiesMetadata,
  useSPFxProperties,
  useSPFxUserInfo,
  type IWebPartShellProperties,
} from '@apvee/spfx-react-toolkit';

interface IExampleWebPartProperties extends IWebPartShellProperties {
  description: string;
}

function Content(): React.ReactElement {
  const { displayName } = useSPFxUserInfo();
  const { properties } = useSPFxProperties<IExampleWebPartProperties>();
  return <p>Hello {displayName}! {properties?.description}</p>;
}

export default class ExampleWebPart extends BaseClientSideWebPart<IExampleWebPartProperties> {
  public render(): void {
    ReactDom.render(
      <WebPartShell instance={this} appearance="outline" size="medium" titlePosition="outside">
        <Content />
      </WebPartShell>,
      this.domElement
    );
  }

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      ...getWebPartShellPropertiesMetadata(),
      description: { isSearchablePlainText: true },
    };
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ groups: [{ groupName: 'Title', groupFields: [
      PropertyPaneToggle('showTitle', { label: 'Show title', checked: this.properties.showTitle ?? true }),
    ] }] }] };
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
```

Initialize these defaults in the web part manifest's `preconfiguredEntries[].properties`. Leave `titleHTML` omitted until an author edits the title:

```json
{
  "title": "My web part",
  "showTitle": true,
  "description": "Built with SPFx React Toolkit."
}
```

### Shell presentation

| Option | Default | Behavior |
| --- | --- | --- |
| `appearance` | `subtle` | No Card wrapper with `subtle`; `filled`, `filled-alternative` and `outline` render a Fluent Card. |
| `size` | `medium` | Native Card size: `small`, `medium` or `large`; ignored with `subtle`. |
| `titlePosition` | `outside` | `inside` places the complete title/action header inside the Card; ignored with `subtle`. |
| Persisted `showTitle` property | `true` when omitted | `false` hides the entire title/editor/action header and its gap. |

`surface` / `padding` control outer host spacing; Card size controls inner spacing. `appearance`, `size` and `titlePosition` are React props: exposing them in the consumer's Property Pane is optional. See [WebPartShell](../api/controls/web-part-shell.md) for title persistence, host detection, actions and state retention.

For direct provider composition, choose `SPFxWebPartProvider`, `SPFxApplicationCustomizerProvider`, `SPFxFieldCustomizerProvider` or `SPFxListViewCommandSetProvider` for the matching SPFx host. See [Providers](../api/core/providers.md) and [Introduction](quick-start.md). Generic React controls and utility hooks document their individual provider requirements.

## Direct provider composition

Use this alternative when you own the surrounding layout and only need SPFx context. It does not supply the Shell title or Fluent theme wrapper.

```tsx
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPFxWebPartProvider, useSPFxUserInfo } from '@apvee/spfx-react-toolkit';

const Greeting: React.FC = () => {
  const { displayName } = useSPFxUserInfo();
  return <div>Hello {displayName}!</div>;
};

export default class GreetingWebPart extends BaseClientSideWebPart<{}> {
  public render(): void {
    ReactDom.render(
      React.createElement(SPFxWebPartProvider, { instance: this }, React.createElement(Greeting)),
      this.domElement
    );
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
```

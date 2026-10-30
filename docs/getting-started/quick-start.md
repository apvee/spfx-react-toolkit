# First SPFx web part

[Documentation home](../INDEX.md) · [Installation](installation.md) · [Choose a provider](choosing-a-provider.md)

After installing the package and its compatible peers, use the following in a web part `.tsx` file. The provider receives the actual web part instance; its child can then read the current user's data through a toolkit hook.

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

The provider belongs inside `render()` and the matching React tree is unmounted from `onDispose()`. Each provider owns the runtime state for its host instance; mount separate providers for separate SPFx instances. Choose another provider for an extension host rather than passing that instance to the web part provider.

## Next steps

- [Read and update web part properties](../api/hooks/properties.md).
- [Select an HTTP client](../guides/http-clients.md).
- [Choose browser, user, tenant or site storage](../guides/storage.md).
- [Compose typed classes with useSx](../guides/styling.md).
- [Find an API by name](../api/INDEX.md).

To execute the repository's interactive samples, follow [development setup](../development/setup.md). Local builds and examples do not establish authenticated SharePoint-host behavior; use the [validation checklists](../development/sharepoint-validation/INDEX.md) for that separate step.
